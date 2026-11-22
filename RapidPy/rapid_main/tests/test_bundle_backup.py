"""Verify accepted backup contents and the desktop completion path."""
import hashlib
import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from data_viewer.data_loading import load_magic_directory
from rapid_main.config import AppConfig
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.io.bundle_backup import backup_measurement
from rapid_main.io.bundle_history import ARCHIVE_PAYLOAD, HISTORY_DIRECTORY, open_bundle_archive
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.panels.measurement import MeasurementPanel


class BundleBackupTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        profile = patch.dict(os.environ, {'RAPID_CONFIG': str(self.root / 'config.json')})
        profile.start()
        self.addCleanup(profile.stop)
        self.destination = self.root / 'backup folder'
        self.destination.mkdir()

    def measure(self, label='NRM'):
        worker = MeasurementWorker(SpecimenMeta('A'), [label], self.root / 'source', NoCommBackend())
        worker.run()
        self.assertIs(worker._pending_run_result, False)
        return worker._publish_dir

    def test_backup_copies_every_indexed_artifact_once_and_opens_in_data_review(self):
        source = self.measure()
        index_bytes = (source / 'artifact_index.json').read_bytes()
        index = json.loads(index_bytes)
        result = backup_measurement(source, self.destination)
        archive = Path(result['archive'])
        self.assertIn('SIMULATED', archive.parts)
        self.assertEqual(result, backup_measurement(source, self.destination))
        self.assertEqual(list(archive.parent.iterdir()), [archive])
        self.assertFalse((self.destination / 'RapidPy backups' / 'Measurements').exists())
        with open_bundle_archive(archive, expected_id=result['archive_id']) as payload:
            self.assertEqual(payload.read('artifact_index.json'), index_bytes)
            for entry in index['artifacts']:
                if entry['exists']:
                    self.assertEqual(hashlib.sha256(payload.read(entry['relative_path'])).hexdigest(), entry['sha256'])
        self.assertEqual(load_magic_directory(archive).specimens[0].steps[0].demag_label, 'NRM')

    def test_different_or_tampered_publication_cannot_create_a_backup(self):
        source = self.measure()
        index_path = source / 'artifact_index.json'
        original = index_path.read_bytes()
        with self.assertRaisesRegex(ValueError, 'changed before backup started'):
            backup_measurement(source, self.destination, expected_index_sha256='0' * 64)
        index_path.write_text(json.dumps({**json.loads(original), 'simulated': False}), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'provenance and artifact index disagree'):
            backup_measurement(source, self.destination)
        index_path.write_bytes(original)
        (source / 'measurements.txt').write_text('changed', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'changed or is missing: measurements.txt'):
            backup_measurement(source, self.destination)
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_retry_copies_the_original_verified_publication_after_remeasurement(self):
        source = self.measure()
        original_index = (source / 'artifact_index.json').read_bytes()
        original_digest = hashlib.sha256(original_index).hexdigest()
        self.assertEqual(self.measure('AF20'), source)
        self.assertEqual(load_magic_directory(source).specimens[0].steps[0].demag_label, 'AF20')

        result = backup_measurement(source, self.destination, expected_index_sha256=original_digest)
        self.assertEqual(result, backup_measurement(source, self.destination, expected_index_sha256=original_digest))
        with open_bundle_archive(result['archive']) as payload:
            self.assertEqual(payload.read('artifact_index.json'), original_index)
        self.assertEqual(load_magic_directory(result['archive']).specimens[0].steps[0].demag_label, 'NRM')

        historical, = (source / HISTORY_DIRECTORY).iterdir()
        (historical / ARCHIVE_PAYLOAD).write_bytes(b'damaged')
        another_destination = self.root / 'another destination'
        another_destination.mkdir()
        with self.assertRaisesRegex(ValueError, 'archive ZIP is damaged'):
            backup_measurement(source, another_destination, expected_index_sha256=original_digest)
        self.assertEqual(list(another_destination.iterdir()), [])

    def test_older_unhashed_quicklook_is_reported_and_excluded(self):
        source = self.measure()
        index_path = source / 'artifact_index.json'
        index = json.loads(index_path.read_bytes())
        quicklook, = [entry for entry in index['artifacts'] if entry['relative_path'] == 'quicklook.json']
        quicklook['exists'] = True
        (source / 'quicklook.json').write_text('{}', encoding='utf-8')
        index_path.write_text(json.dumps(index), encoding='utf-8')

        result = backup_measurement(source, self.destination)
        self.assertEqual(result['omitted'], ['quicklook.json'])
        with open_bundle_archive(result['archive']) as payload:
            self.assertNotIn('quicklook.json', payload.namelist())
            self.assertIn('measurements.txt', payload.namelist())

    def test_desktop_run_reports_backups_and_retains_primary_on_destination_failure(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        window = QtWidgets.QMainWindow()
        self.addCleanup(window.deleteLater)
        panel = MeasurementPanel(window)
        window.config = AppConfig()
        window.config.general.data_dir = str(self.root / 'source')
        window._sequence_labels = ['NRM']
        window.measurement_backend = NoCommBackend
        runs = []
        statuses = []
        panel.sample_run_finished.connect(lambda aborted, name: runs.append((aborted, name)))
        panel._backups.status_changed.connect(lambda text, warning: statuses.append((text, warning)))

        def wait_for(predicate):
            timer = QtCore.QElapsedTimer()
            timer.start()
            while not predicate() and timer.elapsed() < 15_000:
                app.processEvents()
                QtCore.QThread.msleep(5)
            self.assertTrue(predicate(), statuses)

        window.config.general.backup_dir = str(self.destination)
        self.assertTrue(panel.start_measurement_for_sample('A'))
        wait_for(lambda: bool(runs) and any(text.startswith('Backup saved') for text, _ in statuses))
        archive, = (self.destination / 'RapidPy backups' / 'SIMULATED').iterdir()
        with open_bundle_archive(archive) as payload:
            self.assertIn('quicklook.json', payload.namelist())
        self.assertEqual(runs, [(False, 'A')])

        window.config.general.backup_dir = str(self.root / 'missing')
        self.assertTrue(panel.start_measurement_for_sample('B'))
        wait_for(lambda: len(runs) == 2 and any(text.startswith('Backup failed for B') for text, _ in statuses))
        self.assertEqual(runs[-1], (False, 'B'))
        self.assertTrue(statuses[-1][1])
        self.assertIn('The primary measurement is saved', panel._backup_status.text())
        self.assertTrue((self.root / 'source' / 'B' / 'SIMULATED' / 'measurements.txt').is_file())

        reopened_window = QtWidgets.QMainWindow()
        self.addCleanup(reopened_window.deleteLater)
        reopened_window.config = window.config
        reopened_window.config.general.backup_dir = str(self.destination)
        reopened_panel = MeasurementPanel(reopened_window)
        wait_for(lambda: not reopened_panel._retry_backup_btn.isHidden())
        self.assertEqual(reopened_panel._retry_backup_btn.text(), 'Retry Pending Backups (1)')
        reopened_panel._retry_backup_btn.click()
        self.assertIn('original backup folder', reopened_panel._backup_status.text())
        self.assertEqual(list(archive.parent.iterdir()), [archive])

        recovered = self.root / 'missing'
        recovered.mkdir()
        reopened_window.config.general.backup_dir = str(recovered)
        reopened_panel._retry_backup_btn.click()
        wait_for(lambda: reopened_panel._retry_backup_btn.isHidden())
        self.assertIn('Backup saved for B', reopened_panel._backup_status.text())
        restored, = (recovered / 'RapidPy backups' / 'SIMULATED').iterdir()
        self.assertEqual(load_magic_directory(restored).specimens[0].steps[0].demag_label, 'NRM')
        self.assertEqual(list((self.root / 'measurement-backup-queue').iterdir()), [])
