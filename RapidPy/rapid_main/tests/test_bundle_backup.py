"""Verify accepted backup contents and the desktop completion path."""
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from PySide6 import QtCore, QtWidgets

from data_viewer.data_loading import load_magic_directory
from rapid_main.config import AppConfig
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.io.bundle_backup import backup_measurement
from rapid_main.io.bundle_history import open_bundle_archive
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.panels.measurement import MeasurementPanel


class BundleBackupTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.destination = self.root / 'backup folder'
        self.destination.mkdir()

    def measure(self):
        worker = MeasurementWorker(SpecimenMeta('A'), ['NRM'], self.root / 'source', NoCommBackend())
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
