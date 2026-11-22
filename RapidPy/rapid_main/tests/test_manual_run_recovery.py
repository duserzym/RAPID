"""Check the manual restart guard across real process boundaries in No-Comm."""
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from PySide6 import QtWidgets

from rapid_main.config import AppConfig
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.manual_run_recovery import acknowledge_interruption, reconcile_manual_run
from rapid_main.measurement_checkpoint import pending_checkpoint
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.panels.measurement import MeasurementPanel


_CRASH = '''
import os
from pathlib import Path
import sys
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main import measurement_worker

destination, cut = Path(sys.argv[1]), sys.argv[2]

class Backend(NoCommBackend):
    def set_demag_step(self, label):
        if cut == 'treatment' and label == 'AF10':
            os._exit(91)
        super().set_demag_step(label)

if cut in {'publication', 'checkpoint'}:
    measurement_worker.clear_manual_run = lambda attempt: os._exit(91)

worker = measurement_worker.MeasurementWorker(
    SpecimenMeta('A'), ['NRM', 'AF10'], destination, Backend(),
    run_id='interrupted-attempt', checkpoint_profile='fixture-profile')
if cut == 'checkpoint':
    worker.request_checkpoint()
worker.run()
raise RuntimeError(f'Expected process exit, got {worker._error_messages}')
'''


class RecordingBackend(NoCommBackend):
    def __init__(self):
        super().__init__()
        self.preflights = 0
        self.treatments = []

    def preflight(self):
        self.preflights += 1
        return super().preflight()

    def set_demag_step(self, label):
        self.treatments.append(label)
        super().set_demag_step(label)


class ManualRunRecoveryTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.data = Path(temp.name)
        self.specimen = self.data / 'A'
        self.output = self.specimen / 'SIMULATED'

    def crash_after(self, cut):
        env = {**os.environ, 'PYTHONPATH': os.pathsep.join(path for path in sys.path if path)}
        process = subprocess.run([sys.executable, '-c', _CRASH, str(self.specimen), cut], env=env,
                                 capture_output=True, text=True, timeout=30)
        self.assertEqual(process.returncode, 91, process.stderr)

    def new_worker(self, backend, *, saved=None):
        return MeasurementWorker(SpecimenMeta('A'), ['NRM', 'AF10'], self.specimen, backend,
                                 run_id='new-attempt', checkpoint_profile='fixture-profile', continue_from=saved)

    def test_process_exit_preserves_accepted_output_and_blocks_fresh_treatment(self):
        first = MeasurementWorker(SpecimenMeta('A'), ['AF25'], self.specimen, NoCommBackend(), run_id='accepted')
        first.run()
        self.assertFalse(first._pending_run_result, first._error_messages)
        original = (self.output / 'measurements.txt').read_bytes()
        self.crash_after('treatment')

        backend = RecordingBackend()
        worker = self.new_worker(backend)
        worker.run()
        self.assertIn('Review Interrupted Run', worker._error_messages[0])
        self.assertEqual(backend.preflights, 0)
        self.assertEqual((self.output / 'measurements.txt').read_bytes(), original)
        attempt = reconcile_manual_run(self.output, 'A')
        summary = attempt.describe()
        self.assertIn('1. NRM — Readings captured', summary)
        self.assertIn('2. AF10 — Acquisition incomplete', summary)
        staged = self.output / attempt.staging
        raw = attempt.progress.directory / '00001-started.json'
        self.assertTrue(staged.is_dir())
        digest = raw.read_bytes()

        receipt = acknowledge_interruption(attempt, operator='Fixture custodian', notes='Original specimen set aside.')
        self.assertEqual(json.loads(receipt.read_bytes())['disposition'], 'operator_reviewed_no_replay')
        self.assertEqual(raw.read_bytes(), digest)
        self.assertTrue((receipt.parent / 'unpublished' / 'measurements.txt').is_file())
        self.assertFalse(staged.exists())
        self.assertEqual((self.output / 'measurements.txt').read_bytes(), original)
        self.assertIsNone(reconcile_manual_run(self.output, 'A'))

        next_backend = RecordingBackend()
        fresh = self.new_worker(next_backend)
        fresh.run()
        self.assertFalse(fresh._pending_run_result, fresh._error_messages)
        self.assertEqual(next_backend.treatments, ['NRM', 'AF10'])

    def test_leftover_marker_after_committed_publication_is_reconciled(self):
        self.crash_after('publication')
        original = (self.output / 'measurements.txt').read_bytes()
        self.assertIsNone(reconcile_manual_run(self.output, 'A'))
        self.assertEqual((self.output / 'measurements.txt').read_bytes(), original)
        self.assertFalse(list(self.output.glob('.rapidpy-pending-*')))

    def test_changed_committed_output_cannot_clear_the_interruption_automatically(self):
        self.crash_after('publication')
        output = self.output / 'measurements.txt'
        modified = output.read_bytes() + b'changed'
        output.write_bytes(modified)
        self.assertIsNotNone(reconcile_manual_run(self.output, 'A'))
        backend = RecordingBackend()
        self.new_worker(backend).run()
        self.assertEqual(backend.preflights, 0)
        self.assertEqual(output.read_bytes(), modified)

    def test_leftover_marker_after_planned_stop_does_not_block_saved_continuation(self):
        self.crash_after('checkpoint')
        saved = pending_checkpoint(self.output)
        self.assertIsNotNone(saved)
        self.assertIsNone(reconcile_manual_run(self.output, 'A'))
        backend = RecordingBackend()
        worker = self.new_worker(backend, saved=saved)
        worker.run()
        self.assertFalse(worker._pending_run_result, worker._error_messages)
        self.assertEqual(backend.treatments, ['AF10'])

    def test_live_measure_requires_an_explicit_completed_review(self):
        self.crash_after('treatment')
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        window = QtWidgets.QMainWindow()
        self.addCleanup(window.deleteLater)
        window.config = AppConfig()
        window.config.general.data_dir = str(self.data)
        backend = RecordingBackend()
        window.measurement_backend = lambda: backend
        window._sequence_labels = ['NRM', 'AF10']
        panel = MeasurementPanel(window)
        window.setCentralWidget(panel)
        panel.set_specimen_context('A')
        with patch.object(QtWidgets.QMessageBox, 'warning') as warning:
            self.assertFalse(panel.start_measurement_for_sample('A'))
            self.assertIn('Review Interrupted Run', warning.call_args.args[2])

        with patch.object(QtWidgets.QMessageBox, 'exec', return_value=QtWidgets.QMessageBox.StandardButton.Cancel):
            panel._review_run_btn.click()
        self.assertIsNotNone(reconcile_manual_run(self.output, 'A'))
        with (patch.object(QtWidgets.QMessageBox, 'exec', return_value=QtWidgets.QMessageBox.StandardButton.Ok),
              patch.object(QtWidgets.QInputDialog, 'getMultiLineText', return_value=('Specimen set aside.', True))):
            panel._review_run_btn.click()
        self.assertIsNone(reconcile_manual_run(self.output, 'A'))
        self.assertEqual(backend.preflights, 0)
        window.close()
        app.processEvents()
