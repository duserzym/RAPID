"""Exercise saved-step recovery through the worker and Live Measure."""
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import threading
import unittest
from uuid import uuid4
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets
from data_viewer.data_loading import load_magic_directory
from rapid_main.config import AppConfig
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.holder_state import HolderStatus
from rapid_main.magnetometer import BlockAudit, BracketedMeasurementBlock
from rapid_main.manual_run_recovery import acknowledge_interruption, reconcile_manual_run
from rapid_main.measurement_checkpoint import pending_checkpoint
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.panels.measurement import MeasurementPanel


class RecordingBackend(NoCommBackend):
    def __init__(self, *, fail_step=False, fail_return=False):
        super().__init__()
        self.treatments = []
        self.preflights = 0
        self.fail_step = fail_step
        self.fail_return = fail_return

    def preflight(self):
        self.preflights += 1
        return super().preflight()

    def set_demag_step(self, label):
        self.treatments.append(label)
        if self.fail_step:
            raise RuntimeError('indeterminate treatment')
        super().set_demag_step(label)

    def return_to_safe_state(self):
        if self.fail_return:
            raise RuntimeError('unverified return')


class MeasurementCheckpointTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.output = self.root / 'SIMULATED'
        self.labels = ['NRM', 'AF10', 'AF10']

    def run_worker(self, backend, *, checkpoint=False, saved=None, profile='station-a', labels=None):
        worker = MeasurementWorker(SpecimenMeta('A'), labels or self.labels, self.root, backend,
                                   run_id='reopened' if saved else 'original', checkpoint_profile=profile,
                                   continue_from=saved)
        if checkpoint:
            worker.request_checkpoint()
        worker.run()
        return worker

    def test_restarts_exact_prefix_and_preserves_repeated_plan_positions(self):
        backend = RecordingBackend()
        original = self.run_worker(backend, checkpoint=True)
        self.assertTrue(original.checkpoint_saved, original._error_messages)
        self.assertEqual(backend.treatments, ['NRM'])
        self.assertFalse((self.output / 'measurements.txt').exists())
        self.assertFalse((self.output / 'quicklook.json').exists())
        checkpoint = pending_checkpoint(self.output)

        resumed_backend = RecordingBackend()
        middle = self.run_worker(resumed_backend, saved=checkpoint, checkpoint=True)
        self.assertTrue(middle.checkpoint_saved, middle._error_messages)
        self.assertEqual(resumed_backend.treatments, ['AF10'])
        later = pending_checkpoint(self.output)
        self.assertNotEqual(later, checkpoint)
        final_backend = RecordingBackend()
        continued = self.run_worker(final_backend, saved=later)
        self.assertFalse(continued._pending_run_result, continued._error_messages)
        self.assertEqual(final_backend.treatments, ['AF10'])
        self.assertEqual(continued._run_id, 'original')
        self.assertEqual([step['label'] for step in continued._progress.accepted_steps()], self.labels)
        self.assertIsNone(pending_checkpoint(self.output))
        self.assertEqual([step.demag_label for step in load_magic_directory(self.output).specimens[0].steps], self.labels)
        quicklook = json.loads((self.output / 'quicklook.json').read_bytes())
        self.assertEqual(quicklook['labels'], self.labels)
        provenance = json.loads((self.output / 'provenance.json').read_bytes())
        self.assertEqual(provenance['continued_from_saved_step'], 2)
        index = json.loads((self.output / 'artifact_index.json').read_bytes())
        self.assertTrue(any(item['relative_path'].endswith('/communication.tsv')
                            for item in index['artifacts']))

    def test_rejects_changed_profile_and_unresolved_subsequent_intent_before_io(self):
        original = self.run_worker(RecordingBackend(), checkpoint=True)
        self.assertTrue(original.checkpoint_saved, original._error_messages)
        checkpoint = pending_checkpoint(self.output)
        changed = RecordingBackend()
        self.run_worker(changed, saved=checkpoint, profile='station-b')
        self.assertEqual(changed.preflights, 0)
        failed = RecordingBackend(fail_step=True)
        continuation = self.run_worker(failed, saved=checkpoint)
        self.assertTrue(continuation._pending_run_result)
        self.assertEqual(failed.treatments, ['AF10'])
        retry = RecordingBackend()
        rejected = self.run_worker(retry, saved=checkpoint)
        self.assertIn('Review Interrupted Run', rejected._error_messages[0])
        self.assertEqual(retry.preflights, 0)
        self.assertEqual(pending_checkpoint(self.output), checkpoint)

    def test_save_preserves_accepted_output_and_detects_an_edit_before_restoring(self):
        previous = MeasurementWorker(SpecimenMeta('A'), ['AF20'], self.root, RecordingBackend(), run_id='previous')
        previous.run()
        self.assertFalse(previous._pending_run_result)
        index = (self.output / 'artifact_index.json').read_bytes()
        measurements = self.output / 'measurements.txt'
        content = measurements.read_bytes()

        worker = self.run_worker(RecordingBackend(), checkpoint=True)
        self.assertTrue(worker.checkpoint_saved, worker._error_messages)
        self.assertEqual((self.output / 'artifact_index.json').read_bytes(), index)
        self.assertEqual(measurements.read_bytes(), content)
        measurements.write_bytes(content + b'changed')
        retry = RecordingBackend()
        rejected = self.run_worker(retry, saved=pending_checkpoint(self.output))
        self.assertIn('accepted specimen output changed', rejected._error_messages[0])
        self.assertEqual(retry.preflights, 0)
        self.assertEqual(measurements.read_bytes(), content + b'changed')

    def test_failed_safe_return_cannot_create_a_checkpoint(self):
        original = self.run_worker(RecordingBackend(fail_return=True), checkpoint=True)
        self.assertFalse(original.checkpoint_saved)
        self.assertIn('unverified return', original._error_messages[0])
        self.assertIsNone(pending_checkpoint(self.output))

    def test_hardware_contract_requires_the_original_holder_before_preflight(self):
        class RoutedFixture(RecordingBackend):
            simulated = False

            def __init__(self, holder='original-holder', *, velocity=0):
                super().__init__()
                self.holder = holder
                self.velocity = velocity

            def holder_status(self):
                return HolderStatus(True, True, record_version=self.holder, is_up=True)

            def verify_checkpoint_stop(self):
                observed = [dict(position_raw=1, velocity_register=self.velocity) for _ in range(2)]
                return dict(safe_state_confirmed=True, simulated=False, axes={'updown': observed, 'turning': observed})

            def read_squid(self):
                return BracketedMeasurementBlock((0., 0., 0.), ((2., 0., 1.),)*4, (0., 0., 0.),
                    audit=BlockAudit(block_id=uuid4().hex, sample_name='A', run_id='original',
                                    treatment_label=self.treatments[-1], holder_record_id=self.holder))

        failed = self.run_worker(RoutedFixture(velocity=1), checkpoint=True, labels=['NRM', 'AF10'], profile='moving')
        self.assertFalse(failed.checkpoint_saved)
        self.assertIn('moving at updown', failed._error_messages[0])
        self.assertIsNone(pending_checkpoint(self.root))
        attempt = reconcile_manual_run(self.root, 'A')
        acknowledge_interruption(attempt, operator='Fixture operator', notes='The first fixture specimen was set aside.')
        first = self.run_worker(RoutedFixture(), checkpoint=True, labels=['NRM', 'AF10'])
        self.assertTrue(first.checkpoint_saved, first._error_messages)
        checkpoint = pending_checkpoint(self.root)
        changed = RoutedFixture('replacement-holder')
        self.run_worker(changed, saved=checkpoint, labels=['NRM', 'AF10'])
        self.assertEqual(changed.preflights, 0)
        expected = RoutedFixture()
        restored = self.run_worker(expected, saved=checkpoint, labels=['NRM', 'AF10'])
        self.assertFalse(restored._pending_run_result, restored._error_messages)
        self.assertEqual(expected.treatments, ['AF10'])

    def test_live_measure_reopens_a_saved_run_in_a_fresh_window(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        gate = threading.Event()

        class AwaitRequestBackend(RecordingBackend):
            def set_demag_step(self, label):
                if not gate.wait(5):
                    raise RuntimeError('test did not request checkpoint')
                super().set_demag_step(label)

        def make_window(backend):
            window = QtWidgets.QMainWindow()
            self.addCleanup(window.deleteLater)
            window.config = AppConfig()
            window.config.general.data_dir = str(self.root)
            window._sequence_labels = ['NRM', 'AF10']
            window.measurement_backend = lambda: backend
            panel = MeasurementPanel(window)
            window.setCentralWidget(panel)
            panel.set_specimen_context('A')
            return window, panel

        def wait_until_settled(panel):
            deadline = QtCore.QElapsedTimer()
            deadline.start()
            while panel.is_active() and deadline.elapsed() < 15_000:
                app.processEvents()
                QtCore.QThread.msleep(5)
            self.assertFalse(panel.is_active())

        first = AwaitRequestBackend()
        window, panel = make_window(first)
        with (patch('rapid_main.panels.measurement.QtWidgets.QMessageBox.critical') as error,
              patch('rapid_main.panels.measurement.QtWidgets.QMessageBox.question',
                    return_value=QtWidgets.QMessageBox.StandardButton.Yes)):
            try:
                panel._start_btn.click()
                self.assertTrue(panel._save_run_btn.isEnabled())
                panel._save_run_btn.click()
            finally:
                gate.set()
            wait_until_settled(panel)
            self.assertEqual(first.treatments, ['NRM'])
            self.assertIsNotNone(pending_checkpoint(self.root / 'A' / 'SIMULATED'))
            window.close()
            second = RecordingBackend()
            reopened, active = make_window(second)
            active._continue_run_btn.click()
            wait_until_settled(active)
            self.assertEqual(second.treatments, ['AF10'])
            self.assertEqual([step.demag_label for step in active._completed_steps], ['NRM', 'AF10'])
            error.assert_not_called()
            reopened.close()
