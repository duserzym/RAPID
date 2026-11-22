"""Desktop queue preparation and cancellation before the hardware owner starts."""
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.config import AppConfig
from rapid_main.queue_compiler import QueueOptions, QueueSample
from rapid_main.queue_preparation import validate_measurement_output


class QueuePreparationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.original_quit_policy = cls.app.quitOnLastWindowClosed()
        cls.app.setQuitOnLastWindowClosed(False)

    @classmethod
    def tearDownClass(cls):
        cls.app.setQuitOnLastWindowClosed(cls.original_quit_policy)

    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name)
        config = AppConfig()
        config.general.nocomm = True
        config.general.data_dir = str(self.output)
        with patch('rapid_main.app.AppConfig.load', return_value=config):
            self.window = MainWindow()
        self.window._sequence_labels = ['NRM']
        self.window._sample_queue.load_rows([])
        self.entered, self.release = threading.Event(), threading.Event()
        self.thread_ids = []
        self.addCleanup(self.cleanup_window)

    def cleanup_window(self):
        self.release.set()
        if self.window._queue_preparation is not None:
            self.window.cancel_queue_run()
            self.until(lambda: self.window._queue_preparation is None)
        self.window.deleteLater()
        QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
        self.app.processEvents()

    def until(self, condition):
        deadline = time.monotonic() + 10
        while not condition() and time.monotonic() < deadline:
            self.app.processEvents()
            threading.Event().wait(.005)
        self.assertTrue(condition())

    def blocked_path_check(self, *args, **kwargs):
        self.thread_ids.append(threading.get_ident())
        self.entered.set()
        if not self.release.wait(8):
            raise TimeoutError('Test did not release the background path check.')
        return validate_measurement_output(*args, **kwargs)

    def test_desktop_checks_paths_off_thread_and_launches_the_captured_plan(self):
        samples = [QueueSample('S1', 'F1', 1)]
        options = QueueOptions()
        callbacks = []
        with (patch('rapid_main.queue_preparation.validate_measurement_output', side_effect=self.blocked_path_check),
              patch.object(self.window, '_start_prepared_queue') as start):
            self.assertTrue(self.window.request_queue_run(samples, options))
            self.until(self.entered.is_set)
            self.assertEqual(self.window._workflow_state, 'preflight')
            self.assertEqual(len(self.thread_ids), 1)
            self.assertNotEqual(self.thread_ids[0], threading.get_ident())
            QtCore.QTimer.singleShot(0, lambda: callbacks.append(True))
            self.until(lambda: bool(callbacks))
            self.assertFalse(self.window._queue_active)
            self.assertIsNone(self.window._queue_lease)
            start.assert_not_called()
            self.assertFalse(self.window.request_queue_run(samples, options))
            samples[0].sample_name = 'changed'
            options.do_return = False
            self.release.set()
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_called_once()
            prepared = start.call_args.args[0]
            self.assertEqual(prepared.samples[0].sample_name, 'S1')
            self.assertTrue(prepared.options.do_return)
            self.assertIn('Goto', [command.command_type for command in prepared.commands])

    def test_header_halt_cancels_preparation_without_claiming_hardware(self):
        with (patch('rapid_main.queue_preparation.validate_measurement_output', side_effect=self.blocked_path_check),
              patch.object(self.window, '_start_prepared_queue') as start,
              patch.object(QtWidgets.QMessageBox, 'critical') as alert):
            self.assertTrue(self.window.request_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions()))
            self.until(self.entered.is_set)
            self.window._on_header_halt()
            self.assertTrue(self.window._has_active_automation())
            self.release.set()
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_not_called()
            alert.assert_not_called()
        self.assertFalse(self.window._has_active_automation())
        self.assertIsNone(self.window._queue_lease)
        self.assertIn('halted by user', self.window._sb_status.text())

    def test_shutdown_waits_for_the_canceled_preparation_thread_to_exit(self):
        with (patch('rapid_main.queue_preparation.validate_measurement_output', side_effect=self.blocked_path_check),
              patch.object(self.window, '_start_prepared_queue') as start,
              patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes)):
            self.assertTrue(self.window.request_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions()))
            self.until(self.entered.is_set)
            event = QtGui.QCloseEvent()
            self.window.closeEvent(event)
            self.assertFalse(event.isAccepted())
            self.assertTrue(self.window._shutdown_cleanup_requested)
            self.assertIsNotNone(self.window._queue_preparation)
            self.release.set()
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_not_called()

    def test_settings_changed_during_validation_invalidate_the_result(self):
        with (patch('rapid_main.queue_preparation.validate_measurement_output', side_effect=self.blocked_path_check),
              patch.object(self.window, '_start_prepared_queue') as start):
            self.assertTrue(self.window.request_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions()))
            self.until(self.entered.is_set)
            self.window.config.general.data_dir = str(self.output / 'changed')
            self.release.set()
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_not_called()
        self.assertIn('settings changed', self.window._sb_status.text())

    def test_editing_a_queue_row_during_real_run_button_validation_invalidates_the_result(self):
        panel = self.window._sample_queue
        panel.add_sample(position='1', name='S1')
        with (patch('rapid_main.queue_preparation.validate_measurement_output', side_effect=self.blocked_path_check),
              patch.object(self.window, '_start_prepared_queue') as start):
            panel._run_btn.click()
            self.until(self.entered.is_set)
            panel._table.item(0, 2).setText('changed')
            self.release.set()
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_not_called()
        self.assertIsNone(self.window._queue_lease)
        self.assertIn('Queue or settings changed', self.window._sb_status.text())

    def test_invalid_worker_input_reports_error_on_ui_thread_without_claiming_hardware(self):
        alert_threads = []

        def alert(*args):
            alert_threads.append(threading.get_ident())

        with (patch.object(self.window, '_start_prepared_queue') as start,
              patch.object(QtWidgets.QMessageBox, 'critical', side_effect=alert)):
            self.assertTrue(self.window.request_queue_run([QueueSample('../bad', 'F1', 1)], QueueOptions()))
            self.until(lambda: self.window._queue_preparation is None)
            start.assert_not_called()
        self.assertEqual(alert_threads, [threading.get_ident()])
        self.assertEqual(self.window._workflow_state, 'error')
        self.assertIsNone(self.window._queue_lease)
