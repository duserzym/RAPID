"""Native queue ownership persists through asynchronous command recovery."""
import threading
import time
import unittest
from unittest.mock import patch
from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.queue_compiler import QueueCommand
from rapid_main.queue_command_worker import QueueCommandWorker


class Backend:
    queue_commands_require_worker = True

    def __init__(self):
        self.started = threading.Event()
        self.release = threading.Event()
        self.recovering = threading.Event()
        self.recovery_release = threading.Event()
        self.recovery_release.set()
        self.calls = []
        self.check = None
        self.cleanup_error = False

    def set_halt_check(self, check):
        self.check = check

    def move(self):
        self.calls.append(('move', threading.get_ident()))
        self.started.set()
        if not self.release.wait(5):
            raise RuntimeError('test movement timed out')

    def return_to_safe_state(self):
        self.calls.append(('cleanup', threading.get_ident()))
        self.recovering.set()
        if not self.recovery_release.wait(5):
            raise RuntimeError('test cleanup timed out')
        if self.cleanup_error:
            raise RuntimeError('stop could not be verified')


class QueueWorkerUnitTests(unittest.TestCase):
    def test_pre_cancel_does_not_initialize_or_recover_hardware(self):
        backend = Backend()
        worker = QueueCommandWorker(backend, backend.move)
        worker.stop()
        worker.run()
        self.assertFalse(worker.ok)
        self.assertEqual(backend.calls, [])

    def test_cancel_after_action_returns_still_requires_recovery(self):
        backend = Backend()
        worker = QueueCommandWorker(backend, lambda: worker.stop())
        worker.run()
        self.assertFalse(worker.ok)
        self.assertEqual([name for name, _ in backend.calls], ['cleanup'])
        self.assertIsNone(backend.check)

    def test_hook_clear_failure_still_emits_terminal_signal(self):
        class BrokenHook(Backend):
            def set_halt_check(self, check):
                if check is None:
                    raise RuntimeError('hook failure')
        worker = QueueCommandWorker(BrokenHook(), lambda: None)
        emitted = []
        worker.settled.connect(lambda: emitted.append(True))
        worker.run()
        self.assertFalse(worker.ok)
        self.assertEqual(emitted, [True])


class QueueWorkerWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.quit_policy = cls.app.quitOnLastWindowClosed()
        cls.app.setQuitOnLastWindowClosed(False)

    @classmethod
    def tearDownClass(cls):
        cls.app.setQuitOnLastWindowClosed(cls.quit_policy)

    def setUp(self):
        QtCore.QSettings('RAPID', 'RapidPy-rapid_main').clear()
        self.window = MainWindow()
        self.backend = Backend()
        self.window._measurement_backend = self.backend
        self.window._queue_active = True
        self.window._queue_lease = self.window.acquire_device('changer', 'queue_workflow')
        self.window._queue_paused = True

    def wait_for(self, predicate):
        deadline = time.monotonic() + 6
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.005)
        self.assertTrue(predicate())

    def tearDown(self):
        self.backend.release.set()
        self.backend.recovery_release.set()
        if self.window._queue_command_worker is not None:
            self.window.cancel_queue_run()
            self.wait_for(lambda: self.window._queue_command_thread is None)
        self.window._queue_active = False
        self.window._release_queue_lease()
        self.window.deleteLater()
        self.app.processEvents()

    def start_command(self):
        lease = self.window.acquire_device('measurement', 'queue_command')
        self.window._start_queue_command_worker(QueueCommand('Goto', hole=1), self.backend.move, [lease])
        self.wait_for(self.backend.started.is_set)

    def test_halt_retains_both_leases_until_worker_cleanup_finishes(self):
        self.backend.recovery_release.clear()
        self.start_command()
        self.window.halt_measurement()
        self.backend.release.set()
        self.wait_for(self.backend.recovering.is_set)
        self.assertTrue(self.window._ownership.is_owned('changer'))
        self.assertTrue(self.window._ownership.is_owned('measurement'))
        self.assertTrue(self.window._has_active_automation())
        self.backend.recovery_release.set()
        self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertFalse(self.window._ownership.is_owned('changer'))
        self.assertFalse(self.window._ownership.is_owned('measurement'))
        self.assertFalse(self.window._queue_active)
        self.assertTrue(all(tid != threading.get_ident() for _, tid in self.backend.calls))

    def test_gui_events_and_resume_do_not_advance_while_motion_is_live(self):
        self.start_command()
        fired = []
        timer = QtCore.QTimer(self.window)
        timer.setSingleShot(True)
        timer.timeout.connect(lambda: fired.append(True))
        timer.start(0)
        self.wait_for(lambda: bool(fired))
        self.window._queue_pos = 3
        self.window._queue_paused = False
        self.window._run_next_queue_command()
        self.assertEqual(self.window._queue_pos, 3)
        self.window._queue_paused = True
        self.backend.release.set()
        self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertTrue(self.window._queue_active)
        self.assertEqual([name for name, _ in self.backend.calls], ['move'])

    def test_terminal_cleanup_runs_in_worker_and_failure_is_reported(self):
        self.backend.cleanup_error = True
        self.backend.recovery_release.clear()
        self.window._finalize_queue_run('complete', reason='Queue complete.')
        self.wait_for(self.backend.recovering.is_set)
        self.assertTrue(self.window._ownership.is_owned('changer'))
        self.backend.recovery_release.set()
        self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertEqual(self.window._workflow_state, 'error')
        self.assertIn('stop could not be verified', self.window._sb_status.text())
        self.assertTrue(all(tid != threading.get_ident() for _, tid in self.backend.calls))

    def test_measurement_terminal_signal_does_not_allow_overlapping_cleanup(self):
        panel = self.window._measurement
        entered, release = threading.Event(), threading.Event()
        class EarlyTerminal(QtCore.QThread):
            terminal = QtCore.Signal(bool)
            def run(self):
                self.terminal.emit(True)
                entered.set()
                release.wait(5)
            def halt(self):
                pass
        thread = EarlyTerminal(panel)
        panel._worker = thread
        panel._leases = [self.window.acquire_device('measurement', 'measurement_run')]
        thread.terminal.connect(panel._on_run_finished)
        thread.finished.connect(panel._settle_measurement_worker)
        thread.start()
        try:
            self.wait_for(entered.is_set)
            with patch.object(QtWidgets.QMessageBox, 'critical', return_value=None):
                panel._on_error('transport failed; cleanup still running')
            self.window.halt_measurement()
            self.assertTrue(panel.is_active())
            self.assertTrue(self.window._ownership.is_owned('measurement'))
            self.assertTrue(self.window._ownership.is_owned('changer'))
            self.assertFalse(self.backend.recovering.is_set())
            release.set()
            self.wait_for(self.backend.recovering.is_set)
            self.wait_for(lambda: self.window._queue_command_thread is None)
            self.assertFalse(panel.is_active())
            self.assertFalse(self.window._ownership.is_owned('measurement'))
            self.assertFalse(self.window._ownership.is_owned('changer'))
        finally:
            release.set()
            if panel._worker is not None:
                self.wait_for(lambda: panel._worker is None)
