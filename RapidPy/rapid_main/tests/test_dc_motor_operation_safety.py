from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtCore, QtWidgets
from rapid_main.diagnostic_services import DCMotorBackendAdapter
from rapid_main.dialogs.dc_motors import DCMotorDialog, MotorCommandWorker
from rapidpy_common.hardware import HardwareError, MotorControllerConfig, MotorSerialClient, MoveResult
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError


class DCMotorOperationTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.store = HardwareSafetyStore(Path(directory.name) / 'safety.json')
        self.motor = Mock(spec=MotorSerialClient)
        self.motor.config = MotorControllerConfig()
        self.motor.is_connected = True
        self.motor.read_registers.return_value = (1000, 0)
        self.motor.read_position.return_value = 1000
        self.motor.move_motor.return_value = MoveResult(1000, 1000, True)
        self.motor.turning_motor_spin.return_value = MoveResult(1000, 0, True)
        with patch('rapid_main.diagnostic_services.MotorSerialClient', return_value=self.motor):
            self.backend = DCMotorBackendAdapter(port='COM3', baud=9600)
        self.backend._safety_store = self.store
        self.backend.connect('COM3', 9600)
        self.motor.reset_mock()

    def move(self, **kwargs):
        return self.backend.move_motor('Changer (X)', target=1000, speed=1200, **kwargs)

    def test_motion_claims_before_output_and_publishes_stopped_evidence(self):
        def move(*args, **kwargs):
            self.assertEqual(self.store.pending()['plan']['axes'], ['Changer (X)'])
            with self.assertRaises(HardwareSafetyError):
                with HardwareSafetyStore(self.store.path).operation_lease():
                    pass
            return MoveResult(1000, 1000, True)
        self.motor.move_motor.side_effect = move
        self.assertEqual(self.move(), (1000, 1000, True))
        self.motor.stop.assert_called_once_with(self.backend._axes['Changer (X)'])
        self.assertIsNone(self.store.pending())
        record = self.store.read()['record']
        self.assertEqual(record['station_profile']['helper'], 'rapid_main_dc_motors')
        self.assertEqual(record['cleanup_observations'][0]['readbacks'][0]['velocity_register'], 0)
        self.assertEqual(len(list((self.store.path.parent / 'hardware_diagnostics').glob('motion-*/artifact_index.json'))), 1)

    def test_failed_acknowledgement_stops_every_composite_axis_and_retains_latch(self):
        self.motor.home_xy_to_center.return_value = (MoveResult(0, 0, True), MoveResult(0, 0, True))
        self.motor.stop.side_effect = [HardwareError('no acknowledgement'), None, None]
        with self.assertRaisesRegex(HardwareSafetyError, 'acknowledgement'):
            self.backend.home_xy_to_center()
        self.assertEqual(self.motor.stop.call_count, 3)
        self.assertEqual(self.motor.read_registers.call_count, 6)
        self.motor.halt.assert_called_once()
        self.assertIsNotNone(self.store.pending())
        self.assertFalse(self.backend.operation_active)

    def test_original_station_only_recovery_never_replays_motion(self):
        self.motor.read_registers.return_value = (1000, 1)
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertTrue(self.backend.can_recover_pending)
        self.motor.reset_mock()
        with self.assertRaisesRegex(HardwareError, 'profile'):
            self.backend.connect('COM4', 9600)
        self.motor.connect.assert_not_called()
        self.backend.connect('COM3', 9600)
        self.motor.read_registers.return_value = (1000, 0)
        self.assertTrue(self.backend.recover_stop().safe_state_confirmed)
        self.motor.move_motor.assert_not_called()
        self.assertIsNone(self.store.pending())

    def test_foreign_pending_blocks_connection_before_output(self):
        self.store.begin('af', {}, {})
        with self.assertRaises(HardwareError):
            self.backend.connect('COM3', 9600)
        self.motor.connect.assert_not_called()
        self.assertFalse(self.backend.can_recover_pending)

    def test_precancel_does_not_initialize_or_claim(self):
        self.backend.set_cancel_check(lambda: True)
        with self.assertRaisesRegex(HardwareError, 'cancelled before'):
            self.move()
        self.assertEqual(self.motor.mock_calls, [])
        self.assertIsNone(self.store.pending())

    def test_stop_arriving_at_native_return_does_not_report_success(self):
        def move(*args, **kwargs):
            self.backend.request_stop()
            return MoveResult(1000, 1000, True)
        self.motor.move_motor.side_effect = move
        with self.assertRaisesRegex(HardwareError, 'stopped'):
            self.move()
        self.motor.stop.assert_called_once()
        self.assertIsNone(self.store.pending())
        self.assertIn('stopped', self.store.read()['record']['error'])

    def test_native_success_does_not_excuse_short_target(self):
        self.motor.move_motor.return_value = MoveResult(1000, 600, True)
        with self.assertRaisesRegex(HardwareError, 'target=1000 final=600'):
            self.move()
        self.assertIsNone(self.store.pending())

    def test_spin_waits_for_actual_target_and_verified_velocity(self):
        self.assertEqual(self.backend.spin_turning(speed_rps=1., duration_s=2.), (1000, 1000, True))
        self.motor.wait_for_motor_stop.assert_called_once_with(self.backend._axes['Turning'], timeout_s=17.)
        self.motor.stop.assert_called_once()

    def test_publication_failure_retains_restart_latch(self):
        with patch('rapidpy_common.motor_diagnostic_safety.publish_diagnostic_record', side_effect=HardwareSafetyError('disk full')):
            with self.assertRaisesRegex(HardwareSafetyError, 'disk full'):
                self.move()
        self.assertIsNotNone(self.store.pending())
        self.assertFalse(self.backend.operation_active)

    def test_nonblocking_move_owns_station_until_terminal_cleanup(self):
        entered, release = threading.Event(), threading.Event()
        def settle(axis):
            entered.set()
            if not release.wait(5):
                raise HardwareError('test settling timeout')
        self.motor.wait_for_motor_stop.side_effect = settle
        try:
            self.move(wait_for_stop=False)
            self.assertTrue(entered.wait(2))
            self.assertTrue(self.backend.operation_active)
            self.assertIsNotNone(self.store.pending())
            with self.assertRaises(HardwareError):
                self.backend.disconnect()
            with self.assertRaises(HardwareError):
                self.move()
            start = time.monotonic()
            self.backend.request_stop()
            self.assertLess(time.monotonic() - start, .1)
        finally:
            release.set()
            self.backend._background_thread.join(5)
        self.assertFalse(self.backend.operation_active)
        self.assertIsNone(self.store.pending())

        self.assertIn('stopped', self.store.read()['record']['error'])
        self.assertIsNone(self.backend._last_operation_result)

    def test_cancel_native_home_before_coordinates_are_reset(self):
        native = MotorSerialClient()
        native._motion_cancel_check = lambda: True
        native.zero_target_pos = Mock()
        native.query_ascii = Mock()
        with self.assertRaisesRegex(HardwareError, 'cancelled'):
            native.home_to_top(self.backend._axes['Up/Down'])
        native.zero_target_pos.assert_not_called()
        native.query_ascii.assert_not_called()

    def test_cancel_arriving_during_poll_blocks_motion_opcode(self):
        native = MotorSerialClient()
        stopped = threading.Event()
        native._motion_cancel_check = stopped.is_set
        native.poll_motor = lambda axis: stopped.set()
        native.clear_poll_status = Mock()
        native.query_ascii = Mock()
        with self.assertRaisesRegex(HardwareError, 'cancelled'):
            native.move_motor(self.backend._axes['Changer (X)'], 1000, 1200)
        native.query_ascii.assert_not_called()

    def test_original_dc_panel_can_open_for_recovery_without_unlocking_other_controls(self):
        from types import SimpleNamespace
        from rapid_main.app import MainWindow
        from rapid_main.device_ownership import DeviceOwnershipManager, DeviceOwnershipError
        window = SimpleNamespace(_measurement_backend=SimpleNamespace(has_unresolved_hardware_fault=True),
            _dc_motor_backend=self.backend, _ownership=DeviceOwnershipManager())
        self.store.begin('motion_diagnostic', {'axes': ['Changer (X)']}, self.backend._safety_profile())
        lease = MainWindow.acquire_device(window, 'changer', 'dc_motors_panel')
        lease.release()
        for resource, owner in [('measurement', 'dc_motors_panel'), ('changer', 'some_other_panel')]:
            with self.assertRaises(DeviceOwnershipError):
                MainWindow.acquire_device(window, resource, owner)


class DCMotorWorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.previous_quit_policy = cls.qt_app.quitOnLastWindowClosed()
        cls.qt_app.setQuitOnLastWindowClosed(False)

    @classmethod
    def tearDownClass(cls):
        cls.qt_app.setQuitOnLastWindowClosed(cls.previous_quit_policy)

    def test_precancel_worker_never_calls_action_or_reports_success(self):
        action, backend = Mock(), Mock()
        worker = MotorCommandWorker(backend, 'move', action)
        success, failure, settled = [], [], []
        worker.succeeded.connect(lambda *args: success.append(args))
        worker.failed.connect(lambda *args: failure.append(args))
        worker.settled.connect(lambda: settled.append(True))
        worker.stop()
        worker.run()
        action.assert_not_called()
        self.assertEqual(success, [])
        self.assertIn('cancelled', failure[0][1])
        self.assertEqual(settled, [True])

    def test_close_keeps_live_worker_until_stop_cleanup_finishes(self):
        from rapid_main.diagnostic_services import DCMotorNoCommBackend
        backend = DCMotorNoCommBackend(port='COM1', baud=9600)
        backend.connect('COM1', 9600)
        entered, release = threading.Event(), threading.Event()
        def action():
            entered.set()
            release.wait(5)
        dialog = DCMotorDialog(backend=backend)
        dialog._connected = True
        dialog._refresh_connection_state()
        dialog.show()
        try:
            dialog._start_command('test move', action)
            self.assertTrue(entered.wait(2))
            self.assertTrue(dialog.stop_motion_btn.isEnabled())
            dialog.close()
            self.assertTrue(dialog.isVisible())
            self.assertTrue(dialog._command_thread.isRunning())
            self.assertTrue(backend.is_connected())
            release.set()
            deadline = time.monotonic() + 5
            while dialog.isVisible() and time.monotonic() < deadline:
                self.qt_app.processEvents()
                time.sleep(.01)
            self.assertFalse(dialog.isVisible())
            self.assertIsNone(dialog._command_thread)
            self.assertFalse(backend.is_connected())
            self.assertIn('failed or stopped', dialog._console.toPlainText())
        finally:
            release.set()
            dialog._telemetry_thread.stop_monitoring()
            dialog._telemetry_thread.wait(5000)
            if dialog._command_thread is not None:
                dialog._command_thread.quit()
                dialog._command_thread.wait(5000)
            self.qt_app.processEvents()
            dialog.close()
            dialog.deleteLater()

    def test_main_window_shutdown_waits_for_owned_motor_worker(self):
        from rapid_main.app import MainWindow
        from rapid_main.diagnostic_services import DCMotorNoCommBackend
        from rapid_main.device_ownership import DeviceOwnershipError
        from rapid_main.config import AppConfig
        config = AppConfig()
        config.general.nocomm = True
        with patch.object(AppConfig, 'load', return_value=config):
            window = MainWindow()
        window._confirm_shutdown = lambda **kwargs: True
        backend = DCMotorNoCommBackend(port='COM1', baud=9600)
        backend.connect('COM1', 9600)
        entered, release = threading.Event(), threading.Event()
        def action():
            entered.set()
            release.wait(5)
        window.show()
        window._run_owned_dialog('changer', 'dc_motors_panel', lambda parent: DCMotorDialog(parent, backend=backend), modal=False)
        dialog = window._owned_dialogs['changer']
        dialog._connected = True
        try:
            dialog._start_command('test move', action)
            self.assertTrue(entered.wait(2))
            window.close()
            self.assertTrue(window.isVisible())
            self.assertTrue(dialog._command_thread.isRunning())
            self.assertTrue(backend.is_connected())
            with self.assertRaisesRegex(DeviceOwnershipError, 'Shutdown cleanup'):
                window.acquire_device('measurement', 'new_run')
            release.set()
            deadline = time.monotonic() + 5
            while window.isVisible() and time.monotonic() < deadline:
                self.qt_app.processEvents()
                time.sleep(.01)
            self.assertFalse(window.isVisible())
            self.assertEqual(window._owned_dialog_leases, {})
            self.assertEqual(window._owned_dialogs, {})
            self.assertFalse(backend.is_connected())
        finally:
            release.set()
            deadline = time.monotonic() + 5
            while window._owned_dialogs and time.monotonic() < deadline:
                self.qt_app.processEvents()
                time.sleep(.01)
            window.close()
            window.deleteLater()
