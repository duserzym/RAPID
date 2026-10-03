from dataclasses import asdict
from pathlib import Path
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from updown_control import app as lift
from rapidpy_common.hardware import MoveResult
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError
from rapidpy_common.motor_diagnostic_safety import (motor_diagnostic_operation, recover_motor_diagnostic,
                                                 verify_stopped_in_place)
from rapid_main.config import AppConfig
from rapid_main.hardware_contracts import QueueHardwareBackend


class LiftFixture:
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.store = HardwareSafetyStore(Path(self.directory.name) / 'safety.json')
        self.profile = lift._load_settings_profile(lift.DEFAULT_SETTINGS_PATH)
        self.motor = Mock()
        self.motor.is_connected = True
        self.motor.read_registers.return_value = (1000, 0)
        self.motor.move_motor.return_value = MoveResult(1000, 1000, True)
        self.motor.read_position.return_value = 1000
        self.motor.check_internal_status.return_value = 0
        with patch.object(lift, 'MotorSerialClient', return_value=self.motor):
            self.controller = lift.UpDownController(self.profile, safety_store=self.store)
        self.controller.connect('COM5')
        self.motor.reset_mock()

    def worker(self, **overrides):
        options = dict(controller=self.controller, squid=Mock(), calibration=lift.SquidCalibration(),
                       baseline_raw=(0., 0., 0.), target_positions=[1000], settle_s=0,
                       velocity_raw=2500000, sample_height_cm=2.54, counts_per_cm=1000,
                       target_meas_center_raw=1000, safe_min_raw=0, safe_max_raw=2000)
        options.update(overrides)
        options['squid'].read_moment.return_value = (.1, .2, .3, .4)
        return lift.ScanWorker(**options)


class LiftDiagnosticTests(LiftFixture, unittest.TestCase):
    def test_publication_failure_retains_latch_after_stop(self):
        with patch('rapidpy_common.motor_diagnostic_safety.publish_diagnostic_record', side_effect=HardwareSafetyError('disk full')):
            with self.assertRaisesRegex(HardwareSafetyError, 'disk full'):
                self.controller.move_to_raw(1000, 100)
        self.motor.stop.assert_called_once()
        self.assertIsNotNone(self.store.pending())

    def test_move_claims_before_native_output_and_finishes_after_stop(self):
        def move(*args, **kwargs):
            self.assertEqual(self.store.pending()['family'], 'motion_diagnostic')
            return MoveResult(1000, 1000, True)
        self.motor.move_motor.side_effect = move
        self.assertTrue(self.controller.move_to_raw(1000, 100).success)
        self.motor.stop.assert_called_once_with(self.profile.updown_axis)
        self.assertEqual(self.motor.read_registers.call_count, 2)
        self.assertIsNone(self.store.pending())
        self.assertFalse(self.controller.operation_active)
        state = self.store.read()
        self.assertEqual(state['record']['observations'][0]['result']['final_position'], 1000)
        folders = list((self.store.path.parent / 'hardware_diagnostics').glob('motion-*'))
        self.assertEqual(len(folders), 1)
        self.assertTrue((folders[0] / 'artifact_index.json').is_file())

    def test_zero_target_does_not_bypass_position_verification(self):
        with self.assertRaisesRegex(lift.HardwareError, 'target'):
            self.controller.move_to_raw(0, 100)
        self.assertIsNone(self.store.pending())
        self.assertIn('target', self.store.read()['record']['error'])

    def test_active_main_treatment_blocks_move_before_native_output(self):
        self.store.begin('pulse', {}, {})
        with self.assertRaises(HardwareSafetyError):
            self.controller.move_to_raw(1000, 100)
        self.assertEqual(self.motor.mock_calls, [])
        self.assertFalse(self.controller.operation_active)

    def test_main_pending_blocks_connection_before_broadcast_setup(self):
        self.store.begin('af', {}, {})
        with self.assertRaises(HardwareSafetyError):
            self.controller.connect('COM5')
        self.motor.connect.assert_not_called()

    def test_cleanup_failure_retains_restart_latch_and_prevents_next_move(self):
        self.motor.read_registers.return_value = (1000, 1)
        with self.assertRaisesRegex(HardwareSafetyError, 'motion'):
            self.controller.move_to_raw(1000, 100)
        self.assertIsNotNone(HardwareSafetyStore(self.store.path).pending())
        self.motor.reset_mock()
        with self.assertRaises(HardwareSafetyError):
            self.controller.move_to_raw(1000, 100)
        self.assertEqual(self.motor.mock_calls, [])

    def test_recovery_requires_original_port_and_configuration_without_homing(self):
        profile = self.controller.safety_profile()
        self.store.begin('motion_diagnostic', {'action': 'scan'}, profile)
        self.controller._port = 'COM6'
        with self.assertRaises(HardwareSafetyError):
            self.controller.recover_stopped_in_place()
        self.assertEqual(self.motor.mock_calls, [])
        self.controller._port = 'COM5'
        record = self.controller.recover_stopped_in_place()
        self.assertTrue(record.safe_state_confirmed)
        self.motor.home_to_top.assert_not_called()
        self.motor.zero_target_pos.assert_not_called()
        self.motor.move_motor.assert_not_called()
        self.assertIsNone(self.store.pending())

    def test_recovery_without_prior_latch_still_retains_failed_stop(self):
        self.motor.stop.side_effect = RuntimeError('stop link lost')
        with self.assertRaisesRegex(HardwareSafetyError, 'stop link lost'):
            self.controller.recover_stopped_in_place()
        self.motor.halt.assert_called_once()
        self.assertIsNotNone(self.store.pending())

    def test_station_changes_disconnect_and_second_thread_blocked_while_owned(self):
        failures = []
        with self.controller.motion_operation({'action': 'scan'}):
            for action in (lambda: self.controller.connect('COM6'), self.controller.disconnect,
                           lambda: self.controller.apply_settings_profile(self.profile)):
                with self.assertRaises(HardwareSafetyError):
                    action()
            def second_thread():
                try:
                    self.controller.move_to_raw(1000, 100)
                except HardwareSafetyError as exc:
                    failures.append(str(exc))
            thread = threading.Thread(target=second_thread)
            thread.start()
            thread.join(2)
            self.assertFalse(thread.is_alive())
            self.assertEqual(len(failures), 1)
            self.motor.move_motor.assert_not_called()

    def test_main_routes_recovery_to_original_helper_before_motion(self):
        self.store.begin('motion_diagnostic', {}, self.controller.safety_profile())
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._config.general.nocomm = False
        backend._safety_store = self.store
        backend._client = Mock()
        self.assertTrue(backend.has_unresolved_hardware_fault)
        with self.assertRaisesRegex(RuntimeError, 'original helper'):
            backend.return_to_safe_state()
        self.assertEqual(backend._client.mock_calls, [])

    def test_scan_owns_whole_acquisition_and_cleans_before_result(self):
        worker = self.worker()
        results, failures = [], []
        worker.scan_complete.connect(lambda result: (self.assertIsNone(self.store.pending()), results.append(result)))
        worker.scan_failed.connect(failures.append)
        def moment(*args):
            self.assertEqual(self.store.pending()['plan']['action'], 'measurement_z_scan')
            self.assertTrue(self.controller.operation_active)
            return (.1, .2, .3, .4)
        worker._squid.read_moment.side_effect = moment
        worker.run()
        self.assertEqual(failures, [])
        self.assertEqual(len(results), 1)
        self.motor.stop.assert_called_once()
        self.assertEqual(len(self.store.read()['record']['observations']), 2)

    def test_scan_cancellation_before_start_does_not_initialize_hardware(self):
        worker = self.worker()
        results, failures = [], []
        worker.scan_complete.connect(results.append)
        worker.scan_failed.connect(failures.append)
        worker.request_stop()
        worker.run()
        self.assertEqual(results, [])
        self.assertEqual(len(failures), 1)
        self.assertEqual(self.motor.mock_calls, [])
        self.assertIsNone(self.store.pending())

    def test_cancel_during_moment_acquisition_never_reports_success(self):
        worker = self.worker()
        results, failures = [], []
        worker.scan_complete.connect(results.append)
        worker.scan_failed.connect(failures.append)
        def moment(*args):
            worker.request_stop()
            return (.1, .2, .3, .4)
        worker._squid.read_moment.side_effect = moment
        worker.run()
        self.assertEqual(results, [])
        self.assertEqual(len(failures), 1)
        self.motor.halt.assert_called_once()
        self.assertIsNone(self.store.pending())
        self.assertIn('cancelled', self.store.read()['record']['error'])

    def test_queued_scan_blocks_ui_binding_changes_before_worker_claims(self):
        window = SimpleNamespace(_scan_worker=object(), _append=Mock())
        for method in (lift.MainWindow._connect_motor, lift.MainWindow._disconnect_motor,
                       lift.MainWindow._connect_squid, lift.MainWindow._disconnect_squid,
                       lift.MainWindow._reload_settings_profile, lift.MainWindow._save_settings_file,
                       lift.MainWindow._take_squid_baseline, lift.MainWindow._recover_lift_stop):
            method(window)
        self.assertFalse(lift.MainWindow._require_motor(window))
        self.assertEqual(window._append.call_count, 9)
        self.assertEqual(self.motor.mock_calls, [])

    def test_closing_scan_failure_does_not_open_a_modal_dialog(self):
        window = SimpleNamespace(_close_after_cleanup=True, _set_motion_fault=Mock(),
                                 scan_result_label=Mock(), _append=Mock())
        with patch.object(lift.QtWidgets.QMessageBox, 'warning') as warning:
            lift.MainWindow._handle_scan_failed(window, 'Scan cancelled.')
        warning.assert_not_called()

    def test_scan_invalid_bounds_fail_before_any_move(self):
        worker = self.worker(target_positions=[1000, 3000])
        worker.run()
        self.assertEqual(self.motor.mock_calls, [])

    def test_nonfinite_scan_readings_fail_after_checked_motor_cleanup(self):
        worker = self.worker()
        worker._squid.read_moment.return_value = (float('nan'), .2, .3, .4)
        results, failures = [], []
        worker.scan_complete.connect(results.append)
        worker.scan_failed.connect(failures.append)
        worker.run()
        self.assertEqual(results, [])
        self.assertIn('nonfinite', failures[0])
        self.assertIsNone(self.store.pending())

    def test_close_keeps_live_scan_and_connections_until_completion(self):
        worker, event = Mock(), Mock()
        worker.isRunning.return_value = True
        window = SimpleNamespace(_scan_worker=worker, _close_after_cleanup=False, _append=Mock(),
                                 controller=self.controller, squid=Mock(), vacuum=Mock())
        lift.MainWindow.closeEvent(window, event)
        event.ignore.assert_called_once()
        worker.request_stop.assert_called_once()
        worker.wait.assert_not_called()
        self.assertTrue(window._close_after_cleanup)
        self.motor.disconnect.assert_not_called()


class StopReadbackTests(unittest.TestCase):
    def test_stationary_positions_with_nonzero_velocity_are_unsafe(self):
        motor = Mock()
        motor.read_registers.return_value = (0, 65536)
        observations, error = verify_stopped_in_place(motor, lift.MotorAxisConfig('lift', 3, 16), sleep=lambda _: None)
        self.assertEqual(len(observations), 2)
        self.assertIn('motion', error)


class LiftCompactUiTests(LiftFixture, unittest.TestCase):
    def test_compact_motion_tab_exposes_recovery_and_preserves_wide_widgets(self):
        app = lift.QtWidgets.QApplication.instance() or lift.QtWidgets.QApplication([])
        font = app.font()
        self.addCleanup(lambda: app.setFont(font))
        font_path = Path('C:/Windows/Fonts/segoeui.ttf')
        if font_path.exists():
            font_id = lift.QtGui.QFontDatabase.addApplicationFont(str(font_path))
            self.addCleanup(lambda: lift.QtGui.QFontDatabase.removeApplicationFont(font_id))
        app.setFont(lift.QtGui.QFont('Segoe UI', 10))
        settings = lift.UpDownSettings(settings_path=str(lift.DEFAULT_SETTINGS_PATH))
        with patch.object(lift, 'load_settings', new=lambda: settings), \
             patch.object(lift, 'save_settings', new=lambda settings: None), \
             patch.object(lift.MainWindow, '_autodetect_squid_port', new=lambda self: None):
            window = lift.MainWindow()
            window._fit_to_screen(SimpleNamespace(availableGeometry=lambda: lift.QtCore.QRect(0, 0, 800, 800)))
            window.show()
            window._compact_tabs.setCurrentIndex(3)
            app.processEvents()
            page = window._compact_tabs.widget(3)
            page.ensureWidgetVisible(window.recover_stop_btn)
            app.processEvents()
            self.assertLessEqual(window.width(), 736)
            self.assertLessEqual(window.height(), 720)
            self.assertTrue(window.recover_stop_btn.isVisible())
            origin = window.recover_stop_btn.mapTo(page.viewport(), lift.QtCore.QPoint(0, 0))
            self.assertTrue(page.viewport().rect().contains(origin))
            columns = window._columns
            window._set_compact_layout(False)
            self.assertEqual(window._columns, columns)
            self.assertEqual(window._compact_tabs.count(), 0)
            window._set_compact_layout(True)
            self.assertEqual(window._compact_tabs.count(), 4)
            window.close()

    def test_zero_velocity_with_drifting_positions_is_unsafe(self):
        motor = Mock()
        motor.read_registers.side_effect = [(10, 0), (11, 0)]
        _, error = verify_stopped_in_place(motor, lift.MotorAxisConfig('lift', 3, 16), sleep=lambda _: None)
        self.assertIn('motion', error)
