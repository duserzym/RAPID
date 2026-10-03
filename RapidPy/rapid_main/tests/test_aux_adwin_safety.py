import math
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from types import SimpleNamespace

from rapidpy_common.adwin_af import (AdwinAFController, AdwinBoardConfig, AdwinCoilLimits, AdwinDenseCaptureRequest,
                                     AdwinDenseCaptureResult, AdwinError, validate_dense_capture_request)
from rapidpy_common.af_diagnostic_safety import ManualAdwinDiagnostic
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError
from af_clip_test import app as clip
from adwin_comms import app as comms


class DiagnosticFixture:
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / 'safety.json'
        self.env = patch.dict(os.environ, RAPID_SAFETY_STATE=str(self.path))
        self.env.start()
        self.addCleanup(self.env.stop)
        self.store = HardwareSafetyStore(self.path)
        self.ctrl = Mock()
        self.ctrl.board = AdwinBoardConfig()
        self.ctrl.limits = AdwinCoilLimits()
        self.ctrl.test_version.return_value = 9
        self.ctrl.get_digout.return_value = 0
        self.ctrl.set_dac.return_value = None


class ManualOutputTests(DiagnosticFixture, unittest.TestCase):
    def test_output_lifetime_holds_lease_until_explicit_cleanup(self):
        session = ManualAdwinDiagnostic(self.ctrl)
        session.write({'dac': 1, 'voltage': 2}, lambda: self.ctrl.set_dac(1, 2))
        self.assertTrue(session.active)
        self.assertEqual(self.store.pending()['family'], 'af_diagnostic')
        self.ctrl.recover_safe_field.assert_called_once()
        with self.assertRaises(HardwareSafetyError):
            ManualAdwinDiagnostic(self.ctrl).write({'dac': 2}, lambda: self.fail('Overlapping output'))
        session.write({'dac': 2, 'voltage': 0}, lambda: None)
        session.observe({'adc': .2})
        session.close()
        self.assertFalse(session.active)
        self.assertIsNone(self.store.pending())
        self.assertEqual(self.ctrl.recover_safe_field.call_count, 2)
        self.assertEqual(len(self.store.read()['record']['observations']), 3)

    def test_cleanup_addresses_original_station_after_mutation(self):
        session = ManualAdwinDiagnostic(self.ctrl)
        session.write({'dac': 1}, lambda: None)
        self.ctrl.board.board_num = 2
        with self.assertRaises(HardwareSafetyError):
            session.write({}, lambda: self.fail('Mutated board output'))
        self.ctrl.recover_safe_field.side_effect = lambda: self.assertEqual(self.ctrl.board.board_num, 1)
        session.close()

    def test_failure_recovers_before_propagation(self):
        session = ManualAdwinDiagnostic(self.ctrl)
        with self.assertRaisesRegex(RuntimeError, 'write failed'):
            session.write({}, Mock(side_effect=RuntimeError('write failed')))
        self.assertFalse(session.active)
        self.assertIsNone(self.store.pending())
        self.assertIn('write failed', self.store.read()['record']['error'])

    def test_failed_cleanup_remains_durable(self):
        session = ManualAdwinDiagnostic(self.ctrl)
        session.write({}, lambda: None)
        self.ctrl.recover_safe_field.side_effect = RuntimeError('DAC still high')
        with self.assertRaises(HardwareSafetyError):
            session.close()
        self.assertIsNotNone(HardwareSafetyStore(self.path).pending())

    def test_invalid_request_does_not_acquire_or_output(self):
        action = Mock()
        with self.assertRaises(ValueError):
            ManualAdwinDiagnostic(self.ctrl).write({'voltage': math.nan}, action)
        action.assert_not_called()
        self.assertIsNone(self.store.pending())


class AuxiliaryWorkerTests(DiagnosticFixture, unittest.TestCase):
    def capture(self, request, **kwargs):
        self.assertEqual(self.store.pending()['family'], 'af_diagnostic')
        validate_dense_capture_request(request)
        t = [i / request.io_rate_hz for i in range(1000)]
        values = [request.amplitude_v * math.sin(2 * math.pi * request.sine_freq_hz * x) for x in t]
        return AdwinDenseCaptureResult(t, values, values, 1000, 1000, 0, 999,
                                      1 / request.io_rate_hz, request.io_rate_hz / request.sine_freq_hz, 0, 1000)

    def test_full_clipping_range_and_zero_baseline_cleanup_before_success(self):
        config = clip.ClipTestConfig(scan_points=3)
        worker = clip.AutoClipWorker(clip.BackendConfig(), self.ctrl.limits, 'axial', config)
        results, settled, failed = [], [], []
        worker.result_ready.connect(lambda result: (self.assertIsNone(self.store.pending()), results.append(result)))
        worker.settled.connect(lambda: settled.append(True))
        worker.failed.connect(failed.append)
        self.ctrl.run_dense_loopback.side_effect = self.capture
        with patch.object(clip, 'AdwinAFController', return_value=self.ctrl):
            worker.run()
        self.assertEqual(failed, [])
        self.assertEqual(len(results), 1)
        self.assertEqual(settled, [True])
        requests = [call.args[0] for call in self.ctrl.run_dense_loopback.call_args_list]
        self.assertEqual([r.amplitude_v for r in requests], [0, 5, 10, 10, 5, 0])
        self.assertTrue(all(r.diagnostic_ceiling_v == 10 for r in requests))
        self.ctrl.recover_safe_field.assert_called_once()

    def test_capture_failure_settles_without_success_after_cleanup(self):
        worker = comms.SineLoopbackWorker(self.ctrl, 5, 1, 1, 1000, 1, 1)
        success, failed, settled = [], [], []
        worker.finished.connect(lambda: success.append(True))
        worker.failed.connect(failed.append)
        worker.settled.connect(lambda: settled.append(True))
        self.ctrl.run_dense_loopback.side_effect = RuntimeError('capture failed')
        worker.run()
        self.assertEqual(success, [])
        self.assertEqual(settled, [True])
        self.assertIn('capture failed', failed[0])
        self.ctrl.recover_safe_field.assert_called_once()
        self.assertIsNone(self.store.pending())

    def test_cancel_during_returned_capture_never_reports_success(self):
        worker = comms.SineLoopbackWorker(self.ctrl, 5, 1, 1, 1000, 1, 1)
        success, failed = [], []
        worker.finished.connect(lambda: success.append(True))
        worker.capture_ready.connect(lambda *args: success.append(True))
        worker.failed.connect(failed.append)
        def capture(request, **kwargs):
            result = self.capture(request)
            worker.stop()
            return result
        self.ctrl.run_dense_loopback.side_effect = capture
        worker.run()
        self.assertEqual(success, [])
        self.assertEqual(len(failed), 1)
        self.ctrl.recover_safe_field.assert_called_once()
        self.assertIsNone(self.store.pending())

    def test_cancelled_workers_do_not_initialize_outputs_or_report_success(self):
        workers = [comms.SineLoopbackWorker(self.ctrl, 5, 1, 1, 1000, 1, 1),
                   comms.SelfTestWorker(self.ctrl, 1, 1),
                   clip.AutoClipWorker(clip.BackendConfig(), self.ctrl.limits, 'axial', clip.ClipTestConfig())]
        for worker in workers:
            success, failed = [], []
            signal = getattr(worker, 'all_done', None) or worker.finished
            signal.connect(lambda *args: success.append(True))
            worker.failed.connect(failed.append)
            worker.stop()
            worker.run()
            self.assertEqual(success, [])
            self.assertEqual(len(failed), 1)
        self.assertEqual(self.ctrl.mock_calls, [])

    def test_connect_is_read_only_and_boot_is_guarded_cleanup_before_connected(self):
        for force in (False, True):
            with self.subTest(force=force):
                self.ctrl.reset_mock()
                worker = comms.AdwinConnectWorker(comms.AdwinCommsConfig(), force_reboot=force)
                connected, failures = [], []
                worker.connected.connect(lambda *args: (self.assertIsNone(self.store.pending()), connected.append(args)))
                worker.failed.connect(lambda *args: failures.append(args))
                self.ctrl.boot_board.side_effect = lambda: self.assertIsNotNone(self.store.pending())
                with patch.object(comms, 'AdwinAFController', return_value=self.ctrl):
                    worker.run()
                self.assertEqual(failures, [])
                self.assertEqual(len(connected), 1)
                self.assertEqual(self.ctrl.boot_board.call_count, int(force))
                self.assertEqual(self.ctrl.recover_safe_field.call_count, int(force))

    def test_unfinished_main_treatment_prevents_boot(self):
        self.store.begin('pulse', {}, {})
        worker = comms.AdwinConnectWorker(comms.AdwinCommsConfig(), force_reboot=True)
        failures = []
        worker.failed.connect(lambda *args: failures.append(args))
        with patch.object(comms, 'AdwinAFController', return_value=self.ctrl):
            worker.run()
        self.assertEqual(len(failures), 1)
        self.ctrl.boot_board.assert_not_called()


class DiagnosticRangeTests(unittest.TestCase):
    def test_zero_dac_return_still_checks_native_error_channel(self):
        controller = object.__new__(AdwinAFController)
        controller.board = AdwinBoardConfig()
        controller._dll = Mock()
        controller._dll.Set_DAC.return_value = 0
        controller._raise_if_error = Mock(side_effect=AdwinError('native DAC error'))
        with self.assertRaisesRegex(AdwinError, 'native DAC error'):
            controller.set_dac(1, 0)
        controller._raise_if_error.assert_called_once()

    def test_invalid_manual_dac_requests_fail_before_native_writes(self):
        controller = object.__new__(AdwinAFController)
        controller._dll = Mock()
        for channel, voltage in ((3, 1), (True, 1), (1, 11), (1, math.nan), (1, True)):
            with self.subTest(channel=channel, voltage=voltage), self.assertRaises(AdwinError):
                controller.set_dac(channel, voltage)
        self.assertEqual(controller._dll.mock_calls, [])

    def test_zero_baseline_requires_explicit_trial_ceiling(self):
        request = AdwinDenseCaptureRequest(10, 0, 1000, .1)
        with self.assertRaises(AdwinError):
            validate_dense_capture_request(request)
        request.diagnostic_ceiling_v = 10
        validate_dense_capture_request(request)
        for ceiling in (0, 11, math.nan, True):
            request.diagnostic_ceiling_v = ceiling
            with self.assertRaises(AdwinError):
                validate_dense_capture_request(request)


class AuxiliaryUiTests(DiagnosticFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.app = comms.QtWidgets.QApplication.instance() or comms.QtWidgets.QApplication([])
        self.original_font = self.app.font()
        self.addCleanup(lambda: self.app.setFont(self.original_font))
        font_path = Path('C:/Windows/Fonts/segoeui.ttf')
        if font_path.exists():
            font_id = comms.QtGui.QFontDatabase.addApplicationFont(str(font_path))
            self.addCleanup(lambda: comms.QtGui.QFontDatabase.removeApplicationFont(font_id))
        self.app.setFont(comms.QtGui.QFont('Segoe UI', 10))
        for module, names in ((comms, ('CONFIG_PATH',)), (clip, ('COIL_CONFIG_PATH', 'BACKEND_CONFIG_PATH', 'CLIP_TEST_CONFIG_PATH'))):
            for name in names:
                override = patch.object(module, name, Path(self.directory.name) / (name + '.json'))
                override.start()
                self.addCleanup(override.stop)

    def make_comms(self):
        with patch.object(comms.AdwinCommsApp, '_try_init_dll', new=lambda self: None):
            window = comms.AdwinCommsApp()
        window._ctrl, window._booted = self.ctrl, True
        window._update_hw_enabled()
        self.addCleanup(window.close)
        return window

    def test_manual_outputs_lock_station_settings_and_reset_releases(self):
        window = self.make_comms()
        window._spin_dac_v.setValue(2)
        window._write_dac()
        self.assertTrue(window._outputs_owned())
        self.assertFalse(window._spin_board.isEnabled())
        self.assertTrue(window._btn_recover_outputs.isEnabled())
        window._run_sig()
        self.assertIsNone(window._worker)
        window._all_relays_off()
        self.assertFalse(window._outputs_owned())
        self.assertTrue(window._spin_board.isEnabled())
        self.assertIsNone(self.store.pending())

    def test_close_retains_live_worker_until_actual_thread_settlement(self):
        window = self.make_comms()
        worker, thread = Mock(), Mock()
        thread.isRunning.return_value = True
        window._worker, window._worker_thread = worker, thread
        event = Mock()
        window.closeEvent(event)
        event.ignore.assert_called_once()
        worker.stop.assert_called_once()
        self.assertIs(window._worker_thread, thread)
        window._cleanup_worker()
        self.assertIs(window._worker_thread, thread)
        thread.wait.assert_not_called()
        thread.terminate.assert_not_called()
        thread.isRunning.return_value = False
        with patch.object(comms.QtCore.QTimer, 'singleShot') as schedule:
            window._cleanup_worker()
        self.assertIsNone(window._worker_thread)
        schedule.assert_called_once_with(0, window.close)

    def test_compact_helpers_keep_controls_and_plots_available(self):
        for window in (self.make_comms(), clip.MainWindow()):
            with self.subTest(window=type(window).__name__):
                screen = SimpleNamespace(availableGeometry=lambda: comms.QtCore.QRect(0, 0, 800, 800))
                window._fit_to_screen(screen)
                window.show()
                self.app.processEvents()
                self.assertLessEqual(window.width(), 736)
                self.assertLessEqual(window.height(), 720)
                self.assertEqual(window._compact_tabs.count(), 2)
                self.assertIs(window._compact_tabs.widget(0), window._controls_scroll)
                window._compact_tabs.setCurrentIndex(1)
                self.app.processEvents()
                self.assertTrue(window._plot_card.isVisible())
                window.close()
