from dataclasses import asdict
import hashlib
import json
import os
from pathlib import Path
import tempfile
from unittest.mock import Mock, patch
import unittest

from rapidpy_common.adwin_af import AdwinBoardConfig, AdwinCoilLimits, AdwinRampResult, AdwinDenseCaptureResult
from rapidpy_common.af_diagnostic_safety import af_diagnostic_operation, recover_af_diagnostic, station_profile
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError
from rapid_main.hardware_contracts import QueueHardwareBackend
from rapid_main.config import AppConfig


class AfDiagnosticFixture:
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "hardware_safety.json"
        self.store = HardwareSafetyStore(self.path)
        self.controller = Mock()
        self.controller.board = AdwinBoardConfig()
        self.controller.limits = AdwinCoilLimits()
        self.operation = {"action": "capture", "voltage": .5}


class AfDiagnosticSafetyTests(AfDiagnosticFixture, unittest.TestCase):
    def test_latch_exists_before_diagnostic_output_and_verified_cleanup_clears(self):
        with af_diagnostic_operation(self.controller, self.operation, store=self.store):
            self.assertEqual(HardwareSafetyStore(self.path).pending()["family"], "af_diagnostic")
            self.controller.run_ramp()
        self.assertIsNone(self.store.pending())
        self.assertEqual([call[0] for call in self.controller.mock_calls], ["run_ramp", "recover_safe_field"])
        self.assertEqual(self.store.read()["record"]["schema"], "rapidpy.af.diagnostic.v1")
        folders = list((self.path.parent / "hardware_diagnostics").iterdir())
        self.assertEqual(len(folders), 1)
        index = json.loads((folders[0] / "artifact_index.json").read_text())
        self.assertEqual(index["artifacts"][0]["sha256"], hashlib.sha256((folders[0] / "record.json").read_bytes()).hexdigest())

    def test_failed_cleanup_survives_restart_and_does_not_allow_next_diagnostic(self):
        self.controller.recover_safe_field.side_effect = RuntimeError("relay readback high")
        with self.assertRaisesRegex(HardwareSafetyError, "relay readback high"):
            with af_diagnostic_operation(self.controller, self.operation, store=self.store):
                self.controller.run_ramp()
        restarted = HardwareSafetyStore(self.path)
        self.assertIsNotNone(restarted.pending())
        self.controller.reset_mock()
        with self.assertRaises(HardwareSafetyError):
            with af_diagnostic_operation(self.controller, self.operation, store=restarted):
                self.fail("Unsafe diagnostic started.")
        self.assertEqual(self.controller.mock_calls, [])

    def test_aborted_diagnostic_retains_failure_evidence_and_checks_cleanup(self):
        with self.assertRaises(InterruptedError):
            with af_diagnostic_operation(self.controller, self.operation, store=self.store):
                raise InterruptedError("operator stop")
        self.controller.recover_safe_field.assert_called_once()
        state = self.store.read()
        self.assertEqual(state["status"], "verified")
        self.assertIn("operator stop", state["record"]["error"])

    def test_main_treatment_latch_blocks_diagnostic_before_outputs(self):
        self.store.begin("pulse", {"field": 50}, {"board": 1})
        with self.assertRaises(HardwareSafetyError):
            with af_diagnostic_operation(self.controller, self.operation, store=self.store):
                self.controller.run_ramp()
        self.assertEqual(self.controller.mock_calls, [])

    def test_main_operation_lease_blocks_helper_even_before_begin(self):
        with self.store.operation_lease():
            with self.assertRaises(HardwareSafetyError):
                with af_diagnostic_operation(self.controller, self.operation, store=HardwareSafetyStore(self.path)):
                    self.controller.run_ramp()
        self.assertEqual(self.controller.mock_calls, [])

    def test_restart_helper_recovery_stops_outputs_without_boot_or_ramp(self):
        self.store.begin("af_diagnostic", self.operation, station_profile(self.controller))
        record = recover_af_diagnostic(self.controller, store=HardwareSafetyStore(self.path))
        self.assertTrue(record.safe_state_confirmed)
        self.assertIsNone(self.store.pending())
        self.controller.recover_safe_field.assert_called_once()
        self.controller.boot_board.assert_not_called()
        self.controller.run_ramp.assert_not_called()

    def test_helper_cannot_recover_main_treatment(self):
        self.store.begin("arm", {}, {"different": "profile"})
        with self.assertRaisesRegex(HardwareSafetyError, "main app"):
            recover_af_diagnostic(self.controller, store=self.store)
        self.controller.recover_safe_field.assert_not_called()

    def test_changed_board_blocks_helper_recovery_without_io(self):
        self.store.begin("af_diagnostic", self.operation, station_profile(self.controller))
        self.controller.board.board_num = 2
        with self.assertRaisesRegex(HardwareSafetyError, "Station wiring"):
            recover_af_diagnostic(self.controller, store=self.store)
        self.controller.recover_safe_field.assert_not_called()

    def test_main_app_routes_helper_recovery_to_tuner_without_io(self):
        self.store.begin("af_diagnostic", self.operation, station_profile(self.controller))
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._config.general.nocomm = False
        backend._safety_store = self.store
        backend._client = Mock()
        self.assertTrue(backend.has_unresolved_hardware_fault)
        with self.assertRaisesRegex(RuntimeError, "ADwin helper"):
            backend.return_to_safe_state()
        self.assertEqual(backend._client.mock_calls, [])

    def test_journal_failure_before_diagnostic_causes_no_output(self):
        with patch("rapidpy_common.hardware_safety.os.replace", side_effect=OSError("disk full")):
            with self.assertRaises(HardwareSafetyError):
                with af_diagnostic_operation(self.controller, self.operation, store=self.store):
                    self.controller.run_ramp()
        self.assertEqual(self.controller.mock_calls, [])

    def test_evidence_publication_failure_keeps_latch_despite_physical_cleanup(self):
        with patch("rapidpy_common.af_diagnostic_safety.publish_diagnostic_record", side_effect=HardwareSafetyError("evidence disk full")):
            with self.assertRaisesRegex(HardwareSafetyError, "evidence disk full"):
                with af_diagnostic_operation(self.controller, self.operation, store=self.store):
                    self.controller.run_ramp()
        self.controller.recover_safe_field.assert_called_once()
        self.assertIsNotNone(HardwareSafetyStore(self.path).pending())


class AfTunerWorkerSafetyTests(AfDiagnosticFixture, unittest.TestCase):
    # Worker-specific tests use the real helper entry points, with only the
    # physical controller injected.
    def setUp(self):
        super().setUp()
        from af_tuner import app as tuner
        self.tuner = tuner
        self.environment = patch.dict(os.environ, {"RAPID_SAFETY_STATE": str(self.path)})
        self.environment.start()
        self.addCleanup(self.environment.stop)
        self.controller_factory = patch.object(tuner, "AdwinAFController", return_value=self.controller)
        self.controller_factory.start()
        self.addCleanup(self.controller_factory.stop)
        self.controller.run_dense_loopback.return_value = AdwinDenseCaptureResult([0., .001], [0., .5], [0., .3],
                                                                                 2, 2, 0, 1, .001, 12.5, 0, 2)

    def sweep(self):
        return self.tuner.AutoTuneWorker(self.tuner.BackendConfig(), AdwinCoilLimits(), 800., 810., 10.,
                                          100, 10000., .5, .5, "axial")

    def capture(self):
        return self.tuner.DenseCaptureWorker(self.tuner.BackendConfig(), AdwinCoilLimits(), "axial", 800., .5,
                                             100, 10000., 1., 1., 10, "test capture")

    def observe(self, worker):
        self.failures, self.finished = [], []
        worker.failed.connect(self.failures.append)
        worker.finished.connect(lambda *_: self.finished.append(self.store.pending()))

    def test_sweep_holds_lease_for_all_ramps_and_emits_after_cleanup(self):
        worker = self.sweep()
        self.observe(worker)
        def ramp(request, *, should_cancel):
            self.assertFalse(should_cancel())
            self.assertEqual(self.store.pending()["plan"]["action"], "auto_tune_sweep")
            with self.assertRaises(HardwareSafetyError):
                with HardwareSafetyStore(self.path).operation_lease():
                    self.fail("Another hardware owner entered during a sweep.")
            return AdwinRampResult(10, 10, 1, 8, .3, .5, 1., .0001, 12.5)
        self.controller.run_ramp.side_effect = ramp
        worker.run()
        self.assertEqual(self.failures, [])
        self.assertEqual(self.finished, [None])
        self.assertEqual(self.controller.run_ramp.call_count, 2)
        self.controller.recover_safe_field.assert_called_once()
        self.assertEqual(len(self.store.read()["record"]["observations"]), 2)

    def test_sweep_cancel_inside_ramp_checks_cleanup_and_emits_no_success(self):
        worker = self.sweep()
        self.observe(worker)
        def ramp(_request, *, should_cancel):
            worker.abort()
            self.assertTrue(should_cancel())
            raise InterruptedError("operator stop")
        self.controller.run_ramp.side_effect = ramp
        worker.run()
        self.assertEqual(self.finished, [])
        self.assertIn("operator stop", self.failures[0])
        self.controller.recover_safe_field.assert_called_once()
        self.assertIsNone(self.store.pending())

    def test_capture_journals_before_relay_and_cleans_before_ready_signal(self):
        worker = self.capture()
        self.observe(worker)
        ready = []
        worker.capture_ready.connect(lambda *_: ready.append(self.store.pending()))
        self.controller.set_af_relays.side_effect = lambda *_args, **_kwargs: self.assertIsNotNone(self.store.pending())
        worker.run()
        self.assertEqual(self.failures, [])
        self.assertEqual(ready, [None])
        self.assertEqual(self.finished, [None])
        self.controller.recover_safe_field.assert_called_once()
        self.assertEqual(self.store.read()["record"]["observations"][0]["capture"]["dac_v"], [0., .5])

    def test_capture_cleanup_failure_emits_no_ready_or_success(self):
        worker = self.capture()
        self.observe(worker)
        ready = []
        worker.capture_ready.connect(ready.append)
        self.controller.recover_safe_field.side_effect = RuntimeError("relay stuck")
        worker.run()
        self.assertEqual(ready, [])
        self.assertEqual(self.finished, [])
        self.assertIn("relay stuck", self.failures[0])
        self.assertIsNotNone(self.store.pending())

    def test_pre_cancelled_workers_do_not_connect_or_write(self):
        for worker in (self.sweep(), self.capture()):
            self.observe(worker)
            if isinstance(worker, self.tuner.AutoTuneWorker):
                worker.abort()
            else:
                worker.stop()
            worker.run()
            self.assertEqual(self.finished, [])
            self.assertIn("before output", self.failures[0])
        self.assertEqual(self.controller.mock_calls, [])

    def fake_window(self):
        from types import SimpleNamespace
        backend = self.tuner.BackendConfig(process_file="sineout.T91")
        return SimpleNamespace(_backend_config=backend, _ctrl=None, _connected=False, _last_version=0,
                               _build_backend_from_widgets=lambda: backend, _runtime_limits=lambda: AdwinCoilLimits(),
                               _active_coil_name=lambda: "axial", _set_backend_status=Mock(), _append=Mock(),
                               _update_comm_snapshot=Mock(), _set_comm_summary=Mock(), _set_relay_status=Mock())

    def test_ui_connection_only_probes_existing_board(self):
        window = self.fake_window()
        self.controller.test_version.return_value = 9
        self.controller.get_digout.return_value = 0
        self.tuner.MainWindow._connect_backend(window)
        self.assertTrue(window._connected)
        self.controller.boot_board.assert_not_called()
        self.controller.set_af_relays.assert_not_called()
        self.controller.recover_safe_field.assert_not_called()
        self.assertFalse(self.path.exists())

    def test_ui_relay_test_returns_outputs_off_and_preserves_readback(self):
        window = self.fake_window()
        window._ctrl = self.controller
        self.controller.set_af_relays.return_value = 1
        self.controller.get_digout.return_value = 1
        self.tuner.MainWindow._apply_relays(window)
        self.controller.recover_safe_field.assert_called_once()
        window._update_comm_snapshot.assert_called_once_with(log_result=False, relay_word=0)
        self.assertEqual(self.store.read()["record"]["observations"], [{"selected_relay_word": 1}])

    def test_ui_close_waits_for_active_worker_cleanup_without_thread_termination(self):
        from types import SimpleNamespace
        worker = self.capture()
        thread, event = Mock(), Mock()
        thread.isRunning.return_value = True
        window = SimpleNamespace(_worker_thread=thread, _worker=worker, _queued_capture_freq=800.,
                                 _close_after_cleanup=False, _append=Mock())
        self.tuner.MainWindow.closeEvent(window, event)
        self.assertTrue(worker._stop)
        self.assertTrue(window._close_after_cleanup)
        self.assertIsNone(window._queued_capture_freq)
        event.ignore.assert_called_once()
        event.accept.assert_not_called()
        thread.terminate.assert_not_called()

    def test_ui_cleanup_closes_without_launching_queued_waveform(self):
        from types import SimpleNamespace
        window = SimpleNamespace(_task_kind="sweep", _worker_thread=Mock(), _worker=Mock(),
                                 _set_busy=Mock(), _set_coil_locked=Mock(), _queued_capture_freq=800.,
                                 _close_after_cleanup=True, close=Mock(), _start_capture=Mock())
        with patch.object(self.tuner.QtCore.QTimer, "singleShot") as schedule:
            self.tuner.MainWindow._cleanup_worker(window)
        schedule.assert_called_once_with(0, window.close)
        window._start_capture.assert_not_called()
        self.assertIsNone(window._worker_thread)


class AfTunerUiSafetyTests(AfDiagnosticFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        from af_tuner import app as tuner
        self.tuner = tuner
        self.application = tuner.QtWidgets.QApplication.instance() or tuner.QtWidgets.QApplication([])
        original_font = self.application.font()
        self.addCleanup(lambda: self.application.setFont(original_font))
        font_path = Path("C:/Windows/Fonts/segoeui.ttf")
        if font_path.exists():
            font_id = tuner.QtGui.QFontDatabase.addApplicationFont(str(font_path))
            self.addCleanup(lambda: tuner.QtGui.QFontDatabase.removeApplicationFont(font_id))
        self.application.setFont(tuner.QtGui.QFont("Segoe UI", 10))
        for name in ("COIL_CONFIG_PATH", "BACKEND_CONFIG_PATH", "AUTOTUNE_CONFIG_PATH", "CLIP_CAPTURE_PATH"):
            override = patch.object(tuner, name, Path(self.directory.name) / (name + ".json"))
            override.start()
            self.addCleanup(override.stop)
        factory = patch.object(tuner, "AdwinAFController", side_effect=AssertionError("UI construction attempted hardware I/O"))
        factory.start()
        self.addCleanup(factory.stop)
        self.window = tuner.MainWindow()
        self.addCleanup(self.window.close)
        self.window.show()
        self.application.processEvents()

    def screen(self, width, height):
        from types import SimpleNamespace
        return SimpleNamespace(availableGeometry=lambda: self.tuner.QtCore.QRect(0, 0, width, height))

    def test_compact_screen_exposes_controls_and_plots_without_horizontal_clip(self):
        self.window._fit_to_screen(self.screen(800, 800))
        self.application.processEvents()
        self.assertEqual(self.window._compact_tabs.count(), 2)
        self.assertTrue(self.window._compact_layout)
        self.assertEqual(self.window._controls_scroll.horizontalScrollBar().maximum(), 0)
        self.assertLessEqual(self.window.width(), 800)
        self.assertLessEqual(self.window.height(), 800)
        self.assertIsNone(self.window._ctrl)

    def test_wide_compact_transitions_preserve_both_widgets_without_duplicate_tabs(self):
        controls, plot = self.window._controls_scroll, self.window._plot_card
        for width, compact in ((2000, False), (800, True), (2000, False), (800, True)):
            self.window._fit_to_screen(self.screen(width, 1100))
            self.application.processEvents()
            self.assertEqual(self.window._compact_layout, compact)
            self.assertEqual(self.window._compact_tabs.count(), 2 if compact else 0)
            self.assertIs(self.window._controls_scroll, controls)
            self.assertIs(self.window._plot_card, plot)

    def test_recovery_button_is_visible_within_compact_controls_viewport(self):
        self.window._fit_to_screen(self.screen(800, 800))
        scroll = self.window._controls_scroll
        scroll.ensureWidgetVisible(self.window.relays_off_btn)
        self.application.processEvents()
        point = self.window.relays_off_btn.mapTo(scroll.viewport(), self.tuner.QtCore.QPoint(0, 0))
        self.assertGreaterEqual(point.x(), 0)
        self.assertLessEqual(point.x() + self.window.relays_off_btn.width(), scroll.viewport().width())
        self.assertGreaterEqual(point.y(), 0)
        self.assertLessEqual(point.y() + self.window.relays_off_btn.height(), scroll.viewport().height())
