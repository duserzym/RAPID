from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow as RapidMainWindow
from rapid_main.device_ownership import DeviceOwnershipManager
from rapid_main.vrm import (
    VRM_CONTEXT_ENV,
    VrmLaunchContextError,
    build_vrm_launch_context,
    vrm_launch_block_reason,
    write_vrm_launch_context,
)
from vrm_logger.session_manifest import (
    build_vrm_output_manifest,
    load_handoff_context,
    vrm_manifest_path,
    write_vrm_output_manifest,
)
from vrm_logger.app import MainWindow as VrmWindow


class VrmIntegrationTests(unittest.TestCase):
    def _context(self, **overrides: object) -> dict[str, object]:
        values: dict[str, object] = {
            "launch_id": "launch-7",
            "active_automation": False,
            "nocomm": True,
            "run_id": "",
            "sample_id": "VRM_DEMO",
            "operator": "operator-a",
            "software_version": "4.0.0+abc123",
            "config_hash": "config-7",
        }
        values.update(overrides)
        return build_vrm_launch_context(**values)

    def test_rapid_main_atomically_writes_complete_launch_context(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            path = write_vrm_launch_context(root, context=self._context())
            payload = json.loads(path.read_text(encoding="utf-8"))
            leftovers = list(root.glob("*.tmp"))

        self.assertEqual(payload["schema"], "rapidpy.vrm.launch_context.v1")
        self.assertFalse(payload["active_automation"])
        self.assertTrue(payload["nocomm"])
        self.assertEqual(payload["sample_id"], "VRM_DEMO")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertEqual(payload["software_version"], "4.0.0+abc123")
        self.assertEqual(leftovers, [])

    def test_launch_context_rejects_active_automation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(VrmLaunchContextError, "automation is active"):
                write_vrm_launch_context(
                    tmp,
                    context=self._context(active_automation=True),
                )

    def test_vrm_launch_block_reason_reports_ownership(self) -> None:
        self.assertIn(
            "queue_workflow",
            vrm_launch_block_reason(
                active_automation=False,
                measurement_owner="queue_workflow",
            ),
        )
        self.assertIn(
            "live measurement",
            vrm_launch_block_reason(active_automation=True),
        )
        self.assertEqual(vrm_launch_block_reason(active_automation=False), "")

    def test_vrm_logger_validates_handoff_context(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            handoff_path = write_vrm_launch_context(tmp, context=self._context())
            handoff = load_handoff_context({VRM_CONTEXT_ENV: str(handoff_path)})

        self.assertEqual(handoff["association_status"], "validated_handoff")
        self.assertEqual(handoff["run_association_status"], "unassociated")
        self.assertEqual(handoff["source_app"], "rapid_main")
        self.assertEqual(handoff["sample_id"], "VRM_DEMO")

    def test_handoff_claims_run_association_only_with_explicit_run_id(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = write_vrm_launch_context(
                tmp,
                context=self._context(run_id="rapid-run-42"),
            )
            handoff = load_handoff_context({VRM_CONTEXT_ENV: str(path)})

        self.assertEqual(handoff["association_status"], "validated_handoff")
        self.assertEqual(handoff["run_association_status"], "associated")

    def test_invalid_and_missing_handoffs_are_explicit(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            invalid_path = Path(tmp) / "invalid.json"
            invalid_path.write_text("{}", encoding="utf-8")
            invalid = load_handoff_context({VRM_CONTEXT_ENV: str(invalid_path)})
        standalone = load_handoff_context({})

        self.assertEqual(invalid["association_status"], "invalid")
        self.assertIn("unsupported handoff schema", invalid["error"])
        self.assertEqual(standalone["association_status"], "standalone")
        self.assertEqual(load_handoff_context({VRM_CONTEXT_ENV: ""})["association_status"], "invalid")

    def test_practice_window_records_a_simulated_session_without_serial_access(self) -> None:
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        self.addCleanup(app.setStyleSheet, app.styleSheet())
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            handoff = write_vrm_launch_context(root, context=self._context(intended_output_dir=root))
            live_csv = root / "vrm_log.csv"
            live_csv.write_bytes(b"existing live output\n")
            with (patch.dict("os.environ", {VRM_CONTEXT_ENV: str(handoff)}),
                  patch("vrm_logger.app.load_config") as load,
                  patch("vrm_logger.app.save_config") as save,
                  patch("vrm_logger.app.SquidSerialClient") as serial,
                  patch("vrm_logger.app.list_ports.comports") as ports):
                window = VrmWindow()
                try:
                    self.assertIn("Simulated", window.windowTitle())
                    self.assertFalse(window.connect_btn.isEnabled())
                    self.assertEqual(Path(window.file_edit.text()), live_csv)
                    window.interval_spin.setValue(0.05)
                    window.baseline_btn.click()
                    window.start_btn.click()
                    timer = QtCore.QElapsedTimer()
                    timer.start()
                    while len(window._time) < 2 and timer.elapsed() < 5000:
                        app.processEvents()
                        QtCore.QThread.msleep(5)
                    self.assertGreaterEqual(len(window._time), 2)
                    window.stop_btn.click()
                    timer.restart()
                    while window._thread is not None and timer.elapsed() < 5000:
                        app.processEvents()
                        QtCore.QThread.msleep(5)
                    self.assertIsNone(window._thread)
                finally:
                    window.close()
                    window.deleteLater()
                    app.processEvents()
                serial.assert_not_called()
                ports.assert_not_called()
                load.assert_not_called()
                save.assert_not_called()

            output, = (root / "SIMULATED").glob("*/vrm_log.csv")
            manifest, = output.parent.glob("*.vrm.json")
            payload = json.loads(manifest.read_bytes())
            self.assertTrue(payload["provenance"]["simulated"])
            self.assertEqual(payload["device"]["port"], "")
            self.assertEqual(payload["handoff_context"]["sample_id"], "VRM_DEMO")
            self.assertEqual(payload["outcome"], "operator_stopped")
            self.assertGreaterEqual(payload["row_count"], 2)
            self.assertEqual(payload["output_sha256"], hashlib.sha256(output.read_bytes()).hexdigest())
            self.assertEqual(live_csv.read_bytes(), b"existing live output\n")

    def test_invalid_window_handoff_keeps_serial_unavailable(self) -> None:
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        self.addCleanup(app.setStyleSheet, app.styleSheet())
        with tempfile.TemporaryDirectory() as directory:
            bad_context = Path(directory) / "missing.json"
            with (patch.dict("os.environ", {VRM_CONTEXT_ENV: str(bad_context)}),
                  patch("vrm_logger.app.load_config") as load,
                  patch("vrm_logger.app.SquidSerialClient") as serial,
                  patch("vrm_logger.app.list_ports.comports") as ports):
                window = VrmWindow()
                try:
                    self.assertIn("VRM launch blocked", window.status_label.text())
                    self.assertFalse(window.connect_btn.isEnabled())
                    self.assertFalse(window.baseline_btn.isEnabled())
                    window._connect()
                    window._refresh_ports()
                finally:
                    window.close()
                    window.deleteLater()
                    app.processEvents()
                serial.assert_not_called()
                ports.assert_not_called()
                load.assert_not_called()

    def test_vrm_logger_finalizes_hashed_immutable_session_sidecar(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            csv_path = root / "run.csv"
            csv_bytes = b"time_s,x\n1.0,2.0\n"
            csv_path.write_bytes(csv_bytes)
            manifest_path = write_vrm_output_manifest(
                csv_path,
                session_start_epoch=1_700_000_000.0,
                session_end_epoch=1_700_000_010.0,
                session_id="session-1",
                outcome="operator_stopped",
                row_count=1,
                write_mode="new",
                interval_s=2.5,
                spacing_mode="Log",
                display_unit="Moment",
                baseline_volts=(0.1, 0.2, 0.3),
                calibration={"x": 1.0, "y": 2.0, "z": 3.0, "range_factor": 1e-6},
                handoff_context={"association_status": "associated"},
                device_port="COM7",
            )
            payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            with self.assertRaises(FileExistsError):
                write_vrm_output_manifest(
                    csv_path,
                    session_start_epoch=1_700_000_000.0,
                    session_end_epoch=1_700_000_010.0,
                    session_id="session-1",
                    outcome="operator_stopped",
                    row_count=1,
                    write_mode="append",
                    interval_s=2.5,
                    spacing_mode="Log",
                    display_unit="Moment",
                    baseline_volts=None,
                    calibration={},
                )

        self.assertEqual(manifest_path.name, "run.session-1.vrm.json")
        self.assertEqual(payload["schema"], "rapidpy.vrm.session_manifest.v2")
        self.assertEqual(payload["outcome"], "operator_stopped")
        self.assertEqual(payload["row_count"], 1)
        self.assertEqual(payload["device"]["port"], "COM7")
        self.assertFalse(payload["provenance"]["simulated"])
        self.assertEqual(payload["output_size_bytes"], len(csv_bytes))
        self.assertEqual(payload["output_sha256"], hashlib.sha256(csv_bytes).hexdigest())

    def test_manifest_path_keeps_legacy_lookup_and_supports_unique_session(self) -> None:
        self.assertEqual(vrm_manifest_path(Path("example.csv")), Path("example.vrm.json"))
        self.assertEqual(
            vrm_manifest_path(Path("example.csv"), "session-2"),
            Path("example.session-2.vrm.json"),
        )

    def test_manifest_requires_truthful_error_and_valid_lifecycle(self) -> None:
        common = dict(
            output_csv=Path("run.csv"),
            session_start_epoch=1_700_000_000.0,
            session_end_epoch=1_700_000_001.0,
            session_id="session-3",
            outcome="error",
            row_count=0,
            write_mode="new",
            interval_s=1.0,
            spacing_mode="Linear",
            display_unit="Volts",
            baseline_volts=None,
            calibration={"x": 1.0, "y": 1.0, "z": 1.0, "range_factor": 1.0},
            handoff_context={},
        )
        with self.assertRaisesRegex(ValueError, "requires an error message"):
            build_vrm_output_manifest(**common)
        payload = build_vrm_output_manifest(**common, error="serial reply timed out")
        self.assertEqual(payload["error"], "serial reply timed out")

    def test_main_window_vrm_launch_refuses_active_automation_without_override(self) -> None:
        class FakeWindow:
            def __init__(self) -> None:
                self._ownership = DeviceOwnershipManager()
                self.status = ""

            def _has_active_automation(self) -> bool:
                return True

            def set_status(self, value: str) -> None:
                self.status = value

        window = FakeWindow()
        with patch("rapid_main.app.QtWidgets.QMessageBox.warning") as warning:
            RapidMainWindow._launch_vrm(window)

        self.assertIn("active", window.status.lower())
        warning.assert_called_once()
        self.assertFalse(window._ownership.is_owned("measurement"))
        self.assertFalse(window._ownership.is_owned("squid"))

    def test_main_window_releases_retained_squid_clients_before_handoff(self) -> None:
        class _MeasurementBackend:
            def __init__(self) -> None:
                self.released = False

            def release_squid_for_external_tool(self) -> None:
                self.released = True

        class _DiagnosticBackend:
            def __init__(self) -> None:
                self.connected = True

            def is_connected(self) -> bool:
                return self.connected

            def disconnect(self) -> None:
                self.connected = False

        class _Window:
            def __init__(self) -> None:
                self._measurement_backend = _MeasurementBackend()
                self._squid_backend = _DiagnosticBackend()

        window = _Window()
        RapidMainWindow._release_squid_connections_for_vrm(window)

        self.assertTrue(window._measurement_backend.released)
        self.assertFalse(window._squid_backend.connected)

if __name__ == "__main__":
    unittest.main()
