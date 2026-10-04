from __future__ import annotations

import math
import unittest
from typing import cast
from types import SimpleNamespace

from PySide6 import QtCore, QtWidgets

from rapid_main.dialogs.dc_motors import DCMotorDialog
from rapid_main.diagnostic_services import DCMotorNoCommBackend
from rapidpy_common.hardware import MotorTelemetry


class _FakeScreen:
    def __init__(self, available: QtCore.QRect) -> None:
        self._available = available

    def availableGeometry(self) -> QtCore.QRect:
        return self._available


class _TestAppMixin:
    @classmethod
    def setUpClass(cls) -> None:
        if QtWidgets.QApplication.instance() is None:
            cls._qt_app = QtWidgets.QApplication([])
        else:
            cls._qt_app = None

    @classmethod
    def tearDownClass(cls) -> None:
        if cls._qt_app is not None:
            cls._qt_app.quit()
            cls._qt_app = None


class DCMotorDialogTelemetryTest(_TestAppMixin, unittest.TestCase):
    @staticmethod
    def _backend() -> DCMotorNoCommBackend:
        backend = DCMotorNoCommBackend(port="COM1", baud=9600)
        backend.connect("COM1", 9600)
        return backend

    def _build_dialog(self) -> DCMotorDialog:
        backend = self._backend()
        dialog = DCMotorDialog(backend=backend)
        dialog._connected = True
        return dialog

    def _close_dialog(self, dialog: DCMotorDialog) -> None:
        dialog._telemetry_thread.stop_monitoring()
        dialog._telemetry_thread.wait(1000)
        dialog.close()
        dialog.deleteLater()

    def test_telemetry_stream_updates_input_output_and_torque_fields(self) -> None:
        dialog = self._build_dialog()
        try:
            self.assertEqual(dialog._values_grid.count(), len(dialog._value_grid_tiles))
            axis = dialog._selected_axis()
            sample = MotorTelemetry(
                timestamp=12.5,
                axis_name=axis,
                target_position=18_500,
                actual_position=18_490,
                position_error=10,
                velocity_1=42,
                velocity_2=39,
                actual_torque=512,
            )

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 18,500")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 18,490")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -10")
            self.assertEqual(dialog.error_value.text(), "Error +10")
            self.assertEqual(dialog.velocity_value.text(), "Velocity1 +42 SAV")
            self.assertEqual(dialog.velocity2_value.text(), "Velocity2 +39 SAV")
            self.assertIn("Torque +512", dialog.torque_value.text())
            self.assertIn("% FS", dialog.torque_value.text())
            self.assertGreater(dialog.torque_value.text().count("("), 0)
            self.assertTrue(dialog._traces["time"])
            self.assertTrue(dialog._traces["torque"])
            self.assertTrue(dialog._torque_seen)
        finally:
            self._close_dialog(dialog)

    def test_missing_torque_telemetry_stays_live_with_na_placeholder(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = SimpleNamespace(
                timestamp=12.5,
                axis_name=axis,
                target_position=12_345,
                actual_position=12_300,
                position_error=-45,
                velocity_1=20,
                velocity_2=19,
            )

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 12,345")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 12,300")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -45")
            self.assertEqual(dialog.error_value.text(), "Error -45")
            self.assertEqual(dialog.velocity_value.text(), "Velocity1 +20 SAV")
            self.assertEqual(dialog.velocity2_value.text(), "Velocity2 +19 SAV")
            self.assertEqual(dialog.torque_value.text(), "Torque -- (N/A)")
            self.assertFalse(dialog._torque_seen)
            self.assertEqual(len(dialog._traces["torque"]), 1)
            self.assertTrue(math.isnan(dialog._traces["torque"][0]))

            dialog._refresh_plots()
            dialog._clear_traces()
            self.assertEqual(dialog.target_value.text(), "Input command --")
            self.assertEqual(dialog.torque_value.text(), "Torque -- (N/A)")
            self.assertFalse(dialog._torque_seen)
        finally:
            self._close_dialog(dialog)

    def test_partial_telemetry_without_optional_velocity_or_error(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = SimpleNamespace(
                timestamp=12.5,
                axis_name=axis,
                target_position=9000,
                actual_position=8990,
            )

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 9,000")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 8,990")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -10")
            self.assertEqual(dialog.error_value.text(), "Error --")
            self.assertEqual(dialog.velocity_value.text(), "Velocity1 --")
            self.assertEqual(dialog.velocity2_value.text(), "Velocity2 --")
            self.assertEqual(dialog.torque_value.text(), "Torque -- (N/A)")
            self.assertEqual(len(dialog._traces["error"]), 1)
            self.assertEqual(len(dialog._traces["velocity_1"]), 1)
            self.assertEqual(len(dialog._traces["velocity_2"]), 1)
            self.assertTrue(math.isnan(dialog._traces["error"][0]))
            self.assertTrue(math.isnan(dialog._traces["velocity_1"][0]))
            self.assertTrue(math.isnan(dialog._traces["velocity_2"][0]))
        finally:
            self._close_dialog(dialog)

    def test_torque_alias_field_updates_live_torque_readout(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = SimpleNamespace(
                timestamp=12.5,
                axis_name=axis,
                target_position=9000,
                actual_position=8996,
                position_error=4,
                velocity_1=11,
                velocity_2=9,
                feedback_torque=321,
            )

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -4")
            self.assertIn("321", dialog.torque_value.text())
            self.assertTrue(dialog._torque_seen)
            self.assertEqual(dialog.torque_value.text().count("("), 1)
        finally:
            self._close_dialog(dialog)

    def test_telemetry_alias_fields_for_command_output_are_supported(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = {
                "timestamp": "45.0",
                "axis_name": axis,
                "command": 24_000,
                "feedback": 23_998,
                "position_error": -2,
                "velocity_1": 120,
                "velocity_2": 118,
                "feedback_torque": 250,
            }
            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 24,000")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 23,998")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -2")
            self.assertIn("250", dialog.torque_value.text())
            self.assertIn("% FS", dialog.torque_value.text())
            self.assertTrue(dialog._torque_seen)
        finally:
            self._close_dialog(dialog)

    def test_telemetry_supports_axis_and_timestamp_alias_fields(self) -> None:
        dialog = self._build_dialog()
        try:
            selected = dialog._selected_axis()
            sample = {
                "time": "45.5",
                "axis": selected.lower(),
                "position_cmd": 33_200,
                "feedback_position": 33_140,
                "position_error": -60,
                "velocity_1": 12,
                "velocity_2": 11,
                "motor_torque": 88,
            }

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 33,200")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 33,140")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -60")
            self.assertIn("88", dialog.torque_value.text())
            self.assertTrue(dialog._traces["torque"])
        finally:
            self._close_dialog(dialog)

    def test_telemetry_samples_without_timestamp_still_update(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = {
                "axis_name": axis,
                "command": 19_000,
                "feedback": 18_995,
            }

            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 19,000")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 18,995")
            self.assertTrue(dialog._sample_tick)
        finally:
            self._close_dialog(dialog)

    def test_telemetry_panel_compacts_value_grid_when_area_is_narrow(self) -> None:
        dialog = self._build_dialog()
        try:
            wide = _FakeScreen(QtCore.QRect(0, 0, 1500, 900))
            compact = _FakeScreen(QtCore.QRect(0, 0, 1200, 900))
            narrow = _FakeScreen(QtCore.QRect(0, 0, 540, 700))

            dialog.resize(1400, 800)
            dialog._fit_to_screen(wide)
            self.assertEqual(dialog._values_per_row, 3)
            self.assertEqual(dialog._splitter.orientation(), QtCore.Qt.Orientation.Horizontal)

            dialog.resize(900, 700)
            dialog._fit_to_screen(compact)
            self.assertEqual(dialog._values_per_row, 2)
            self.assertEqual(dialog._splitter.orientation(), QtCore.Qt.Orientation.Vertical)

            dialog._fit_to_screen(narrow)
            self.assertEqual(dialog._values_per_row, 1)
            self.assertEqual(dialog._splitter.orientation(), QtCore.Qt.Orientation.Vertical)
            self.assertEqual(dialog._values_grid.count(), len(dialog._value_grid_tiles))
        finally:
            self._close_dialog(dialog)

    def test_dict_like_telemetry_sample_updates_live_feedback(self) -> None:
        dialog = self._build_dialog()
        try:
            axis = dialog._selected_axis()
            sample = {
                "timestamp": "13.75",
                "axis_name": axis,
                "target_position": "18600",
                "actual_position": "18595",
                "position_error": "5",
                "velocity_1": "55",
                "velocity_2": "50",
                "torque_adc": 401.2,
            }
            dialog._on_telemetry(sample)

            self.assertEqual(dialog.target_value.text(), "Input command 18,600")
            self.assertEqual(dialog.actual_value.text(), "Output feedback 18,595")
            self.assertEqual(dialog.io_delta_value.text(), "Input-output delta -5")
            self.assertEqual(dialog.error_value.text(), "Error +5")
            self.assertIn("401", dialog.torque_value.text())
        finally:
            self._close_dialog(dialog)

    def test_telemetry_ignores_samples_from_unselected_axis(self) -> None:
        dialog = self._build_dialog()
        try:
            selected = dialog._selected_axis()
            if selected == "Turning":
                unselected_axis = "Changer (X)"
            else:
                unselected_axis = "Turning"

            sample = {
                "timestamp": 10.0,
                "axis_name": unselected_axis,
                "target_position": 1200,
                "actual_position": 1180,
                "position_error": 20,
                "velocity_1": 5,
                "velocity_2": 6,
                "actual_torque": 99,
            }
            dialog._on_telemetry(cast(dict[str, object], sample))

            self.assertEqual(dialog.target_value.text(), "Input command --")
            self.assertFalse(dialog._traces["time"])
            self.assertFalse(dialog._traces["torque"])
            self.assertEqual(dialog._sample_tick, 0)
        finally:
            self._close_dialog(dialog)


if __name__ == "__main__":
    unittest.main(verbosity=2)
