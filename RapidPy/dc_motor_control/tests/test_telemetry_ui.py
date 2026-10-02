from __future__ import annotations

import math
import unittest

from PySide6 import QtCore, QtWidgets

from rapidpy_common.hardware import MotorTelemetry

from dc_motor_control.app import MainWindow


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


class DCMotorControlTelemetryUiTest(_TestAppMixin, unittest.TestCase):
    def _build_window(self) -> MainWindow:
        window = MainWindow()
        # Keep threads quiet and avoid any window mapping side effects in tests.
        return window

    def _close_window(self, window: MainWindow) -> None:
        window._plot_timer.stop()
        window._telemetry_thread.stop_monitoring()
        window._telemetry_thread.wait(1000)
        window.client.disconnect()
        window.close()
        window.deleteLater()

    class _PlotCounter:
        def __init__(self) -> None:
            self.calls = 0

        def setData(self, *_args: object, **_kwargs: object) -> None:
            self.calls += 1

    def test_main_window_telemetry_updates_live_input_output_and_torque(self) -> None:
        window = self._build_window()
        try:
            axis = window._selected_axis()
            sample = MotorTelemetry(
                timestamp=4.2,
                axis_name=axis.name,
                target_position=12_345,
                actual_position=12_300,
                position_error=-45,
                velocity_1=24,
                velocity_2=21,
                actual_torque=512,
            )

            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 12,345")
            self.assertEqual(window.actual_value.text(), "Output feedback 12,300")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -45")
            self.assertEqual(window.error_value.text(), "Error -45")
            self.assertEqual(window.velocity_value.text(), "Velocity1 +24 SAV")
            self.assertEqual(window.velocity2_value.text(), "Velocity2 +21 SAV")
            self.assertIn("Torque", window.torque_value.text())
            self.assertIn("% FS", window.torque_value.text())
            self.assertTrue(window._torque_seen)
            self.assertTrue(window._traces["time"])
            self.assertTrue(window._traces["torque"])
        finally:
            self._close_window(window)

    def test_main_window_telemetry_shows_na_when_torque_missing(self) -> None:
        window = self._build_window()
        try:
            axis = window._selected_axis()
            sample = MotorTelemetry(
                timestamp=4.2,
                axis_name=axis.name,
                target_position=7_700,
                actual_position=7_699,
                position_error=1,
                velocity_1=-4,
                velocity_2=-3,
                actual_torque=None,
            )

            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 7,700")
            self.assertEqual(window.actual_value.text(), "Output feedback 7,699")
            self.assertEqual(window.error_value.text(), "Error +1")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -1")
            self.assertEqual(window.velocity_value.text(), "Velocity1 -4 SAV")
            self.assertEqual(window.velocity2_value.text(), "Velocity2 -3 SAV")
            self.assertEqual(window.torque_value.text(), "Torque -- (N/A)")
            self.assertEqual(len(window._traces["torque"]), 1)
            self.assertTrue(math.isnan(window._traces["torque"][0]))
            self.assertFalse(window._torque_seen)

            window._refresh_plots()
            window._clear_traces()
            self.assertEqual(window.torque_value.text(), "Torque -- (N/A)")
            self.assertFalse(window._traces["torque"])
            self.assertFalse(window._torque_seen)
        finally:
            self._close_window(window)

    def test_main_window_handles_partial_telemetry_samples(self) -> None:
        window = self._build_window()
        try:
            axis = window._selected_axis()
            sample = {
                "timestamp": 4.2,
                "axis_name": axis.name,
                "target_position": 9000,
                "actual_position": 8995,
            }
            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 9,000")
            self.assertEqual(window.actual_value.text(), "Output feedback 8,995")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -5")
            self.assertEqual(window.velocity_value.text(), "Velocity1 --")
            self.assertEqual(window.velocity2_value.text(), "Velocity2 --")
            self.assertEqual(window.torque_value.text(), "Torque -- (N/A)")
            self.assertFalse(window._torque_seen)
            self.assertTrue(math.isnan(window._traces["velocity_1"][0]))
            self.assertTrue(math.isnan(window._traces["velocity_2"][0]))
            self.assertTrue(math.isnan(window._traces["error"][0]))
            self.assertEqual(len(window._traces["error"]), 1)
        finally:
            self._close_window(window)

    def test_main_window_telemetry_alias_fields_for_command_output_are_supported(self) -> None:
        window = self._build_window()
        try:
            axis = window._selected_axis()
            sample = {
                "timestamp": 8.2,
                "axis_name": axis.name,
                "command_position": 11_000,
                "position_feedback": 10_970,
                "position_error": 30,
                "velocity_1": 64,
                "velocity_2": 61,
                "feedback_torque": 420,
            }

            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 11,000")
            self.assertEqual(window.actual_value.text(), "Output feedback 10,970")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -30")
            self.assertEqual(window.error_value.text(), "Error +30")
            self.assertIn("420", window.torque_value.text())
            self.assertIn("% FS", window.torque_value.text())
            self.assertTrue(window._torque_seen)
            self.assertTrue(window._traces["time"])
            self.assertTrue(window._traces["torque"])
        finally:
            self._close_window(window)

    def test_main_window_telemetry_axis_aliases_and_missing_timestamp_still_update(self) -> None:
        window = self._build_window()
        try:
            selected = window._selected_axis()
            sample = {
                "time": "45.25",
                "axis": selected.name.lower(),
                "command": 15_000,
                "feedback": 14_980,
                "position_error": -20,
                "velocity_1": 30,
                "velocity_2": 28,
                "motor_torque": 312,
            }

            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 15,000")
            self.assertEqual(window.actual_value.text(), "Output feedback 14,980")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -20")
            self.assertEqual(window.error_value.text(), "Error -20")
            self.assertIn("312", window.torque_value.text())
            self.assertIn("% FS", window.torque_value.text())
            self.assertTrue(window._traces["time"])
        finally:
            self._close_window(window)

    def test_main_window_telemetry_without_timestamp_still_updates_traces(self) -> None:
        window = self._build_window()
        try:
            selected = window._selected_axis()
            sample = {
                "motor_axis": selected.name,
                "command_position": 19_000,
                "position_feedback": 18_950,
                "position_error": 50,
            }

            window._on_telemetry(sample)

            self.assertEqual(window.target_value.text(), "Input command 19,000")
            self.assertEqual(window.actual_value.text(), "Output feedback 18,950")
            self.assertEqual(window.io_delta_value.text(), "Input-output delta -50")
            self.assertTrue(window._traces["time"])
            self.assertGreater(window._sample_tick, 0)
        finally:
            self._close_window(window)

    def test_main_window_fit_to_screen_keeps_telemetry_grid_compact(self) -> None:
        window = self._build_window()
        try:
            compact_screen = _FakeScreen(QtCore.QRect(0, 0, 1180, 860))
            narrow_screen = _FakeScreen(QtCore.QRect(0, 0, 520, 640))

            window._fit_to_screen(compact_screen)
            self.assertEqual(window._values_per_row, 2)
            self.assertEqual(window._splitter.orientation(), QtCore.Qt.Orientation.Vertical)
            self.assertLess(window.width(), compact_screen.availableGeometry().width())

            window._fit_to_screen(narrow_screen)
            self.assertEqual(window._values_per_row, 1)
            self.assertLess(window.width(), narrow_screen.availableGeometry().width())
        finally:
            self._close_window(window)

    def test_main_window_fit_to_screen_does_not_fill_monitor_width(self) -> None:
        window = self._build_window()
        try:
            large_screen = _FakeScreen(QtCore.QRect(0, 0, 2400, 1400))
            window._fit_to_screen(large_screen)
            self.assertLess(window.width(), 2400)
        finally:
            self._close_window(window)

    def test_plot_refresh_is_throttled_by_latest_sample(self) -> None:
        window = self._build_window()
        try:
            if not window._curves:
                self.skipTest("pyqtgraph is not installed; fallback plotting state is active")
            counter = self._PlotCounter()
            window._curves = {name: counter for name in window._curves}

            # Seed traces with a single sample, but keep the render tick behind.
            window._traces["time"].extend([1.0])
            window._traces["target"].append(1.0)
            window._traces["actual"].append(1.0)
            window._traces["error"].append(0.0)
            window._traces["velocity_1"].append(0.0)
            window._traces["velocity_2"].append(0.0)
            window._traces["io_delta"].append(0.0)
            window._traces["torque"].append(0.0)
            window._sample_tick = 1
            window._last_render_tick = 0

            window._refresh_plots()
            first_count = counter.calls
            self.assertGreater(first_count, 0)

            # Repeated refresh without a new telemetry sample must not
            # append further redraw work.
            window._refresh_plots()
            self.assertEqual(counter.calls, first_count)

            # A new sample should permit one additional redraw pass.
            window._on_telemetry(
                MotorTelemetry(
                    timestamp=2.0,
                    axis_name=window._selected_axis().name,
                    target_position=2,
                    actual_position=2,
                )
            )
            window._refresh_plots()
            self.assertGreater(counter.calls, first_count)
        finally:
            self._close_window(window)


if __name__ == "__main__":
    unittest.main(verbosity=2)
