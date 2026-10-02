from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtWidgets

from rapid_main.panels.measurement import MeasurementPanel, MeasurementTraceWidget
from rapid_main.panels.sample_queue import SampleQueuePanel
from rapid_main.panels.settings_panel import SettingsPanel
from rapid_main.printing import print_widget_snapshot


class _CancelledPrintDialog(QtWidgets.QDialog):
    def exec(self) -> int:
        return int(QtWidgets.QDialog.DialogCode.Rejected)


class TestCompletedUiActions(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_print_helper_returns_false_when_operator_cancels(self) -> None:
        widget = QtWidgets.QLabel("report")
        printed = print_widget_snapshot(
            widget,
            widget,
            title="test",
            dialog_factory=lambda _printer, parent: _CancelledPrintDialog(parent),
        )
        self.assertFalse(printed)
        widget.deleteLater()

    def test_measurement_print_button_invokes_real_print_path(self) -> None:
        panel = MeasurementPanel()
        with patch(
            "rapid_main.panels.measurement.print_widget_snapshot",
            return_value=True,
        ) as print_snapshot:
            panel._print_btn.click()
        print_snapshot.assert_called_once()
        self.assertIs(print_snapshot.call_args.args[0], panel)
        self.assertIs(print_snapshot.call_args.args[1], panel)
        panel.deleteLater()

    def test_queue_print_button_invokes_real_print_path(self) -> None:
        panel = SampleQueuePanel()
        with patch(
            "rapid_main.panels.sample_queue.print_widget_snapshot",
            return_value=True,
        ) as print_snapshot:
            panel._print_btn.click()
        print_snapshot.assert_called_once()
        self.assertIs(print_snapshot.call_args.args[1], panel._table)
        panel.deleteLater()

    def test_dependency_free_measurement_plot_renders_data(self) -> None:
        plot = MeasurementTraceWidget()
        plot.resize(560, 360)
        plot.set_series(
            [1.0, 2.0, 3.0],
            [1.0e-7, 8.0e-8, 5.0e-8],
            [15.0, 28.0, 42.0],
            [50.0, 35.0, 22.0],
        )
        plot.show()
        QtWidgets.QApplication.processEvents()

        self.assertEqual(plot._series["step"], [1.0, 2.0, 3.0])
        self.assertFalse(plot.grab().isNull())
        plot.clear()
        self.assertEqual(plot._series["moment"], [])
        plot.deleteLater()

    def test_settings_browse_button_updates_its_path_field(self) -> None:
        panel = SettingsPanel()
        button = panel._path_browse_buttons["Data folder:"]
        with patch(
            "rapid_main.panels.settings_panel.QtWidgets.QFileDialog.getExistingDirectory",
            return_value="C:/RAPID/data",
        ):
            button.click()
        self.assertEqual(panel._data_dir.text(), str(Path("C:/RAPID/data")))
        panel.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
