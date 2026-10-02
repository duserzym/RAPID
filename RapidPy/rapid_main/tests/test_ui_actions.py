from __future__ import annotations

import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PySide6 import QtWidgets

from rapid_main.panels.measurement import MeasurementPanel, MeasurementTraceWidget
from rapid_main.panels.dashboard import DashboardPanel
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

    def test_sample_index_selection_adds_real_queue_row_and_metadata(self) -> None:
        from rapid_main.app import MainWindow

        registrations = object()
        dialog = SimpleNamespace(
            exec=Mock(return_value=QtWidgets.QDialog.DialogCode.Accepted),
            selected_record={
                "sample_name": "SPEC-42",
                "depth_cm": "18.2",
                "formation": "Basalt Unit",
                "location": "Site Z",
            },
            registrations=registrations,
            source_path=Path("C:/samples/site-z.sam"),
        )

        class _Table:
            @staticmethod
            def rowCount() -> int:
                return 0

        queue = SimpleNamespace(_table=_Table(), add_sample=Mock())
        measurement = SimpleNamespace(set_specimen_context=Mock())
        controller = SimpleNamespace(
            _sample_queue=queue,
            _measurement=measurement,
            _sequence_labels=["NRM", "AF20"],
            sample_registrations=None,
            _nav_select=Mock(),
            _save_queue_state=Mock(),
            set_status=Mock(),
        )

        with (
            patch("rapid_main.app.SampleSelectDialog", return_value=dialog),
            patch(
                "rapid_main.app.QtWidgets.QInputDialog.getText",
                return_value=("B12", True),
            ),
        ):
            loaded = MainWindow._load_sample_from_index(controller)

        self.assertTrue(loaded)
        queue.add_sample.assert_called_once_with(
            position="B12",
            name="SPEC-42",
            sample_set="Basalt Unit",
            treatment="NRM → AF20",
        )
        measurement.set_specimen_context.assert_called_once_with(
            "SPEC-42",
            depth="18.2",
            treatment="NRM → AF20",
        )
        self.assertIs(controller.sample_registrations, registrations)
        controller._nav_select.assert_called_once_with(1)
        controller._save_queue_state.assert_called_once_with()

    def test_dashboard_load_sample_action_requests_index_workflow(self) -> None:
        dashboard = DashboardPanel()
        emitted: list[bool] = []
        dashboard.load_sample_requested.connect(lambda: emitted.append(True))
        load_button = next(
            button
            for button in dashboard.findChildren(QtWidgets.QPushButton)
            if "Load Sample" in button.text()
        )

        load_button.click()

        self.assertEqual(emitted, [True])
        dashboard.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
