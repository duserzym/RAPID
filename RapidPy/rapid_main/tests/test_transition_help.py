from __future__ import annotations

import unittest
from unittest.mock import Mock

from PySide6 import QtWidgets

from rapid_main.app import MainWindow
from rapid_main.dialogs.transition_help import TRANSITION_ENTRIES, TransitionHelpDialog


class TransitionHelpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_map_covers_common_measurement_and_hardware_forms(self) -> None:
        legacy = " ".join(entry.vb6 for entry in TRANSITION_ENTRIES)
        for form in ("frmMeasure", "frmChanger", "frmSquid", "frmVacuum", "frmIRMARM"):
            self.assertIn(form, legacy)
        self.assertTrue(all(entry.status for entry in TRANSITION_ENTRIES))

    def test_search_filters_across_form_task_destination_and_status(self) -> None:
        dialog = TransitionHelpDialog()

        dialog._search.setText("vacuum")

        visible = [
            row
            for row in range(dialog._table.rowCount())
            if not dialog._table.isRowHidden(row)
        ]
        self.assertEqual(len(visible), 1)
        self.assertIn("frmVacuum", dialog._table.item(visible[0], 0).text())
        dialog.deleteLater()

    def test_open_selected_emits_routable_destination(self) -> None:
        dialog = TransitionHelpDialog()
        destinations: list[str] = []
        dialog.destination_requested.connect(destinations.append)
        dialog._search.setText("frmMeasure")

        dialog._open_selected()

        self.assertEqual(destinations, ["measure"])
        dialog.deleteLater()

    def test_no_search_results_disable_open_destination(self) -> None:
        dialog = TransitionHelpDialog()

        dialog._search.setText("definitely-not-a-legacy-control")

        self.assertFalse(dialog._open_button.isEnabled())
        self.assertFalse(dialog._table.selectionModel().selectedRows())
        dialog._search.setText("vacuum")
        self.assertTrue(dialog._open_button.isEnabled())
        dialog.deleteLater()

    def test_main_window_routes_panel_and_diagnostic_destinations(self) -> None:
        controller = type("Controller", (), {})()
        controller.navigate_to = Mock()
        controller._launch_vacuum = Mock()
        for name in (
            "_launch_dc_motors", "_launch_squid", "_launch_irm", "_launch_af",
            "_launch_debug_console", "_launch_step_monitor", "_launch_vrm",
            "_launch_webcam", "_launch_login", "_launch_startup_guide",
        ):
            setattr(controller, name, Mock())

        MainWindow._open_transition_destination(controller, "queue")
        MainWindow._open_transition_destination(controller, "vacuum")

        controller.navigate_to.assert_called_once_with("queue")
        controller._launch_vacuum.assert_called_once_with()


if __name__ == "__main__":
    unittest.main(verbosity=2)
