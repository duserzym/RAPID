from __future__ import annotations

import unittest
from unittest import mock

from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.diagnostic_services import DiagnosticStatusLine
from rapid_main.panels.dashboard import DashboardPanel


class _QtTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


class DashboardDiagnosticTests(_QtTestCase):
    def test_dashboard_renders_live_simulated_unavailable_and_fault_states(self) -> None:
        panel = DashboardPanel()
        stamp = QtCore.QDateTime.fromString("2026-10-02T10:30:00", QtCore.Qt.DateFormat.ISODate)
        panel.update_diagnostics(
            [
                DiagnosticStatusLine("SQUID", True, False, "COM4 ready"),
                DiagnosticStatusLine("Vacuum", False, True, "No-comm simulation"),
                DiagnosticStatusLine("DC Motors", False, False, "DC motors unavailable: driver missing"),
                DiagnosticStatusLine(
                    "AF Demag", False, False, "status unavailable", True, "connection read failed"
                ),
                DiagnosticStatusLine("IRM/ARM", False, False, "Disconnected"),
            ],
            refreshed_at=stamp,
        )

        self.assertEqual(panel._instrument_status_labels["SQUID"].objectName(), "instOk")
        self.assertIn("Connected", panel._instrument_status_labels["SQUID"].text())
        self.assertEqual(panel._instrument_status_labels["Vacuum"].objectName(), "instSim")
        self.assertIn("SIMULATED", panel._instrument_status_labels["Vacuum"].text())
        self.assertEqual(panel._instrument_status_labels["DC Motors"].objectName(), "instErr")
        self.assertIn("Unavailable", panel._instrument_status_labels["DC Motors"].text())
        self.assertEqual(panel._instrument_status_labels["AF Demag"].objectName(), "instErr")
        self.assertIn("connection read failed", panel._instrument_status_labels["AF Demag"].toolTip())
        self.assertIn("10:30:00", panel._diagnostic_refresh_label.text())

    def test_dashboard_refresh_button_emits_request(self) -> None:
        panel = DashboardPanel()
        emitted: list[bool] = []
        panel.refresh_diagnostics_requested.connect(lambda: emitted.append(True))
        button = next(
            button
            for button in panel.findChildren(QtWidgets.QPushButton)
            if "Refresh" in button.text()
        )
        button.click()
        self.assertEqual(emitted, [True])

    def test_dashboard_reflows_cards_between_compact_and_wide_widths(self) -> None:
        panel = DashboardPanel()
        panel.resize(520, 900)
        panel.show()
        QtWidgets.QApplication.processEvents()
        third = panel._instrument_cards[2]
        compact_position = panel._instrument_grid.getItemPosition(
            panel._instrument_grid.indexOf(third)
        )
        self.assertTrue(panel._dashboard_compact)
        self.assertEqual(compact_position[:2], (1, 0))

        panel.resize(1200, 900)
        QtWidgets.QApplication.processEvents()
        fifth = panel._instrument_cards[4]
        wide_position = panel._instrument_grid.getItemPosition(
            panel._instrument_grid.indexOf(fifth)
        )
        self.assertFalse(panel._dashboard_compact)
        self.assertEqual(wide_position[:2], (0, 4))
        self.assertEqual(
            panel._mid_grid.getItemPosition(panel._mid_grid.indexOf(panel._actions_card))[:2],
            (0, 1),
        )
        panel.deleteLater()


class MainShellStateTests(_QtTestCase):
    def test_dashboard_current_run_tracks_main_window_state(self) -> None:
        window = MainWindow()
        window.set_sample("HBK-42")
        window.set_step("AF50")
        window.set_flow_state("measuring")

        self.assertEqual(window._dashboard._run_flow.text(), "Measuring")
        self.assertEqual(window._dashboard._run_sample.text(), "HBK-42")
        self.assertEqual(window._dashboard._run_step.text(), "AF50")
        self.assertEqual(window._dashboard._run_treat.text(), "AF50")
        window.deleteLater()

    def test_no_comm_mode_does_not_claim_workflow_is_running(self) -> None:
        window = MainWindow()
        window.set_flow_state("idle")
        window._on_nocomm_toggled(True)

        self.assertEqual(window._workflow_state, "idle")
        self.assertIn("Idle", window._flow_lbl.text())
        self.assertTrue(window._nocomm_btn.isChecked() or window.config.general.nocomm)
        window.deleteLater()

    def test_declining_shutdown_never_halts_active_work(self) -> None:
        window = MainWindow()
        window._sequence.confirm_discard_changes = mock.Mock(return_value=True)
        window._has_active_automation = mock.Mock(return_value=True)
        window.halt_measurement = mock.Mock()
        with mock.patch.object(
            QtWidgets.QMessageBox,
            "question",
            return_value=QtWidgets.QMessageBox.StandardButton.No,
        ):
            accepted = window._confirm_shutdown(prompt=True, on_close=True)

        self.assertFalse(accepted)
        window.halt_measurement.assert_not_called()
        window.deleteLater()

    def test_status_override_is_visual_only_and_preserves_workflow_state(self) -> None:
        window = MainWindow()
        window.set_flow_state("paused")
        window._toggle_status_override(True)

        self.assertEqual(window._workflow_state, "paused")
        self.assertEqual(window._dashboard._run_flow.text(), "Paused")
        self.assertEqual(window._flow_lbl.objectName(), "flowOverride")
        window._toggle_status_override(False)
        self.assertEqual(window._flow_lbl.objectName(), "flowPaused")
        window.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
