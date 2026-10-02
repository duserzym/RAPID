from __future__ import annotations

import unittest

from PySide6 import QtGui, QtWidgets

from rapid_main.diagnostic_services import DCMotorNoCommBackend, UnavailableBackend
from rapid_main.dialogs.dc_motors import DCMotorDialog
from rapid_main.dialogs.debug_console import DebugConsoleDialog
from rapid_main.dialogs.step_monitor import StepMonitorDialog
from rapid_main.glass_theme import apply_main_glass_theme


class RemainingDialogGlassTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        apply_main_glass_theme(cls._app)

    def _close_motor_dialog(self, dialog: DCMotorDialog) -> None:
        dialog._telemetry_thread.stop_monitoring()
        dialog._telemetry_thread.wait(1000)
        dialog.close()
        dialog.deleteLater()
        self._app.processEvents()

    def test_debug_console_uses_shared_contract_and_escapes_backend_text(self) -> None:
        dialog = DebugConsoleDialog()
        try:
            self.assertEqual(dialog.objectName(), "glassDialog")
            self.assertTrue(dialog.accessibleName())
            self.assertFalse(dialog.styleSheet())
            self.assertTrue(dialog._console.accessibleName())
            self.assertTrue(dialog._refresh_btn.accessibleName())
            self.assertTrue(dialog._copy_btn.accessibleName())
            self.assertTrue(dialog._clear_btn.accessibleName())
            self.assertTrue(dialog._close_btn.accessibleName())

            dialog.append("ERROR", "backend returned <b>unsafe-looking text</b>")
            plain = dialog._console.toPlainText()
            self.assertIn("[ERROR", plain)
            self.assertIn("<b>unsafe-looking text</b>", plain)
        finally:
            dialog.deleteLater()

    def test_step_monitor_has_semantic_progress_and_qtgui_screen_contract(self) -> None:
        dialog = StepMonitorDialog()
        try:
            self.assertEqual(dialog.objectName(), "glassDialog")
            self.assertEqual(dialog._state_lbl.property("status"), "neutral")
            self.assertTrue(dialog._step_lbl.accessibleName())
            self.assertTrue(dialog._progress.accessibleName())
            self.assertTrue(dialog._log.accessibleName())

            dialog.update_step("Measure", "AF 20 mT", "Magnetometer", 12, 10)
            self.assertEqual(dialog._progress.maximum(), 10)
            self.assertEqual(dialog._progress.value(), 10)
            self.assertEqual(dialog._count_lbl.text(), "10 / 10")
            self.assertEqual(dialog._state_lbl.property("status"), "ready")
            self.assertTrue(dialog._state_lbl.text().startswith("READY —"))

            screen = self._app.primaryScreen()
            self.assertIsInstance(screen, QtGui.QScreen)
            dialog._fit_to_screen(screen)
        finally:
            dialog.deleteLater()

    def test_dc_motor_unavailable_backend_fails_closed_in_words(self) -> None:
        backend = UnavailableBackend("DC motors", "driver not installed")
        dialog = DCMotorDialog(backend=backend)  # type: ignore[arg-type]
        try:
            self.assertEqual(dialog.objectName(), "glassDialog")
            self.assertTrue(dialog.accessibleName())
            self.assertFalse(dialog.connect_btn.isEnabled())
            self.assertFalse(dialog.move_btn.isEnabled())
            self.assertEqual(dialog._status.property("status"), "unavailable")
            self.assertTrue(dialog._status.text().startswith("UNAVAILABLE —"))
            self.assertIn("driver not installed", dialog._status.text())
        finally:
            self._close_motor_dialog(dialog)

    def test_dc_motor_simulation_and_disconnect_states_are_explicit(self) -> None:
        backend = DCMotorNoCommBackend(port="COM1", baud=9600)
        backend.connect("COM1", 9600)
        dialog = DCMotorDialog(backend=backend)
        try:
            dialog._connected = True
            dialog._refresh_connection_state()
            self.assertEqual(dialog._status.property("status"), "simulated")
            self.assertTrue(dialog._status.text().startswith("SIMULATED —"))
            self.assertTrue(dialog.move_btn.isEnabled())
            self.assertIn("physical hardware", dialog.move_btn.accessibleDescription())
            self.assertTrue(dialog.target_value.accessibleName())

            dialog._disconnect()
            self.assertTrue(dialog.connect_btn.isEnabled())
            self.assertFalse(dialog.move_btn.isEnabled())
            self.assertEqual(dialog._status.property("status"), "simulated")
        finally:
            self._close_motor_dialog(dialog)

    def test_compact_remaining_dialog_actions_stay_visible(self) -> None:
        dialogs = [
            (DebugConsoleDialog(), (520, 360), ("_refresh_btn", "_close_btn")),
            (StepMonitorDialog(), (440, 360), ("_clear_btn", "_close_btn")),
        ]
        try:
            for dialog, size, action_names in dialogs:
                with self.subTest(dialog=type(dialog).__name__):
                    dialog.resize(*size)
                    dialog.show()
                    self._app.processEvents()
                    self.assertLessEqual(dialog.width(), size[0])
                    self.assertLessEqual(dialog.height(), size[1])
                    self.assertFalse(dialog.grab().isNull())
                    for name in action_names:
                        self.assertTrue(getattr(dialog, name).isVisibleTo(dialog))
        finally:
            for dialog, _, _ in dialogs:
                dialog.hide()
                dialog.deleteLater()
            self._app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
