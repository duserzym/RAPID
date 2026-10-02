from __future__ import annotations

import unittest

from PySide6 import QtCore, QtWidgets

from rapid_main.config import SquidConfig
from rapid_main.diagnostic_services import (
    IrmArmNoCommBackend,
    SquidNoCommBackend,
    UnavailableBackend,
    VacuumNoCommBackend,
)
from rapid_main.dialogs.irm_arm import IrmArmDialog
from rapid_main.dialogs.squid_comm import SquidCommDialog
from rapid_main.dialogs.vacuum import VacuumDialog
from rapid_main.glass_theme import apply_main_glass_theme, set_semantic_status


class HardwareDialogGlassTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        apply_main_glass_theme(cls._app)

    def tearDown(self) -> None:
        for widget in list(self._app.topLevelWidgets()):
            if isinstance(widget, (VacuumDialog, IrmArmDialog, SquidCommDialog)):
                widget.close()
                widget.deleteLater()
        self._app.processEvents()

    def test_semantic_status_has_text_property_and_accessibility(self) -> None:
        label = QtWidgets.QLabel()
        set_semantic_status(
            label,
            "Pump command failed",
            "error",
            accessible_name="Vacuum system status",
        )

        self.assertEqual(label.property("status"), "error")
        self.assertEqual(label.text(), "ERROR — Pump command failed")
        self.assertEqual(label.accessibleName(), "Vacuum system status")
        self.assertIn("Error state", label.accessibleDescription())
        with self.assertRaises(ValueError):
            set_semantic_status(label, "bad", "mystery")

    def test_hardware_dialogs_use_shared_theme_and_non_color_status(self) -> None:
        dialogs = [
            VacuumDialog(backend=VacuumNoCommBackend()),
            IrmArmDialog(backend=IrmArmNoCommBackend()),
            SquidCommDialog(backend=SquidNoCommBackend()),
        ]

        for dialog in dialogs:
            with self.subTest(dialog=type(dialog).__name__):
                self.assertEqual(dialog.objectName(), "glassDialog")
                self.assertTrue(dialog.accessibleName())
                self.assertFalse(dialog.styleSheet())
                local_styles = [
                    widget
                    for widget in dialog.findChildren(QtWidgets.QWidget)
                    if widget.styleSheet()
                ]
                self.assertEqual(local_styles, [])
                status = dialog._status_lbl  # type: ignore[attr-defined]
                self.assertEqual(status.property("status"), "simulated")
                self.assertTrue(status.text().startswith("SIMULATED —"))
                self.assertTrue(status.accessibleName())

    def test_dialog_actions_have_accessible_names(self) -> None:
        vacuum = VacuumDialog(backend=VacuumNoCommBackend())
        irm = IrmArmDialog(backend=IrmArmNoCommBackend())
        squid = SquidCommDialog(backend=SquidNoCommBackend())

        controls = [
            vacuum._pump_btn,
            vacuum._close_btn,
            vacuum._target_spin,
            vacuum._warn_spin,
            irm._mode,
            irm._apply_btn,
            irm._reset_btn,
            irm._close_btn,
            squid._port,
            squid._baud,
            squid._range,
            squid._samples,
            squid._settle,
            squid._test_btn,
            squid._save_btn,
            squid._cancel_btn,
        ]
        for control in controls:
            with self.subTest(control=control):
                self.assertTrue(control.accessibleName())

    def test_unavailable_vacuum_dialog_opens_fail_closed(self) -> None:
        dialog = VacuumDialog(
            backend=UnavailableBackend("Vacuum", "COM4 did not answer")  # type: ignore[arg-type]
        )

        self.assertEqual(dialog._status_lbl.property("status"), "unavailable")
        self.assertIn("UNAVAILABLE —", dialog._status_lbl.text())
        self.assertIn("COM4 did not answer", dialog._status_lbl.text())
        self.assertFalse(dialog._pump_btn.isEnabled())

    def test_unavailable_field_dialogs_report_unavailable_not_ready(self) -> None:
        irm = IrmArmDialog(
            backend=UnavailableBackend("IRM/ARM", "ADWIN driver missing")  # type: ignore[arg-type]
        )
        squid = SquidCommDialog(
            backend=UnavailableBackend("SQUID", "COM7 did not answer")  # type: ignore[arg-type]
        )

        self.assertEqual(irm._status_lbl.property("status"), "unavailable")
        self.assertTrue(irm._status_lbl.text().startswith("UNAVAILABLE —"))
        self.assertFalse(irm._apply_btn.isEnabled())
        self.assertFalse(irm._reset_btn.isEnabled())
        self.assertEqual(squid._status_lbl.property("status"), "unavailable")
        self.assertTrue(squid._status_lbl.text().startswith("UNAVAILABLE —"))

    def test_squid_settings_load_and_apply_to_shared_config(self) -> None:
        cfg = SquidConfig(
            port="COM7",
            baud=19200,
            range_label="100×",
            samples_per_pos=12,
            settle_time=2.5,
        )
        dialog = SquidCommDialog(backend=SquidNoCommBackend(cfg))
        self.assertEqual(dialog._port.currentText(), "COM7")
        self.assertEqual(dialog._baud.currentText(), "19200")
        self.assertEqual(dialog._range.currentText(), "100×")
        self.assertEqual(dialog._samples.value(), 12)
        self.assertAlmostEqual(dialog._settle.value(), 2.5)

        dialog._port.setCurrentText("COM3")
        dialog._baud.setCurrentText("4800")
        dialog._range.setCurrentText("10×")
        dialog._samples.setValue(6)
        dialog._settle.setValue(1.2)
        dialog._accept_settings()

        self.assertEqual(cfg.port, "COM3")
        self.assertEqual(cfg.baud, 4800)
        self.assertEqual(cfg.range_label, "10×")
        self.assertEqual(cfg.samples_per_pos, 6)
        self.assertAlmostEqual(cfg.settle_time, 1.2)

    def test_dialogs_fit_compact_width_without_horizontal_clipping(self) -> None:
        dialogs = [
            VacuumDialog(backend=VacuumNoCommBackend()),
            IrmArmDialog(backend=IrmArmNoCommBackend()),
            SquidCommDialog(backend=SquidNoCommBackend()),
        ]

        for dialog in dialogs:
            with self.subTest(dialog=type(dialog).__name__):
                dialog.resize(360, 520)
                dialog.show()
                self._app.processEvents()
                self.assertLessEqual(dialog.width(), 420)
                self.assertLessEqual(dialog.height(), 520)
                self.assertFalse(dialog.grab().isNull())
                for child in dialog.findChildren(QtWidgets.QWidget):
                    if not child.isVisible() or child.isWindow():
                        continue
                    left = child.mapTo(dialog, QtCore.QPoint(0, 0)).x()
                    right = child.mapTo(dialog, QtCore.QPoint(child.width(), 0)).x()
                    self.assertGreaterEqual(
                        left,
                        -1,
                        f"{type(dialog).__name__}: {type(child).__name__} clips left",
                    )
                    self.assertLessEqual(
                        right,
                        dialog.width() + 1,
                        f"{type(dialog).__name__}: {type(child).__name__} clips right",
                    )

                if isinstance(dialog, VacuumDialog):
                    actions = [dialog._pump_btn, dialog._close_btn]
                elif isinstance(dialog, IrmArmDialog):
                    actions = [dialog._apply_btn, dialog._reset_btn, dialog._close_btn]
                else:
                    actions = [dialog._test_btn, dialog._save_btn, dialog._cancel_btn]
                    self.assertGreaterEqual(
                        dialog._settings_scroll.verticalScrollBar().maximum(), 0
                    )
                for action in actions:
                    self.assertIsNotNone(action)
                    assert action is not None
                    top = action.mapTo(dialog, QtCore.QPoint(0, 0)).y()
                    bottom = action.mapTo(dialog, QtCore.QPoint(0, action.height())).y()
                    self.assertTrue(action.isVisibleTo(dialog))
                    self.assertGreaterEqual(top, -1)
                    self.assertLessEqual(bottom, dialog.height() + 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
