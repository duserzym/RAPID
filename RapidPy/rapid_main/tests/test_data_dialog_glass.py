from __future__ import annotations

import unittest
from unittest import mock

from PySide6 import QtGui, QtWidgets

from rapid_main.dialogs.plots import PlotsDialog
from rapid_main.dialogs.sample_select import SampleSelectDialog
from rapid_main.glass_theme import apply_main_glass_theme


class _FakeWebcamWidget(QtWidgets.QWidget):
    pass


class DataDialogGlassTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        apply_main_glass_theme(cls._app)

    def tearDown(self) -> None:
        for widget in list(self._app.topLevelWidgets()):
            if widget.objectName() == "glassDialog":
                widget.hide()
                widget.deleteLater()
        self._app.processEvents()

    def test_plot_and_sample_dialogs_use_shared_contract(self) -> None:
        plots = PlotsDialog()
        samples = SampleSelectDialog()

        for dialog in (plots, samples):
            with self.subTest(dialog=type(dialog).__name__):
                self.assertEqual(dialog.objectName(), "glassDialog")
                self.assertTrue(dialog.accessibleName())
                self.assertFalse(dialog.styleSheet())

        self.assertTrue(plots._tabs.accessibleName())
        self.assertTrue(plots._demo_btn.accessibleName())
        self.assertTrue(plots._export_btn.accessibleName())
        self.assertFalse(plots._export_btn.isEnabled())
        self.assertTrue(plots._close_btn.accessibleName())
        self.assertEqual(plots._demo_lbl.property("status"), "neutral")
        self.assertTrue(samples._search.accessibleName())
        self.assertTrue(samples._load_btn.accessibleName())
        self.assertTrue(samples._table.accessibleName())
        self.assertTrue(samples._select_btn.accessibleName())
        self.assertTrue(samples._cancel_btn.accessibleName())

    def test_plot_status_distinguishes_real_and_simulated_data_in_words(self) -> None:
        dialog = PlotsDialog()
        dialog._load_demo()
        self.assertEqual(dialog._demo_lbl.property("status"), "simulated")
        self.assertTrue(dialog._demo_lbl.text().startswith("SIMULATED —"))

        dialog.set_data([1.0], [0.0], [0.0], ["NRM"])
        self.assertEqual(dialog._demo_lbl.property("status"), "ready")
        self.assertTrue(dialog._demo_lbl.text().startswith("READY —"))

    def test_webcam_dialog_uses_glass_and_qtgui_screen_contract(self) -> None:
        import rapid_main.dialogs.webcam_dialog as webcam_module

        with (
            mock.patch.object(webcam_module, "_WEBCAM_AVAILABLE", True),
            mock.patch.object(
                webcam_module,
                "WebcamWidget",
                _FakeWebcamWidget,
                create=True,
            ),
        ):
            dialog = webcam_module.WebcamDialog()

        self.assertEqual(dialog.objectName(), "glassDialog")
        self.assertTrue(dialog.accessibleName())
        self.assertTrue(dialog._webcam.accessibleName())
        screen = self._app.primaryScreen()
        self.assertIsInstance(screen, QtGui.QScreen)
        dialog._fit_to_screen(screen)

    def test_compact_plot_and_sample_actions_remain_visible(self) -> None:
        cases = [
            (PlotsDialog(), (480, 420), ("_demo_btn", "_export_btn", "_close_btn")),
            (SampleSelectDialog(), (460, 360), ("_load_btn", "_select_btn", "_cancel_btn")),
        ]

        for dialog, size, action_names in cases:
            with self.subTest(dialog=type(dialog).__name__):
                dialog.resize(*size)
                dialog.show()
                self._app.processEvents()
                self.assertLessEqual(dialog.width(), size[0])
                self.assertLessEqual(dialog.height(), size[1])
                self.assertFalse(dialog.grab().isNull())
                for name in action_names:
                    action = getattr(dialog, name)
                    self.assertIsNotNone(action)
                    self.assertTrue(action.isVisibleTo(dialog))


if __name__ == "__main__":
    unittest.main(verbosity=2)
