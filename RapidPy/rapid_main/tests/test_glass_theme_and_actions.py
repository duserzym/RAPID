from __future__ import annotations

import unittest

from PySide6 import QtWidgets

from rapid_main.glass_theme import (
    GlassBackdrop,
    MAIN_GLASS_QSS,
    apply_main_glass_theme,
    install_glass_elevation,
)


class TestMainGlassTheme(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_main_theme_has_glass_accessibility_and_state_contracts(self) -> None:
        original = self._app.styleSheet()
        try:
            self._app.setStyleSheet("")
            apply_main_glass_theme(self._app)
            applied = self._app.styleSheet()
            self.assertIn("QWidget#appBackdrop", applied)
            self.assertIn("rgba(255, 255, 255, 214)", applied)
            self.assertIn("QPushButton:focus", applied)
            self.assertIn("QPushButton:disabled", applied)
            self.assertIn("QPushButton#navBtn:checked", applied)
            self.assertIn("QDialog#glassDialog", applied)
            self.assertIn('QLabel#statusText[status="error"]', applied)
            self.assertIn('QLabel#statusText[status="simulated"]', applied)
            self.assertIn(MAIN_GLASS_QSS, applied)
        finally:
            self._app.setStyleSheet(original)

    def test_backdrop_renders_and_card_elevation_is_idempotent(self) -> None:
        root = GlassBackdrop()
        root.resize(640, 420)
        card = QtWidgets.QFrame(root)
        card.setObjectName("card")
        card.setGeometry(40, 40, 220, 140)
        root.show()
        QtWidgets.QApplication.processEvents()

        self.assertFalse(root.grab().isNull())
        self.assertEqual(install_glass_elevation(root), 1)
        self.assertIsNotNone(card.graphicsEffect())
        self.assertEqual(install_glass_elevation(root), 0)
        root.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
