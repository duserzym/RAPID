"""The RAPID icon reaches the taskbar/title bar (source and frozen) and the UI."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from PySide6 import QtWidgets

from rapid_main.branding import APP_USER_MODEL_ID, app_icon, brand_pixmap, splash_screen
from rapid_main.startup import main_assets_dir
from rapidpy_common.ui import set_app_icon

_APP = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


class BrandingTests(unittest.TestCase):
    def test_icon_has_taskbar_and_title_bar_sizes(self):
        sizes = {size.width() for size in app_icon().availableSizes()}
        self.assertTrue({16, 24, 32, 48, 256} <= sizes, sizes)
        self.assertFalse(brand_pixmap(30).isNull())
        self.assertIsNotNone(splash_screen())

    def test_frozen_build_finds_packaged_asset_folder(self):
        # Regression: frozen runs only looked in _MEIPASS/assets, where the
        # rapid_main_assets icon does not live, so no window icon was set.
        widget = QtWidgets.QWidget()
        with tempfile.TemporaryDirectory() as empty_meipass, \
                mock.patch.object(sys, "frozen", True, create=True), \
                mock.patch.object(sys, "_MEIPASS", empty_meipass, create=True):
            self.assertTrue(set_app_icon(widget, "rapid_main_icon.ico", main_assets_dir()))
        self.assertFalse(widget.windowIcon().isNull())
        widget.deleteLater()

    def test_missing_icon_reports_false(self):
        widget = QtWidgets.QWidget()
        self.assertFalse(set_app_icon(widget, "nope.ico", Path(tempfile.gettempdir()) / "no-such-assets"))
        widget.deleteLater()

    def test_main_window_shows_the_icon(self):
        from rapid_main.app import MainWindow

        window = MainWindow()
        try:
            self.assertFalse(window.windowIcon().isNull())
            self.assertFalse(window._header_icon.pixmap().isNull())
            self.assertEqual(APP_USER_MODEL_ID, "UMN.IRM.RAPID.Main")
        finally:
            window._confirm_shutdown = lambda **kwargs: True
            window.close()
            window.deleteLater()


if __name__ == "__main__":
    unittest.main()
