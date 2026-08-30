from __future__ import annotations

import unittest

from PySide6 import QtWidgets

from rapid_main.app import MainWindow, _QSETTINGS_SHOW_STARTUP_GUIDE
from rapid_main.dialogs import StartupGuideDialog


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


class TestStartupGuide(_TestAppMixin, unittest.TestCase):
    def test_dialog_emits_preference_and_navigation_requests(self) -> None:
        dialog = StartupGuideDialog(show_at_startup=True)
        preferences: list[bool] = []
        routes: list[str] = []
        dialog.show_at_startup_changed.connect(preferences.append)
        dialog.open_settings_requested.connect(lambda: routes.append("settings"))
        dialog.open_queue_requested.connect(lambda: routes.append("queue"))

        dialog._show_at_startup.setChecked(False)
        dialog._settings_button.click()
        dialog._queue_button.click()

        self.assertEqual(preferences, [False])
        self.assertEqual(routes, ["settings", "queue"])
        dialog.deleteLater()

    def test_main_window_guide_routes_and_persists_its_preference(self) -> None:
        window = MainWindow()
        settings = window._settings
        had_value = settings.contains(_QSETTINGS_SHOW_STARTUP_GUIDE)
        previous = settings.value(_QSETTINGS_SHOW_STARTUP_GUIDE)
        try:
            window._set_show_startup_guide(False)
            self.assertFalse(window._show_startup_guide_enabled())

            window._launch_startup_guide()
            dialog = window._startup_guide_dialog
            self.assertIsNotNone(dialog)
            self.assertFalse(dialog._show_at_startup.isChecked())

            dialog.open_settings_requested.emit()
            self.assertEqual(window._stack.currentIndex(), 4)
            dialog.open_queue_requested.emit()
            self.assertEqual(window._stack.currentIndex(), 1)
        finally:
            if had_value:
                settings.setValue(_QSETTINGS_SHOW_STARTUP_GUIDE, previous)
            else:
                settings.remove(_QSETTINGS_SHOW_STARTUP_GUIDE)
            if window._startup_guide_dialog is not None:
                window._startup_guide_dialog.close()
            window.deleteLater()