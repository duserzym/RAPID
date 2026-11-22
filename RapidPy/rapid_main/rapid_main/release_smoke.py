"""Offline release check with temporary configuration and Qt preferences."""
from __future__ import annotations

from importlib import import_module
import json
import os
from pathlib import Path
import tempfile


def _check_transverse_probe(window, app):
    from PySide6 import QtCore, QtGui

    actions = [action for action in window.menuBar().findChildren(QtGui.QAction)
               if action.text() == "Transverse Probe Scan"]
    if len(actions) != 1:
        return False
    actions[0].trigger()
    dialog = window._owned_dialogs.get("transverse_probe")
    if dialog is None:
        return False

    def wait_until(predicate):
        timer = QtCore.QElapsedTimer()
        timer.start()
        while not predicate() and timer.elapsed() < 5000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        return predicate()

    try:
        dialog.start_button.click()
        dialog.pause_button.click()
        if not wait_until(lambda: dialog.message.text() == "Paused between readings."):
            return False
        dialog.pause_button.click()
        if not wait_until(lambda: not dialog.active) or not dialog.apply_button.isEnabled() or dialog.table.rowCount() < 3:
            return False
        dialog.apply_button.click()
        return wait_until(lambda: not dialog.active) and dialog.message.text().startswith("Simulated angle")
    finally:
        dialog.stop()
        wait_until(lambda: not dialog.active)
        dialog.close()
        app.processEvents()


def run_smoke_test() -> int:
    from PySide6 import QtCore, QtWidgets
    from rapid_main.config import AppConfig
    from rapid_main.package_launch import HELPER_MODULES

    with tempfile.TemporaryDirectory(prefix="rapidpy-smoke-") as directory:
        root = Path(directory)
        previous_config = os.environ.get("RAPID_CONFIG")
        os.environ["RAPID_CONFIG"] = str(root / "config.json")
        config = AppConfig()
        config.general.nocomm = True
        config.general.data_dir = str(root / "data")
        config.save()
        previous_format = QtCore.QSettings.defaultFormat()
        QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
        QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.UserScope, str(root))
        original_settings = QtCore.QSettings
        class IsolatedSettings(original_settings):
            def __init__(self, *args, **kwargs):
                if len(args) == 2 and all(isinstance(arg, str) for arg in args):
                    super().__init__(original_settings.IniFormat, original_settings.UserScope, *args, **kwargs)
                else:
                    super().__init__(*args, **kwargs)
        QtCore.QSettings = IsolatedSettings
        window = None
        try:
            # Constructing this window uses only explicit No-Communication backends.
            from rapid_main.app import MainWindow
            app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
            app.setQuitOnLastWindowClosed(False)
            window = MainWindow()
            window.show()
            app.processEvents()
            helpers = {}
            for name in HELPER_MODULES:
                helpers[name] = callable(import_module(f"{name}.app").main)
            transverse_probe = _check_transverse_probe(window, app)
            report = {
                "schema": "rapidpy.release_smoke.v1", "simulated": True,
                "ok": window.config.general.nocomm and all(helpers.values()) and transverse_probe,
                "transverse_probe": transverse_probe,
                "panels": window._stack.count(), "helpers": helpers,
                "width": window.width(), "height": window.height(),
            }
            print(json.dumps(report, indent=2, sort_keys=True))
            return 0 if report["ok"] else 2
        finally:
            if window is not None:
                window._confirm_shutdown(prompt=False)
                window.hide()
                window.deleteLater()
                app.processEvents()
            QtCore.QSettings.setDefaultFormat(previous_format)
            QtCore.QSettings = original_settings
            if previous_config is None:
                os.environ.pop("RAPID_CONFIG", None)
            else:
                os.environ["RAPID_CONFIG"] = previous_config
