"""Regressions found by an operator click-through of the packaged app."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from PySide6 import QtCore, QtWidgets

from rapid_main.device_ownership import DeviceOwnershipError, DeviceOwnershipManager, busy_message

_APP = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def _pump(times: int = 5) -> None:
    for _ in range(times):
        _APP.processEvents()


class BusyMessageTests(unittest.TestCase):
    def test_same_owner_reopening_is_explained(self):
        text = busy_message("changer", "dc_motors_panel", "dc_motors_panel")
        self.assertIn("DC Motors window is already open", text)
        self.assertNotIn("not 'dc_motors_panel'", text)

    def test_other_owner_names_the_window_to_close(self):
        registry = DeviceOwnershipManager()
        registry.acquire("changer", "dc_motors_panel")
        with self.assertRaises(DeviceOwnershipError) as caught:
            registry.acquire("changer", "irm_panel")
        self.assertIn("The sample changer is in use by the DC Motors window", str(caught.exception))


class DefaultTilingTests(unittest.TestCase):
    def test_pairs_only_when_both_panels_stay_legible(self):
        from rapid_main.app import _default_tiling

        narrow, narrow_focus = _default_tiling(1060)
        self.assertEqual([tree.get("leaf") for tree in narrow.values()],
                         ["dashboard", "measure", "queue", "sequence", "settings", "calibration"])
        wide, wide_focus = _default_tiling(1600)
        self.assertEqual(wide[1]["split"], "h")
        self.assertEqual(wide_focus[1], "dashboard")

    def test_tile_scrolls_instead_of_crushing_a_wide_panel(self):
        from rapidpy_common.tiling import TileScrollArea

        panel = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(panel)
        for _ in range(4):
            box = QtWidgets.QLabel("x")
            box.setMinimumWidth(200)
            layout.addWidget(box)
        area = TileScrollArea()
        area.setWidget(panel)
        area.resize(400, 300)
        area.show()
        _pump()
        self.assertGreaterEqual(panel.width(), 800)  # legible width kept, the tile scrolls
        area.resize(790, 300)
        _pump()
        self.assertEqual(area.horizontalScrollBarPolicy(), QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        area.deleteLater()


class ErrorLogTests(unittest.TestCase):
    def test_unhandled_exception_is_written_and_reported(self):
        from rapid_main import error_log

        with tempfile.TemporaryDirectory() as tmp:
            notes = []
            with mock.patch.object(error_log, "log_directory", lambda: Path(tmp)):
                error_log._STATE["notify"] = notes.append
                try:
                    raise ValueError("boom")
                except ValueError:
                    error_log.record_exception(*sys.exc_info())
                text = (Path(tmp) / "rapid_main_errors.log").read_text(encoding="utf-8")
            error_log._STATE["notify"] = None
        self.assertIn("ValueError: boom", text)
        self.assertIn("details saved to", notes[0])


class MainWindowFixTests(unittest.TestCase):
    def _window(self):
        from rapid_main.app import MainWindow

        window = MainWindow()
        window._stack.set_animation_duration(0)
        window.show()
        _pump()
        return window

    def _dispose(self, window):
        window._confirm_shutdown = lambda **kwargs: True
        window.close()
        window.deleteLater()
        _pump()

    def test_status_override_menu_toggles_without_error(self):
        window = self._window()
        try:
            window._status_override_action.trigger()
            self.assertTrue(window._status_override_active)
            window._status_override_action.trigger()
            self.assertFalse(window._status_override_active)
        finally:
            self._dispose(window)

    def test_reopening_an_open_tool_focuses_it_instead_of_device_busy(self):
        window = self._window()
        warnings = []
        try:
            with mock.patch.object(QtWidgets.QMessageBox, "warning", lambda *a, **k: warnings.append(a)):
                factory = lambda owner: QtWidgets.QDialog(owner)  # noqa: E731
                window._run_owned_dialog("vacuum", "vacuum_panel", factory, modal=False)
                first = window._owned_dialogs["vacuum"]
                window._stack.switch_workspace(4)
                window._run_owned_dialog("vacuum", "vacuum_panel", factory, modal=False)
            self.assertEqual(warnings, [])
            self.assertIs(window._owned_dialogs["vacuum"], first)
            self.assertEqual(window._stack.focused_key(), "tool:vacuum")
        finally:
            dialog = window._owned_dialogs.get("vacuum")
            if dialog is not None:
                dialog.close()
            _pump()
            self._dispose(window)

    def test_login_is_prefilled_with_last_session(self):
        from rapid_main.dialogs.login import LoginDialog

        dialog = LoginDialog()
        dialog.set_operator_name("YZ")
        dialog.set_operator_email("yz@umn.edu")
        dialog.set_nocomm(True)
        self.assertEqual((dialog.operator_name, dialog.operator_email, dialog.nocomm), ("YZ", "yz@umn.edu", True))
        dialog.deleteLater()


if __name__ == "__main__":
    unittest.main()
