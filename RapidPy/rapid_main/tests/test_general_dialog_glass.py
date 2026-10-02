from __future__ import annotations

import unittest
from unittest import mock

from PySide6 import QtCore, QtWidgets

from rapid_main.dialogs.about import AboutDialog
from rapid_main.dialogs.login import LoginDialog
from rapid_main.dialogs.startup_guide import StartupGuideDialog
from rapid_main.dialogs.transition_help import TransitionHelpDialog
from rapid_main.glass_theme import apply_main_glass_theme


class GeneralDialogGlassTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        apply_main_glass_theme(cls._app)

    def tearDown(self) -> None:
        for widget in list(self._app.topLevelWidgets()):
            if isinstance(
                widget,
                (LoginDialog, AboutDialog, StartupGuideDialog, TransitionHelpDialog),
            ):
                widget.close()
                widget.deleteLater()
        self._app.processEvents()

    def test_dialogs_use_shared_glass_without_local_styles(self) -> None:
        dialogs = [
            LoginDialog(),
            AboutDialog(),
            StartupGuideDialog(),
            TransitionHelpDialog(),
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

    def test_primary_controls_have_accessible_names(self) -> None:
        login = LoginDialog()
        about = AboutDialog()
        guide = StartupGuideDialog()
        transition = TransitionHelpDialog()
        controls = [
            login._name_edit,
            login._nocomm_chk,
            login._ok_button,
            login._cancel_button,
            about._github_btn,
            about._ok_btn,
            guide._show_at_startup,
            guide._settings_button,
            guide._queue_button,
            guide._close_button,
            transition._search,
            transition._table,
            transition._open_button,
            transition._close_button,
        ]
        for control in controls:
            self.assertIsNotNone(control)
            assert control is not None
            with self.subTest(control=type(control).__name__):
                self.assertTrue(control.accessibleName())

    def test_login_requires_identity_and_accepts_named_operator(self) -> None:
        dialog = LoginDialog()
        with mock.patch.object(QtWidgets.QMessageBox, "warning") as warning:
            dialog._on_accept()
        warning.assert_called_once()
        self.assertNotEqual(dialog.result(), QtWidgets.QDialog.DialogCode.Accepted)

        dialog._name_edit.setCurrentText("AB")
        dialog._on_accept()
        self.assertEqual(dialog.operator_name, "AB")
        self.assertEqual(dialog.result(), QtWidgets.QDialog.DialogCode.Accepted)

    def test_compact_dialog_actions_remain_visible(self) -> None:
        cases = [
            (LoginDialog(), (340, 300)),
            (AboutDialog(), (360, 520)),
            (StartupGuideDialog(), (360, 440)),
            (TransitionHelpDialog(), (460, 400)),
        ]

        for dialog, size in cases:
            with self.subTest(dialog=type(dialog).__name__):
                dialog.resize(*size)
                dialog.show()
                self._app.processEvents()
                self.assertLessEqual(dialog.width(), size[0])
                self.assertLessEqual(dialog.height(), size[1])
                self.assertFalse(dialog.grab().isNull())

                if isinstance(dialog, LoginDialog):
                    actions = [dialog._ok_button, dialog._cancel_button]
                elif isinstance(dialog, AboutDialog):
                    actions = [dialog._github_btn, dialog._ok_btn]
                elif isinstance(dialog, StartupGuideDialog):
                    actions = [
                        dialog._settings_button,
                        dialog._queue_button,
                        dialog._close_button,
                    ]
                else:
                    actions = [dialog._open_button, dialog._close_button]

                for action in actions:
                    self.assertIsNotNone(action)
                    assert action is not None
                    top_left = action.mapTo(dialog, QtCore.QPoint(0, 0))
                    bottom_right = action.mapTo(
                        dialog,
                        QtCore.QPoint(action.width(), action.height()),
                    )
                    self.assertTrue(action.isVisibleTo(dialog))
                    self.assertGreaterEqual(top_left.x(), -1)
                    self.assertGreaterEqual(top_left.y(), -1)
                    self.assertLessEqual(bottom_right.x(), dialog.width() + 1)
                    self.assertLessEqual(bottom_right.y(), dialog.height() + 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
