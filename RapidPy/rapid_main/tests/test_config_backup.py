from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PySide6 import QtWidgets

from rapid_main.config import (
    CONFIG_BACKUP_SCHEMA,
    AppConfig,
    read_config_backup,
    write_config_backup,
)
from rapid_main.panels.settings_panel import SettingsPanel
from rapid_main.offline_profile import offline_profile


class ConfigBackupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_versioned_backup_round_trips_all_config_sections(self) -> None:
        config = AppConfig()
        config.general.operator = "Operator A"
        config.squid.port = "COM17"
        config.motion.zero_pos = -1200
        config.susceptibility.port = "COM7"

        with tempfile.TemporaryDirectory() as temporary:
            path = write_config_backup(config, Path(temporary) / "backup.json")
            payload = json.loads(path.read_text(encoding="utf-8"))
            restored = read_config_backup(path)

        self.assertEqual(payload["schema"], CONFIG_BACKUP_SCHEMA)
        self.assertIn("created_at", payload)
        self.assertEqual(restored.general.operator, "Operator A")
        self.assertEqual(restored.squid.port, "COM17")
        self.assertEqual(restored.motion.zero_pos, -1200)
        self.assertEqual(restored.susceptibility.port, "COM7")

    def test_backup_reader_rejects_unknown_or_incomplete_schema(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "bad.json"
            path.write_text('{"schema":"unknown","config":{}}', encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Unsupported"):
                read_config_backup(path)

            path.write_text(
                json.dumps({"schema": CONFIG_BACKUP_SCHEMA, "config": {"general": []}}),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ValueError, "missing sections"):
                read_config_backup(path)

    def test_config_save_is_atomic_and_leaves_no_temporary_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            path = root / "config.json"
            config = AppConfig()
            config.general.lab_name = "Atomic Lab"

            config.save(path)

            self.assertEqual(AppConfig.load(path).general.lab_name, "Atomic Lab")
            self.assertEqual(list(root.glob("*.tmp")), [])

    def test_restore_persists_only_when_closing_and_reloads_on_next_launch(self) -> None:
        from rapid_main.app import MainWindow

        with tempfile.TemporaryDirectory() as temporary, offline_profile(Path(temporary)) as profile:
            root = Path(temporary)
            backup = root / "backup.json"
            active_config = AppConfig.default_path()
            restored = AppConfig.load(active_config)
            restored.general.operator = "Restored Operator"
            restored.squid.port = "COM22"
            write_config_backup(restored, backup)

            window = MainWindow()
            window._set_show_startup_guide(False)
            window.show()
            original = window.config
            backend = window.measurement_backend()
            original_bytes = active_config.read_bytes()
            self.assertTrue(original.general.nocomm)
            try:
                with (
                    patch.object(QtWidgets.QFileDialog, "getOpenFileName", return_value=(str(backup), "")),
                    patch.object(QtWidgets.QMessageBox, "question", return_value=QtWidgets.QMessageBox.StandardButton.Yes),
                    patch.object(QtWidgets.QMessageBox, "critical") as error,
                ):
                    with patch.object(window._sequence, "confirm_discard_changes", return_value=False):
                        self.assertFalse(window._settings_panel._restore_backup())
                    with patch.object(AppConfig, "save", side_effect=OSError("Disk unavailable")):
                        self.assertFalse(window._settings_panel._restore_backup())
                    error.assert_called_once()
                    self.assertTrue(window.isVisible())
                    self.assertIs(window.config, original)
                    self.assertEqual(active_config.read_bytes(), original_bytes)
                    self.assertTrue(window._settings_panel._restore_backup())
                self.assertFalse(window.isVisible())
                self.assertIs(window.config, original)
                self.assertIs(window.measurement_backend(), backend)
                self.assertEqual(window.config.general.operator, profile.general.operator)

                reopened = MainWindow()
                try:
                    self.assertTrue(reopened.config.general.nocomm)
                    self.assertEqual(reopened.config.general.operator, "Restored Operator")
                    self.assertEqual(reopened.config.squid.port, "COM22")
                    self.assertIsNot(reopened.measurement_backend(), backend)
                finally:
                    reopened._shutdown_cleanup_requested = True
                    reopened.close()
                    reopened.deleteLater()
            finally:
                if window.isVisible():
                    window._shutdown_cleanup_requested = True
                    window.close()
                window.deleteLater()

    def test_restore_is_blocked_while_workflow_is_active(self) -> None:
        window = QtWidgets.QMainWindow()
        window.config = AppConfig()  # type: ignore[attr-defined]
        window._operating_mode_change_blocker = Mock(return_value="Finish active work first.")  # type: ignore[attr-defined]
        panel = SettingsPanel()
        window.setCentralWidget(panel)

        with (
            patch.object(QtWidgets.QMessageBox, "warning") as warning,
            patch.object(QtWidgets.QFileDialog, "getOpenFileName") as choose_file,
        ):
            self.assertFalse(panel._restore_backup())

        warning.assert_called_once()
        choose_file.assert_not_called()
        window.deleteLater()

    def test_save_and_import_are_blocked_before_changing_active_settings(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            settings = root / "config.json"
            legacy = root / "legacy.ini"
            legacy.write_text("[Program]\nNoCommMode=False\nLastLogin=Imported operator\n", encoding="utf-8")
            window = QtWidgets.QMainWindow()
            window.config = AppConfig()  # type: ignore[attr-defined]
            window.config.general.operator = "Active operator"
            window.config.general.nocomm = True
            window._operating_mode_change_blocker = Mock(return_value="Finish active work first.")  # type: ignore[attr-defined]
            window._on_nocomm_toggled = Mock()  # type: ignore[attr-defined]
            panel = SettingsPanel(window)
            window.setCentralWidget(panel)
            panel.load_from_config(window.config)
            panel._operator.setText("Edited operator")
            panel._nocomm_chk.setChecked(False)
            try:
                with (
                    patch.dict(os.environ, {"RAPID_CONFIG": str(settings)}),
                    patch.object(QtWidgets.QFileDialog, "getOpenFileName", return_value=(str(legacy), "")) as choose,
                    patch.object(QtWidgets.QMessageBox, "warning") as warn,
                    patch.object(QtWidgets.QMessageBox, "information"),
                ):
                    window.config.save()
                    original = settings.read_bytes()
                    panel._save()
                    panel._import_vb6_ini()
                    choose.assert_not_called()
                    window._operating_mode_change_blocker.side_effect = ["", "Finish active work first."]
                    panel._import_vb6_ini()
                self.assertEqual(settings.read_bytes(), original)
                self.assertEqual(window.config.general.operator, "Active operator")
                self.assertTrue(window.config.general.nocomm)
                self.assertEqual(warn.call_count, 3)
                choose.assert_called_once()
                window._on_nocomm_toggled.assert_not_called()
            finally:
                window.deleteLater()

    def test_import_applies_to_the_window_only_after_settings_are_saved(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            settings = root / "config.json"
            legacy = root / "legacy.ini"
            legacy.write_text("[Program]\nNoCommMode=True\nLastLogin=Imported operator\n", encoding="utf-8")
            window = QtWidgets.QMainWindow()
            window.config = AppConfig()  # type: ignore[attr-defined]
            original = window.config
            original.general.nocomm = False
            window._nocomm_btn = QtWidgets.QCheckBox(window)  # type: ignore[attr-defined]
            window._operating_mode_change_blocker = Mock(return_value="")  # type: ignore[attr-defined]
            window._on_nocomm_toggled = Mock()  # type: ignore[attr-defined]
            window._nocomm_btn.toggled.connect(window._on_nocomm_toggled)
            panel = SettingsPanel(window)
            window.setCentralWidget(panel)
            try:
                with (
                    patch.dict(os.environ, {"RAPID_CONFIG": str(settings)}),
                    patch.object(QtWidgets.QFileDialog, "getOpenFileName", return_value=(str(legacy), "")),
                    patch.object(QtWidgets.QMessageBox, "critical") as error,
                    patch.object(QtWidgets.QMessageBox, "information"),
                ):
                    original.save()
                    before = settings.read_bytes()
                    with patch.object(AppConfig, "save", side_effect=OSError("Disk unavailable")):
                        panel._import_vb6_ini()
                    self.assertIs(window.config, original)
                    self.assertEqual(settings.read_bytes(), before)
                    self.assertFalse(window._nocomm_btn.isChecked())
                    window._on_nocomm_toggled.assert_not_called()
                    error.assert_called_once()

                    panel._import_vb6_ini()
                    self.assertEqual(window.config.general.operator, "Imported operator")
                    self.assertTrue(AppConfig.load(settings).general.nocomm)
                    self.assertTrue(window._nocomm_btn.isChecked())
                    window._on_nocomm_toggled.assert_called_once_with(True)
            finally:
                window.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
