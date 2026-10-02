from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PySide6 import QtWidgets

from rapid_main.config import (
    CONFIG_BACKUP_SCHEMA,
    AppConfig,
    read_config_backup,
    write_config_backup,
)
from rapid_main.panels.settings_panel import SettingsPanel


class ConfigBackupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_versioned_backup_round_trips_all_config_sections(self) -> None:
        config = AppConfig()
        config.general.operator = "Operator A"
        config.squid.port = "COM17"
        config.motion.zero_pos = -1200

        with tempfile.TemporaryDirectory() as temporary:
            path = write_config_backup(config, Path(temporary) / "backup.json")
            payload = json.loads(path.read_text(encoding="utf-8"))
            restored = read_config_backup(path)

        self.assertEqual(payload["schema"], CONFIG_BACKUP_SCHEMA)
        self.assertIn("created_at", payload)
        self.assertEqual(restored.general.operator, "Operator A")
        self.assertEqual(restored.squid.port, "COM17")
        self.assertEqual(restored.motion.zero_pos, -1200)

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

    def test_restore_persists_validated_config_and_requires_restart(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            backup = root / "backup.json"
            active_config = root / "active.json"
            restored = AppConfig()
            restored.general.operator = "Restored Operator"
            restored.squid.port = "COM22"
            write_config_backup(restored, backup)

            window = QtWidgets.QMainWindow()
            window.config = AppConfig()  # type: ignore[attr-defined]
            window._has_active_workflow = Mock(return_value=False)  # type: ignore[attr-defined]
            window._on_nocomm_toggled = Mock()  # type: ignore[attr-defined]
            window._estimator = SimpleNamespace(step_times={})  # type: ignore[attr-defined]
            window.set_status = Mock()  # type: ignore[attr-defined]
            panel = SettingsPanel()
            window.setCentralWidget(panel)

            with (
                patch.dict(os.environ, {"RAPID_CONFIG": str(active_config)}),
                patch.object(
                    QtWidgets.QFileDialog,
                    "getOpenFileName",
                    return_value=(str(backup), ""),
                ),
                patch.object(
                    QtWidgets.QMessageBox,
                    "question",
                    return_value=QtWidgets.QMessageBox.StandardButton.Yes,
                ),
                patch.object(QtWidgets.QMessageBox, "information"),
            ):
                self.assertTrue(panel._restore_backup())

            self.assertEqual(window.config.general.operator, "Restored Operator")  # type: ignore[attr-defined]
            self.assertEqual(AppConfig.load(active_config).squid.port, "COM22")
            self.assertEqual(panel._operator.text(), "Restored Operator")
            self.assertTrue(window._settings_restart_required)  # type: ignore[attr-defined]
            self.assertFalse(panel._restart_notice.isHidden())
            window._on_nocomm_toggled.assert_not_called()  # type: ignore[attr-defined]
            window.deleteLater()

    def test_restore_is_blocked_while_workflow_is_active(self) -> None:
        window = QtWidgets.QMainWindow()
        window.config = AppConfig()  # type: ignore[attr-defined]
        window._has_active_workflow = Mock(return_value=True)  # type: ignore[attr-defined]
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
