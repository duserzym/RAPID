from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.dialogs.sample_select import SampleSelectDialog
from rapid_main.panels.sample_queue import SampleQueuePanel
from rapid_main.sample_index_registry import IndexQueueDefaults, read_registered_indexes


class SampleSelectDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_dialog_starts_empty_without_demo_samples(self) -> None:
        dialog = SampleSelectDialog()

        self.assertEqual(dialog._table.rowCount(), 0)
        self.assertIsNone(dialog.selected_sample)
        self.assertIn("No sample index loaded", dialog._source_lbl.text())
        self.assertEqual(dialog._source_lbl.property("status"), "neutral")
        self.assertFalse(dialog._select_btn.isEnabled())
        dialog.deleteLater()

    def test_csv_loader_populates_real_rows_and_missing_columns(self) -> None:
        dialog = SampleSelectDialog()
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "samples.csv"
            path.write_text(
                "Sample Name,Depth (cm),Formation,Location\n"
                "SPEC-1,12.5,Unit A,Site 1\n"
                "SPEC-2\n",
                encoding="utf-8",
            )

            self.assertTrue(dialog._load_csv(path))

        self.assertEqual(dialog._table.rowCount(), 2)
        self.assertEqual(dialog._table.item(0, 0).text(), "SPEC-1")
        self.assertEqual(dialog._table.item(0, 3).text(), "Site 1")
        self.assertEqual(dialog._table.item(1, 0).text(), "SPEC-2")
        self.assertEqual(dialog._table.item(1, 3).text(), "")
        dialog._table.selectRow(0)
        self.assertTrue(dialog._select_btn.isEnabled())
        self.assertEqual(
            dialog.selected_record,
            {
                "sample_name": "SPEC-1",
                "depth_cm": "12.5",
                "formation": "Unit A",
                "location": "Site 1",
            },
        )
        self.assertIsNotNone(dialog.registrations)
        self.assertEqual(dialog.registrations.names, ["SPEC-1", "SPEC-2"])
        dialog._search.setText("does-not-match")
        self.assertIsNone(dialog.selected_sample)
        self.assertFalse(dialog._select_btn.isEnabled())
        dialog.deleteLater()

    def test_csv_loader_rejects_header_only_file(self) -> None:
        dialog = SampleSelectDialog()
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "empty.csv"
            path.write_text("Sample Name,Depth (cm),Formation,Location\n", encoding="utf-8")
            with patch.object(QtWidgets.QMessageBox, "information") as information:
                self.assertFalse(dialog._load_csv(path))

        information.assert_called_once()
        self.assertEqual(dialog._table.rowCount(), 0)
        dialog.deleteLater()

    def test_registry_reopens_current_source_and_persists_edits_and_removal(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            index = root / 'field.sam'
            index.write_text('A\n', encoding='ascii')
            original = index.read_bytes()

            def open_dialog():
                return SampleSelectDialog(settings=QtCore.QSettings(str(root / 'registry.ini'), QtCore.QSettings.IniFormat))

            dialog = open_dialog()
            self.assertTrue(dialog._open_source(index))
            dialog._treatment.setText('SUSC')
            dialog._direction.setCurrentText('Up then Down')
            dialog._blocks.setValue(5)
            dialog._register_btn.click()
            self.assertEqual(index.read_bytes(), original)

            index.write_text('B\nC\n', encoding='ascii')
            reopened = open_dialog()
            reopened._indexes.setCurrentIndex(1)
            reopened._open_registration(1)
            self.assertEqual(reopened.registrations.names, ['B', 'C'])
            self.assertEqual(reopened.queue_defaults, IndexQueueDefaults('SUSC', True, True, 5))
            reopened._direction.setCurrentText('Down only')
            reopened._register_btn.click()
            edited = open_dialog()
            saved, _ = read_registered_indexes(edited._registry_settings)
            self.assertEqual(saved, {index.resolve(): IndexQueueDefaults('SUSC', False, False, 5)})

            index.unlink()
            unavailable = open_dialog()
            unavailable._indexes.setCurrentIndex(1)
            with patch.object(QtWidgets.QMessageBox, 'warning'):
                unavailable._open_registration(1)
            self.assertIsNone(unavailable.source_path)
            self.assertFalse(unavailable._select_btn.isEnabled())
            unavailable._remove_btn.click()
            cleared = open_dialog()
            self.assertEqual(cleared._indexes.count(), 1)

    def test_invalid_sources_and_conflicting_pending_presets_are_rejected(self) -> None:
        dialog = SampleSelectDialog()
        with tempfile.TemporaryDirectory() as temporary:
            index = Path(temporary) / 'field.sam'
            for rows in ('../outside\n', 'A\na\n'):
                index.write_text(rows, encoding='ascii')
                with patch.object(QtWidgets.QMessageBox, 'warning') as warning:
                    self.assertFalse(dialog._open_source(index))
                warning.assert_called_once()
            self.assertIsNone(dialog.source_path)
            self.assertEqual(dialog._table.rowCount(), 0)

            queue = SampleQueuePanel()
            queue.add_sample(position='A1', name='A', treatment='NRM -> AF20', source_file=str(index), avg_steps=2)
            queue.check_index_queue_settings(str(index), 'NRM → AF20', do_up=True, do_both=False, avg_steps=2)
            before = queue.row_snapshot()
            with self.assertRaisesRegex(ValueError, 'different settings'):
                queue.check_index_queue_settings(str(index), 'NRM → AF20', do_up=True, do_both=False, avg_steps=3)
            self.assertEqual(queue.row_snapshot(), before)
            with self.assertRaisesRegex(ValueError, 'exactly one treatment'):
                IndexQueueDefaults(do_both=True).resolve_treatment('NRM → AF20')


if __name__ == "__main__":
    unittest.main(verbosity=2)
