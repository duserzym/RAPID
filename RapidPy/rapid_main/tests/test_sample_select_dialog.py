from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtWidgets

from rapid_main.dialogs.sample_select import SampleSelectDialog


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


if __name__ == "__main__":
    unittest.main(verbosity=2)
