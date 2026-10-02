from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from PySide6 import QtWidgets

from rapid_main.io.sequence_io import SequenceFormatError, load_sequence_strict
from rapid_main.panels.sequence import SequencePanel


class SequenceDocumentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_loaded_labels_remain_the_active_saveable_sequence(self) -> None:
        panel = SequencePanel()
        labels = ["NRM", "AF25", "AF50"]
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "loaded.json"
            output = Path(tmp) / "saved.json"
            panel.load_labels(labels, source_path=source)

            self.assertEqual(panel.generate_labels(), labels)
            self.assertFalse(panel.has_unsaved_changes())
            self.assertTrue(panel.save_to_path(output))
            self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["steps"], labels)

    def test_hawaiian_preset_contains_documented_af_steps_and_is_dirty(self) -> None:
        panel = SequencePanel()
        panel._preset_hawaiian()
        self.assertEqual(
            panel.generate_labels(),
            ["NRM", "AF25", "AF50", "AF100", "AF200", "AF400", "AF800"],
        )
        self.assertTrue(panel.has_unsaved_changes())
        self.assertIn("Unsaved", panel._preview_header.text())

    def test_strict_loader_reports_malformed_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.json"
            path.write_text('{"steps": ["NRM", 3]}', encoding="utf-8")
            with self.assertRaisesRegex(SequenceFormatError, "must be text"):
                load_sequence_strict(path)

    def test_manual_change_after_load_marks_document_dirty(self) -> None:
        panel = SequencePanel()
        panel.load_labels(["NRM", "AF25"], source_path=Path("source.json"))
        panel._chk_susc.setChecked(True)
        QtWidgets.QApplication.processEvents()
        self.assertTrue(panel.has_unsaved_changes())
        self.assertNotEqual(panel.generate_labels(), ["NRM", "AF25"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
