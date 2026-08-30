from __future__ import annotations

import unittest

from PySide6 import QtWidgets

from rapid_main.panels.sequence import SequencePanel
from rapid_main.rockmag import compile_rockmag_routine, rockmag_the_works


class SequenceRockmagTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_rockmag_works_preset_uses_compiled_routine_labels(self) -> None:
        panel = SequencePanel()
        expected = compile_rockmag_routine(rockmag_the_works()).to_queue_labels()

        panel._preset_works()

        self.assertEqual(panel.generate_labels(), expected)
        self.assertIn("IRM-BF", panel.generate_labels())
        self.assertIn("SUSC", panel.generate_labels())

    def test_manual_change_after_rockmag_preset_returns_to_control_generated_labels(self) -> None:
        panel = SequencePanel()
        panel._preset_works()

        panel._chk_susc.setChecked(False)
        QtWidgets.QApplication.processEvents()

        labels = panel.generate_labels()
        self.assertNotEqual(labels, compile_rockmag_routine(rockmag_the_works()).to_queue_labels())
        self.assertNotIn("SUSC", labels)


if __name__ == "__main__":
    unittest.main(verbosity=2)
