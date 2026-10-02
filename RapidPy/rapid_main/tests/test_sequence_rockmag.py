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
        self.assertIsNotNone(panel._compiled_routine_plan)
        self.assertEqual(
            panel._compiled_routine_plan.to_queue_labels(),
            panel.generate_labels(),
        )

    def test_hawaiian_preset_retains_executable_rockmag_plan(self) -> None:
        panel = SequencePanel()

        panel._preset_hawaiian()

        self.assertEqual(
            panel.generate_labels(),
            ["NRM", "AF25", "AF50", "AF100", "AF200", "AF400", "AF800"],
        )
        self.assertIsNotNone(panel._compiled_routine_plan)
        self.assertEqual(panel._compiled_routine_plan.spec.name, "Hawaiian AF Preset")

    def test_manual_change_after_rockmag_preset_returns_to_control_generated_labels(self) -> None:
        panel = SequencePanel()
        panel._preset_works()

        panel._chk_susc.setChecked(False)
        QtWidgets.QApplication.processEvents()

        labels = panel.generate_labels()
        self.assertNotEqual(labels, compile_rockmag_routine(rockmag_the_works()).to_queue_labels())
        self.assertNotIn("SUSC", labels)
        self.assertIsNone(panel._compiled_routine_plan)

    def test_compiled_plan_is_handed_to_main_shell_and_cleared_after_manual_edit(self) -> None:
        class _Host(QtWidgets.QMainWindow):
            def __init__(self) -> None:
                super().__init__()
                self.labels: list[str] = []
                self.plan = None

            def load_sequence_labels(self, labels: list[str]) -> None:
                self.labels = list(labels)

            def set_rockmag_routine_plan(self, plan) -> None:
                self.plan = plan

        host = _Host()
        panel = SequencePanel(host)
        host.setCentralWidget(panel)

        panel._preset_hawaiian()

        self.assertEqual(host.labels, panel.generate_labels())
        self.assertIs(host.plan, panel._compiled_routine_plan)

        panel._chk_susc.setChecked(True)
        QtWidgets.QApplication.processEvents()

        self.assertIsNone(host.plan)


if __name__ == "__main__":
    unittest.main(verbosity=2)
