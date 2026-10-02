from __future__ import annotations

import unittest

from PySide6 import QtWidgets

from rapid_main.panels.sequence import SequencePanel
from rapid_main.thermal import compile_thermal_routine


class SequenceThermalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_reviewed_thermal_plan_is_handed_to_main_shell(self) -> None:
        class _Host(QtWidgets.QMainWindow):
            def __init__(self) -> None:
                super().__init__()
                self.labels: list[str] = []
                self.rockmag_plan = object()
                self.thermal_plan = None

            def load_sequence_labels(self, labels: list[str]) -> None:
                self.labels = list(labels)

            def set_rockmag_routine_plan(self, plan) -> None:
                self.rockmag_plan = plan

            def set_thermal_routine_plan(self, plan) -> None:
                self.thermal_plan = plan

        plan = compile_thermal_routine([100.0, 200.0], name="Manual oven sequence")
        host = _Host()
        panel = SequencePanel(host)
        host.setCentralWidget(panel)

        panel.install_thermal_plan(plan)

        self.assertEqual(panel.generate_labels(), ["TT100", "TT200"])
        self.assertEqual(host.labels, ["TT100", "TT200"])
        self.assertIs(host.thermal_plan, plan)
        self.assertIsNone(host.rockmag_plan)
        self.assertTrue(panel.has_unsaved_changes())

    def test_manual_edit_invalidates_thermal_plan_identity(self) -> None:
        panel = SequencePanel()
        plan = compile_thermal_routine([100.0], name="Manual oven sequence")
        panel.install_thermal_plan(plan)

        panel._chk_nrm.setChecked(True)
        QtWidgets.QApplication.processEvents()

        self.assertIsNone(panel._compiled_thermal_plan)
        self.assertNotEqual(panel.generate_labels(), plan.labels)


if __name__ == "__main__":
    unittest.main(verbosity=2)
