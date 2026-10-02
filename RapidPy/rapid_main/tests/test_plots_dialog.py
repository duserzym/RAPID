from __future__ import annotations

import math
import json
import tempfile
import unittest
from pathlib import Path

from PySide6 import QtWidgets

from rapid_main.dialogs.plots import PlotsDialog, build_quicklook_summary, write_quicklook_json


class PlotsDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_set_data_records_quicklook_plot_contract(self) -> None:
        dialog = PlotsDialog()

        dialog.set_data(
            north=[1.0, 0.0],
            east=[0.0, 1.0],
            up=[0.0, -1.0],
            labels=["NRM", "AF20"],
        )

        data = dialog._last_plot_data
        self.assertEqual(data["labels"], ["NRM", "AF20"])
        self.assertEqual(data["down"], [-0.0, 1.0])
        self.assertEqual(data["declination"], [0.0, 90.0])
        self.assertEqual(data["inclination"], [-0.0, 45.0])
        self.assertAlmostEqual(data["intensity"][0], 1.0)
        self.assertAlmostEqual(data["intensity"][1], math.sqrt(2.0))
        dialog.deleteLater()

    def test_dialog_starts_empty_without_implicit_demo_evidence(self) -> None:
        dialog = PlotsDialog()

        self.assertEqual(dialog.quicklook_summary()["step_count"], 0)
        self.assertIn("No measurement data", dialog._demo_lbl.text())
        dialog.deleteLater()

    def test_simulated_example_is_explicitly_marked(self) -> None:
        dialog = PlotsDialog()

        dialog._load_demo()

        self.assertGreater(dialog.quicklook_summary()["step_count"], 0)
        self.assertEqual(dialog._demo_lbl.property("status"), "simulated")
        self.assertIn("SIMULATED", dialog._demo_lbl.text())
        self.assertIn("not hardware evidence", dialog._demo_lbl.text())
        dialog.deleteLater()

    def test_real_data_replaces_simulated_example_marker(self) -> None:
        dialog = PlotsDialog()
        dialog._load_demo()

        dialog.set_data([1.0], [0.0], [0.0], ["NRM"])

        self.assertEqual(dialog._demo_lbl.property("status"), "ready")
        self.assertIn("Measurement data: 1 step", dialog._demo_lbl.text())
        self.assertNotIn("SIMULATED", dialog._demo_lbl.text())
        dialog.deleteLater()

    def test_quicklook_rejects_mismatched_vector_lengths(self) -> None:
        with self.assertRaisesRegex(ValueError, "equal lengths"):
            build_quicklook_summary([1.0], [0.0, 1.0], [0.0], ["NRM"])

    def test_quicklook_rejects_non_finite_values(self) -> None:
        with self.assertRaisesRegex(ValueError, "not finite"):
            build_quicklook_summary([float("nan")], [0.0], [0.0], ["NRM"])

    def test_quicklook_summary_can_be_written_as_json_artifact(self) -> None:
        dialog = PlotsDialog()
        dialog.set_data(
            north=[1.0],
            east=[0.0],
            up=[0.0],
            labels=["NRM"],
        )

        with tempfile.TemporaryDirectory() as td:
            path = dialog.write_quicklook_json(Path(td) / "quicklook.json")
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["step_count"], 1)
        self.assertEqual(payload["labels"], ["NRM"])
        self.assertEqual(payload["vectors"]["north"], [1.0])
        self.assertEqual(payload["intensity"], [1.0])
        dialog.deleteLater()

    def test_quicklook_summary_helper_does_not_require_dialog(self) -> None:
        summary = build_quicklook_summary(
            north=[0.0],
            east=[1.0],
            up=[-1.0],
            labels=["AF20"],
        )

        with tempfile.TemporaryDirectory() as td:
            path = write_quicklook_json(Path(td) / "quicklook.json", summary)
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["step_count"], 1)
        self.assertEqual(payload["labels"], ["AF20"])
        self.assertEqual(payload["vectors"]["down"], [1.0])
        self.assertEqual(payload["declination"], [90.0])
        self.assertEqual(payload["inclination"], [45.0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
