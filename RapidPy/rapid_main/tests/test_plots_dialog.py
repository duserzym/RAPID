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
