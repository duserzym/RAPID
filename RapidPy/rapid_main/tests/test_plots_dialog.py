from __future__ import annotations

import math
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from PySide6 import QtWidgets

from rapid_main.dialogs.plots import (
    PlotsDialog,
    build_quicklook_summary,
    write_quicklook_csv,
    write_quicklook_json,
)


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
        self.assertIn("principal_axis", data["analysis"])
        self.assertEqual(dialog._analysis_table.rowCount(), 11)
        self.assertTrue(dialog._analysis_table.accessibleName())
        dialog.deleteLater()

    def test_analysis_tab_reports_principal_axis_and_decay_evidence(self) -> None:
        dialog = PlotsDialog()
        dialog.set_data(
            north=[4.0, 3.0, 2.0, 1.0],
            east=[4.0, 3.0, 2.0, 1.0],
            up=[0.0, 0.0, 0.0, 0.0],
            labels=["NRM", "AF10", "AF20", "AF30"],
        )

        analysis = dialog.quicklook_summary()["analysis"]
        principal = analysis["principal_axis"]
        decay = analysis["decay"]
        self.assertAlmostEqual(principal["declination_deg"], 45.0)
        self.assertAlmostEqual(principal["variance_fraction"], 1.0)
        self.assertAlmostEqual(principal["rms_perpendicular"], 0.0)
        self.assertTrue(decay["monotonic_nonincreasing"])
        self.assertEqual(dialog._analysis_status.property("status"), "ready")
        self.assertTrue(dialog._analysis_status.text().startswith("READY —"))
        dialog.deleteLater()

    def test_analysis_tab_warns_for_non_monotonic_sequence(self) -> None:
        dialog = PlotsDialog()
        dialog.set_data(
            north=[3.0, 1.0, 2.0],
            east=[0.0, 0.0, 0.0],
            up=[0.0, 0.0, 0.0],
            labels=["NRM", "AF10", "AF20"],
        )

        self.assertEqual(dialog._analysis_status.property("status"), "warning")
        self.assertIn("not monotonic", dialog._analysis_status.text())
        dialog.deleteLater()

    def test_principal_axis_uses_north_east_down_inclination_convention(self) -> None:
        summary = build_quicklook_summary(
            north=[0.0, 0.0],
            east=[0.0, 0.0],
            up=[-2.0, -1.0],
            labels=["NRM", "AF10"],
        )

        principal = summary["analysis"]["principal_axis"]
        self.assertEqual(principal["coordinate_convention"], "north_east_down")
        self.assertAlmostEqual(principal["inclination_deg"], 90.0)

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
        self.assertTrue(dialog.quicklook_summary()["provenance"]["simulated"])
        self.assertTrue(dialog._export_btn.isEnabled())
        dialog.deleteLater()

    def test_real_data_replaces_simulated_example_marker(self) -> None:
        dialog = PlotsDialog()
        dialog._load_demo()

        dialog.set_data([1.0], [0.0], [0.0], ["NRM"])

        self.assertEqual(dialog._demo_lbl.property("status"), "ready")
        self.assertIn("Measurement data: 1 step", dialog._demo_lbl.text())
        self.assertNotIn("SIMULATED", dialog._demo_lbl.text())
        self.assertFalse(dialog.quicklook_summary()["provenance"]["simulated"])
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
        self.assertFalse(payload["provenance"]["simulated"])
        self.assertIn("principal_axis", payload["analysis"])
        dialog.deleteLater()

    def test_simulated_json_and_csv_exports_retain_provenance(self) -> None:
        dialog = PlotsDialog()
        dialog._load_demo()

        with tempfile.TemporaryDirectory() as td:
            json_path = dialog.write_quicklook_json(Path(td) / "quicklook.json")
            csv_path = dialog.write_quicklook_csv(Path(td) / "quicklook.csv")
            payload = json.loads(json_path.read_text(encoding="utf-8"))
            csv_text = csv_path.read_text(encoding="utf-8")

        self.assertTrue(payload["provenance"]["simulated"])
        self.assertIn("not hardware evidence", payload["provenance"]["statement"])
        self.assertIn("simulated,provenance_statement", csv_text.splitlines()[0])
        self.assertIn("pca_declination_deg", csv_text.splitlines()[0])
        self.assertIn(",true,Example plot data; not hardware evidence", csv_text)
        dialog.deleteLater()

    def test_csv_export_rejects_mismatched_fields(self) -> None:
        summary = build_quicklook_summary([1.0], [0.0], [0.0], ["NRM"])
        summary["intensity"] = []
        with tempfile.TemporaryDirectory() as td:
            with self.assertRaisesRegex(ValueError, "equal lengths"):
                write_quicklook_csv(Path(td) / "bad.csv", summary)

    def test_operator_export_action_writes_selected_csv(self) -> None:
        dialog = PlotsDialog()
        dialog.set_data([1.0], [0.0], [0.0], ["NRM"])

        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "operator-export"
            with (
                mock.patch.object(
                    QtWidgets.QFileDialog,
                    "getSaveFileName",
                    return_value=(str(target), "CSV table (*.csv)"),
                ),
                mock.patch.object(QtWidgets.QMessageBox, "information") as info,
            ):
                dialog._export_data()
            written = target.with_suffix(".csv")
            self.assertTrue(written.exists())
            self.assertIn("NRM", written.read_text(encoding="utf-8"))
            info.assert_called_once()
        dialog.deleteLater()

    def test_atomic_export_cleans_temporary_file_on_publish_failure(self) -> None:
        summary = build_quicklook_summary([1.0], [0.0], [0.0], ["NRM"])
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "quicklook.json"
            with mock.patch("rapid_main.dialogs.plots.os.replace", side_effect=OSError("denied")):
                with self.assertRaisesRegex(OSError, "denied"):
                    write_quicklook_json(target, summary)
            self.assertFalse(target.exists())
            self.assertEqual(list(Path(td).glob("*.tmp")), [])

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
