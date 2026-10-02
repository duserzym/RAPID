from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from PySide6 import QtWidgets

from rapid_main.panels.calibration import CalibrationCenterPanel
from rapid_main.calibration_registry import CalibrationRegistry


class CalibrationCenterPanelTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if QtWidgets.QApplication.instance() is None:
            cls._qt_app = QtWidgets.QApplication([])
        else:
            cls._qt_app = None

    @classmethod
    def tearDownClass(cls) -> None:
        if cls._qt_app is not None:
            cls._qt_app.quit()
            cls._qt_app = None

    def test_irm_voltage_operator_path_writes_artifact(self) -> None:
        panel = CalibrationCenterPanel()
        try:
            with tempfile.TemporaryDirectory() as td:
                panel.set_procedure("irm_voltage")
                panel._irm_field_1.setValue(0.0)
                panel._irm_voltage_1.setValue(0.0)
                panel._irm_field_2.setValue(1000.0)
                panel._irm_voltage_2.setValue(5.0)
                panel._irm_context.setText("bench-context")
                panel._irm_operator.setText("operator-a")
                panel._irm_artifact_dir.setText(td)

                panel._run_irm_voltage()

                artifacts = list(Path(td).glob("irm_voltage_calibration_*.json"))
                self.assertEqual(len(artifacts), 1)
                payload = json.loads(artifacts[0].read_text(encoding="utf-8"))

            self.assertEqual(payload["procedure_id"], "calibration/irm-voltage")
            self.assertEqual(payload["run_context"], "bench-context")
            self.assertEqual(payload["operator"], "operator-a")
            self.assertAlmostEqual(payload["slope_v_per_mT"], 0.005)
            self.assertIn("artifact recorded", panel._status.text())
        finally:
            panel.deleteLater()

    def test_gaussmeter_baseline_manual_path_writes_artifact(self) -> None:
        panel = CalibrationCenterPanel()
        try:
            with tempfile.TemporaryDirectory() as td:
                panel.set_procedure("gaussmeter_baseline")
                panel.set_mode("manual")
                panel._manual_value.setValue(0.99)
                panel._expected.setValue(1.0)
                panel._tolerance.setValue(0.02)
                panel._baseline_context.setText("bench-context")
                panel._baseline_operator.setText("operator-a")
                panel._baseline_artifact_dir.setText(td)

                panel._run_manual()

                artifacts = list(Path(td).glob("gaussmeter_baseline_*.json"))
                self.assertEqual(len(artifacts), 1)
                payload = json.loads(artifacts[0].read_text(encoding="utf-8"))

            self.assertEqual(payload["procedure_id"], "calibration/gaussmeter_baseline")
            self.assertEqual(payload["run_context"], "bench-context")
            self.assertEqual(payload["operator"], "operator-a")
            self.assertEqual(payload["status"], "PASS")
            self.assertTrue(payload["hardware_validation_required"])
            self.assertIn("artifact recorded", panel._status.text())
        finally:
            panel.deleteLater()

    def test_thermal_operator_path_writes_planning_artifact(self) -> None:
        panel = CalibrationCenterPanel()
        recorded_plans = []
        panel.thermal_plan_recorded.connect(recorded_plans.append)
        try:
            with tempfile.TemporaryDirectory() as td:
                panel.set_procedure("thermal_routine")
                panel._thermal_name.setText("Bench thermal check")
                panel._thermal_targets.setText("100, 200")
                panel._thermal_hold_seconds.setValue(45)
                panel._thermal_context.setText("bench-context")
                panel._thermal_operator.setText("operator-a")
                panel._thermal_artifact_dir.setText(td)

                panel._run_thermal_routine()

                artifacts = list(Path(td).glob("thermal_routine_plan_*.json"))
                self.assertEqual(len(artifacts), 1)
                payload = json.loads(artifacts[0].read_text(encoding="utf-8"))

            self.assertEqual(payload["procedure_id"], "thermal/routine-planning")
            self.assertEqual(payload["routine_name"], "Bench thermal check")
            self.assertEqual(payload["labels"], ["TT100", "TT200"])
            self.assertEqual(payload["run_context"], "bench-context")
            self.assertEqual(payload["operator"], "operator-a")
            self.assertTrue(payload["hardware_validation_required"])
            self.assertIn("artifact recorded", panel._status.text())
            self.assertEqual(len(recorded_plans), 1)
            self.assertEqual(recorded_plans[0].to_queue_labels(), ["TT100", "TT200"])
            self.assertIn("manual-external", panel._run_thermal_btn.text().lower())
        finally:
            panel.deleteLater()

    def test_passed_artifact_can_be_approved_and_invalidated_from_panel(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            registry = CalibrationRegistry(root / "registry")
            panel = CalibrationCenterPanel(registry=registry)
            try:
                panel.set_procedure("gaussmeter_baseline")
                panel.set_mode("manual")
                panel._manual_value.setValue(1.0)
                panel._expected.setValue(1.0)
                panel._tolerance.setValue(0.02)
                panel._baseline_operator.setText("operator-a")
                panel._baseline_artifact_dir.setText(str(root / "artifacts"))
                panel._run_manual()

                self.assertTrue(panel._approve_artifact_btn.isEnabled())
                panel._registry_valid_days.setValue(30)
                panel._registry_reason.setText("Reference standard verified.")
                panel._approve_last_artifact()

                records = registry.records()
                self.assertEqual(len(records), 1)
                self.assertEqual(registry.state_for(records[0].record_id).status, "active")
                self.assertEqual(panel._registry_table.rowCount(), 1)
                self.assertIn("Approved and activated", panel._registry_status.text())

                panel._registry_table.selectRow(0)
                panel._registry_reason.setText("Post-check failed.")
                panel._invalidate_selected_record()
                self.assertEqual(registry.state_for(records[0].record_id).status, "invalidated")
                self.assertIn("Invalidated", panel._registry_status.text())
            finally:
                panel.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
