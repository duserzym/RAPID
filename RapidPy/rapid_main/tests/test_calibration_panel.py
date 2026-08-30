from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from PySide6 import QtWidgets

from rapid_main.panels.calibration import CalibrationCenterPanel


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
        finally:
            panel.deleteLater()


if __name__ == "__main__":
    unittest.main(verbosity=2)
