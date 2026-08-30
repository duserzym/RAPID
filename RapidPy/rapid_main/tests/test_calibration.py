from __future__ import annotations

import json
import tempfile
import unittest

from rapid_main.calibration import (
    IrmVoltageCalibrationPoint,
    run_manual_calibration,
    write_calibration_result_artifact,
    fit_irm_voltage_calibration,
    write_irm_voltage_calibration_artifact,
)


class IrmVoltageCalibrationTests(unittest.TestCase):
    def test_fit_maps_irm_field_to_voltage_request(self) -> None:
        fit = fit_irm_voltage_calibration(
            [
                IrmVoltageCalibrationPoint(0.0, 0.1),
                IrmVoltageCalibrationPoint(500.0, 2.6),
                IrmVoltageCalibrationPoint(1000.0, 5.1),
            ],
            max_voltage_v=6.0,
        )

        self.assertAlmostEqual(fit.slope_v_per_mT, 0.005)
        self.assertAlmostEqual(fit.intercept_v, 0.1)
        self.assertAlmostEqual(fit.r_squared, 1.0)
        request = fit.request_for_field(750.0)
        self.assertAlmostEqual(request.voltage_v, 3.85)
        self.assertTrue(request.within_limits)

    def test_fit_can_force_zero_intercept_for_legacy_voltage_tables(self) -> None:
        fit = fit_irm_voltage_calibration(
            [(250.0, 1.25), (500.0, 2.5), (1000.0, 5.0)],
            force_zero_intercept=True,
        )

        self.assertEqual(fit.intercept_v, 0.0)
        self.assertAlmostEqual(fit.voltage_for_field(300.0), 1.5)
        self.assertTrue(fit.force_zero_intercept)

    def test_request_reports_voltage_limit_without_clamping(self) -> None:
        fit = fit_irm_voltage_calibration([(0.0, 0.0), (1000.0, 5.0)], max_voltage_v=4.0)

        request = fit.request_for_field(1000.0)

        self.assertAlmostEqual(request.voltage_v, 5.0)
        self.assertFalse(request.within_limits)

    def test_fit_rejects_unsafe_or_degenerate_inputs(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least two points"):
            fit_irm_voltage_calibration([(100.0, 1.0)])
        with self.assertRaisesRegex(ValueError, "distinct fields"):
            fit_irm_voltage_calibration([(100.0, 1.0), (100.0, 2.0)])
        with self.assertRaisesRegex(ValueError, "non-negative"):
            IrmVoltageCalibrationPoint(-1.0, 1.0)
        with self.assertRaisesRegex(ValueError, "max_voltage_v"):
            fit_irm_voltage_calibration([(0.0, 0.0), (1.0, 1.0)], max_voltage_v=0.0)

    def test_irm_voltage_fit_writes_timestamped_acceptance_artifact(self) -> None:
        fit = fit_irm_voltage_calibration([(0.0, 0.0), (1000.0, 5.0)], max_voltage_v=10.0)
        with tempfile.TemporaryDirectory() as td:
            path = write_irm_voltage_calibration_artifact(
                f"{td}/irm_voltage.json",
                fit,
                run_context="bench-irm-2026-07-15",
                operator="operator-a",
                notes="acceptance run",
            )
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["procedure_id"], "calibration/irm-voltage")
        self.assertEqual(payload["run_context"], "bench-irm-2026-07-15")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertAlmostEqual(payload["slope_v_per_mT"], 0.005)
        self.assertEqual(len(payload["points"]), 2)

    def test_generic_calibration_result_writes_acceptance_artifact(self) -> None:
        result = run_manual_calibration(0.99, expected=1.0, tolerance=0.02)
        with tempfile.TemporaryDirectory() as td:
            path = write_calibration_result_artifact(
                f"{td}/gaussmeter_baseline.json",
                result,
                run_context="bench-gaussmeter-2026-07-15",
                operator="operator-a",
            )
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["procedure_id"], "gaussmeter/squid-baseline")
        self.assertEqual(payload["run_context"], "bench-gaussmeter-2026-07-15")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertEqual(payload["status"], "PASS")
        self.assertTrue(payload["hardware_validation_required"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
