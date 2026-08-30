from __future__ import annotations

import unittest
import json
from pathlib import Path
import tempfile

from rapid_main.thermal import (
    ThermalSafetyLimits,
    ThermalStep,
    compile_thermal_routine,
    write_thermal_routine_artifact,
)


class ThermalRoutineTests(unittest.TestCase):
    def test_thermal_step_parses_labels_and_reports_kelvin(self) -> None:
        step = ThermalStep.from_label("TH 400")

        self.assertEqual(step.label, "TH400")
        self.assertEqual(step.temperature_c, 400.0)
        self.assertAlmostEqual(step.treatment_temp_k(), 673.15)

    def test_compile_thermal_routine_produces_queue_block_and_estimate(self) -> None:
        limits = ThermalSafetyLimits(
            max_temperature_c=500.0,
            ambient_temperature_c=25.0,
            cool_to_c=50.0,
            ramp_rate_c_per_min=25.0,
            default_hold_seconds=120,
        )

        plan = compile_thermal_routine([100.0, 200.0], name="TT check", limits=limits)
        block = plan.to_measurement_block()

        self.assertEqual(plan.labels, ["TT100", "TT200"])
        self.assertEqual(plan.max_temperature_c, 200.0)
        self.assertTrue(plan.requires_cooldown)
        self.assertEqual(block.block_type, "thermal")
        self.assertEqual(block.to_queue_labels(), ["TT100", "TT200"])
        self.assertEqual(block.metadata["requires_cooldown"], "True")
        self.assertEqual(int(block.metadata["estimated_seconds"]), plan.estimated_seconds())

    def test_compile_thermal_routine_respects_hold_override(self) -> None:
        limits = ThermalSafetyLimits(ramp_rate_c_per_min=60.0, default_hold_seconds=600)

        plan = compile_thermal_routine([85.0], limits=limits, hold_seconds=30)

        self.assertEqual(plan.labels, ["TT85"])
        self.assertEqual(plan.estimated_seconds(), 125)

    def test_thermal_plan_writes_operator_evidence_artifact(self) -> None:
        limits = ThermalSafetyLimits(max_temperature_c=500.0, ramp_rate_c_per_min=25.0)
        plan = compile_thermal_routine(
            [100.0, 200.0],
            name="Bench thermal check",
            limits=limits,
            hold_seconds=45,
        )

        with tempfile.TemporaryDirectory() as td:
            path = write_thermal_routine_artifact(
                Path(td) / "thermal.json",
                plan,
                run_context="bench-context",
                operator="operator-a",
                timestamp_iso="2026-07-15T00:00:00+00:00",
            )
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["procedure_id"], "thermal/routine-planning")
        self.assertEqual(payload["status"], "PLANNED")
        self.assertEqual(payload["run_context"], "bench-context")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertEqual(payload["labels"], ["TT100", "TT200"])
        self.assertEqual(payload["queue_block"]["block_type"], "thermal")
        self.assertTrue(payload["hardware_validation_required"])
        self.assertIn("hardware acceptance", payload["hardware_validation_statement"])

    def test_thermal_plan_rejects_unsafe_or_degenerate_inputs(self) -> None:
        with self.assertRaisesRegex(ValueError, "invalid thermal label"):
            ThermalStep.from_label("AF20")
        with self.assertRaisesRegex(ValueError, "non-negative"):
            ThermalStep(-1.0)
        with self.assertRaisesRegex(ValueError, "at least one"):
            compile_thermal_routine([])
        with self.assertRaisesRegex(ValueError, "exceeds max"):
            compile_thermal_routine([800.0], limits=ThermalSafetyLimits(max_temperature_c=700.0))
        with self.assertRaisesRegex(ValueError, "ramp rate"):
            ThermalSafetyLimits(ramp_rate_c_per_min=0.0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
