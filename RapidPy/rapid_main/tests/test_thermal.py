from __future__ import annotations

import unittest
import json
from pathlib import Path
import tempfile

from rapid_main.thermal import (
    ThermalSafetyLimits,
    ThermalStep,
    compile_thermal_routine,
    thermal_integration_decision,
    thermal_run_artifact,
    write_thermal_integration_decision,
    write_thermal_routine_artifact,
    write_thermal_run_artifact,
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
        self.assertEqual(payload["schema"], "rapidpy.thermal.plan.v1")
        self.assertEqual(payload["status"], "PLANNED")
        self.assertEqual(payload["run_context"], "bench-context")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertEqual(payload["labels"], ["TT100", "TT200"])
        self.assertEqual(payload["queue_labels"], ["TT100", "TT200"])
        self.assertEqual(payload["queue_block"]["block_type"], "thermal")
        self.assertTrue(payload["hardware_validation_required"])
        self.assertIn("hardware acceptance", payload["hardware_validation_statement"])

    def test_integration_decision_matches_versioned_repository_record(self) -> None:
        expected = thermal_integration_decision()
        repository_record = Path(__file__).parents[3] / "docs" / "thermal_integration_decision.json"

        self.assertEqual(json.loads(repository_record.read_text(encoding="utf-8")), expected)
        self.assertEqual(expected["status"], "MANUAL_EXTERNAL_ONLY")
        self.assertEqual(expected["automated_live_dispatch"], "BLOCKED")
        self.assertIn("VB6/modThermal.bas", [item["path"] for item in expected["legacy_source_evidence"]])

        with tempfile.TemporaryDirectory() as td:
            path = write_thermal_integration_decision(Path(td) / "decision.json")
            self.assertEqual(json.loads(path.read_text(encoding="utf-8")), expected)
            self.assertEqual(list(Path(td).glob("*.tmp")), [])

    def test_thermal_run_evidence_is_tied_to_exact_compiled_plan(self) -> None:
        plan = compile_thermal_routine([100.0, 200.0], name="External oven sequence")
        context = plan.to_artifact(
            run_context="thermal-run-1",
            operator="operator-a",
            timestamp_iso="2026-10-02T12:00:00+00:00",
        )
        payload = thermal_run_artifact(
            context,
            run_id="thermal-run-1",
            sample="THERMAL_1",
            operator="operator-a",
            labels_requested=plan.labels,
            completed_labels=["TT100"],
            errors=["operator halted before next external treatment"],
            aborted=True,
            simulated=False,
            final_phase="HALTED",
            timestamp_iso="2026-10-02T12:30:00+00:00",
        )

        self.assertEqual(payload["schema"], "rapidpy.thermal.run.v1")
        self.assertEqual(payload["status"], "ABORTED")
        self.assertEqual(payload["automation_status"], "MANUAL_EXTERNAL_ONLY")
        self.assertEqual(payload["labels_completed"], ["TT100"])
        self.assertFalse(payload["simulated"])

        with self.assertRaisesRegex(ValueError, "do not match"):
            thermal_run_artifact(
                context,
                run_id="bad",
                sample="THERMAL_1",
                operator="operator-a",
                labels_requested=["TT300"],
                completed_labels=[],
                aborted=True,
                simulated=False,
                final_phase="PREFLIGHT",
            )

        with tempfile.TemporaryDirectory() as td:
            path = write_thermal_run_artifact(
                Path(td) / "thermal_run.json",
                context,
                run_id="thermal-run-1",
                sample="THERMAL_1",
                operator="operator-a",
                labels_requested=plan.labels,
                completed_labels=[],
                aborted=True,
                simulated=False,
                final_phase="PREFLIGHT",
            )
            self.assertTrue(path.exists())
            self.assertEqual(list(Path(td).glob("*.tmp")), [])

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
