from __future__ import annotations

import json
import tempfile
import unittest

from rapid_main.rockmag import (
    RockmagRoutineSpec,
    compile_rockmag_routine,
    rockmag_routine_artifact,
    rockmag_af_demag,
    rockmag_the_works,
    write_rockmag_routine_artifact,
)


class RockmagRoutineTests(unittest.TestCase):
    def test_the_works_template_compiles_to_runner_labels_and_family_blocks(self) -> None:
        spec = rockmag_the_works(
            af_fields_mT=(20.0, 40.0),
            irm_fields_g=(100.0,),
            arm_fields_g=(50.0,),
        )

        plan = compile_rockmag_routine(spec)

        self.assertEqual(plan.spec.name, "Rockmag the Works")
        self.assertEqual(plan.labels, ["NRM", "AF20", "AF40", "IRM100", "ARM50", "IRM-BF", "SUSC"])
        self.assertEqual(plan.to_queue_labels(), plan.labels)
        self.assertEqual(
            plan.blocks.block_names(),
            [
                "Rockmag the Works: NRM",
                "Rockmag the Works: AF",
                "Rockmag the Works: IRM",
                "Rockmag the Works: ARM",
                "Rockmag the Works: BACKFIELD",
                "Rockmag the Works: SUSCEPTIBILITY",
            ],
        )
        self.assertEqual(plan.blocks.blocks[1].metadata["template"], "the-works")

    def test_af_template_formats_fractional_fields_and_optional_nrm(self) -> None:
        spec = rockmag_af_demag((2.5, 5.0), include_nrm=False)
        plan = compile_rockmag_routine(spec)

        self.assertEqual(plan.labels, ["AF2.5", "AF5"])
        self.assertEqual(plan.blocks.block_names(), ["Rockmag AF demag: AF"])

    def test_repeated_spec_preserves_repeat_boundaries_in_labels(self) -> None:
        spec = RockmagRoutineSpec("Repeat check", ("NRM", "AF10"), repeats=2)
        plan = compile_rockmag_routine(spec)

        self.assertEqual(plan.labels, ["NRM", "AF10", "REPEAT2", "NRM", "AF10"])
        self.assertEqual(plan.step_count, 5)

    def test_plan_artifact_preserves_queue_labels_and_hardware_caveat(self) -> None:
        plan = compile_rockmag_routine(rockmag_af_demag((5.0, 10.0)))

        payload = rockmag_routine_artifact(
            plan,
            run_context="bench-context",
            operator="operator-a",
            notes="artifact test",
            timestamp_iso="2026-07-15T12:00:00+00:00",
        )

        self.assertEqual(payload["procedure_id"], "rockmag/routine-planning")
        self.assertEqual(payload["status"], "PLANNED")
        self.assertEqual(payload["run_context"], "bench-context")
        self.assertEqual(payload["operator"], "operator-a")
        self.assertEqual(payload["labels"], ["NRM", "AF5", "AF10"])
        self.assertEqual(payload["queue_labels"], payload["labels"])
        self.assertTrue(payload["hardware_validation_required"])
        self.assertEqual(payload["blocks"][0]["block_type"], "rockmag")
        json.dumps(payload)

    def test_write_plan_artifact_outputs_json_sidecar(self) -> None:
        plan = compile_rockmag_routine(
            rockmag_the_works(af_fields_mT=(20.0,), irm_fields_g=(), arm_fields_g=())
        )

        with tempfile.TemporaryDirectory() as tmp:
            path = write_rockmag_routine_artifact(
                f"{tmp}/rockmag_plan.json",
                plan,
                run_context="queue-demo",
                operator="operator-b",
            )
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["routine_name"], "Rockmag the Works")
        self.assertEqual(payload["run_context"], "queue-demo")
        self.assertIn("Rockmag the Works: AF", [block["name"] for block in payload["blocks"]])

    def test_specs_reject_empty_or_unsafe_inputs(self) -> None:
        with self.assertRaisesRegex(ValueError, "name"):
            RockmagRoutineSpec(" ", ("NRM",))
        with self.assertRaisesRegex(ValueError, "at least one label"):
            RockmagRoutineSpec("Empty", (" ",))
        with self.assertRaisesRegex(ValueError, "repeats"):
            RockmagRoutineSpec("Bad repeat", ("NRM",), repeats=0)
        with self.assertRaisesRegex(ValueError, "non-negative"):
            rockmag_af_demag((-1.0,))


if __name__ == "__main__":
    unittest.main(verbosity=2)
