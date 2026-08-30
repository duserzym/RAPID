from __future__ import annotations

import unittest

from rapid_main.data_model import AngleVsFieldCollection, AngleVsFieldPoint
from rapid_main.transverse import plan_transverse_auto_position


class TransverseAutoPositionTests(unittest.TestCase):
    def test_plan_selects_peak_field_angle_and_shortest_clockwise_move(self) -> None:
        collection = AngleVsFieldCollection(
            [
                AngleVsFieldPoint(10.0, 2.0),
                AngleVsFieldPoint(35.0, 8.0),
                AngleVsFieldPoint(90.0, 4.0),
            ]
        )

        plan = plan_transverse_auto_position(collection, current_angle_deg=20.0, tolerance_deg=0.25)

        self.assertEqual(plan.target_angle_deg, 35.0)
        self.assertAlmostEqual(plan.relative_move_deg, 15.0)
        self.assertEqual(plan.direction, "CW")
        self.assertTrue(plan.move_required)
        self.assertTrue(plan.allowed)
        self.assertIn("Move transverse probe CW 15.000 deg", plan.command_summary)

    def test_plan_uses_shortest_counterclockwise_wraparound_move(self) -> None:
        collection = AngleVsFieldCollection(
            [
                AngleVsFieldPoint(350.0, 9.0),
                AngleVsFieldPoint(120.0, 1.0),
            ]
        )

        plan = plan_transverse_auto_position(collection, current_angle_deg=5.0)

        self.assertAlmostEqual(plan.relative_move_deg, -15.0)
        self.assertEqual(plan.direction, "CCW")

    def test_plan_holds_when_already_within_tolerance(self) -> None:
        collection = AngleVsFieldCollection([AngleVsFieldPoint(45.2, 10.0)])

        plan = plan_transverse_auto_position(collection, current_angle_deg=45.0, tolerance_deg=0.5)

        self.assertFalse(plan.move_required)
        self.assertEqual(plan.direction, "NONE")
        self.assertIn("Hold transverse probe", plan.command_summary)

    def test_plan_blocks_low_confidence_without_losing_target_evidence(self) -> None:
        collection = AngleVsFieldCollection(
            [
                AngleVsFieldPoint(0.0, 10.0),
                AngleVsFieldPoint(90.0, 10.0),
            ]
        )

        plan = plan_transverse_auto_position(
            collection,
            current_angle_deg=90.0,
            min_confidence=0.5,
        )

        self.assertFalse(plan.allowed)
        self.assertFalse(plan.move_required)
        self.assertEqual(plan.direction, "BLOCKED")
        self.assertIn("confidence", plan.reason)
        self.assertIn("BLOCKED", plan.command_summary)

    def test_plan_rejects_empty_scan_or_invalid_thresholds(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least one"):
            plan_transverse_auto_position(AngleVsFieldCollection([]), current_angle_deg=0.0)
        with self.assertRaisesRegex(ValueError, "tolerance"):
            plan_transverse_auto_position(
                AngleVsFieldCollection([AngleVsFieldPoint(0.0, 1.0)]),
                current_angle_deg=0.0,
                tolerance_deg=-0.1,
            )
        with self.assertRaisesRegex(ValueError, "min_confidence"):
            plan_transverse_auto_position(
                AngleVsFieldCollection([AngleVsFieldPoint(0.0, 1.0)]),
                current_angle_deg=0.0,
                min_confidence=1.5,
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
