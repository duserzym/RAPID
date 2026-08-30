from __future__ import annotations

import unittest

from rapid_main.dac import (
    DacChannelConfig,
    DacCommandPlanner,
    execute_dac_command,
    inspect_dac_startup,
)


class FakeDacAdapter:
    def __init__(self) -> None:
        self.calls: list[tuple[int, float]] = []

    def write_voltage(self, channel: int, voltage: float) -> str:
        self.calls.append((channel, voltage))
        return f"ch{channel}={voltage}"


class DacCommandTests(unittest.TestCase):
    def test_planner_accepts_in_range_voltage_without_clamping(self) -> None:
        planner = DacCommandPlanner([DacChannelConfig(0, "IRM axial", 0.0, 5.0)])

        command = planner.plan_write(0, 3.25, purpose="irm-pulse")

        self.assertTrue(command.allowed)
        self.assertEqual(command.channel_number, 0)
        self.assertEqual(command.voltage_v, 3.25)
        self.assertIn("IRM-PULSE IRM axial ch0: 3.25 V", command.command_summary)

    def test_planner_blocks_out_of_range_voltage_without_adapter_call(self) -> None:
        planner = DacCommandPlanner([DacChannelConfig(1, "IRM transverse", -2.0, 2.0)])
        adapter = FakeDacAdapter()

        command = planner.plan_write(1, 2.5)

        self.assertFalse(command.allowed)
        self.assertIn("outside -2..2 V", command.reason)
        with self.assertRaisesRegex(RuntimeError, "BLOCKED"):
            execute_dac_command(adapter, command)
        self.assertEqual(adapter.calls, [])

    def test_planner_emits_safe_zero_command(self) -> None:
        planner = DacCommandPlanner([DacChannelConfig(2, "AF monitor", -10.0, 10.0, safe_v=0.1)])
        adapter = FakeDacAdapter()

        command = planner.plan_zero(2)
        result = execute_dac_command(adapter, command)

        self.assertTrue(command.allowed)
        self.assertEqual(command.voltage_v, 0.1)
        self.assertEqual(result, "ch2=0.1")
        self.assertEqual(adapter.calls, [(2, 0.1)])

    def test_planner_rejects_invalid_channel_configuration(self) -> None:
        with self.assertRaisesRegex(ValueError, "non-negative"):
            DacChannelConfig(-1)
        with self.assertRaisesRegex(ValueError, "below maximum"):
            DacChannelConfig(0, minimum_v=1.0, maximum_v=1.0)
        with self.assertRaisesRegex(ValueError, "safe voltage"):
            DacChannelConfig(0, minimum_v=0.0, maximum_v=1.0, safe_v=2.0)
        with self.assertRaisesRegex(ValueError, "unique"):
            DacCommandPlanner([DacChannelConfig(0), DacChannelConfig(0)])

    def test_execute_requires_adapter_write_voltage_surface(self) -> None:
        planner = DacCommandPlanner([DacChannelConfig(0)])
        command = planner.plan_zero(0)

        with self.assertRaisesRegex(TypeError, "write_voltage"):
            execute_dac_command(object(), command)

    def test_startup_report_documents_planning_only_retained_interface(self) -> None:
        planner = DacCommandPlanner(
            [
                DacChannelConfig(0, "IRM axial", 0.0, 10.0, safe_v=0.0),
                DacChannelConfig(1, "IRM transverse", 0.0, 10.0, safe_v=0.1),
            ]
        )

        report = inspect_dac_startup(planner)

        self.assertTrue(report.ok)
        self.assertFalse(report.adapter_present)
        self.assertFalse(report.adapter_ready)
        self.assertEqual(len(report.safe_zero_commands), 2)
        self.assertIn("planning-only", report.warnings[0])
        self.assertIn("SAFE-ZERO IRM transverse ch1: 0.1 V", report.safe_zero_summary)

    def test_startup_report_blocks_adapter_without_write_voltage(self) -> None:
        planner = DacCommandPlanner([DacChannelConfig(0)])

        report = inspect_dac_startup(planner, adapter=object())

        self.assertFalse(report.ok)
        self.assertTrue(report.adapter_present)
        self.assertFalse(report.adapter_ready)
        self.assertIn("write_voltage", report.blockers[0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
