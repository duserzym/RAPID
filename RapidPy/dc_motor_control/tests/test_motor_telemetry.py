from __future__ import annotations

import unittest

from rapidpy_common.hardware import MotorAxisConfig, MotorSerialClient


class _FakeMotorClient(MotorSerialClient):
    def __init__(self) -> None:
        super().__init__()
        self.commands: list[str] = []

    def query_ascii(self, command: str) -> str:
        self.commands.append(command)
        if command.endswith("12 0 1 6 7"):
            return "@01 12 0000 03E8 0000 03DE FFFF FFF6 0064 005A"
        if command.endswith("12 9"):
            return "@01 12 FFFF FF06"
        raise AssertionError(f"Unexpected command: {command}")


class _FakeMotorClientNoTorque(MotorSerialClient):
    def __init__(self) -> None:
        super().__init__()
        self.commands: list[str] = []

    def query_ascii(self, command: str) -> str:
        self.commands.append(command)
        if command.endswith("12 0 1 6 7"):
            return "@01 12 0000 03E8 0000 03DE FFFF FFF6 0064 005A"
        if command.endswith("12 9"):
            raise RuntimeError("register not implemented")
        raise AssertionError(f"Unexpected command: {command}")


class MotorTelemetryTests(unittest.TestCase):
    def test_read_telemetry_decodes_dedicated_registers(self) -> None:
        client = _FakeMotorClient()
        axis = MotorAxisConfig("Turning", 2, 1)

        sample = client.read_telemetry(axis)

        self.assertEqual(sample.axis_name, "Turning")
        self.assertEqual(sample.target_position, 1000)
        self.assertEqual(sample.actual_position, 990)
        self.assertEqual(sample.position_error, -10)
        self.assertEqual(sample.velocity_1, 100)
        self.assertEqual(sample.velocity_2, 90)
        self.assertEqual(sample.actual_torque, -250)
        self.assertEqual(client.commands, ["@1 12 0 1 6 7", "@1 12 9"])

    def test_missing_torque_readback_falls_back_to_none(self) -> None:
        client = _FakeMotorClientNoTorque()
        axis = MotorAxisConfig("Turning", 2, 1)

        sample = client.read_telemetry(axis)

        self.assertIsNone(sample.actual_torque)
        self.assertEqual(client.commands, ["@1 12 0 1 6 7", "@1 12 9"])

        client.commands.clear()
        sample2 = client.read_telemetry(axis)

        self.assertIsNone(sample2.actual_torque)
        self.assertEqual(client.commands, ["@1 12 0 1 6 7"])

    def test_register_parser_sign_extends_32_bit_values(self) -> None:
        values = MotorSerialClient._parse_register_values(
            "@01 12 FFFF FFFF 7FFF FFFF", 2
        )
        self.assertEqual(values, (-1, 2_147_483_647))


if __name__ == "__main__":
    unittest.main()
