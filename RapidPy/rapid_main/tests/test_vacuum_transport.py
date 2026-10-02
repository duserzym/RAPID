"""Tests for the legacy vacuum controller's exact serial evidence."""
from __future__ import annotations

import unittest
from unittest import mock

from updown_control.app import VacuumCommunicationError, VacuumController


class _FakeVacuumSerial:
    def __init__(self, responses: dict[str, str]) -> None:
        self.is_open = True
        self.rts = False
        self.writes: list[bytes] = []
        self._responses = dict(responses)
        self._buffer = b""

    def reset_input_buffer(self) -> None:
        self._buffer = b""

    def reset_output_buffer(self) -> None:
        return None

    def write(self, payload: bytes) -> int:
        self.writes.append(payload)
        if payload != b"\r":
            command = payload.decode("ascii")
            response = self._responses.get(command)
            if response is not None:
                self._buffer = (response + "\r").encode("ascii")
        return len(payload)

    def flush(self) -> None:
        return None

    def read(self, size: int = 1) -> bytes:
        chunk, self._buffer = self._buffer[:size], self._buffer[size:]
        return chunk


class VacuumControllerTransportTests(unittest.TestCase):
    def test_enable_records_exact_commands_and_acknowledgements(self) -> None:
        events: list[tuple[str, str, str]] = []
        controller = VacuumController(trace=lambda *event: events.append(event))
        controller._serial = _FakeVacuumSerial(  # type: ignore[attr-defined]
            {"10MFF": "MOTOR-ON", "10VFF": "VALVE-OPEN"}
        )

        with mock.patch("updown_control.app.time.sleep"):
            controller.set_enabled(True)

        commands = [
            payload.decode("ascii")
            for payload in controller._serial.writes  # type: ignore[attr-defined]
            if payload != b"\r"
        ]
        self.assertEqual(commands, ["E", "10MFF", "O", "10VFF"])
        self.assertEqual(
            [(direction, payload) for direction, payload, _detail in events],
            [
                ("TX", "E"),
                ("TX", "10MFF"),
                ("RX", "MOTOR-ON"),
                ("TX", "O"),
                ("TX", "10VFF"),
                ("RX", "VALVE-OPEN"),
            ],
        )
        self.assertTrue(controller.is_enabled)

    def test_missing_or_unterminated_response_fails_without_changing_state(self) -> None:
        events: list[tuple[str, str, str]] = []
        controller = VacuumController(trace=lambda *event: events.append(event))
        controller._serial = _FakeVacuumSerial({})  # type: ignore[attr-defined]

        with mock.patch("updown_control.app.time.sleep"):
            with self.assertRaises(VacuumCommunicationError):
                controller.set_motor_power(True)

        self.assertFalse(controller.is_enabled)
        self.assertEqual(events[-1][0], "ERROR")
        self.assertIn("no reply", events[-1][2])


if __name__ == "__main__":
    unittest.main(verbosity=2)
