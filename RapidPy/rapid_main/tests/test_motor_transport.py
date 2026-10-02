from __future__ import annotations

import unittest
from unittest import mock

from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorSerialClient


class _FakeSerial:
    def __init__(self, replies: list[bytes]) -> None:
        self.is_open = True
        self.dtr = False
        self.rts = False
        self.replies = list(replies)
        self.writes: list[bytes] = []

    def write(self, payload: bytes) -> None:
        self.writes.append(bytes(payload))

    def flush(self) -> None:
        return None

    def read_until(self, terminator: bytes) -> bytes:
        self.terminator = terminator
        return self.replies.pop(0) if self.replies else b""


class MotorTransportEvidenceTests(unittest.TestCase):
    def _client(self, replies: list[bytes]):
        events: list[tuple[str, str, str]] = []
        client = MotorSerialClient(
            trace=lambda direction, payload, detail: events.append(
                (direction, payload, detail)
            )
        )
        serial = _FakeSerial(replies)
        client._serial = serial
        client._port = "COM3"
        return client, serial, events

    def test_query_records_exact_command_and_cr_terminated_reply(self) -> None:
        client, serial, events = self._client([b"@1 ACK 0000\r"])

        reply = client.query_ascii("@1 0")

        self.assertEqual(reply, "@1 ACK 0000")
        self.assertEqual(serial.writes, [b"@1 0\r\n"])
        self.assertEqual(
            [(direction, payload) for direction, payload, _detail in events],
            [("TX", "@1 0\r\n"), ("RX", "@1 ACK 0000\r")],
        )
        self.assertIn("@1 0", events[-1][2])

    def test_unterminated_partial_reply_fails_closed_and_records_error(self) -> None:
        client, _serial, events = self._client([b"@1 ACK 0000"])

        with self.assertRaisesRegex(HardwareError, "without CR terminator"):
            client.query_ascii("@1 0")

        self.assertEqual([event[0] for event in events], ["TX", "ERROR"])
        self.assertEqual(events[-1][1], "@1 ACK 0000")
        self.assertIn("@1 0", events[-1][2])

    def test_non_ascii_reply_fails_closed_and_records_error(self) -> None:
        client, _serial, events = self._client([b"@1 ACK \xff\r"])

        with self.assertRaisesRegex(HardwareError, "not valid ASCII"):
            client.query_ascii("@1 0")

        self.assertEqual([event[0] for event in events], ["TX", "ERROR"])
        self.assertIn("not valid ASCII", events[-1][2])

    def test_parse_failure_is_recorded_after_the_raw_reply(self) -> None:
        client, _serial, events = self._client([b"malformed\r"])
        axis = MotorAxisConfig("ChangerX", 1, 1)

        with self.assertRaisesRegex(HardwareError, "Unable to parse position"):
            client.read_position(axis)

        self.assertEqual([event[0] for event in events], ["TX", "RX", "ERROR"])
        self.assertEqual(events[1][1], "malformed\r")
        self.assertIn("position parse failed", events[-1][2])

    def test_connect_setup_failure_closes_partially_open_port(self) -> None:
        events: list[tuple[str, str, str]] = []

        class _SetupFailSerial:
            def __init__(self, **kwargs) -> None:
                self.kwargs = kwargs
                self.is_open = True
                self.closed = False
                self.dtr = False
                self.rts = False

            def reset_input_buffer(self) -> None:
                return None

            def reset_output_buffer(self) -> None:
                return None

            def write(self, payload: bytes) -> None:
                del payload
                raise OSError("write unavailable")

            def close(self) -> None:
                self.closed = True
                self.is_open = False

        created: list[_SetupFailSerial] = []

        def _serial_factory(**kwargs):
            serial = _SetupFailSerial(**kwargs)
            created.append(serial)
            return serial

        client = MotorSerialClient(
            trace=lambda direction, payload, detail: events.append(
                (direction, payload, detail)
            )
        )
        with mock.patch("rapidpy_common.hardware.serial.Serial", _serial_factory):
            with self.assertRaisesRegex(OSError, "write unavailable"):
                client.connect("COM3", 9600)

        self.assertTrue(created[0].closed)
        self.assertIsNone(client._serial)
        self.assertFalse(client.is_connected)
        self.assertEqual([event[0] for event in events], ["ERROR", "ERROR"])
        self.assertIn("motor setup", events[-1][2])


if __name__ == "__main__":
    unittest.main()
