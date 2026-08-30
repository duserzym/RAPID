from __future__ import annotations

import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

from rapid_main.communication_log import (
    CommunicationDirection,
    CommunicationLogger,
    LoggingTransportBridge,
    normalize_payload,
)


def _clock() -> datetime:
    return datetime(2026, 7, 15, 12, 30, 5, 123000, tzinfo=timezone.utc)


class FakeTransport:
    def __init__(self) -> None:
        self.connected: tuple[str, int] | None = None
        self.writes: list[bytes] = []
        self.extra_value = "delegated"

    def connect(self, port: str, baudrate: int = 9600) -> None:
        self.connected = (port, baudrate)

    def disconnect(self) -> None:
        self.connected = None

    def write(self, payload: bytes) -> int:
        self.writes.append(payload)
        return len(payload)

    def read(self) -> bytes:
        return b"OK\r\n"


class FailingTransport(FakeTransport):
    def read(self) -> bytes:
        raise RuntimeError("serial timeout")


class CommunicationLogTests(unittest.TestCase):
    def test_payload_normalization_is_single_line_and_bounded(self) -> None:
        self.assertEqual(normalize_payload(b"MEAS?\r\n"), "MEAS?\\r\\n")
        self.assertEqual(normalize_payload("A\tB\nC"), "A\\tB\\nC")
        self.assertEqual(normalize_payload("abcdef", max_chars=3), "abc...<truncated 3 chars>")

    def test_logger_records_exportable_transcript_lines(self) -> None:
        logger = CommunicationLogger("SQUID", port="COM4", clock=_clock)

        logger.sent(b"READ?\r\n")
        logger.received("1.0,2.0,3.0")
        logger.error("checksum mismatch", payload="BAD")

        lines = logger.transcript.lines()

        self.assertEqual(lines[0], "timestamp\tchannel\tdirection\tport\tpayload\tdetail")
        self.assertIn("2026-07-15T12:30:05.123+00:00\tSQUID\tTX\tCOM4\tREAD?\\r\\n\t", lines[1])
        self.assertIn("\tRX\tCOM4\t1.0,2.0,3.0\t", lines[2])
        self.assertIn("\tERROR\tCOM4\tBAD\tchecksum mismatch", lines[3])

    def test_logger_writes_transcript_file(self) -> None:
        logger = CommunicationLogger("VACUUM", port="COM2", clock=_clock)
        logger.info("connect")

        with tempfile.TemporaryDirectory() as td:
            path = logger.write_text(Path(td) / "comm.tsv")
            text = path.read_text(encoding="utf-8")

        self.assertIn("timestamp\tchannel\tdirection\tport\tpayload\tdetail", text)
        self.assertIn("\tVACUUM\tINFO\tCOM2\t\tconnect", text)

    def test_transport_bridge_logs_connect_write_and_read_without_changing_results(self) -> None:
        transport = FakeTransport()
        logger = CommunicationLogger("MOTOR", clock=_clock)
        bridge = LoggingTransportBridge(transport, logger)

        bridge.connect("COM3", baudrate=19200)
        written = bridge.write(b"TPOS\r")
        response = bridge.read()

        self.assertEqual(transport.connected, ("COM3", 19200))
        self.assertEqual(written, 5)
        self.assertEqual(response, b"OK\r\n")
        self.assertEqual(bridge.extra_value, "delegated")
        self.assertEqual(
            [event.direction for event in logger.transcript.events],
            [CommunicationDirection.INFO, CommunicationDirection.TX, CommunicationDirection.RX],
        )
        self.assertEqual(logger.transcript.events[0].port, "COM3")

    def test_transport_bridge_logs_errors_and_reraises(self) -> None:
        logger = CommunicationLogger("SQUID", port="COM1", clock=_clock)
        bridge = LoggingTransportBridge(FailingTransport(), logger)

        with self.assertRaisesRegex(RuntimeError, "serial timeout"):
            bridge.read()

        self.assertEqual(logger.transcript.events[-1].direction, CommunicationDirection.ERROR)
        self.assertIn("read failed: serial timeout", logger.transcript.events[-1].detail)


if __name__ == "__main__":
    unittest.main(verbosity=2)
