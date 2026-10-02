from __future__ import annotations

import unittest

from rapid_main.communication_log import CommunicationDirection
from rapid_main.susceptibility_transport import (
    SusceptibilitySerialClient,
    SusceptibilityTransportConfig,
    SusceptibilityTransportError,
)


class _Clock:
    def __init__(self) -> None:
        self.value = 0.0

    def __call__(self) -> float:
        self.value += 0.01
        return self.value


class _Serial:
    def __init__(self, responses: list[bytes], **kwargs) -> None:
        self.responses = list(responses)
        self.kwargs = kwargs
        self.is_open = True
        self.writes: list[bytes] = []
        self.reset_input_calls = 0
        self.reset_output_calls = 0
        self.closed = False

    def reset_input_buffer(self) -> None:
        self.reset_input_calls += 1

    def reset_output_buffer(self) -> None:
        self.reset_output_calls += 1

    def write(self, payload: bytes) -> int:
        self.writes.append(payload)
        return len(payload)

    def flush(self) -> None:
        return

    def read(self, size: int = 1) -> bytes:
        del size
        if not self.responses:
            return b""
        chunk = self.responses[0]
        if len(chunk) == 1:
            self.responses.pop(0)
            return chunk
        self.responses[0] = chunk[1:]
        return chunk[:1]

    def close(self) -> None:
        self.closed = True
        self.is_open = False


class TestSusceptibilitySerialClient(unittest.TestCase):
    def _client(self, responses: list[bytes], **config_overrides):
        holder: dict[str, _Serial] = {}

        def factory(**kwargs):
            holder["serial"] = _Serial(responses, **kwargs)
            return holder["serial"]

        config = SusceptibilityTransportConfig(
            port="COM7",
            response_timeout_s=0.2,
            **config_overrides,
        )
        client = SusceptibilitySerialClient(config, serial_factory=factory, clock=_Clock())
        client.connect()
        return client, holder["serial"]

    def test_zero_uses_exact_vb6_command_and_cr_reply(self) -> None:
        client, transport = self._client([b"O", b"K", b"\r"])

        response = client.zero()

        self.assertEqual(response, "OK\r")
        self.assertEqual(transport.writes, [b"Z\r\n"])
        events = client.communication_events()
        self.assertEqual(
            [event.direction for event in events],
            [CommunicationDirection.INFO, CommunicationDirection.TX, CommunicationDirection.RX],
        )
        self.assertEqual(events[1].payload, "Z\\r\\n")
        self.assertEqual(events[2].payload, "OK\\r")

    def test_measure_applies_configured_scale_factor(self) -> None:
        client, transport = self._client([b"1", b"2", b".", b"5", b"\r"], scale_factor=0.1)

        value = client.measure()

        self.assertAlmostEqual(value, 1.25)
        self.assertEqual(transport.writes, [b"M\r\n"])

    def test_empty_response_fails_closed_and_records_error(self) -> None:
        client, _transport = self._client([])

        with self.assertRaisesRegex(SusceptibilityTransportError, "empty response"):
            client.measure()

        self.assertEqual(client.communication_events()[-1].direction, CommunicationDirection.ERROR)

    def test_unterminated_partial_response_fails_closed(self) -> None:
        client, _transport = self._client([b"1", b"2"])

        with self.assertRaisesRegex(SusceptibilityTransportError, "unterminated partial"):
            client.measure()

    def test_non_ascii_response_fails_closed(self) -> None:
        client, _transport = self._client([b"\xff", b"\r"])

        with self.assertRaisesRegex(SusceptibilityTransportError, "non-ASCII"):
            client.measure()

    def test_malformed_and_nonfinite_measurements_fail_closed(self) -> None:
        malformed, _ = self._client([b"N", b"O", b"P", b"E", b"\r"])
        with self.assertRaisesRegex(SusceptibilityTransportError, "non-numeric"):
            malformed.measure()

        nonfinite, _ = self._client([b"n", b"a", b"n", b"\r"])
        with self.assertRaisesRegex(SusceptibilityTransportError, "non-finite"):
            nonfinite.measure()

    def test_close_is_idempotent(self) -> None:
        client, transport = self._client([])

        client.close()
        client.close()

        self.assertTrue(transport.closed)
        self.assertFalse(client.is_connected)

    def test_invalid_configuration_is_rejected_before_open(self) -> None:
        with self.assertRaisesRegex(SusceptibilityTransportError, "port is not configured"):
            SusceptibilitySerialClient(SusceptibilityTransportConfig(port=""))
        with self.assertRaisesRegex(SusceptibilityTransportError, "scale factor"):
            SusceptibilitySerialClient(
                SusceptibilityTransportConfig(port="COM7", scale_factor=0.0)
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
