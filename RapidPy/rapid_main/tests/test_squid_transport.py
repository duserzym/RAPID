"""Tests for the live 2G transport, motion adapters, and backend facade."""
from __future__ import annotations

import unittest
from datetime import datetime, timezone

from rapid_main.acquisition import (
    AcquisitionConfig,
    AcquisitionError,
    BlockContext,
    BracketedAcquisitionService,
    MotionOutcome,
    RecoveryFailedError,
    RecoveryRecord,
    TransportReadError,
)
from rapid_main.communication_log import CommunicationDirection, CommunicationLogger
from rapid_main.config import AppConfig, CalibrationConfig, SquidConfig
from rapid_main.magnetometer import ZeroPairValidation
from rapid_main.squid_transport import (
    BracketedSquidBackend,
    RawSquidTransport,
    SquidTransportConfig,
    SquidTransportError,
    _range_label,
    acquisition_config_from_app_config,
    wrap_turning_position,
)

from tests.acquisition_fakes import (
    FakeClock,
    FakeSquidTransport,
    FakeTurning,
    FakeVertical,
    counts_for,
)

CALIBRATION = (0.09000563, 0.106, 0.066)
TURNING_FULL_ROTATION = -172_800


class _FakeSquidClient:
    """Stand-in for ``updown_control.app.RawSquidClient``."""

    def __init__(self, samples, *, connected: bool = True) -> None:
        self.is_connected = connected
        self._samples = list(samples)
        self.calls: list[tuple[str, object]] = []
        self._index = -1

    def clear_and_reset(self, axis: str = "A") -> tuple[str, ...]:
        self.calls.append(("clear_and_reset", axis))
        return (f"{axis}CLP", f"{axis}RC")

    def set_range(self, axis: str = "A", range_label: str = "1") -> tuple[str, ...]:
        self.calls.append(("set_range", (axis, range_label)))
        return (f"{axis}CR{range_label}",)

    def latch(self, axis: str = "A", *, settle_s: float = 0.0) -> tuple[str, ...]:
        self.calls.append(("latch", (axis, settle_s)))
        self._index += 1
        return (f"{axis}LC", f"{axis}LD")

    def read_axis(self, axis: str, *, range_value: float = 1.0):
        self.calls.append(("read_axis", (axis, range_value)))
        counts, dvm = self._samples[self._index]["XYZ".index(axis)]

        class _Sample:
            def __init__(self) -> None:
                self.axis = axis
                self.counts = counts
                self.dvm = dvm
                self.range_value = range_value
                self.count_command = f"{axis}SC"
                self.count_reply = f"{counts:g}"
                self.data_command = f"{axis}SD"
                self.data_reply = f"{dvm:g}"

        return _Sample()


class RawSquidTransportTests(unittest.TestCase):
    def _transport(self, *, connected: bool = True) -> RawSquidTransport:
        client = _FakeSquidClient(
            [counts_for((1.0, 2.0, 3.0), CALIBRATION)] * 6, connected=connected
        )
        return RawSquidTransport(
            client, config=SquidTransportConfig(port="COM1", baud=1200, settle_delay_s=1.25)
        )

    def test_latch_mints_an_identity_and_stamps_axis_replies(self) -> None:
        transport = self._transport()

        latch = transport.latch("A", settle=True)
        reply = transport.read_axis("X")

        self.assertTrue(latch.latch_id)
        self.assertEqual(latch.commands, ("ALC", "ALD"))
        self.assertEqual(reply.latch_id, latch.latch_id)
        self.assertEqual(reply.count_command, "XSC")
        self.assertEqual(reply.data_command, "XSD")
        self.assertIn(("latch", ("A", 1.25)), transport.client.calls)

    def test_second_latch_produces_a_different_identity(self) -> None:
        transport = self._transport()

        first = transport.latch("A")
        second = transport.latch("A")

        self.assertNotEqual(first.latch_id, second.latch_id)

    def test_read_before_latch_is_refused(self) -> None:
        transport = self._transport()

        with self.assertRaisesRegex(SquidTransportError, "before a latch"):
            transport.read_axis("X")

    def test_disconnected_transport_never_returns_a_value(self) -> None:
        transport = self._transport(connected=False)

        with self.assertRaisesRegex(SquidTransportError, "not connected"):
            transport.latch("A")
        with self.assertRaisesRegex(SquidTransportError, "not connected"):
            transport.clear_and_reset_counters("A")

    def test_drives_a_complete_block_through_the_acquisition_service(self) -> None:
        client = _FakeSquidClient(
            [
                counts_for((0.0, 0.0, 0.0), CALIBRATION),
                counts_for((1.0, 2.0, 3.0), CALIBRATION),
                counts_for((2.0, -1.0, 3.0), CALIBRATION),
                counts_for((-1.0, -2.0, 3.0), CALIBRATION),
                counts_for((-2.0, 1.0, 3.0), CALIBRATION),
                counts_for((0.0, 0.0, 0.0), CALIBRATION),
            ]
        )
        transport = RawSquidTransport(client, config=SquidTransportConfig(port="COM1"))
        service = BracketedAcquisitionService(
            transport,
            FakeVertical(),
            FakeTurning(),
            config=AcquisitionConfig(
                zero_position=-25886,
                measurement_position=-30607,
                axis_calibration=CALIBRATION,
                arc_delay_s=0.0,
            ),
            clock=FakeClock(),
        )

        block = service.acquire().block

        self.assertEqual(len({obs.latch_id for obs in block.observations}), 6)
        self.assertAlmostEqual(block.positions[0][2], 3.0)
        self.assertEqual(("set_range", ("A", "1")), client.calls[0])

    def test_records_exact_commands_and_replies_in_transport_order(self) -> None:
        client = _FakeSquidClient(
            [((0.0, -1.0), (0.0, -2.0), (0.0, -3.0))]
        )
        logger = CommunicationLogger(
            "SQUID-2G",
            port="COM7",
            clock=lambda: datetime(2026, 10, 2, tzinfo=timezone.utc),
        )
        transport = RawSquidTransport(
            client,
            config=SquidTransportConfig(port="COM7"),
            communication_logger=logger,
        )

        transport.clear_and_reset_counters("A")
        transport.set_range("A", "H")
        latch = transport.latch("A")
        for axis in "XYZ":
            transport.read_axis(axis)

        events = transport.communication_events()
        self.assertIsInstance(events, tuple)
        self.assertEqual(
            [(event.direction.value, event.payload) for event in events],
            [
                ("TX", "ACLP"),
                ("TX", "ARC"),
                ("TX", "ACRH"),
                ("TX", "ALC"),
                ("TX", "ALD"),
                ("TX", "XSC"),
                ("RX", "0"),
                ("TX", "XSD"),
                ("RX", "-1"),
                ("TX", "YSC"),
                ("RX", "0"),
                ("TX", "YSD"),
                ("RX", "-2"),
                ("TX", "ZSC"),
                ("RX", "0"),
                ("TX", "ZSD"),
                ("RX", "-3"),
            ],
        )
        self.assertTrue(all(event.port == "COM7" for event in events))
        self.assertTrue(all(latch.latch_id in event.detail for event in events[3:]))
        with self.assertRaises(AttributeError):
            events.append(events[0])  # type: ignore[attr-defined]

    def test_adapter_error_is_recorded_and_original_exception_propagates(self) -> None:
        failure = RuntimeError("serial read failed")

        class _FailingClient(_FakeSquidClient):
            def read_axis(self, axis: str, *, range_value: float = 1.0):
                raise failure

        client = _FailingClient([counts_for((1.0, 2.0, 3.0), CALIBRATION)])
        transport = RawSquidTransport(client, config=SquidTransportConfig(port="COM9"))
        transport.latch("A")

        with self.assertRaises(RuntimeError) as raised:
            transport.read_axis("X")

        self.assertIs(raised.exception, failure)
        event = transport.communication_events()[-1]
        self.assertEqual(event.direction, CommunicationDirection.ERROR)
        self.assertEqual(event.payload, "X")
        self.assertIn("serial read failed", event.detail)


class TurningPositionWrapTests(unittest.TestCase):
    def test_full_negative_rotation_wraps_to_zero(self) -> None:
        # After MotorTurn_360 the encoder sits at -full_rotation.
        self.assertEqual(wrap_turning_position(-TURNING_FULL_ROTATION, TURNING_FULL_ROTATION), 0)

    def test_position_already_inside_the_window_is_unchanged(self) -> None:
        self.assertEqual(wrap_turning_position(120, TURNING_FULL_ROTATION), 120)

    def test_positive_full_rotation_convention_is_supported(self) -> None:
        self.assertEqual(wrap_turning_position(-172_800, 172_800), 0)

    def test_zero_full_rotation_is_rejected(self) -> None:
        with self.assertRaises(Exception):
            wrap_turning_position(10, 0)


class RangeLabelTests(unittest.TestCase):
    def test_maps_ui_labels_onto_2g_control_rates(self) -> None:
        self.assertEqual(_range_label("1x"), "1")
        self.assertEqual(_range_label("10x"), "T")
        self.assertEqual(_range_label("100x"), "H")
        self.assertEqual(_range_label("1000x"), "E")
        self.assertEqual(_range_label("Flux counting"), "F")
        self.assertEqual(_range_label(""), "1")

    def test_acquisition_config_reads_calibration_and_squid_settings(self) -> None:
        config = AppConfig(
            squid=SquidConfig(range_label="100x", settle_time=2.25),
            calibration=CalibrationConfig(cal_x=2.0, cal_y=3.0, cal_z=4.0, range_factor=1e-6),
        )

        acquisition = acquisition_config_from_app_config(
            config, zero_position=-25886, measurement_position=-30607
        )

        self.assertEqual(acquisition.axis_calibration, (2.0, 3.0, 4.0))
        self.assertEqual(acquisition.range_factor, 1e-6)
        self.assertEqual(acquisition.range_label, "H")
        self.assertEqual(acquisition.holder_range_label, "1")
        self.assertEqual(acquisition.settle_delay_s, 2.25)


class BracketedSquidBackendTests(unittest.TestCase):
    def _backend(self, **kwargs) -> tuple[BracketedSquidBackend, FakeVertical, FakeTurning]:
        observations = [
            counts_for((0.0, 0.0, 0.0), CALIBRATION),
            counts_for((1.0, 2.0, 3.0), CALIBRATION),
            counts_for((2.0, -1.0, 3.0), CALIBRATION),
            counts_for((-1.0, -2.0, 3.0), CALIBRATION),
            counts_for((-2.0, 1.0, 3.0), CALIBRATION),
            counts_for((0.0, 0.0, 0.0), CALIBRATION),
        ]
        clock = FakeClock()
        vertical = FakeVertical()
        turning = FakeTurning()
        service = BracketedAcquisitionService(
            FakeSquidTransport(observations, clock=clock),
            vertical,
            turning,
            config=AcquisitionConfig(
                zero_position=-25886,
                measurement_position=-30607,
                axis_calibration=CALIBRATION,
                arc_delay_s=0.0,
            ),
            clock=clock,
        )
        return BracketedSquidBackend(service, **kwargs), vertical, turning

    def test_read_squid_returns_a_whole_block_with_injected_holder(self) -> None:
        holder = (
            (0.01, 0.0, 0.0),
            (0.0, 0.01, 0.0),
            (-0.01, 0.0, 0.0),
            (0.0, -0.01, 0.0),
        )
        backend, _vertical, _turning = self._backend(
            holder_provider=lambda: holder,
            direction_provider=lambda: False,
            context_provider=lambda: BlockContext(sample_name="RS01A", run_id="run-9"),
        )

        block = backend.read_squid()

        self.assertEqual(block.holder_positions, holder)
        self.assertFalse(block.is_up)
        self.assertEqual(block.audit.sample_name, "RS01A")
        self.assertEqual(block.audit.run_id, "run-9")
        self.assertIsNotNone(backend.last_acquisition)

    def test_simulated_backend_marks_every_block(self) -> None:
        backend, _vertical, _turning = self._backend(simulated=True)

        block = backend.read_squid()

        self.assertTrue(backend.simulated)
        self.assertTrue(block.simulated)
        self.assertTrue(block.audit.simulated)

    def test_recovery_records_each_attempt(self) -> None:
        backend, vertical, turning = self._backend()
        validation = ZeroPairValidation(
            valid=False,
            deltas=(0.09, 0.0, 0.0),
            discontinuous_axes=("X",),
            flux_step_axes=("X",),
            reason="discontinuous bracketing zeros",
        )

        first = backend.recover_flux_count_discontinuity(validation)
        second = backend.recover_flux_count_discontinuity(validation)

        self.assertEqual(first.attempt, 1)
        self.assertEqual(second.attempt, 2)
        self.assertEqual(len(backend.recovery_records), 2)
        self.assertEqual(first.detail, "discontinuous bracketing zeros")
        self.assertEqual(vertical.moves, [(-25886, 2), (-25886, 2)])
        self.assertEqual(turning.angles, [0.0, 0.0])

    def test_return_to_safe_state_lifts_and_squares_the_turning_axis(self) -> None:
        backend, vertical, turning = self._backend()

        backend.return_to_safe_state()

        self.assertEqual(vertical.moves, [(-25886, 2)])
        self.assertEqual(turning.angles, [0.0])

    def test_communication_event_snapshot_is_forwarded_immutably(self) -> None:
        logger = CommunicationLogger("SQUID-2G", port="COM4")
        logger.sent("ALC", detail="latch")
        backend, _vertical, _turning = self._backend(
            communication_events_provider=lambda: tuple(logger.transcript.events)
        )

        events = backend.communication_events()

        self.assertIsInstance(events, tuple)
        self.assertEqual(events[0].payload, "ALC")
        logger.sent("ALD", detail="latch")
        self.assertEqual(len(events), 1)
        self.assertEqual(len(backend.communication_events()), 2)

    def test_simulated_backend_cannot_publish_events_as_live_transport_evidence(self) -> None:
        logger = CommunicationLogger("SQUID-2G", port="SIMULATED")
        logger.sent("ALC", detail="synthetic latch")
        backend, _vertical, _turning = self._backend(
            simulated=True,
            communication_events_provider=lambda: tuple(logger.transcript.events),
        )

        self.assertEqual(backend.communication_events(), ())

    def test_transport_failure_recovers_then_restarts_the_whole_block(self) -> None:
        class _Service:
            def __init__(self) -> None:
                self.acquire_calls = 0
                self.recoveries: list[tuple[str, int, float]] = []

            def acquire(self, **_kwargs):
                self.acquire_calls += 1
                if self.acquire_calls == 1:
                    raise TransportReadError("position-2 axis Y timed out")
                return type("Acquisition", (), {"block": "fresh-block"})()

            def recover_transport_failure(self, detail, *, attempt, backoff_s):
                self.recoveries.append((detail, attempt, backoff_s))
                return RecoveryRecord(
                    attempt=attempt,
                    started_iso="2026-10-02T00:00:00+00:00",
                    completed_iso="2026-10-02T00:00:01+00:00",
                    validation=None,
                    commands=(),
                    detail=detail,
                )

        service = _Service()
        backend = BracketedSquidBackend(  # type: ignore[arg-type]
            service,
            transport_retries=2,
            transport_retry_backoff_s=0.5,
        )

        block = backend.read_squid()

        self.assertEqual(block, "fresh-block")
        self.assertEqual(service.acquire_calls, 2)
        self.assertEqual(service.recoveries, [("position-2 axis Y timed out", 1, 0.5)])
        self.assertEqual(len(backend.transport_recovery_records), 1)

    def test_transport_retry_exhaustion_never_returns_a_block(self) -> None:
        class _Service:
            def __init__(self) -> None:
                self.acquire_calls = 0
                self.recovery_calls = 0
                self.backoffs: list[float] = []

            def acquire(self, **_kwargs):
                self.acquire_calls += 1
                raise TransportReadError("persistent timeout")

            def recover_transport_failure(self, detail, *, attempt, backoff_s):
                self.recovery_calls += 1
                self.backoffs.append(backoff_s)
                return RecoveryRecord(
                    attempt=attempt,
                    started_iso="start",
                    completed_iso="complete",
                    validation=None,
                    commands=(),
                    detail=detail,
                )

        service = _Service()
        backend = BracketedSquidBackend(service, transport_retries=2)  # type: ignore[arg-type]

        with self.assertRaisesRegex(TransportReadError, "persistent timeout"):
            backend.read_squid()

        self.assertEqual(service.acquire_calls, 3)
        self.assertEqual(service.recovery_calls, 2)
        self.assertEqual(service.backoffs, [0.25, 0.5])
        self.assertIsNone(backend.last_acquisition)

    def test_transport_recovery_failure_stops_without_another_read(self) -> None:
        class _Service:
            def __init__(self) -> None:
                self.acquire_calls = 0

            def acquire(self, **_kwargs):
                self.acquire_calls += 1
                raise TransportReadError("timeout")

            def recover_transport_failure(self, detail, *, attempt, backoff_s):
                del detail, attempt, backoff_s
                raise RecoveryFailedError("counter reset refused")

        service = _Service()
        backend = BracketedSquidBackend(service, transport_retries=2)  # type: ignore[arg-type]

        with self.assertRaisesRegex(RecoveryFailedError, "counter reset refused"):
            backend.read_squid()

        self.assertEqual(service.acquire_calls, 1)
        self.assertEqual(backend.transport_recovery_records, ())

    def test_motion_or_other_acquisition_failure_is_not_retried(self) -> None:
        class _Service:
            def __init__(self) -> None:
                self.acquire_calls = 0

            def acquire(self, **_kwargs):
                self.acquire_calls += 1
                raise AcquisitionError("turning verification failed")

        service = _Service()
        backend = BracketedSquidBackend(service, transport_retries=2)  # type: ignore[arg-type]

        with self.assertRaisesRegex(AcquisitionError, "turning verification failed"):
            backend.read_squid()

        self.assertEqual(service.acquire_calls, 1)


class MotionOutcomeTests(unittest.TestCase):
    def test_motion_outcome_defaults_are_explicit(self) -> None:
        outcome = MotionOutcome(target=1.0, actual=1.0, ok=True)

        self.assertEqual(outcome.detail, "")


class _FakeSerialPort:
    """Minimal serial stand-in that answers queued replies one byte at a time."""

    def __init__(self, replies: list[str]) -> None:
        self.is_open = True
        self.writes: list[str] = []
        self.input_resets = 0
        self._replies = list(replies)
        self._buffer = b""

    def reset_input_buffer(self) -> None:
        self.input_resets += 1
        self._buffer = b""

    def write(self, payload: bytes) -> int:
        text = payload.decode("ascii").strip("\r")
        if text:
            self.writes.append(text)
            if text.endswith(("SC", "SD")) and self._replies:
                self._buffer += (self._replies.pop(0) + "\r").encode("ascii")
        return len(payload)

    def flush(self) -> None:
        return None

    def read(self, size: int = 1) -> bytes:
        if not self._buffer:
            return b""
        chunk, self._buffer = self._buffer[:size], self._buffer[size:]
        return chunk


class RawSquidClientAtomicCommandTests(unittest.TestCase):
    """The extended client must expose VB6's atomic 2G operations."""

    def _client(self, replies: list[str]):
        from updown_control.app import RawSquidClient

        client = RawSquidClient()
        client._serial = _FakeSerialPort(replies)  # type: ignore[attr-defined]
        return client

    def test_read_axis_queries_counter_before_dvm_and_keeps_both(self) -> None:
        client = self._client(["1", "-0.5"])

        sample = client.read_axis("X")

        self.assertEqual(client._serial.writes, ["XSC", "XSD"])  # type: ignore[attr-defined]
        self.assertEqual(sample.counts, 1.0)
        self.assertEqual(sample.dvm, -0.5)
        self.assertEqual(sample.count_command, "XSC")
        self.assertEqual(sample.data_command, "XSD")
        self.assertEqual(sample.count_reply, "1")
        self.assertEqual(sample.data_reply, "-0.5")
        # VB6 getVal: -data - count * range
        self.assertAlmostEqual(sample.raw_value, 0.5 - 1.0)
        self.assertEqual(client._serial.input_resets, 2)  # type: ignore[attr-defined]

    def test_read_axis_discards_stale_input_before_each_query(self) -> None:
        client = self._client(["1", "-0.5"])
        client._serial._buffer = b"999\r"  # type: ignore[attr-defined]

        sample = client.read_axis("X")

        self.assertEqual(sample.count_reply, "1")
        self.assertEqual(sample.data_reply, "-0.5")
        self.assertEqual(client._serial.input_resets, 2)  # type: ignore[attr-defined]

    def test_unterminated_partial_response_is_rejected(self) -> None:
        client = self._client([])
        client._serial._buffer = b"12.5"  # type: ignore[attr-defined]

        with self.assertRaisesRegex(Exception, "terminated SQUID response"):
            client._read_response(timeout_s=0.01)

    def test_latch_issues_latch_count_then_latch_data(self) -> None:
        client = self._client([])

        commands = client.latch("A")

        self.assertEqual(commands, ("ALC", "ALD"))
        self.assertEqual(client._serial.writes, ["ALC", "ALD"])  # type: ignore[attr-defined]

    def test_clear_and_reset_matches_clp_then_rc(self) -> None:
        client = self._client([])

        commands = client.clear_and_reset("A")

        self.assertEqual(commands, ("ACLP", "ARC"))
        self.assertEqual(client._serial.writes, ["ACLP", "ARC"])  # type: ignore[attr-defined]

    def test_set_range_flux_mode_enables_fast_slew_first(self) -> None:
        client = self._client([])

        self.assertEqual(client.set_range("A", "F"), ("ACSE", "ACR1"))
        self.assertEqual(client.set_range("A", "H"), ("ACRH",))
        with self.assertRaises(Exception):
            client.set_range("A", "Q")

    def test_read_xyz_raw_still_returns_combined_values(self) -> None:
        client = self._client(["0", "-1", "0", "-2", "0", "-3"])

        values = client.read_xyz_raw()

        self.assertEqual(values, (1.0, 2.0, 3.0))
        self.assertEqual(
            client._serial.writes,  # type: ignore[attr-defined]
            ["ALC", "ALD", "XSC", "XSD", "YSC", "YSD", "ZSC", "ZSD"],
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
