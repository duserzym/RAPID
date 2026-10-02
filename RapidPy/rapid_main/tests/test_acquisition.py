"""Tests for the VB6-equivalent bracketed SQUID acquisition state machine."""
from __future__ import annotations

import unittest

from rapid_main.acquisition import (
    AcquisitionConfig,
    BlockContext,
    BracketedAcquisitionService,
    MotionVerificationError,
    RecoveryFailedError,
    TransportReadError,
)
from rapid_main.magnetometer import (
    FluxCountDiscontinuityError,
    ObservationIntegrityError,
    reduce_bracketed_measurement,
    validate_block_observations,
)

from tests.acquisition_fakes import (
    FakeClock,
    FakeSquidTransport,
    FakeTurning,
    FakeVertical,
    counts_for,
    counts_with_step,
)

# Archived one-flux-count increments in calibrated 2G raw units.
ARCHIVED_CALIBRATION = (0.09000563, 0.10600000, 0.06600000)

ZERO_POSITION = -25886
MEASUREMENT_POSITION = -30607


def _stable_observations(calibration=ARCHIVED_CALIBRATION):
    """Six coherent observations: two matching zeros around four positions."""

    return [
        counts_for((0.0, 0.0, 0.0), calibration),
        counts_for((1.0, 2.0, 3.0), calibration),
        counts_for((2.0, -1.0, 3.0), calibration),
        counts_for((-1.0, -2.0, 3.0), calibration),
        counts_for((-2.0, 1.0, 3.0), calibration),
        counts_for((0.0, 0.0, 0.0), calibration),
    ]


def _build_service(
    observations=None,
    *,
    calibration=ARCHIVED_CALIBRATION,
    transport_kwargs=None,
    vertical=None,
    turning=None,
    config_kwargs=None,
):
    clock = FakeClock()
    kwargs = dict(transport_kwargs or {})
    kwargs.setdefault("clock", clock)
    transport = FakeSquidTransport(
        observations if observations is not None else _stable_observations(calibration),
        **kwargs,
    )
    vertical = vertical or FakeVertical(start=0)
    turning = turning or FakeTurning()
    config = AcquisitionConfig(
        zero_position=ZERO_POSITION,
        measurement_position=MEASUREMENT_POSITION,
        axis_calibration=calibration,
        range_factor=1.0e-5,
        arc_delay_s=2.5,
        **(config_kwargs or {}),
    )
    service = BracketedAcquisitionService(
        transport,
        vertical,
        turning,
        config=config,
        clock=clock,
        id_factory=lambda prefix: f"{prefix}-test",
    )
    return service, transport, vertical, turning, clock


class BracketedAcquisitionSequenceTests(unittest.TestCase):
    def test_executes_the_vb6_command_order(self) -> None:
        service, transport, vertical, turning, clock = _build_service()

        acquisition = service.acquire(is_up=True)

        kinds = [event.kind for event in acquisition.commands]
        self.assertEqual(
            kinds[:6],
            [
                "turning.rotate",
                "vertical.move",
                "squid.set_range",
                "squid.clear_reset",
                "delay",
                "squid.latch",
            ],
        )
        # Lift order: zero -> measurement -> zero.
        self.assertEqual(
            [target for target, _speed in vertical.moves],
            [ZERO_POSITION, MEASUREMENT_POSITION, ZERO_POSITION],
        )
        # Turn order: reference, re-assert 0, 90, 180, 270, then close at 360.
        self.assertEqual(turning.angles, [0.0, 0.0, 90.0, 180.0, 270.0, 360.0])
        self.assertEqual(turning.references, [0.0])
        # Exactly six latch cycles: two zeros bracketing four orientations.
        self.assertEqual(transport.latch_count, 6)
        self.assertEqual(transport.reset_count, 1)

    def test_zero_before_uses_the_arc_delay_not_the_settling_latch(self) -> None:
        service, transport, _vertical, _turning, clock = _build_service()

        service.acquire()

        # VB6 latches the first zero with withDelay:=False because the ARC
        # delay after CLP/RC has already elapsed.
        self.assertIn(2.5, clock.slept)

    def test_observation_retains_counter_dvm_and_raw_replies(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()

        block = service.acquire().block
        zero_before = block.observations[0]

        self.assertEqual(zero_before.role, "zero-before")
        self.assertEqual([axis.axis for axis in zero_before.axes], ["X", "Y", "Z"])
        self.assertTrue(all(axis.count_command.endswith("SC") for axis in zero_before.axes))
        self.assertTrue(all(axis.data_command.endswith("SD") for axis in zero_before.axes))
        self.assertTrue(all(axis.data_reply for axis in zero_before.axes))
        self.assertEqual(zero_before.latch_commands, ("ALC", "ALD"))
        self.assertEqual(zero_before.vertical_position, ZERO_POSITION)

        position_1 = block.observations[1]
        self.assertEqual(position_1.vertical_position, MEASUREMENT_POSITION)
        self.assertEqual(position_1.turn_angle_deg, 0.0)
        self.assertEqual(block.observations[2].turn_angle_deg, 90.0)
        self.assertEqual(block.observations[3].turn_angle_deg, 180.0)
        self.assertEqual(block.observations[4].turn_angle_deg, 270.0)

    def test_applies_axis_calibration_exactly_once(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()

        block = service.acquire().block

        # getData already calibrated the observation, so reduction must not.
        self.assertEqual(block.axis_calibration, (1.0, 1.0, 1.0))
        self.assertEqual(block.audit.axis_calibration_applied, ARCHIVED_CALIBRATION)
        for axis, expected in zip(block.observations[1].axes, (1.0, 2.0, 3.0)):
            self.assertAlmostEqual(axis.calibrated_value, expected)
        self.assertAlmostEqual(block.positions[0][0], 1.0)
        self.assertAlmostEqual(block.positions[0][2], 3.0)

        result = reduce_bracketed_measurement(block)
        self.assertAlmostEqual(result.mean_raw[2], 3.0)
        self.assertAlmostEqual(result.moment_emu[2], 3.0e-5)

    def test_records_audit_identifiers_and_positions(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()

        block = service.acquire(
            context=BlockContext(
                sample_name="RS01A",
                treatment_label="AF20",
                run_id="run-7",
                operator="opr",
                software_version="rapidpy-test",
                config_hash="cfg-abc",
            )
        ).block
        audit = block.audit

        self.assertEqual(audit.block_id, "block-test")
        self.assertEqual(audit.sample_name, "RS01A")
        self.assertEqual(audit.treatment_label, "AF20")
        self.assertEqual(audit.run_id, "run-7")
        self.assertEqual(audit.zero_position, ZERO_POSITION)
        self.assertEqual(audit.measurement_position, MEASUREMENT_POSITION)
        self.assertEqual(audit.range_factor, 1.0e-5)
        self.assertTrue(audit.started_iso)
        self.assertTrue(audit.completed_iso)
        self.assertTrue(audit.commands)
        self.assertFalse(audit.simulated)
        self.assertFalse(block.simulated)

    def test_holder_block_switches_to_the_one_times_read_range(self) -> None:
        service, transport, _vertical, _turning, _clock = _build_service(
            config_kwargs={"range_label": "T", "holder_range_label": "1"}
        )

        block = service.acquire(context=BlockContext(is_holder_block=True)).block

        self.assertEqual(transport.ranges, [("A", "1")])
        self.assertEqual(block.audit.range_label, "1")
        self.assertTrue(block.audit.is_holder_block)

    def test_emitted_block_passes_observation_validation(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()

        block = service.acquire().block

        validate_block_observations(block)


class BracketedAcquisitionFaultTests(unittest.TestCase):
    def test_unverified_turn_aborts_the_block(self) -> None:
        service, _transport, _vertical, turning, _clock = _build_service(
            turning=FakeTurning(fail_at_angle=180.0)
        )

        with self.assertRaisesRegex(MotionVerificationError, "180"):
            service.acquire()

    def test_unverified_lift_aborts_the_block(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            vertical=FakeVertical(fail_at_target=MEASUREMENT_POSITION)
        )

        with self.assertRaisesRegex(MotionVerificationError, "measurement position"):
            service.acquire()

    def test_mismatched_latch_identity_is_rejected(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={"mismatch_axis": "Z"}
        )

        with self.assertRaisesRegex(TransportReadError, "belongs to latch"):
            service.acquire()

    def test_non_finite_axis_reading_is_rejected(self) -> None:
        observations = _stable_observations()
        observations[2] = ((0.0, float("nan")), (0.0, 0.0), (0.0, 0.0))
        service, _transport, _vertical, _turning, _clock = _build_service(observations)

        with self.assertRaisesRegex(TransportReadError, "not finite"):
            service.acquire()

    def test_empty_counter_reply_is_rejected_as_partial(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={"blank_reply_axis": "Y"}
        )

        with self.assertRaisesRegex(TransportReadError, "counter reply is empty"):
            service.acquire()

    def test_stale_latch_to_read_interval_is_rejected(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={"stale_seconds": 45.0},
            config_kwargs={"max_observation_age_s": 5.0},
        )

        with self.assertRaisesRegex(TransportReadError, "stale"):
            service.acquire()

    def test_transport_read_failure_propagates_without_a_block(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={
                "read_error": TimeoutError("SQUID read timed out"),
                "error_on_read_index": 4,
            }
        )

        with self.assertRaisesRegex(TransportReadError, "SQUID read timed out") as raised:
            service.acquire()
        self.assertIsInstance(raised.exception.__cause__, TimeoutError)

    def test_transport_recovery_returns_to_zero_resets_and_backs_off(self) -> None:
        service, transport, vertical, turning, clock = _build_service()

        record = service.recover_transport_failure(
            "zero-before: axis X read timed out",
            attempt=2,
            backoff_s=0.75,
        )

        self.assertEqual(turning.angles, [0.0])
        self.assertEqual(vertical.moves, [(ZERO_POSITION, 2)])
        self.assertEqual(transport.reset_count, 1)
        self.assertIn(3.25, clock.slept)
        self.assertEqual(record.attempt, 2)
        self.assertEqual(record.validation, None)
        self.assertIn("axis X", record.detail)
        self.assertEqual(
            [event.kind for event in record.commands],
            ["turning.rotate", "vertical.move", "squid.clear_reset", "delay"],
        )

    def test_transport_recovery_reset_failure_is_fatal(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={"reset_error": OSError("reset refused")}
        )

        with self.assertRaisesRegex(RecoveryFailedError, "reset refused"):
            service.recover_transport_failure("read timeout", backoff_s=0.25)


class FluxCountDiscontinuityTests(unittest.TestCase):
    def _acquire_with_step(self, axis: int):
        observations = _stable_observations()
        observations[5] = counts_with_step((0.0, 0.0, 0.0), ARCHIVED_CALIBRATION, axis=axis)
        service, _transport, _vertical, _turning, _clock = _build_service(observations)
        return service.acquire().block

    def test_x_axis_one_count_step_is_rejected(self) -> None:
        block = self._acquire_with_step(0)

        with self.assertRaises(FluxCountDiscontinuityError) as ctx:
            reduce_bracketed_measurement(block)
        self.assertEqual(ctx.exception.validation.discontinuous_axes, ("X",))
        self.assertEqual(ctx.exception.validation.flux_step_axes, ("X",))
        self.assertAlmostEqual(ctx.exception.validation.deltas[0], ARCHIVED_CALIBRATION[0])

    def test_y_axis_one_count_step_is_rejected(self) -> None:
        block = self._acquire_with_step(1)

        with self.assertRaises(FluxCountDiscontinuityError) as ctx:
            reduce_bracketed_measurement(block)
        self.assertEqual(ctx.exception.validation.discontinuous_axes, ("Y",))
        self.assertEqual(ctx.exception.validation.flux_step_axes, ("Y",))
        self.assertAlmostEqual(ctx.exception.validation.deltas[1], ARCHIVED_CALIBRATION[1])

    def test_z_axis_one_count_step_is_rejected(self) -> None:
        block = self._acquire_with_step(2)

        with self.assertRaises(FluxCountDiscontinuityError) as ctx:
            reduce_bracketed_measurement(block)
        self.assertEqual(ctx.exception.validation.discontinuous_axes, ("Z",))
        self.assertEqual(ctx.exception.validation.flux_step_axes, ("Z",))
        self.assertAlmostEqual(ctx.exception.validation.deltas[2], ARCHIVED_CALIBRATION[2])

    def test_recovery_returns_to_zero_and_resets_the_counter_path(self) -> None:
        service, transport, vertical, turning, _clock = _build_service()

        record = service.recover_flux_count_discontinuity(attempt=2)

        kinds = [event.kind for event in record.commands]
        self.assertEqual(kinds[0], "turning.rotate")
        self.assertEqual(kinds[1], "vertical.move")
        self.assertEqual(kinds[2], "squid.clear_reset")
        self.assertEqual(kinds[-1], "squid.latch")
        self.assertEqual(vertical.moves, [(ZERO_POSITION, 2)])
        self.assertEqual(turning.angles, [0.0])
        self.assertEqual(transport.reset_count, 1)
        self.assertEqual(record.attempt, 2)
        self.assertTrue(record.started_iso)
        self.assertTrue(record.completed_iso)

    def test_recovery_failure_is_reported_not_swallowed(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            transport_kwargs={"reset_error": OSError("counter reset refused")}
        )

        with self.assertRaisesRegex(RecoveryFailedError, "counter reset refused"):
            service.recover_flux_count_discontinuity()

    def test_recovery_motion_failure_halts_instead_of_continuing(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service(
            vertical=FakeVertical(fail_at_target=ZERO_POSITION)
        )

        with self.assertRaises(MotionVerificationError):
            service.recover_flux_count_discontinuity()


class ObservationEvidenceTests(unittest.TestCase):
    def test_block_observation_vector_mismatch_is_rejected(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()
        block = service.acquire().block
        tampered = type(block)(
            zero_before=block.zero_before,
            positions=((9.9, 9.9, 9.9), *block.positions[1:]),  # type: ignore[arg-type]
            zero_after=block.zero_after,
            holder_positions=block.holder_positions,
            is_up=block.is_up,
            range_factor=block.range_factor,
            axis_calibration=block.axis_calibration,
            observations=block.observations,
            audit=block.audit,
        )

        with self.assertRaisesRegex(ObservationIntegrityError, "does not match"):
            reduce_bracketed_measurement(tampered)

    def test_partial_observation_set_is_rejected(self) -> None:
        service, _transport, _vertical, _turning, _clock = _build_service()
        block = service.acquire().block
        truncated = type(block)(
            zero_before=block.zero_before,
            positions=block.positions,
            zero_after=block.zero_after,
            observations=block.observations[:4],
            audit=block.audit,
        )

        with self.assertRaisesRegex(ObservationIntegrityError, "needs 6 observations"):
            reduce_bracketed_measurement(truncated)


if __name__ == "__main__":
    unittest.main(verbosity=2)
