"""Bracketed SQUID acquisition state machine mapped from VB6 ``Measure_ReadSample``.

The legacy application acquires one *block* per averaging cycle:

1. turn to the reference orientation and lift to the zero position;
2. ``CLP``/``RC`` the 2G counters, settle, and latch the **zero-before** read;
3. lower to the measurement position;
4. latch and read four specimen/holder orientations at 0/90/180/270 degrees,
   verifying each turn before the read;
5. lift back to the zero position, complete the 360-degree turn, and latch the
   **zero-after** read.

Only then is the block validated (:func:`rapid_main.magnetometer.validate_zero_pair`)
and reduced.  This module reproduces that sequence behind narrow protocols so
every state transition is deterministic in tests and no changer, lift, turn,
SQUID, or logging operation has to live in UI code.

Design notes that matter for parity and safety
----------------------------------------------

* Axis calibration is applied **once**, at observation build time, exactly like
  VB6 ``frmSQUID.getData`` calling ``Calibrate``.  The emitted
  :class:`~rapid_main.magnetometer.BracketedMeasurementBlock` therefore carries
  ``axis_calibration=(1, 1, 1)`` so :func:`reduce_bracketed_measurement` cannot
  apply it a second time; the applied factors are preserved in the audit record.
* VB6 issues each motion twice (a non-blocking start, then a blocking repeat).
  RapidPy issues one blocking, verified motion instead.  That is a deliberate
  difference: it removes the un-awaited command while keeping the same physical
  order, and it fails loudly when a motion does not reach its target.
* No path in this module fabricates a reading.  A transport that cannot answer
  raises; it never substitutes a synthetic value.
"""
from __future__ import annotations

from contextlib import nullcontext
from dataclasses import dataclass, field, replace
from datetime import datetime, timezone
import itertools
import math
import time
from typing import Any, Callable, Iterable, Protocol, runtime_checkable

from rapid_main.magnetometer import (
    AXIS_NAMES,
    AxisObservation,
    BlockAudit,
    BracketedMeasurementBlock,
    CommandEvent,
    DEFAULT_MAX_OBSERVATION_AGE_S,
    ObservationIntegrityError,
    Position4,
    SquidObservation,
    Vector3,
    ZeroPairValidation,
)

ZERO_HOLDER: Position4 = (
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
)

#: VB6 ``Measure_ReadSample`` reads position ``j`` at ``(j - 1) * 90`` degrees and
#: closes the block with a full 360-degree turn back to the start orientation.
POSITION_ANGLES_DEG: tuple[float, float, float, float] = (0.0, 90.0, 180.0, 270.0)
BLOCK_CLOSING_ANGLE_DEG = 360.0

#: VB6 ``Const Measure_ARCDelay = 2.5`` seconds after an ARC SQUID box command.
DEFAULT_ARC_DELAY_S = 2.5


class AcquisitionError(RuntimeError):
    """Raised when a bracketed acquisition cannot be completed safely."""


class MotionVerificationError(AcquisitionError):
    """Raised when a lift or turn did not reach and hold its commanded target."""


class TransportReadError(AcquisitionError):
    """Raised when the SQUID transport cannot deliver a coherent observation."""


class RecoveryFailedError(AcquisitionError):
    """Raised when the 2G flux-count recovery sequence itself failed."""


@dataclass(frozen=True)
class MotionOutcome:
    """Result of one verified motion command."""

    target: float
    actual: float
    ok: bool
    detail: str = ""


@dataclass(frozen=True)
class LatchResult:
    """One ``ALC``/``ALD`` latch cycle."""

    latch_id: str
    commands: tuple[str, ...] = ()


@dataclass(frozen=True)
class AxisReply:
    """Raw counter and DVM replies for one axis of a latched read."""

    counts: float
    dvm: float
    latch_id: str = ""
    range_value: float = 1.0
    count_command: str = ""
    count_reply: str = ""
    data_command: str = ""
    data_reply: str = ""


@runtime_checkable
class SquidTransport(Protocol):
    """Atomic 2G 581 operations required by the bracketed acquisition."""

    def clear_and_reset_counters(self, axis: str = "A") -> None:
        """Issue VB6 ``CLP``/``ResetCount`` for ``axis``."""

    def set_range(self, axis: str, range_label: str) -> None:
        """Issue VB6 ``ChangeRange`` for ``axis``."""

    def latch(self, axis: str = "A", *, settle: bool = False) -> LatchResult:
        """Latch counter and DVM for ``axis`` and return the latch identity."""

    def read_axis(self, axis: str) -> AxisReply:
        """Read the latched counter and DVM values for one axis."""


@runtime_checkable
class VerticalMotionController(Protocol):
    """Up/down (lift) axis operations."""

    def move_to(self, position: int, *, speed_index: int = 0) -> MotionOutcome:
        ...

    def position(self) -> int:
        ...


@runtime_checkable
class TurningController(Protocol):
    """Turning (rotation) axis operations."""

    def rotate_to(self, angle_deg: float) -> MotionOutcome:
        ...

    def angle(self) -> float:
        ...

    def set_reference_angle(self, angle_deg: float) -> None:
        """VB6 ``MotorTurn_360`` re-labels 360 degrees as the new zero."""


class AcquisitionClock(Protocol):
    """Injected clock so acquisition timing is deterministic under test."""

    def now(self) -> datetime:
        ...

    def monotonic(self) -> float:
        ...

    def sleep(self, seconds: float) -> None:
        ...


class SystemAcquisitionClock:
    """Default wall/monotonic clock."""

    def now(self) -> datetime:
        return datetime.now(timezone.utc)

    def monotonic(self) -> float:
        return time.monotonic()

    def sleep(self, seconds: float) -> None:
        if seconds > 0:
            time.sleep(float(seconds))


@dataclass(frozen=True)
class AcquisitionConfig:
    """Positions, delays, and calibration for one bracketed acquisition."""

    zero_position: int
    measurement_position: int
    axis_calibration: Vector3 = (1.0, 1.0, 1.0)
    range_factor: float = 1.0e-5
    range_label: str = "1"
    holder_range_label: str = "1"
    arc_delay_s: float = DEFAULT_ARC_DELAY_S
    settle_delay_s: float = 1.0
    max_observation_age_s: float = DEFAULT_MAX_OBSERVATION_AGE_S
    zero_speed_index: int = 2
    measure_speed_index: int = 0
    recovery_settle_cycles: int = 2

    def __post_init__(self) -> None:
        if self.zero_position == self.measurement_position:
            raise ValueError("zero_position and measurement_position must differ")
        for name in ("arc_delay_s", "settle_delay_s", "max_observation_age_s"):
            value = float(getattr(self, name))
            if not math.isfinite(value) or value < 0.0:
                raise ValueError(f"{name} must be finite and non-negative")
        if len(tuple(self.axis_calibration)) != 3:
            raise ValueError("axis_calibration requires exactly three values")
        if any(float(value) == 0.0 for value in self.axis_calibration):
            raise ValueError("axis_calibration values must be non-zero")
        if not math.isfinite(float(self.range_factor)) or float(self.range_factor) == 0.0:
            raise ValueError("range_factor must be finite and non-zero")


@dataclass(frozen=True)
class BlockContext:
    """Identity carried into the block audit record."""

    sample_name: str = ""
    treatment_label: str = ""
    run_id: str = ""
    operator: str = ""
    software_version: str = ""
    config_hash: str = ""
    holder_record_id: str = ""
    holder_recorded_iso: str = ""
    is_holder_block: bool = False
    simulated: bool = False


@dataclass(frozen=True)
class RecoveryRecord:
    """Evidence for one executed flux-count recovery attempt."""

    attempt: int
    started_iso: str
    completed_iso: str
    validation: ZeroPairValidation | None
    commands: tuple[CommandEvent, ...]
    detail: str = ""


@dataclass
class _CommandLog:
    """Ordered command/event record for one acquisition."""

    clock: AcquisitionClock
    events: list[CommandEvent] = field(default_factory=list)

    def record(self, kind: str, detail: str = "", *, reply: str = "") -> "_OpenCommand":
        return _OpenCommand(self, kind, detail, self.clock.now().isoformat(), reply)

    def append(self, event: CommandEvent) -> None:
        self.events.append(event)

    def snapshot(self) -> tuple[CommandEvent, ...]:
        return tuple(self.events)


class _OpenCommand:
    """Context manager that closes one command event with its outcome."""

    def __init__(self, log: _CommandLog, kind: str, detail: str, started_iso: str, reply: str) -> None:
        self._log = log
        self._kind = kind
        self._detail = detail
        self._started_iso = started_iso
        self._reply = reply

    def __enter__(self) -> "_OpenCommand":
        return self

    def set_reply(self, reply: str) -> None:
        self._reply = str(reply)

    def set_detail(self, detail: str) -> None:
        self._detail = str(detail)

    def __exit__(self, exc_type, exc, tb) -> bool:  # noqa: ANN001 - context protocol
        self._log.append(
            CommandEvent(
                index=len(self._log.events),
                kind=self._kind,
                detail=self._detail,
                started_iso=self._started_iso,
                completed_iso=self._log.clock.now().isoformat(),
                ok=exc_type is None,
                reply=self._reply,
            )
        )
        return False


@dataclass(frozen=True)
class BracketedAcquisition:
    """One completed acquisition: the block plus its ordered command record."""

    block: BracketedMeasurementBlock
    commands: tuple[CommandEvent, ...]
    started_iso: str
    completed_iso: str
    #: Optional continuous SQUID traces recorded during the descent, turns and
    #: ascent (:mod:`rapid_main.motion_capture`).  Auxiliary evidence only:
    #: they never enter ``block.observations`` or the reduction.
    motion_traces: tuple[Any, ...] = ()


def _default_id_factory() -> Callable[[str], str]:
    counter = itertools.count(1)

    def make(prefix: str) -> str:
        return f"{prefix}-{next(counter):06d}"

    return make


class BracketedAcquisitionService:
    """Deterministic VB6-equivalent bracketed acquisition.

    All hardware access goes through the injected transport and motion
    protocols, so the full sequence, its failure modes, and its recovery path
    are exercisable without a magnetometer.
    """

    def __init__(
        self,
        transport: SquidTransport,
        vertical: VerticalMotionController,
        turning: TurningController,
        *,
        config: AcquisitionConfig,
        clock: AcquisitionClock | None = None,
        id_factory: Callable[[str], str] | None = None,
        cancel_check: Callable[[], bool] | None = None,
        motion_capture: Any | None = None,
    ) -> None:
        self._transport = transport
        self._vertical = vertical
        self._turning = turning
        self._config = config
        self._clock = clock or SystemAcquisitionClock()
        self._make_id = id_factory or _default_id_factory()
        self._cancel_check = cancel_check
        self._motion_capture = motion_capture
        self._capturing = False

    def set_cancel_check(self, check):
        self._cancel_check = check

    def _check_cancel(self):
        if self._cancel_check is not None and self._cancel_check():
            raise InterruptedError('Bracketed acquisition cancelled; no completed block is available.')

    @property
    def config(self) -> AcquisitionConfig:
        return self._config

    @property
    def motion_capture(self) -> Any | None:
        return self._motion_capture

    def acquire(
        self,
        *,
        is_up: bool = True,
        holder_positions: Position4 | None = None,
        context: BlockContext | None = None,
    ) -> BracketedAcquisition:
        """Run the complete zero/four-orientation/zero sequence.

        The returned block is *not* reduced.  Callers must still run
        :func:`reduce_bracketed_measurement`, which rejects a discontinuous
        zero pair before any accepted-state mutation or output write.
        """

        cfg = self._config
        ctx = context or BlockContext()
        log = _CommandLog(self._clock)
        started_iso = self._clock.now().isoformat()
        block_id = self._make_id("block")
        if self._motion_capture is None:
            return self._acquire_block(cfg, ctx, log, started_iso, block_id, is_up, holder_positions)
        self._motion_capture.begin_block(block_id)
        self._capturing = True
        completed = False
        try:
            acquisition = self._acquire_block(cfg, ctx, log, started_iso, block_id, is_up, holder_positions)
            completed = True
        finally:
            self._capturing = False
            traces = self._motion_capture.end_block(completed=completed)
        return replace(acquisition, motion_traces=tuple(traces))

    def _acquire_block(
        self,
        cfg: AcquisitionConfig,
        ctx: BlockContext,
        log: _CommandLog,
        started_iso: str,
        block_id: str,
        is_up: bool,
        holder_positions: Position4 | None,
    ) -> BracketedAcquisition:
        range_label = cfg.holder_range_label if ctx.is_holder_block else cfg.range_label

        # 1) Reference orientation, then lift to the verified zero position.
        self._turn_to(POSITION_ANGLES_DEG[0], log)
        self._lift_to(cfg.zero_position, cfg.zero_speed_index, log, "zero")

        # 2) 1x read mode for holder blocks, matching VB6 ChangeRange "A","1".
        with log.record("squid.set_range", f"axis=A range={range_label}"):
            self._transport_call(
                f"set range axis=A range={range_label}",
                lambda: self._transport.set_range("A", range_label),
            )

        # 3) Clear and reset the counters before the first zero (VB6 CLP/RC).
        with log.record("squid.clear_reset", "axis=A commands=CLP,RC"):
            self._transport_call(
                "clear/reset counters axis=A",
                lambda: self._transport.clear_and_reset_counters("A"),
            )
        self._delay(cfg.arc_delay_s, log, "post-reset settle")

        # 4) Zero-before. VB6 latches without the extra settling delay here
        #    because the ARC delay above already elapsed.
        zero_before = self._observe(
            role="zero-before",
            sequence_index=0,
            settle=False,
            range_label=range_label,
            turn_angle_deg=self._turning.angle(),
            log=log,
        )

        # 5) Lower to the measurement position and re-assert the 0-degree turn.
        self._lift_to(cfg.measurement_position, cfg.measure_speed_index, log, "measurement", capture="descent")
        self._turn_to(POSITION_ANGLES_DEG[0], log)
        self._delay(cfg.arc_delay_s, log, "pre-measurement settle")

        # 6) Four specimen/holder orientations, each after a verified turn.
        positions: list[SquidObservation] = []
        for index, angle in enumerate(POSITION_ANGLES_DEG):
            if index > 0:
                self._turn_to(angle, log, capture="turn")
            positions.append(
                self._observe(
                    role=f"position-{index + 1}",
                    sequence_index=index + 1,
                    settle=True,
                    range_label=range_label,
                    turn_angle_deg=angle,
                    log=log,
                )
            )

        # 7) Return to zero, close the rotation at 360 degrees, re-label as 0.
        self._lift_to(cfg.zero_position, cfg.measure_speed_index, log, "zero", capture="ascent")
        self._turn_to(BLOCK_CLOSING_ANGLE_DEG, log)
        with log.record("turning.set_reference", "angle=0"):
            self._check_cancel()
            self._turning.set_reference_angle(0.0)

        # 8) Zero-after.
        zero_after = self._observe(
            role="zero-after",
            sequence_index=5,
            settle=True,
            range_label=range_label,
            turn_angle_deg=0.0,
            log=log,
        )

        self._check_cancel()
        completed_iso = self._clock.now().isoformat()
        observations = (zero_before, *positions, zero_after)
        audit = BlockAudit(
            block_id=block_id,
            sample_name=ctx.sample_name,
            treatment_label=ctx.treatment_label,
            run_id=ctx.run_id,
            operator=ctx.operator,
            software_version=ctx.software_version,
            config_hash=ctx.config_hash,
            started_iso=started_iso,
            completed_iso=completed_iso,
            zero_position=int(cfg.zero_position),
            measurement_position=int(cfg.measurement_position),
            range_label=range_label,
            range_factor=float(cfg.range_factor),
            axis_calibration_applied=_as_vector3(cfg.axis_calibration),
            holder_record_id=ctx.holder_record_id,
            holder_recorded_iso=ctx.holder_recorded_iso,
            is_holder_block=bool(ctx.is_holder_block),
            simulated=bool(ctx.simulated),
            commands=log.snapshot(),
        )
        block = BracketedMeasurementBlock(
            zero_before=zero_before.calibrated_vector,
            positions=(
                positions[0].calibrated_vector,
                positions[1].calibrated_vector,
                positions[2].calibrated_vector,
                positions[3].calibrated_vector,
            ),
            zero_after=zero_after.calibrated_vector,
            holder_positions=_coerce_holder(holder_positions),
            is_up=bool(is_up),
            range_factor=float(cfg.range_factor),
            # Calibration was already applied per axis above, exactly once,
            # mirroring VB6 getData -> Calibrate. Reduction must not re-apply it.
            axis_calibration=(1.0, 1.0, 1.0),
            observations=observations,
            audit=audit,
        )
        return BracketedAcquisition(
            block=block,
            commands=log.snapshot(),
            started_iso=started_iso,
            completed_iso=completed_iso,
        )

    def recover_flux_count_discontinuity(
        self,
        validation: ZeroPairValidation | None = None,
        *,
        attempt: int = 1,
    ) -> RecoveryRecord:
        """Move to a safe zero state and reset the 2G counter path.

        The rejected block is *discarded by the caller*; this routine only
        restores the instrument so the complete block can be repeated.  It
        never produces a measurement value.
        """

        cfg = self._config
        log = _CommandLog(self._clock)
        started_iso = self._clock.now().isoformat()
        try:
            self._turn_to(POSITION_ANGLES_DEG[0], log)
            self._lift_to(cfg.zero_position, cfg.zero_speed_index, log, "zero")
            with log.record("squid.clear_reset", "axis=A commands=CLP,RC (recovery)"):
                self._transport.clear_and_reset_counters("A")
            for cycle in range(max(1, int(cfg.recovery_settle_cycles))):
                self._delay(cfg.arc_delay_s, log, f"recovery settle {cycle + 1}")
            with log.record("squid.latch", "axis=A (recovery re-latch)"):
                self._transport.latch("A", settle=True)
        except AcquisitionError:
            raise
        except Exception as exc:  # transport/motion adapters raise their own types
            raise RecoveryFailedError(
                f"flux-count recovery failed on attempt {attempt}: {exc}"
            ) from exc
        return RecoveryRecord(
            attempt=int(attempt),
            started_iso=started_iso,
            completed_iso=self._clock.now().isoformat(),
            validation=validation,
            commands=log.snapshot(),
            detail=(validation.reason if validation is not None else ""),
        )

    def recover_transport_failure(
        self,
        detail: str,
        *,
        attempt: int = 1,
        backoff_s: float = 0.0,
    ) -> RecoveryRecord:
        """Return to zero and prepare a fresh whole-block transport retry.

        The failed latch/read stream is never resumed. The caller must invoke
        :meth:`acquire` again, which starts with a new range/reset/latch cycle.
        """

        delay_s = float(backoff_s)
        if not math.isfinite(delay_s) or delay_s < 0.0:
            raise ValueError("transport recovery backoff must be finite and non-negative")
        cfg = self._config
        log = _CommandLog(self._clock)
        started_iso = self._clock.now().isoformat()
        try:
            self._turn_to(POSITION_ANGLES_DEG[0], log)
            self._lift_to(cfg.zero_position, cfg.zero_speed_index, log, "zero")
            with log.record("squid.clear_reset", "axis=A commands=CLP,RC (transport recovery)"):
                self._transport.clear_and_reset_counters("A")
            self._delay(cfg.arc_delay_s + delay_s, log, "transport recovery backoff")
        except MotionVerificationError:
            raise
        except Exception as exc:
            raise RecoveryFailedError(
                f"transport recovery failed on attempt {attempt}: {exc}"
            ) from exc
        return RecoveryRecord(
            attempt=int(attempt),
            started_iso=started_iso,
            completed_iso=self._clock.now().isoformat(),
            validation=None,
            commands=log.snapshot(),
            detail=str(detail),
        )

    def return_to_safe_state(self) -> tuple[CommandEvent, ...]:
        """Lift to the zero position and square the turning axis."""

        log = _CommandLog(self._clock)
        self._lift_to(self._config.zero_position, self._config.zero_speed_index, log, "zero")
        self._turn_to(POSITION_ANGLES_DEG[0], log)
        return log.snapshot()

    # -- internals ---------------------------------------------------------

    def _delay(self, seconds: float, log: _CommandLog, detail: str) -> None:
        if seconds <= 0:
            return
        with log.record("delay", f"{detail} ({seconds:.3f}s)"):
            if self._cancel_check is None:
                self._clock.sleep(float(seconds))
            else:
                remaining = float(seconds)
                while remaining > 0:
                    self._check_cancel()
                    interval = min(.05, remaining)
                    self._clock.sleep(interval)
                    remaining = max(0., remaining - interval)
                self._check_cancel()

    def _transport_call(self, detail: str, call: Callable[[], object]) -> object:
        """Normalize adapter failures into a retry-classifiable error."""
        self._check_cancel()
        try:
            return call()
        except TransportReadError:
            raise
        except Exception as exc:
            raise TransportReadError(f"{detail}: {exc}") from exc

    def _capture_segment(self, capture: str | None, name: str, target: object, mover: object):
        """Record the SQUID during this motion when a capture plan wants it."""
        if capture is None or not self._capturing or self._motion_capture is None:
            return nullcontext(None)
        return self._motion_capture.segment(capture, name, target=target, mover=mover)

    def _lift_to(
        self, position: int, speed_index: int, log: _CommandLog, name: str, *, capture: str | None = None
    ) -> None:
        self._check_cancel()
        with log.record("vertical.move", f"{name} target={position} speed_index={speed_index}") as event:
            with self._capture_segment(capture, name, int(position), self._vertical) as segment:
                outcome = self._vertical.move_to(int(position), speed_index=int(speed_index))
                if segment is not None:
                    segment.set_outcome(getattr(outcome, "actual", None), bool(getattr(outcome, "ok", False)))
            event.set_reply(f"actual={getattr(outcome, 'actual', '?')}")
            if not getattr(outcome, "ok", False):
                raise MotionVerificationError(
                    f"lift did not reach the {name} position: target={position} "
                    f"actual={getattr(outcome, 'actual', 'unknown')} "
                    f"{getattr(outcome, 'detail', '')}".strip()
                )

    def _turn_to(self, angle_deg: float, log: _CommandLog, *, capture: str | None = None) -> None:
        self._check_cancel()
        with log.record("turning.rotate", f"target_deg={angle_deg:g}") as event:
            with self._capture_segment(capture, f"{angle_deg:g}deg", float(angle_deg), self._turning) as segment:
                outcome = self._turning.rotate_to(float(angle_deg))
                if segment is not None:
                    segment.set_outcome(getattr(outcome, "actual", None), bool(getattr(outcome, "ok", False)))
            event.set_reply(f"actual_deg={getattr(outcome, 'actual', '?')}")
            if not getattr(outcome, "ok", False):
                raise MotionVerificationError(
                    f"turn did not reach {angle_deg:g} degrees: "
                    f"actual={getattr(outcome, 'actual', 'unknown')} "
                    f"{getattr(outcome, 'detail', '')}".strip()
                )

    def _observe(
        self,
        *,
        role: str,
        sequence_index: int,
        settle: bool,
        range_label: str,
        turn_angle_deg: float,
        log: _CommandLog,
    ) -> SquidObservation:
        cfg = self._config
        started = self._clock.now()
        monotonic_start = self._clock.monotonic()
        with log.record("squid.latch", f"{role} axis=A settle={settle}") as event:
            latch = self._transport_call(
                f"{role}: latch failed",
                lambda: self._transport.latch("A", settle=settle),
            )
            latch_id = getattr(latch, "latch_id", "") or ""
            event.set_reply(f"latch_id={latch_id}")
        if not latch_id:
            raise TransportReadError(f"{role}: transport did not return a latch identity")

        axes: list[AxisObservation] = []
        for axis_index, axis_name in enumerate(AXIS_NAMES):
            with log.record("squid.read_axis", f"{role} axis={axis_name}") as event:
                reply = self._transport_call(
                    f"{role}: axis {axis_name} read failed",
                    lambda axis_name=axis_name: self._transport.read_axis(axis_name),
                )
                event.set_reply(
                    f"count={getattr(reply, 'count_reply', '')!r} "
                    f"data={getattr(reply, 'data_reply', '')!r}"
                )
            reply_latch = getattr(reply, "latch_id", "") or ""
            if reply_latch and reply_latch != latch_id:
                raise TransportReadError(
                    f"{role}: axis {axis_name} reply belongs to latch {reply_latch!r}, "
                    f"not {latch_id!r}"
                )
            axes.append(
                AxisObservation(
                    axis=axis_name,
                    counts=float(reply.counts),
                    dvm=float(reply.dvm),
                    range_value=float(getattr(reply, "range_value", 1.0) or 1.0),
                    calibration=float(cfg.axis_calibration[axis_index]),
                    count_command=getattr(reply, "count_command", "") or f"{axis_name}SC",
                    count_reply=getattr(reply, "count_reply", "") or "",
                    data_command=getattr(reply, "data_command", "") or f"{axis_name}SD",
                    data_reply=getattr(reply, "data_reply", "") or "",
                )
            )

        observation = SquidObservation(
            role=role,
            sequence_index=sequence_index,
            axes=(axes[0], axes[1], axes[2]),
            latch_id=latch_id,
            latch_commands=tuple(getattr(latch, "commands", ()) or ("ALC", "ALD")),
            started_iso=started.isoformat(),
            completed_iso=self._clock.now().isoformat(),
            elapsed_s=max(0.0, self._clock.monotonic() - monotonic_start),
            range_label=range_label,
            vertical_position=int(self._vertical.position()),
            turn_angle_deg=float(turn_angle_deg),
        )
        try:
            observation.validate(max_age_s=cfg.max_observation_age_s)
        except ObservationIntegrityError as exc:
            raise TransportReadError(str(exc)) from exc
        return observation


def _as_vector3(values: Iterable[float]) -> Vector3:
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise ValueError("expected exactly three axis values")
    return (items[0], items[1], items[2])


def _coerce_holder(holder_positions: Position4 | None) -> Position4:
    if holder_positions is None:
        return ZERO_HOLDER
    items = tuple(_as_vector3(vector) for vector in holder_positions)
    if len(items) != 4:
        raise ValueError("holder_positions requires exactly four vectors")
    return (items[0], items[1], items[2], items[3])


def with_holder(block: BracketedMeasurementBlock, holder_positions: Position4) -> BracketedMeasurementBlock:
    """Return ``block`` with a different holder correction applied."""

    return replace(block, holder_positions=_coerce_holder(holder_positions))
