"""Auditable Bartington susceptibility acquisition state machine.

The active VB6 path zeros the bridge, moves the specimen centre to
``SCoilPos + SampleHeight / 2``, measures, and either stores the holder value or
subtracts it before applying ``SusceptibilityMomentFactorCGS``.  This module
keeps that math and ordering behind narrow injected protocols.  RapidPy adds a
verified home before and after the read as an intentional safety improvement;
it does not claim that policy as legacy behavior.

Nothing in this module opens a port.  Production adapters and deterministic
tests supply the bridge and vertical-motion implementations.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
from typing import Callable, Mapping, Protocol, runtime_checkable

from rapid_main.acquisition import MotionOutcome
from rapid_main.communication_log import CommunicationEvent

SUSCEPTIBILITY_ACQUISITION_SCHEMA = "rapidpy.susceptibility.acquisition.v1"


class SusceptibilityAcquisitionError(RuntimeError):
    """Raised with immutable evidence when acquisition cannot finish safely."""

    def __init__(self, message: str, record: "SusceptibilityAcquisitionRecord") -> None:
        super().__init__(message)
        self.record = record


@runtime_checkable
class SusceptibilityBridge(Protocol):
    @property
    def simulated(self) -> bool: ...

    def is_connected(self) -> bool: ...

    def zero(self) -> str: ...

    def measure(self) -> float: ...

    def communication_events(self) -> tuple[CommunicationEvent, ...]: ...


@runtime_checkable
class SusceptibilityVerticalMotion(Protocol):
    def position(self) -> int: ...

    def home_to_top(self) -> MotionOutcome: ...

    def move_to(self, position: int, *, speed_index: int = 0) -> MotionOutcome: ...


@dataclass(frozen=True)
class SusceptibilityAcquisitionConfig:
    coil_position: int
    sample_height: int
    moment_factor_cgs: float
    speed_index: int = 0

    @property
    def target_position(self) -> int:
        """VB6 ``Int(SCoilPos + SampleHeight / 2)`` (truncate toward zero)."""

        return int(int(self.coil_position) + int(self.sample_height) / 2)

    def validate(self) -> None:
        coil = int(self.coil_position)
        target = self.target_position
        factor = float(self.moment_factor_cgs)
        if coil == 0:
            raise ValueError("Susceptibility coil position is not configured.")
        if target == 0 or (target > 0) != (coil > 0):
            raise ValueError(
                "The specimen centre would cross the configured susceptibility coil side."
            )
        if not math.isfinite(factor) or factor <= 0.0:
            raise ValueError("Susceptibility moment factor must be finite and greater than zero.")
        if int(self.speed_index) < 0:
            raise ValueError("Susceptibility motion speed index cannot be negative.")


@dataclass(frozen=True)
class SusceptibilityPhase:
    name: str
    started_iso: str
    completed_iso: str
    ok: bool
    detail: str = ""


@dataclass(frozen=True)
class SusceptibilityAcquisitionRecord:
    acquisition_id: str
    sample_id: str
    is_holder: bool
    started_iso: str
    completed_iso: str
    outcome: str
    coil_position: int
    sample_height: int
    target_position: int
    start_position: int | None
    measured_position: int | None
    final_position: int | None
    speed_index: int
    bridge_zero_reply: str = ""
    bridge_scaled_value: float | None = None
    holder_scaled_value: float | None = None
    holder_evidence_id: str = ""
    moment_factor_cgs: float = 0.0
    susceptibility: float | None = None
    safe_state_confirmed: bool = False
    error: str = ""
    safe_return_error: str = ""
    simulated: bool = False
    phases: tuple[SusceptibilityPhase, ...] = ()
    communication_events: tuple[Mapping[str, object], ...] = ()
    schema: str = SUSCEPTIBILITY_ACQUISITION_SCHEMA

    def to_dict(self) -> dict[str, object]:
        payload = asdict(self)
        payload["phases"] = [asdict(phase) for phase in self.phases]
        payload["communication_events"] = [dict(event) for event in self.communication_events]
        return payload


Clock = Callable[[], datetime]
CancelCheck = Callable[[], bool]


class SusceptibilityAcquisitionService:
    """Run one fail-closed home/zero/move/measure/home acquisition."""

    def __init__(
        self,
        bridge: SusceptibilityBridge,
        motion: SusceptibilityVerticalMotion,
        *,
        config: SusceptibilityAcquisitionConfig,
        clock: Clock | None = None,
        should_cancel: CancelCheck | None = None,
        id_factory: Callable[[], str] | None = None,
    ) -> None:
        self._bridge = bridge
        self._motion = motion
        self._config = config
        self._clock = clock or (lambda: datetime.now(timezone.utc))
        self._should_cancel = should_cancel or (lambda: False)
        self._id_factory = id_factory or self._default_id
        self._last_record: SusceptibilityAcquisitionRecord | None = None

    @property
    def last_record(self) -> SusceptibilityAcquisitionRecord | None:
        return self._last_record

    def acquire(
        self,
        *,
        sample_id: str,
        is_holder: bool,
        holder_scaled_value: float | None = None,
        holder_evidence_id: str = "",
    ) -> SusceptibilityAcquisitionRecord:
        started = self._now_iso()
        acquisition_id = str(self._id_factory())
        phases: list[SusceptibilityPhase] = []
        start_position: int | None = None
        measured_position: int | None = None
        final_position: int | None = None
        zero_reply = ""
        scaled: float | None = None
        susceptibility: float | None = None
        primary_error = ""
        safe_return_error = ""
        safe_state = False
        motion_started = False
        event_start = self._event_count()

        try:
            self._config.validate()
            if not str(sample_id).strip():
                raise ValueError("Susceptibility acquisition requires a sample identity.")
            if not bool(self._bridge.is_connected()):
                raise RuntimeError("Susceptibility bridge is not connected.")
            if not is_holder:
                if holder_scaled_value is None or not math.isfinite(float(holder_scaled_value)):
                    raise ValueError("A finite holder susceptibility value is required.")
                if not str(holder_evidence_id).strip():
                    raise ValueError("Holder susceptibility evidence identity is required.")
            start_position = int(self._motion.position())
            self._check_cancel()

            motion_started = True
            self._run_motion_phase("home_before", phases, self._motion.home_to_top)
            self._check_cancel()

            phase_started = self._now_iso()
            zero_reply = str(self._bridge.zero())
            phases.append(
                SusceptibilityPhase("bridge_zero", phase_started, self._now_iso(), True, zero_reply)
            )
            self._check_cancel()

            target = self._config.target_position
            outcome = self._run_motion_phase(
                "move_to_coil",
                phases,
                lambda: self._motion.move_to(target, speed_index=int(self._config.speed_index)),
            )
            measured_position = int(outcome.actual)
            self._check_cancel()

            phase_started = self._now_iso()
            scaled = float(self._bridge.measure())
            if not math.isfinite(scaled):
                raise RuntimeError("Susceptibility bridge returned a non-finite value.")
            phases.append(
                SusceptibilityPhase(
                    "bridge_measure", phase_started, self._now_iso(), True, f"scaled={scaled:.17g}"
                )
            )
            susceptibility = (
                scaled
                if is_holder
                else (scaled - float(holder_scaled_value)) * float(self._config.moment_factor_cgs)
            )
            if not math.isfinite(susceptibility):
                raise RuntimeError("Calculated susceptibility is non-finite.")
        except Exception as exc:
            primary_error = f"{type(exc).__name__}: {exc}"
        finally:
            if motion_started:
                try:
                    outcome = self._run_motion_phase("home_after", phases, self._motion.home_to_top)
                    final_position = int(outcome.actual)
                    safe_state = True
                except Exception as exc:
                    safe_return_error = f"{type(exc).__name__}: {exc}"

        error = primary_error
        if safe_return_error:
            error = (
                f"{primary_error}; safe return failed: {safe_return_error}"
                if primary_error
                else f"Safe return failed: {safe_return_error}"
            )
        succeeded = not error
        record = SusceptibilityAcquisitionRecord(
            acquisition_id=acquisition_id,
            sample_id=str(sample_id),
            is_holder=bool(is_holder),
            started_iso=started,
            completed_iso=self._now_iso(),
            outcome="completed" if succeeded else "failed",
            coil_position=int(self._config.coil_position),
            sample_height=int(self._config.sample_height),
            target_position=self._config.target_position,
            start_position=start_position,
            measured_position=measured_position,
            final_position=final_position,
            speed_index=int(self._config.speed_index),
            bridge_zero_reply=zero_reply,
            bridge_scaled_value=scaled,
            holder_scaled_value=(
                None if is_holder or holder_scaled_value is None else float(holder_scaled_value)
            ),
            holder_evidence_id="" if is_holder else str(holder_evidence_id),
            moment_factor_cgs=float(self._config.moment_factor_cgs),
            susceptibility=susceptibility if succeeded else None,
            safe_state_confirmed=safe_state,
            error=error,
            safe_return_error=safe_return_error,
            simulated=bool(getattr(self._bridge, "simulated", False)),
            phases=tuple(phases),
            communication_events=self._events_since(event_start),
        )
        self._last_record = record
        if not succeeded:
            raise SusceptibilityAcquisitionError(error, record)
        return record

    def _run_motion_phase(
        self,
        name: str,
        phases: list[SusceptibilityPhase],
        operation: Callable[[], MotionOutcome],
    ) -> MotionOutcome:
        started = self._now_iso()
        try:
            outcome = operation()
            if not bool(outcome.ok) or not math.isfinite(float(outcome.actual)):
                raise RuntimeError(
                    outcome.detail
                    or f"motion did not reach target {outcome.target}; actual={outcome.actual}"
                )
        except Exception as exc:
            phases.append(
                SusceptibilityPhase(name, started, self._now_iso(), False, f"{type(exc).__name__}: {exc}")
            )
            raise
        phases.append(
            SusceptibilityPhase(
                name,
                started,
                self._now_iso(),
                True,
                f"target={outcome.target:.17g}; actual={outcome.actual:.17g}",
            )
        )
        return outcome

    def _check_cancel(self) -> None:
        if self._should_cancel():
            raise InterruptedError("Susceptibility acquisition was cancelled.")

    def _event_count(self) -> int:
        try:
            return len(tuple(self._bridge.communication_events()))
        except Exception:
            return 0

    def _events_since(self, start: int) -> tuple[Mapping[str, object], ...]:
        try:
            events = tuple(self._bridge.communication_events())[start:]
        except Exception:
            return ()
        return tuple(_event_payload(event) for event in events)

    def _now_iso(self) -> str:
        value = self._clock()
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc).isoformat()

    def _default_id(self) -> str:
        return "susc-" + self._now_iso().replace(":", "").replace("+00:00", "Z")


def write_susceptibility_acquisition(
    path: str | Path, record: SusceptibilityAcquisitionRecord
) -> Path:
    """Atomically publish one immutable acquisition artifact."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_name(f"{target.name}.tmp-{os.getpid()}")
    text = json.dumps(record.to_dict(), indent=2, sort_keys=True) + "\n"
    try:
        with open(temp, "x", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        if target.exists():
            raise FileExistsError(f"Susceptibility artifact already exists: {target}")
        os.replace(temp, target)
    except Exception:
        try:
            temp.unlink()
        except OSError:
            pass
        raise
    return target


def _event_payload(event: CommunicationEvent) -> Mapping[str, object]:
    timestamp = event.timestamp
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    direction = getattr(event.direction, "value", str(event.direction))
    return {
        "timestamp": timestamp.astimezone(timezone.utc).isoformat(),
        "channel": str(event.channel),
        "direction": str(direction),
        "port": str(event.port),
        "payload": str(event.payload),
        "detail": str(event.detail),
    }
