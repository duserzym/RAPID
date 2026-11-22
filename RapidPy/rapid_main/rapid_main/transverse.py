"""Angle/field analysis and scan planning for the transverse probe."""
from __future__ import annotations

from dataclasses import dataclass
import math

from rapid_main.data_model import (
    AngleVsFieldCollection,
    AngleVsFieldPoint,
    ProbeAngleOptimizationResult,
    ProbeAngleOptimizer,
)


@dataclass(frozen=True)
class TransverseAutoPositionPlan:
    """Auditable result of a transverse probe auto-position calculation."""

    current_angle_deg: float
    target: ProbeAngleOptimizationResult
    relative_move_deg: float
    tolerance_deg: float
    move_required: bool
    direction: str
    allowed: bool
    reason: str = ""

    @property
    def target_angle_deg(self) -> float:
        return self.target.angle_deg

    @property
    def command_summary(self) -> str:
        if not self.allowed:
            return f"BLOCKED: {self.reason}"
        if not self.move_required:
            return f"Hold transverse probe at {self.current_angle_deg:.3f} deg"
        return (
            f"Move transverse probe {self.direction} {abs(self.relative_move_deg):.3f} deg "
            f"to {self.target_angle_deg:.3f} deg"
        )


@dataclass(frozen=True)
class TransverseScanPlan:
    angles_deg: tuple[float, ...]
    field_mT: float
    meter_port: int

    @property
    def bounds(self) -> tuple[float, float]:
        return tuple(sorted((self.angles_deg[0], self.angles_deg[-1])))


def plan_transverse_scan(start_deg, end_deg, pitch_deg, field_mT, meter_port) -> TransverseScanPlan:
    values = start_deg, end_deg, pitch_deg, field_mT
    if not all(math.isfinite(value) for value in values):
        raise ValueError("Scan angles, pitch and field must be finite.")
    span = abs(end_deg - start_deg)
    if max(abs(start_deg), abs(end_deg)) > 360 or not 0 < span <= 360:
        raise ValueError("Choose a travel interval of at most one turn within −360° to 360°.")
    if pitch_deg < 1 or pitch_deg > span / 2:
        raise ValueError("Pitch must be at least 1° and leave at least three scan positions.")
    if field_mT <= 0:
        raise ValueError("Enter a positive calibrated transverse AF field.")
    if type(meter_port) is not int or not 1 <= meter_port <= 256:
        raise ValueError("Select the physical 908A gaussmeter COM port.")
    direction = 1 if end_deg > start_deg else -1
    angles = [start_deg + index * direction * pitch_deg for index in range(math.ceil(span / pitch_deg))]
    if len(angles) >= 3 and abs(angles[-1] - end_deg) < min(1, pitch_deg / 2):
        angles.pop()
    angles.append(end_deg)
    if len(angles) > 181:
        raise ValueError("Choose a larger pitch for the 181-position scan limit.")
    return TransverseScanPlan(tuple(angles), field_mT, meter_port)


@dataclass(frozen=True)
class TransverseSample:
    requested_deg: float
    actual_deg: float
    peak_mT: float


def best_transverse_sample(samples: tuple[TransverseSample, ...]) -> TransverseSample:
    if (len(samples) < 3 or len({sample.actual_deg for sample in samples}) != len(samples)
            or any(not math.isfinite(sample.peak_mT) for sample in samples)):
        raise ValueError("A transverse scan needs at least three valid peak measurements.")
    collection = AngleVsFieldCollection([
        AngleVsFieldPoint(sample.actual_deg, abs(sample.peak_mT)) for sample in samples
    ])
    # Rank magnitudes and retain the signed reading and original encoder travel.
    result = plan_transverse_auto_position(collection, current_angle_deg=samples[-1].actual_deg)
    if math.isclose(max(collection.fields), min(collection.fields), abs_tol=1e-9):
        raise ValueError("The scan did not resolve variation in the measured AF peak.")
    return next(sample for sample in samples if sample.actual_deg == result.target_angle_deg)


def plan_transverse_auto_position(
    collection: AngleVsFieldCollection,
    *,
    current_angle_deg: float,
    tolerance_deg: float = 0.5,
    min_confidence: float = 0.0,
) -> TransverseAutoPositionPlan:
    """Create a safe move plan from angle/field observations."""

    if tolerance_deg < 0.0:
        raise ValueError("transverse auto-position tolerance must be non-negative")
    if not 0.0 <= min_confidence <= 1.0:
        raise ValueError("transverse auto-position min_confidence must be between 0 and 1")

    target = ProbeAngleOptimizer(collection).optimize()
    if target is None:
        raise ValueError("transverse auto-position requires at least one angle/field point")

    current = _normalize_angle(current_angle_deg)
    target_angle = _normalize_angle(target.angle_deg)
    relative = _shortest_delta(current, target_angle)
    move_required = abs(relative) > tolerance_deg
    direction = "NONE" if not move_required else ("CW" if relative > 0.0 else "CCW")

    allowed = target.confidence >= min_confidence
    reason = "" if allowed else (
        f"confidence {target.confidence:.3f} below required {min_confidence:.3f}"
    )
    return TransverseAutoPositionPlan(
        current_angle_deg=current,
        target=target,
        relative_move_deg=relative,
        tolerance_deg=float(tolerance_deg),
        move_required=move_required and allowed,
        direction=direction if allowed else "BLOCKED",
        allowed=allowed,
        reason=reason,
    )


def _normalize_angle(angle_deg: float) -> float:
    value = float(angle_deg) % 360.0
    if value < 0.0:
        value += 360.0
    return value


def _shortest_delta(current_angle_deg: float, target_angle_deg: float) -> float:
    return ((target_angle_deg - current_angle_deg + 180.0) % 360.0) - 180.0
