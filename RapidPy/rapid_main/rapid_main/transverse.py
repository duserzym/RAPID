"""Transverse probe auto-position planning for VB6 parity.

The legacy ``frmTransverseProbeAutoPosition`` workflow combined angle/field
measurements with a motor move. RapidPy separates those concerns: this module
turns an angle-vs-field scan into an auditable positioning plan, while live
motor execution remains behind the existing hardware/ownership layers.
"""
from __future__ import annotations

from dataclasses import dataclass

from rapid_main.data_model import (
    AngleVsFieldCollection,
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
