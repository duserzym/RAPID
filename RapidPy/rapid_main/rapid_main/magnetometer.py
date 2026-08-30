"""Magnetometer reading normalization mapped from VB6 ``modMagnetometer``.

RapidPy hardware backends may return raw SQUID voltages or already calibrated
moment-like vectors depending on the adapter. This module captures the shared
calibration math used by the legacy stack and VRM workflow so measurement,
diagnostic, and logging code can produce auditable reading evidence.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Iterable

from rapid_main.config import CalibrationConfig


Vector3 = tuple[float, float, float]
Position4 = tuple[Vector3, Vector3, Vector3, Vector3]

DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT = 0.02
DEFAULT_AXIS_FLUX_INCREMENTS: Vector3 = (0.090, 0.106, 0.066)


@dataclass(frozen=True)
class MagnetometerCalibration:
    """Axis calibration and background settings for SQUID voltage readings."""

    xcal: float = 1.0
    ycal: float = 1.0
    zcal: float = 1.0
    range_factor: float = 1.0e-5
    background_v: Vector3 = (0.0, 0.0, 0.0)
    subtract_background: bool = False
    saturation_limit_v: float = 9.5
    minimum_signal_emu: float = 0.0

    @classmethod
    def from_config(cls, config: CalibrationConfig) -> "MagnetometerCalibration":
        return cls(
            xcal=float(config.cal_x),
            ycal=float(config.cal_y),
            zcal=float(config.cal_z),
            range_factor=float(config.range_factor),
            background_v=(float(config.bg_x), float(config.bg_y), float(config.bg_z)),
            subtract_background=bool(config.bg_subtract),
        )


@dataclass(frozen=True)
class MagnetometerReading:
    """Calibrated magnetometer reading with quality evidence."""

    raw_volts: Vector3
    corrected_volts: Vector3
    moment_emu: Vector3
    moment_magnitude_emu: float
    flags: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.flags


@dataclass(frozen=True)
class ZeroPairValidation:
    """Validation evidence for the two zeros bracketing a four-position block."""

    valid: bool
    deltas: Vector3
    discontinuous_axes: tuple[str, ...]
    flux_step_axes: tuple[str, ...]
    reason: str = ""


@dataclass(frozen=True)
class BracketedMeasurementBlock:
    """Coherent two-zero/four-position acquisition in calibrated 2G raw units.

    A live backend must capture each vector as a coherent counter+DVM sample.
    The worker accepts this structured payload in addition to legacy moment
    tuples so invalid blocks can be rejected before interpolation or saving.
    """

    zero_before: Vector3
    positions: Position4
    zero_after: Vector3
    holder_positions: Position4 = (
        (0.0, 0.0, 0.0),
        (0.0, 0.0, 0.0),
        (0.0, 0.0, 0.0),
        (0.0, 0.0, 0.0),
    )
    is_up: bool = True
    range_factor: float = 1.0e-5
    axis_calibration: Vector3 = (1.0, 1.0, 1.0)


@dataclass(frozen=True)
class BracketedMeasurementResult:
    """Reduced holder-frame result from one validated bracketed block."""

    validation: ZeroPairValidation
    baseline_adjusted_raw: Position4
    holder_frame_raw: Position4
    mean_raw: Vector3
    moment_emu: Vector3


class FluxCountDiscontinuityError(ValueError):
    """Raised when a block must be rejected instead of interpolated."""

    def __init__(self, message: str, validation: ZeroPairValidation) -> None:
        super().__init__(message)
        self.validation = validation


def validate_zero_pair(
    zero_before: Iterable[float],
    zero_after: Iterable[float],
    *,
    max_continuous_drift: float = DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT,
    axis_flux_increments: Iterable[float] = DEFAULT_AXIS_FLUX_INCREMENTS,
    flux_relative_tolerance: float = 0.20,
) -> ZeroPairValidation:
    """Reject a discontinuous zero pair before linear drift interpolation."""

    before = _coerce_vector(zero_before, "zero_before")
    after = _coerce_vector(zero_after, "zero_after")
    increments = _coerce_vector(axis_flux_increments, "axis_flux_increments")
    if not math.isfinite(max_continuous_drift) or max_continuous_drift <= 0.0:
        raise ValueError("max_continuous_drift must be finite and positive")
    if not math.isfinite(flux_relative_tolerance) or flux_relative_tolerance < 0.0:
        raise ValueError("flux_relative_tolerance must be finite and non-negative")

    deltas = tuple(abs(end - start) for start, end in zip(before, after))
    names = ("X", "Y", "Z")
    discontinuous = tuple(
        axis for axis, delta in zip(names, deltas) if delta > max_continuous_drift
    )
    flux_steps = tuple(
        axis
        for axis, delta, increment in zip(names, deltas, increments)
        if increment > 0.0
        and abs(delta - increment) <= max(0.005, abs(increment) * flux_relative_tolerance)
    )
    valid = not discontinuous
    reason = ""
    if not valid:
        values = ", ".join(
            f"{axis}={delta:.6f}"
            for axis, delta in zip(names, deltas)
            if axis in discontinuous
        )
        count_hint = f"; flux-step-like axes: {', '.join(flux_steps)}" if flux_steps else ""
        reason = (
            f"discontinuous bracketing zeros ({values}; "
            f"continuous-drift limit={max_continuous_drift:.6f}{count_hint})"
        )
    return ZeroPairValidation(
        valid=valid,
        deltas=(deltas[0], deltas[1], deltas[2]),
        discontinuous_axes=discontinuous,
        flux_step_axes=flux_steps,
        reason=reason,
    )


def reduce_bracketed_measurement(
    block: BracketedMeasurementBlock,
    *,
    max_continuous_drift: float = DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT,
) -> BracketedMeasurementResult:
    """Validate, baseline-correct, rotate, and average one 2G block.

    Discontinuous zero pairs and dominant monotonic holder staircases are
    rejected.  No result is returned for callers to save or propagate.
    """

    zero_before = _coerce_vector(block.zero_before, "zero_before")
    zero_after = _coerce_vector(block.zero_after, "zero_after")
    positions = _coerce_position4(block.positions, "positions")
    holder_positions = _coerce_position4(block.holder_positions, "holder_positions")
    calibration = _coerce_vector(block.axis_calibration, "axis_calibration")
    validation = validate_zero_pair(
        zero_before,
        zero_after,
        max_continuous_drift=max_continuous_drift,
    )
    if not validation.valid:
        raise FluxCountDiscontinuityError(validation.reason, validation)

    baseline_adjusted: list[Vector3] = []
    for index, (sample, holder) in enumerate(zip(positions, holder_positions), start=1):
        end_weight = index / 5.0
        start_weight = 1.0 - end_weight
        interpolated = tuple(
            start_weight * start + end_weight * end
            for start, end in zip(zero_before, zero_after)
        )
        baseline_adjusted.append(
            tuple(
                sample_axis - baseline_axis - holder_axis
                for sample_axis, baseline_axis, holder_axis in zip(sample, interpolated, holder)
            )
        )

    staircase_axis = _dominant_monotonic_axis(tuple(baseline_adjusted))
    if staircase_axis is not None:
        axis, span = staircase_axis
        staircase_validation = ZeroPairValidation(
            valid=False,
            deltas=validation.deltas,
            discontinuous_axes=(axis,),
            flux_step_axes=validation.flux_step_axes,
            reason=f"dominant monotonic holder {axis} staircase, span={span:.6f}",
        )
        raise FluxCountDiscontinuityError(staircase_validation.reason, staircase_validation)

    direction = 1.0 if block.is_up else -1.0
    x1, y1, z1 = baseline_adjusted[0]
    x2, y2, z2 = baseline_adjusted[1]
    x3, y3, z3 = baseline_adjusted[2]
    x4, y4, z4 = baseline_adjusted[3]
    holder_frame: Position4 = (
        (x1, -y1 * direction, z1 * direction),
        (y2, x2 * direction, z2 * direction),
        (-x3, y3 * direction, z3 * direction),
        (-y4, -x4 * direction, z4 * direction),
    )
    mean_raw = tuple(sum(vector[axis] for vector in holder_frame) / 4.0 for axis in range(3))
    if not math.isfinite(float(block.range_factor)):
        raise ValueError("range_factor must be finite")
    moment = tuple(
        raw * axis_cal * float(block.range_factor)
        for raw, axis_cal in zip(mean_raw, calibration)
    )
    if not all(math.isfinite(value) for value in (*mean_raw, *moment)):
        raise ValueError("bracketed measurement contains non-finite reduced values")
    return BracketedMeasurementResult(
        validation=validation,
        baseline_adjusted_raw=tuple(baseline_adjusted),  # type: ignore[arg-type]
        holder_frame_raw=holder_frame,
        mean_raw=(mean_raw[0], mean_raw[1], mean_raw[2]),
        moment_emu=(moment[0], moment[1], moment[2]),
    )


def calibrate_magnetometer_reading(
    raw_volts: Iterable[float],
    calibration: MagnetometerCalibration | CalibrationConfig | None = None,
) -> MagnetometerReading:
    """Convert raw SQUID axis voltages to calibrated moment evidence."""

    cal = _coerce_calibration(calibration)
    raw = _coerce_vector(raw_volts, "raw_volts")
    background = _coerce_vector(cal.background_v, "background_v")
    corrected = tuple(raw_axis - bg_axis for raw_axis, bg_axis in zip(raw, background)) if cal.subtract_background else raw
    moment = (
        corrected[0] * float(cal.xcal) * float(cal.range_factor),
        corrected[1] * float(cal.ycal) * float(cal.range_factor),
        corrected[2] * float(cal.zcal) * float(cal.range_factor),
    )
    magnitude = math.sqrt(sum(axis * axis for axis in moment))
    flags = _quality_flags(raw, corrected, moment, magnitude, cal)
    return MagnetometerReading(
        raw_volts=raw,
        corrected_volts=corrected,
        moment_emu=moment,
        moment_magnitude_emu=magnitude,
        flags=flags,
    )


def _coerce_calibration(
    calibration: MagnetometerCalibration | CalibrationConfig | None,
) -> MagnetometerCalibration:
    if calibration is None:
        return MagnetometerCalibration()
    if isinstance(calibration, MagnetometerCalibration):
        return calibration
    if isinstance(calibration, CalibrationConfig):
        return MagnetometerCalibration.from_config(calibration)
    raise TypeError("calibration must be MagnetometerCalibration, CalibrationConfig, or None")


def _coerce_vector(values: Iterable[float], name: str) -> Vector3:
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise ValueError(f"{name} requires exactly three axis values")
    return (items[0], items[1], items[2])


def _coerce_position4(values: Iterable[Iterable[float]], name: str) -> Position4:
    items = tuple(_coerce_vector(value, f"{name}[{index}]") for index, value in enumerate(values))
    if len(items) != 4:
        raise ValueError(f"{name} requires exactly four position vectors")
    return (items[0], items[1], items[2], items[3])


def _dominant_monotonic_axis(
    positions: Position4,
    *,
    minimum_span: float = 0.04,
    dominance_ratio: float = 2.0,
) -> tuple[str, float] | None:
    names = ("X", "Y", "Z")
    values_by_axis = tuple(tuple(position[axis] for position in positions) for axis in range(3))
    spans = tuple(max(values) - min(values) for values in values_by_axis)
    for axis_index, (axis, values, span) in enumerate(zip(names, values_by_axis, spans)):
        increasing = all(left < right for left, right in zip(values, values[1:]))
        decreasing = all(left > right for left, right in zip(values, values[1:]))
        other_span = max(span_value for index, span_value in enumerate(spans) if index != axis_index)
        if (
            (increasing or decreasing)
            and span >= minimum_span
            and span >= dominance_ratio * (other_span + 0.005)
        ):
            return axis, span
    return None


def _quality_flags(
    raw: Vector3,
    corrected: Vector3,
    moment: Vector3,
    magnitude: float,
    calibration: MagnetometerCalibration,
) -> tuple[str, ...]:
    flags: list[str] = []
    values = (*raw, *corrected, *moment, magnitude)
    if not all(math.isfinite(value) for value in values):
        flags.append("non-finite")
    if calibration.range_factor == 0.0:
        flags.append("zero-range-factor")
    if 0.0 in (calibration.xcal, calibration.ycal, calibration.zcal):
        flags.append("zero-axis-calibration")
    if any(abs(value) >= calibration.saturation_limit_v for value in raw):
        flags.append("near-saturation")
    if calibration.minimum_signal_emu > 0.0 and magnitude < calibration.minimum_signal_emu:
        flags.append("below-minimum-signal")
    return tuple(flags)
