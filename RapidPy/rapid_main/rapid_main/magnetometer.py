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
