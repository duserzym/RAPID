from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class CalibrationFactors:
    x: float
    y: float
    z: float


@dataclass(slots=True)
class MeasurementSample:
    time_s: float
    x_volts: float
    y_volts: float
    z_volts: float

    def as_moment(self, factors: CalibrationFactors) -> tuple[float, float, float]:
        return (
            self.x_volts * factors.x,
            self.y_volts * factors.y,
            self.z_volts * factors.z,
        )
