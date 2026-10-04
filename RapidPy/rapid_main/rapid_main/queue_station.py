"""Accepted station geometry for blank-holder and specimen transfer planning.

This module performs no hardware I/O. Queue workers must obtain fresh position
readbacks and persist their transfer stage before commanding motion.
"""
from dataclasses import dataclass
import math


def _integer(value, name):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{name} must be an integer.")
    if not math.isfinite(value) or int(value) != value:
        raise ValueError(f"{name} must be an integer.")
    return int(value)


@dataclass(frozen=True)
class QueueStationGeometry:
    calibration_source: str
    use_xy_table: bool
    slot_min: int
    slot_max: int
    hole_slot: int
    one_step: float

    def __post_init__(self):
        if not isinstance(self.calibration_source, str) or not self.calibration_source.strip():
            raise ValueError("Accepted station calibration is required for queue transfers.")
        if type(self.use_xy_table) is not bool:
            raise ValueError("The station table mode must be explicit.")
        for name in ("slot_min", "slot_max", "hole_slot"):
            object.__setattr__(self, name, _integer(getattr(self, name), name))
        if self.slot_min < 1 or self.slot_max < self.slot_min or self.hole_slot < 1:
            raise ValueError("Invalid station slot range or empty-hole calibration.")
        if (isinstance(self.one_step, bool) or not isinstance(self.one_step, (int, float))
                or not math.isfinite(self.one_step) or self.one_step == 0):
            raise ValueError("Finite nonzero changer counts per slot are required.")
        if self.use_xy_table:
            if not self.slot_min <= self.hole_slot <= self.slot_max:
                raise ValueError("The XY empty hole must be within the accepted slot range.")
        elif self.slot_min != 1:
            # VB6 wraps using SlotMax, not SlotMax - SlotMin + 1. Do not silently
            # reinterpret an unsupported chain origin as a different station.
            raise ValueError("Legacy chain geometry requires SlotMin = 1.")
        elif self.hole_slot > self.slot_max:
            raise ValueError("The chain has no calibrated empty holes.")

    @classmethod
    def from_config(cls, config, *, use_xy_table):
        from .hardware_contracts import _build_motor_controller_config
        if not config.motor_station.calibration_source:
            raise ValueError("Accepted station calibration is required for queue transfers.")
        controller = _build_motor_controller_config(config)
        return cls(config.motor_station.calibration_source, use_xy_table,
                   controller.slot_min, controller.slot_max,
                   config.motor_station.hole_slot, controller.one_step)

    def is_empty(self, slot):
        slot = _integer(slot, "slot")
        if not self.slot_min <= slot <= self.slot_max:
            return False
        return slot == self.hole_slot if self.use_xy_table else slot % self.hole_slot == 0

    def specimen_slot(self, slot):
        slot = _integer(slot, "specimen slot")
        if not self.slot_min <= slot <= self.slot_max or self.is_empty(slot):
            raise ValueError("A specimen must occupy a registered nonempty slot.")
        return slot

    def nearest_empty(self, current_slot):
        current_slot = _integer(current_slot, "current slot")
        if not self.slot_min <= current_slot <= self.slot_max:
            raise ValueError("Current changer position is outside the accepted station range.")
        if self.use_xy_table:
            return self.hole_slot
        # The adjacent holes and the two wrap endpoints cover both directions.
        # Keep the search bounded even for a very large imported slot range.
        lower = current_slot // self.hole_slot * self.hole_slot
        candidates = {self.hole_slot, self.slot_max // self.hole_slot * self.hole_slot}
        candidates.update(slot for slot in (lower, lower + self.hole_slot)
                          if self.slot_min <= slot <= self.slot_max)
        return min(candidates, key=lambda slot: (
            min((slot - current_slot) % self.slot_max, (current_slot - slot) % self.slot_max),
            (slot - current_slot) % self.slot_max))

    def resolve_holder(self, marker, *, current_slot):
        marker = _integer(marker, "holder marker")
        if marker == 0:
            return self.nearest_empty(current_slot)
        if not self.is_empty(marker):
            raise ValueError("Holder measurement requires a calibrated empty hole.")
        return marker

    def slot_from_counts(self, counts):
        counts = _integer(counts, "changer position")
        if not -(2 ** 31) <= counts < 2 ** 31:
            raise ValueError("Changer readback exceeds signed 32-bit controller limits.")
        raw = counts / self.one_step
        if not math.isfinite(raw):
            raise ValueError("Changer calibration cannot resolve this readback.")
        if self.use_xy_table:
            slot = round(raw)
            error = abs(raw - slot)
        else:
            raw %= self.slot_max
            slot = round(raw) % self.slot_max or self.slot_max
            error = min(abs(raw - slot), abs(raw - slot + self.slot_max), abs(raw - slot - self.slot_max))
        if error > 0.02 or not self.slot_min <= slot <= self.slot_max:
            raise ValueError("Changer readback is not aligned with a registered slot.")
        return slot

    def verify_empty_readback(self, target, counts):
        target = _integer(target, "empty-hole target")
        if not self.is_empty(target):
            raise ValueError("Motion target is not a calibrated empty hole.")
        actual = self.slot_from_counts(counts)
        if actual != target:
            raise ValueError("Changer readback does not match the requested empty hole.")
        return actual
