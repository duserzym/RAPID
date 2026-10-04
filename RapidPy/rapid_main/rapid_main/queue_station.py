"""Accepted station geometry for blank-holder and specimen transfer planning.

This module performs no hardware I/O. Queue workers must obtain fresh position
readbacks and persist their transfer stage before commanding motion.
"""
from dataclasses import dataclass
import math


def _integer(value, name):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{name} must be an integer.")
    if isinstance(value, int):
        return value
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
    xy_positions: tuple = ()

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
        positions = []
        for entry in self.xy_positions:
            if not isinstance(entry, (tuple, list)) or len(entry) != 3:
                raise ValueError('XY slots require a slot and two controller coordinates.')
            slot, x, y = (_integer(value, 'XY coordinate') for value in entry)
            if not self.slot_min <= slot <= self.slot_max or any(not -(2**31) <= value < 2**31 for value in (x, y)):
                raise ValueError('XY coordinates exceed the accepted station/controller range.')
            positions.append((slot, x, y))
        if len({entry[0] for entry in positions}) != len(positions):
            raise ValueError('Duplicate XY slot calibration.')
        object.__setattr__(self, 'xy_positions', tuple(sorted(positions)))

    @classmethod
    def from_config(cls, config, *, use_xy_table):
        from .hardware_contracts import _build_motor_controller_config
        if not config.motor_station.calibration_source:
            raise ValueError("Accepted station calibration is required for queue transfers.")
        controller = _build_motor_controller_config(config)
        station = config.motor_station
        if type(station.use_xy_table) is not bool or station.use_xy_table is not use_xy_table:
            raise ValueError('Queue table mode differs from the imported station.')
        coordinates = []
        if use_xy_table:
            if not station.xy_positions or not isinstance(station.xy_home, (list, tuple)) or len(station.xy_home) != 2:
                raise ValueError('Accepted XY home and per-slot coordinates are required.')
            for value in station.xy_home:
                if not -(2**31) <= _integer(value, 'XY home') < 2**31:
                    raise ValueError('Invalid XY home controller counts.')
            for key, values in station.xy_positions.items():
                if not isinstance(key, str) or not key.isdecimal() or str(int(key)) != key:
                    raise ValueError('XY slot keys must be canonical positive integers.')
                if not isinstance(values, (list, tuple)) or len(values) != 2:
                    raise ValueError('XY slots require both controller coordinates.')
                coordinates.append((int(key), *values))
        return cls(config.motor_station.calibration_source, use_xy_table,
                   controller.slot_min, controller.slot_max,
                   config.motor_station.hole_slot, controller.one_step, tuple(coordinates))

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
        if self.use_xy_table:
            raise ValueError('XY slot verification requires both calibrated axis readbacks.')
        counts = _integer(counts, "changer position")
        if not -(2 ** 31) <= counts < 2 ** 31:
            raise ValueError("Changer readback exceeds signed 32-bit controller limits.")
        raw = counts / self.one_step
        if not math.isfinite(raw):
            raise ValueError("Changer calibration cannot resolve this readback.")
        raw %= self.slot_max
        slot = round(raw) % self.slot_max or self.slot_max
        error = min(abs(raw - slot), abs(raw - slot + self.slot_max), abs(raw - slot - self.slot_max))
        if error > 0.02 or not self.slot_min <= slot <= self.slot_max:
            raise ValueError("Changer readback is not aligned with a registered slot.")
        return slot

    def xy_target(self, slot):
        slot = _integer(slot, 'XY slot')
        if not self.use_xy_table:
            raise ValueError('A chain station has no XY slot targets.')
        for candidate, x, y in self.xy_positions:
            if candidate == slot:
                return x, y
        raise ValueError('The requested XY slot has no accepted coordinate pair.')

    def slot_from_xy_counts(self, x, y):
        if not self.use_xy_table:
            raise ValueError('A chain station requires its chain position readback.')
        x, y = _integer(x, 'X readback'), _integer(y, 'Y readback')
        if any(not -(2**31) <= value < 2**31 for value in (x, y)):
            raise ValueError('XY readback exceeds signed 32-bit controller limits.')
        tolerance = abs(self.one_step) * 0.02
        matches = [slot for slot, target_x, target_y in self.xy_positions
                   if abs(x - target_x) <= tolerance and abs(y - target_y) <= tolerance]
        if len(matches) != 1:
            raise ValueError('XY readback must identify exactly one calibrated slot.')
        return matches[0]

    def verify_empty_xy_readback(self, target, x, y):
        if not self.is_empty(target):
            raise ValueError('XY target is not the calibrated empty hole.')
        self.xy_target(target)
        actual = self.slot_from_xy_counts(x, y)
        if actual != target:
            raise ValueError('XY readbacks do not match the requested empty hole.')
        return actual

    def verify_empty_readback(self, target, counts):
        target = _integer(target, "empty-hole target")
        if not self.is_empty(target):
            raise ValueError("Motion target is not a calibrated empty hole.")
        actual = self.slot_from_counts(counts)
        if actual != target:
            raise ValueError("Changer readback does not match the requested empty hole.")
        return actual
