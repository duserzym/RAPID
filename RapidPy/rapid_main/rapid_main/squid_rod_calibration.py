"""Coarse and fine SQUID scans of an operator-mounted magnetic standard."""
from dataclasses import dataclass
import math

from .rod_calibration import RodFitError, RodScanJob


SQUID_ROD_PROCEDURE = 'calibration/squid-rod'
SQUID_ROD_SCHEMA = 'rapidpy.calibration.squid_rod.v1'
AXES = ('X', 'Y', 'Z')


def coil_factors(calibration):
    return tuple(calibration.range_factor * value for value in (calibration.cal_x, calibration.cal_y, calibration.cal_z))


@dataclass(frozen=True)
class SquidRodRequest:
    standard: str
    height_cm: float
    spacing_cm: float
    drift_limit_emu: float
    threshold_percent: float


@dataclass(frozen=True)
class SquidRodPlan:
    request: SquidRodRequest
    zero_pos: int
    meas_pos: int
    counts_per_cm: float
    axis_factors_emu: tuple[float, float, float]
    settle_s: float
    baseline_position: int
    standard_center_offset: float
    positions: tuple[int, ...]
    fine_half_count: int
    readback_tolerance: int

    @property
    def fine_interval(self):
        return self.request.spacing_cm * self.counts_per_cm

    @property
    def deepest_target(self):
        return self.meas_pos + self.readback_tolerance


def plan_squid_rod(request: SquidRodRequest, *, zero_pos: int, meas_pos: int, counts_per_cm: float,
                   axis_factors_emu: tuple[float, float, float], settle_s: float) -> SquidRodPlan:
    quantities = (request.height_cm, request.spacing_cm, request.drift_limit_emu,
                  request.threshold_percent, counts_per_cm, settle_s)
    if not isinstance(request.standard, str) or not request.standard.strip():
        raise ValueError('Enter the magnetic reference standard identity.')
    if any(type(value) not in (int, float) or not math.isfinite(value) or value <= 0 for value in quantities):
        raise ValueError('Use positive finite rod dimensions, fit threshold, drift limit, and lift scale.')
    if (request.threshold_percent >= 100 or settle_s > 60 or len(axis_factors_emu) != 3
            or any(type(value) not in (int, float) or not math.isfinite(value) or value == 0
                   for value in axis_factors_emu)):
        raise ValueError('Verify the SQUID drift/fit settings, settle time, and three calibration factors.')
    if (type(zero_pos) is not int or type(meas_pos) is not int or not -(2**30) < meas_pos < zero_pos < 0):
        raise ValueError('Import the accepted negative SQUID zero and measurement positions.')

    offset = request.height_cm * counts_per_cm / 2
    interval = request.spacing_cm * counts_per_cm
    if not 1 <= offset < -zero_pos or not 4 <= interval < abs(zero_pos):
        raise ValueError('The standard height or scan spacing falls outside the calibrated lift travel.')
    coarse_step = round(10 * interval)
    tolerance = min(150, math.floor(interval / 4))
    deepest_target = meas_pos + tolerance
    baseline = math.floor(zero_pos + offset)
    expected = math.floor(meas_pos + offset)
    if baseline > -math.ceil(interval) or not deepest_target < expected < baseline:
        raise ValueError('The SQUID baseline and scan must fit inside the accepted travel and top-switch margin.')
    height_count = 5 + request.height_cm / request.spacing_cm
    span_count = 4 / request.spacing_cm
    coarse = range(baseline, deepest_target, -coarse_step)
    if len(coarse) + 2 > 500 or height_count >= 502 or span_count > 500:
        raise ValueError('Increase spacing to keep the coarse and fine scans within 1,001 points.')
    fine_count = max(15, math.floor(height_count), math.ceil(span_count) + 1)
    positions = tuple(sorted({*coarse, expected, deepest_target}, reverse=True))
    return SquidRodPlan(request, zero_pos, meas_pos, counts_per_cm, tuple(axis_factors_emu), max(2, settle_s),
                        baseline, offset, positions, fine_count // 2, tolerance)


@dataclass(frozen=True)
class SquidRodReading:
    phase: str
    target: int
    position: int
    raw_xyz: tuple[float, float, float]
    corrected_emu: tuple[float, float, float]


@dataclass(frozen=True)
class SquidRodFit:
    measurement_position: int
    zero_position: int
    selected_axis: str
    fitted_rod_position: int
    fitted_peak_emu: float
    verification_emu: float
    baseline_drift_emu: tuple[float, float, float]


def _primary_peak(readings):
    axis = max(range(3), key=lambda index: max(abs(r.corrected_emu[index]) for r in readings))
    values = [reading.corrected_emu[axis] for reading in readings]
    peak_index = max(range(len(values)), key=lambda index: abs(values[index]))
    if peak_index in (0, len(values) - 1) or abs(values[peak_index]) <= max(abs(values[0]), abs(values[-1])):
        raise RodFitError('The SQUID response must peak inside the scan travel on the strongest axis.')
    peak = values[peak_index]
    return axis, peak_index, peak


def scan_squid_rod(plan, sample, readings, *, cancelled):
    baseline_raw = None

    def take(phase, target):
        if cancelled():
            raise InterruptedError('The SQUID rod scan was canceled.')
        position, raw = sample(target)
        if (type(position) is not int or not plan.meas_pos <= position < 0
                or abs(position - target) > plan.readback_tolerance
                or not isinstance(raw, (list, tuple)) or len(raw) != 3
                or any(type(value) not in (int, float) or not math.isfinite(value) for value in raw)):
            raise ValueError('The SQUID scan point lacks position or three finite axis readings.')
        corrected = (0., 0., 0.) if baseline_raw is None else tuple(
            (value - baseline) * factor for value, baseline, factor in zip(raw, baseline_raw, plan.axis_factors_emu))
        if not all(map(math.isfinite, corrected)):
            raise ValueError('Calibrated SQUID readings are outside the finite measurement range.')
        reading = SquidRodReading(phase, target, position, tuple(raw), corrected)
        readings.append(reading)
        return reading

    baseline_raw = take('baseline', plan.baseline_position).raw_xyz
    coarse = [take('coarse', position) for position in plan.positions]
    axis, index, peak = _primary_peak(coarse)
    center = coarse[index].target
    fine_count = plan.fine_half_count
    targets = [math.floor(center + step * plan.fine_interval) for step in range(fine_count, -fine_count - 1, -1)]
    upper = min(plan.baseline_position, targets[0])
    lower = max(plan.deepest_target, targets[-1])
    fine_positions = sorted({lower, upper, *(target for target in targets if lower <= target <= upper)}, reverse=True)
    if len(fine_positions) < 5 or not lower < center < upper:
        raise RodFitError('The fine SQUID scan cannot bracket the coarse peak within the accepted lift travel.')

    fine = [take('fine', position) for position in fine_positions]
    fine_axis, _, fine_peak = _primary_peak(fine)
    if fine_axis != axis or math.copysign(1, fine_peak) != math.copysign(1, peak):
        raise RodFitError('The coarse and fine SQUID scans disagree on the strongest axis or polarity.')
    threshold_ratio = 1 - plan.request.threshold_percent / 100
    ratios = [reading.corrected_emu[axis] / fine_peak for reading in fine]
    selected = [index for index, ratio in enumerate(ratios) if ratio >= threshold_ratio]
    if (selected[0] == 0 or selected[-1] == len(fine) - 1
            or len(selected) != selected[-1] - selected[0] + 1):
        raise RodFitError('The SQUID fine-scan peak region must be contiguous and clear of the travel limits.')

    fitted = math.fsum(fine[i].position * ratios[i] for i in selected) / math.fsum(ratios[i] for i in selected)
    fitted_target = math.floor(fitted)
    if not lower < fitted_target < upper:
        raise RodFitError('The fitted SQUID center leaves the verified fine scan travel.')
    verification = take('verify_center', fitted_target).corrected_emu[axis]
    drift = take('verify_baseline', plan.baseline_position).corrected_emu
    if max(map(abs, drift)) > plan.request.drift_limit_emu:
        raise RodFitError('The SQUID baseline drift exceeded the configured limit.')
    if abs(verification / fine_peak - 1) > 1 - threshold_ratio:
        raise RodFitError('The independent center reading did not reproduce the SQUID peak.')
    measurement = math.floor(fitted - plan.standard_center_offset)
    zero = plan.zero_pos
    if not 13.5 <= (zero - measurement) / plan.counts_per_cm <= 16.5:
        zero = round(measurement + 15 * plan.counts_per_cm)
    if not -(2**30) < measurement < zero < 0:
        raise RodFitError('The proposed measurement or zero position is outside the negative lift coordinate system.')
    return SquidRodFit(measurement, zero, AXES[axis], fitted_target, fine_peak, verification, drift)


def simulate_squid_rod(plan, readings, *, cancelled):
    center = math.floor(plan.meas_pos + plan.standard_center_offset)
    width = max(plan.fine_interval, min(plan.standard_center_offset / 4, plan.fine_interval * 2))

    def sample(position):
        peak = math.exp(-.5 * ((position - center) / width) ** 2)
        return position, tuple(base + strength * peak for base, strength in zip((.1, -.2, .3), (.9, -.95, 1)))

    return scan_squid_rod(plan, sample, readings, cancelled=cancelled)


class SquidRodScanJob(RodScanJob):
    schema = SQUID_ROD_SCHEMA
    procedure_id = SQUID_ROD_PROCEDURE
    procedure_name = 'SQUID Rod Scan'
    simulate = staticmethod(simulate_squid_rod)
