"""Geometry, acquisition, and publication for mounted SQUID long scans."""
from __future__ import annotations

import csv
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
import tempfile
import time

from .rod_calibration import RodScanResult
from .squid_rod_calibration import coil_factors


UCHANNEL_PROCEDURE = 'measurement/uchannel'


@dataclass(frozen=True)
class CoreGeometry:
    length_cm: float
    first_target: int
    last_target: int
    zero_target: int


@dataclass(frozen=True)
class EprTubeGeometry:
    length_cm: float
    sample_bottom_cm: float
    sample_top_cm: float
    buffer_below_cm: float
    buffer_above_cm: float


@dataclass(frozen=True)
class UChannelRequest:
    specimen: str
    geometry: CoreGeometry | EprTubeGeometry
    assembly_extension_cm: float
    spacing_cm: float
    orientation: str
    four_passes: bool
    positions_per_zero: int
    drift_limit_emu: float


@dataclass(frozen=True)
class UChannelPlan:
    request: UChannelRequest
    carriage_limit: int
    floor_limit: int
    counts_per_cm: float
    axis_factors_emu: tuple[float, float, float]
    range_factor: float
    settle_s: float
    positions: tuple[int, ...]
    zero_target: int
    angles: tuple[int, ...]
    safe_depth: int
    readback_tolerance: int

    @property
    def reading_budget(self):
        group_size = self.request.positions_per_zero or len(self.positions)
        return len(self.angles) * (len(self.positions) + math.ceil(len(self.positions) / group_size) + 1)


def _number(value):
    return type(value) in (int, float) and math.isfinite(value)


def plan_uchannel(request: UChannelRequest, *, meas_pos: int, carriage_limit: int,
                  floor_limit: int, counts_per_cm: float, axis_factors_emu: tuple[float, float, float],
                  range_factor: float, settle_s: float) -> UChannelPlan:
    if not isinstance(request.specimen, str) or not request.specimen.strip():
        raise ValueError('Enter the long-scan specimen identity.')
    quantities = (request.assembly_extension_cm, request.spacing_cm, request.drift_limit_emu,
                  counts_per_cm, range_factor, settle_s)
    if (not all(_number(value) and value > 0 for value in quantities) or settle_s > 60
            or len(axis_factors_emu) != 3 or any(not _number(value) or value == 0 for value in axis_factors_emu)):
        raise ValueError('Verify the dimensions, spacing, drift limit, lift scale, and SQUID calibration.')
    if any(type(value) is not int or not -(2**31) <= value < 0 for value in (meas_pos, carriage_limit, floor_limit)):
        raise ValueError('Import the measurement position, carriage limit, and floor obstruction as negative counts.')
    if (request.orientation not in ('up', 'down') or type(request.four_passes) is not bool
            or type(request.positions_per_zero) is not int or not 0 <= request.positions_per_zero <= 1001):
        raise ValueError('Choose a specimen orientation, pass count, and valid zero interval.')
    extension = request.assembly_extension_cm * counts_per_cm
    interval = request.spacing_cm * counts_per_cm
    if not 1 <= extension < abs(floor_limit) or not 4 <= interval < abs(carriage_limit):
        raise ValueError('The mounted assembly or scan spacing falls outside the station travel.')
    safe_depth = max(carriage_limit, floor_limit + math.ceil(extension))
    tolerance = min(150, math.floor(interval / 4))

    geometry = request.geometry
    if (not isinstance(geometry, (CoreGeometry, EprTubeGeometry)) or not _number(geometry.length_cm)
            or not 0 < geometry.length_cm <= request.assembly_extension_cm):
        raise ValueError('The loaded extension must cover the full core or tube length.')
    if isinstance(geometry, CoreGeometry):
        first, last, zero = geometry.first_target, geometry.last_target, geometry.zero_target
        if any(type(value) is not int for value in (first, last, zero)):
            raise ValueError('Enter the bench-verified core start, end, and zero as integer controller counts.')
    else:
        dimensions = (geometry.sample_bottom_cm, geometry.sample_top_cm, geometry.buffer_below_cm,
                      geometry.buffer_above_cm)
        if (not all(_number(value) and 0 <= value <= geometry.length_cm for value in dimensions)
                or not 0 <= geometry.sample_bottom_cm < geometry.sample_top_cm <= geometry.length_cm):
            raise ValueError('Verify the EPR sample bounds and buffers against the tube length.')
        first = meas_pos + round((geometry.length_cm - geometry.sample_bottom_cm + geometry.buffer_below_cm) * counts_per_cm)
        last = meas_pos + round((geometry.length_cm - geometry.sample_top_cm - geometry.buffer_above_cm) * counts_per_cm)
        zero = first + round(5 * counts_per_cm)
    if (last < safe_depth + tolerance or first - last <= 2 * tolerance
            or zero - first <= 2 * tolerance or zero + tolerance >= 0):
        raise ValueError('The zero and scan span must fit between the top margin and loaded carriage/floor limit.')
    count = math.floor((first - last) / interval) + 1
    if count > 1001:
        raise ValueError('Increase scan spacing to keep the profile within 1,001 positions.')
    points = [round(first - index * interval) for index in range(count)]
    if points[-1] != last:
        if len(points) > 1 and points[-1] - last <= 2 * tolerance:
            points[-1] = last
        else:
            points.append(last)
    if len(points) > 1001:
        raise ValueError('Increase scan spacing to keep the profile within 1,001 positions.')
    angles = (0, 90, 180, 270) if request.four_passes else (0,)
    return UChannelPlan(request, carriage_limit, floor_limit, counts_per_cm, tuple(axis_factors_emu), range_factor,
                        max(2, settle_s), tuple(points), zero, angles, safe_depth, tolerance)


def plan_uchannel_from_config(request, config):
    station = config.motor_station
    return plan_uchannel(request, meas_pos=config.motion.meas_pos, carriage_limit=station.lift_min_counts,
                        floor_limit=station.lift_floor_counts, counts_per_cm=station.lift_counts_per_cm,
                        axis_factors_emu=coil_factors(config.calibration), range_factor=config.calibration.range_factor,
                        settle_s=config.squid.settle_time)


@dataclass(frozen=True)
class ChannelReading:
    phase: str
    requested_angle: int
    actual_angle: float
    target: int
    position: int
    elapsed_s: float
    raw_xyz: tuple[float, float, float]
    moment_emu: tuple[float, float, float]


@dataclass(frozen=True)
class ChannelPoint:
    angle: int
    actual_angle: float
    target: int
    position: int
    distance_cm: float
    moment_emu: tuple[float, float, float]


@dataclass(frozen=True)
class UChannelSummary:
    profile: tuple[ChannelPoint, ...]
    max_baseline_drift_emu: float


def scan_uchannel(plan, sample, orient, readings, *, cancelled, clock=time.perf_counter):
    start = clock()
    profile = []
    max_drift = 0.

    def take(phase, target, angle):
        if cancelled():
            raise InterruptedError('The U-channel scan was canceled.')
        position, actual_angle, raw = sample(target, angle)
        elapsed = clock() - start
        if (type(position) is not int or not plan.safe_depth <= position < 0
                or abs(position - target) > plan.readback_tolerance or not _number(actual_angle)
                or abs(actual_angle - angle) > 5 or not isinstance(raw, (tuple, list)) or len(raw) != 3
                or any(not _number(value) for value in raw) or not math.isfinite(elapsed) or elapsed < 0
                or (readings and elapsed <= readings[-1].elapsed_s)):
            raise ValueError('A long-scan point has an invalid lift, turn, time, or SQUID readback.')
        moment = tuple(value * factor for value, factor in zip(raw, plan.axis_factors_emu))
        if any(not math.isfinite(value) for value in moment):
            raise ValueError('A calibrated SQUID reading exceeds the finite measurement range.')
        reading = ChannelReading(phase, angle, actual_angle, target, position, elapsed, tuple(raw), moment)
        readings.append(reading)
        return reading

    group_size = plan.request.positions_per_zero or len(plan.positions)
    for angle in plan.angles:
        if cancelled():
            raise InterruptedError('The U-channel scan was canceled.')
        orient(angle)
        before = take('zero', plan.zero_target, angle)
        for start_index in range(0, len(plan.positions), group_size):
            group = [take('sample', target, angle) for target in plan.positions[start_index:start_index + group_size]]
            after = take('zero', plan.zero_target, angle)
            drift = tuple(right - left for left, right in zip(before.moment_emu, after.moment_emu))
            max_drift = max(max_drift, max(map(abs, drift)))
            if max_drift > plan.request.drift_limit_emu:
                raise ValueError(f'SQUID baseline drift exceeded the limit during the {angle}° pass.')
            for sample_reading in group:
                fraction = (sample_reading.elapsed_s - before.elapsed_s) / (after.elapsed_s - before.elapsed_s)
                xyz = tuple(measured - (zero + delta * fraction)
                            for measured, zero, delta in zip(sample_reading.moment_emu, before.moment_emu, drift))
                if any(not math.isfinite(value) for value in xyz):
                    raise ValueError('A baseline-corrected point exceeds the finite measurement range.')
                profile.append(ChannelPoint(angle, sample_reading.actual_angle, sample_reading.target,
                                            sample_reading.position,
                                            (plan.positions[0] - sample_reading.position) / plan.counts_per_cm, xyz))
            before = after
    if len(plan.angles) > 1:
        orient(0)
    return UChannelSummary(tuple(profile), max_drift)


def simulate_uchannel(plan, readings, *, cancelled):
    center = (plan.positions[0] + plan.positions[-1]) / 2
    width = (plan.positions[0] - plan.positions[-1]) / 4

    def sample(position, angle):
        strength = 0. if position == plan.zero_target else math.exp(-.5 * ((position - center) / width) ** 2)
        radians = math.radians(angle)
        return position, float(angle), (2 + strength * math.cos(radians), 3 + strength * math.sin(radians), 4 + strength)

    return scan_uchannel(plan, sample, lambda angle: None, readings, cancelled=cancelled)


def _write_csv(path, records):
    with path.open('x', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle)
        writer.writerows(records)
        handle.flush()
        os.fsync(handle.fileno())


def _legacy_rows(plan, readings, started):
    geometry = plan.request.geometry
    is_epr = isinstance(geometry, EprTubeGeometry)
    yield (('EPR Tube' if is_epr else 'Long Core') + ' Scan',)
    yield (started.date().isoformat(),)
    yield (started.strftime('%H:%M:%S UTC'),)
    length = geometry.sample_top_cm - geometry.sample_bottom_cm if is_epr else geometry.length_cm
    yield ('Sample Length (cm):', length)
    yield ('Scan Step size (cm):', plan.request.spacing_cm)
    yield ('Scan Length (cm):', (plan.positions[0] - plan.positions[-1]) / plan.counts_per_cm)
    if is_epr:
        yield ('Sample Top (cm above end of EPR Tube):', geometry.sample_top_cm)
        yield ('Sample Bottom (cm above end of tube):', geometry.sample_bottom_cm)
        yield ('EPR Tube Length (cm)::', geometry.length_cm)
    yield ()
    yield ('Requested Up/Down Pos', 'Actual Up/Down Pos', 'Time (sec)', 'Turn Angle', 'SQUID X', 'SQUID Y', 'SQUID Z')
    for row in readings:
        legacy = tuple(value / plan.range_factor for value in row.moment_emu)
        yield (row.target, row.position, row.elapsed_s, row.actual_angle, *legacy)


def _profile_rows(profile):
    yield ('angle_deg', 'actual_angle_deg', 'target_counts', 'position_counts', 'distance_cm', 'x_emu', 'y_emu', 'z_emu')
    for row in profile:
        yield (row.angle, row.actual_angle, row.target, row.position, row.distance_cm, *row.moment_emu)


class UChannelJob:
    procedure_name = 'U-channel Scan'

    def __init__(self, plan, folder, *, operator, config_hash, scan=None):
        self.plan = plan
        self.path = Path(folder)
        self.operator = operator
        self.config_hash = config_hash
        self.scan = scan

    def run(self, *, cancelled):
        started = datetime.now(timezone.utc)
        simulated = self.scan is None
        readings, evidence = [], {}
        summary, error = None, ''
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            raise FileExistsError(f'The long-scan output already exists at {self.path}.')
        with tempfile.TemporaryDirectory(prefix='.uchannel-', dir=self.path.parent) as temporary:
            staged = Path(temporary) / 'bundle'
            staged.mkdir()
            with (staged / 'scan.json').open('x', encoding='utf-8', newline='\n') as manifest:
                try:
                    if simulated:
                        summary = simulate_uchannel(self.plan, readings, cancelled=cancelled)
                    else:
                        summary = self.scan(self.plan, readings, evidence, cancelled=cancelled)
                    if cancelled():
                        raise InterruptedError('The U-channel scan was canceled.')
                except Exception as exc:
                    summary, error = None, str(exc) or type(exc).__name__
                payload = dict(schema='rapidpy.measurement.uchannel.v1', procedure_id=UCHANNEL_PROCEDURE,
                               operator=self.operator, config_hash=self.config_hash, simulated=simulated,
                               status='complete' if summary is not None else 'failed', started_at=started.isoformat(),
                               completed_at=datetime.now(timezone.utc).isoformat(), plan=asdict(self.plan),
                               station_evidence=evidence, readings=[asdict(row) for row in readings], error=error,
                               max_baseline_drift_emu=summary.max_baseline_drift_emu if summary else None)
                json.dump(payload, manifest, indent=2, allow_nan=False)
                manifest.write('\n')
                manifest.flush()
                os.fsync(manifest.fileno())
            _write_csv(staged / 'readings.csv', _legacy_rows(self.plan, readings, started))
            if summary is not None:
                _write_csv(staged / 'profile.csv', _profile_rows(summary.profile))
            os.replace(staged, self.path)
        return RodScanResult(self.path / 'scan.json', summary, error, simulated, bool(evidence.get('pending_motion_record')))
