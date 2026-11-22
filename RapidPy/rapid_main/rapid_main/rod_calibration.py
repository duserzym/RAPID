"""Bartington calibration rod positioning, fit, and evidence."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
from typing import TYPE_CHECKING
from uuid import uuid4

from PySide6 import QtCore

if TYPE_CHECKING:
    from .squid_rod_calibration import SquidRodFit
    from .uchannel import UChannelSummary


ROD_PROCEDURE = 'calibration/susceptibility-rod'
ROD_SCHEMA = 'rapidpy.calibration.susceptibility_rod.v1'


@dataclass(frozen=True)
class RodScanRequest:
    standard: str
    height_cm: float
    spacing_cm: float
    reference_cgs: float
    threshold_percent: float
    tray_slot: int | None = None


@dataclass(frozen=True)
class RodScanPlan:
    request: RodScanRequest
    coil: int
    travel_limit: int
    counts_per_cm: float
    moment_factor_cgs: float
    center_correction: float
    positions: tuple[int, ...]
    readback_tolerance: int


def plan_rod_scan(request: RodScanRequest, *, coil: int, travel_limit: int,
                  counts_per_cm: float, moment_factor_cgs: float) -> RodScanPlan:
    quantities = (request.height_cm, request.spacing_cm, request.reference_cgs, request.threshold_percent,
                  counts_per_cm, moment_factor_cgs)
    if not isinstance(request.standard, str) or not request.standard.strip():
        raise ValueError('Enter the reference standard identity.')
    if request.tray_slot is not None and (type(request.tray_slot) is not int or request.tray_slot < 1):
        raise ValueError('Enter the registered tray slot of the calibration standard.')
    if any(type(value) not in (int, float) or not math.isfinite(value) or value <= 0 for value in quantities):
        raise ValueError('Enter positive finite standard, scan, and calibration quantities.')
    if request.threshold_percent >= 100:
        raise ValueError('The calibration threshold must be below 100 percent.')
    if (type(coil) is not int or type(travel_limit) is not int
            or not -(2**30) < travel_limit < coil < 0):
        raise ValueError('Import the accepted negative lift and susceptibility coil positions.')

    increment_cm = request.spacing_cm / 2
    if increment_cm < .004:
        raise ValueError('The scan spacing would exceed the supported scan size.')
    fine_count = 5 + request.height_cm / increment_cm
    if fine_count >= 1002:
        raise ValueError('Increase spacing to keep the scan within 1,001 points.')
    count = max(21, math.floor(fine_count), math.ceil(4 / increment_cm) + 1)
    if count % 2 == 0:
        count += 1
    increment = increment_cm * counts_per_cm
    correction = -request.height_cm * counts_per_cm / 2
    if not 1 <= increment <= abs(travel_limit) or not -abs(travel_limit) <= correction < 0:
        raise ValueError('The scan spacing must resolve to distinct controller counts.')
    center = coil - correction
    positions = tuple(math.floor(center - index * increment) for index in range(-(count // 2), (count + 1) // 2))
    if min(positions) < travel_limit or max(positions) > -math.ceil(request.spacing_cm * counts_per_cm):
        raise ValueError('The proposed scan crosses the calibrated lift travel or top-switch margin.')
    tolerance = min(150, (min(a - b for a, b in zip(positions, positions[1:])) - 1) // 2)
    return RodScanPlan(request, coil, travel_limit, counts_per_cm, moment_factor_cgs, correction, positions, tolerance)


@dataclass(frozen=True)
class RodReading:
    target: int
    position: int
    bridge_scaled_value: float
    zero_reply: str


@dataclass(frozen=True)
class RodFit:
    coil: int
    moment_factor_cgs: float
    reference_reading_cgs: float
    verification_reading_cgs: float


class RodFitError(ValueError):
    """The completed acquisition did not support a calibration proposal."""


def scan_rod(plan: RodScanPlan, sample, readings: list[RodReading], *, cancelled) -> RodFit:
    def take(target):
        if cancelled():
            raise InterruptedError('The rod scan was canceled.')
        reading = sample(target)
        if (type(reading.position) is not int or reading.target != target
                or abs(reading.position - target) > plan.readback_tolerance
                or not isinstance(reading.zero_reply, str) or not reading.zero_reply.strip()
                or type(reading.bridge_scaled_value) not in (int, float)
                or not math.isfinite(reading.bridge_scaled_value)):
            raise ValueError('A rod scan point lacks position, zero, or bridge readback evidence.')
        readings.append(reading)
        return reading

    values = [take(position) for position in plan.positions]
    peak = max((value.bridge_scaled_value for value in values), key=abs)
    threshold = plan.request.threshold_percent / 100
    if peak <= 0 or max(abs(values[0].bridge_scaled_value), abs(values[-1].bridge_scaled_value)) >= peak:
        raise RodFitError('The scan must bracket a positive susceptibility peak.')
    near_peak = [index for index, value in enumerate(values) if value.bridge_scaled_value >= peak * (1 - threshold)]
    if (near_peak[0] == 0 or near_peak[-1] == len(values) - 1
            or len(near_peak) != near_peak[-1] - near_peak[0] + 1):
        raise RodFitError('The peak region must be separate from both scan edges and contiguous.')
    weights = [values[index].bridge_scaled_value / peak for index in near_peak]
    center = math.fsum(values[index].position * weight for index, weight in zip(near_peak, weights)) / math.fsum(weights)
    fitted_position = math.floor(center)
    coil = math.floor(center + plan.center_correction)
    if not plan.travel_limit < coil < 0:
        raise RodFitError('The proposed coil position is outside the accepted travel.')
    reference = take(fitted_position).bridge_scaled_value
    if reference <= 0:
        raise RodFitError('The fitted center must produce a positive reference reading.')
    expected = plan.request.reference_cgs
    tolerance = expected * threshold
    factor = plan.moment_factor_cgs
    if abs(reference * factor - expected) > tolerance:
        # VB6 first checks the nominal scale before fitting a new factor.
        factor = 1e-5 if abs(reference * 1e-5 - expected) <= tolerance else expected / reference
    verified = take(fitted_position).bridge_scaled_value * factor
    measured = reference * factor
    if (not math.isfinite(factor) or factor <= 0 or not math.isfinite(measured)
            or not math.isfinite(verified) or abs(verified - expected) > tolerance):
        raise RodFitError('The independent center reading did not verify the proposed calibration.')
    return RodFit(coil, factor, measured, verified)


def simulate_rod_scan(plan, readings, *, cancelled):
    center = plan.coil - plan.center_correction
    width = (max(plan.positions) - min(plan.positions)) / 6
    scale = plan.request.reference_cgs / plan.moment_factor_cgs

    def sample(position):
        raw = scale * math.exp(-.5 * ((position - center) / width) ** 2)
        return RodReading(position, position, raw, 'SIMULATED ZERO')

    return scan_rod(plan, sample, readings, cancelled=cancelled)


@dataclass(frozen=True)
class RodScanResult:
    artifact: Path
    fit: RodFit | SquidRodFit | UChannelSummary | None
    error: str
    simulated: bool
    recovery_required: bool


class RodScanJob:
    schema = ROD_SCHEMA
    procedure_id = ROD_PROCEDURE
    procedure_name = 'Susceptibility Rod Scan'
    simulate = staticmethod(simulate_rod_scan)

    def __init__(self, plan, path, *, operator, config_hash, scan=None):
        self.plan = plan
        self.path = Path(path)
        self.operator = operator
        self.config_hash = config_hash
        self.scan = scan

    def run(self, *, cancelled):
        started = datetime.now(timezone.utc).isoformat()
        simulated = self.scan is None
        readings = []
        fit, error, station_evidence = None, '', {}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            raise FileExistsError(f'The rod scan evidence already exists at {self.path}.')
        temporary = self.path.with_name(f'.{self.path.name}.{uuid4().hex}.tmp')
        handle = temporary.open('x', encoding='utf-8', newline='\n')
        try:
            with handle:
                try:
                    if simulated:
                        fit = self.simulate(self.plan, readings, cancelled=cancelled)
                    else:
                        fit = self.scan(self.plan, readings, station_evidence, cancelled=cancelled)
                    if cancelled():
                        raise InterruptedError('The rod scan was canceled.')
                except Exception as exc:
                    fit = None
                    error = str(exc) or type(exc).__name__
                payload = dict(schema=self.schema, procedure_id=self.procedure_id, procedure_name=self.procedure_name,
                               operator=self.operator, config_hash=self.config_hash, started_at=started,
                               completed_at=datetime.now(timezone.utc).isoformat(), simulated=simulated,
                               passed=fit is not None, plan=asdict(self.plan),
                               samples=[asdict(reading) for reading in readings],
                               proposal=asdict(fit) if fit is not None else None,
                               station_evidence=station_evidence, error=error)
                json.dump(payload, handle, indent=2, allow_nan=False)
                handle.write('\n')
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.path)
        finally:
            temporary.unlink(missing_ok=True)
        return RodScanResult(self.path, fit, error, simulated, bool(station_evidence.get('pending_motion_record')))


class RodScanWorker(QtCore.QThread):
    readings_recorded = QtCore.Signal(int)

    def __init__(self, job, parent=None):
        super().__init__(parent)
        self.job = job
        self.result = None
        self.error = ''

    def run(self):
        try:
            self.result = self.job.run(cancelled=self.isInterruptionRequested)
        except Exception as exc:
            self.error = str(exc) or type(exc).__name__
