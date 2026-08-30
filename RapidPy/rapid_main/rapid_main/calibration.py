"""Calibration workflow primitives for Phase 6 roadmap execution.

This module intentionally keeps the workflow contract small and deterministic:

- A calibration is represented by a small immutable result object.
- Automated runs are deterministic in simulation/backends that are pure/stable.
- Manual runs can be approved explicitly using measured values entered by the operator.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
import math
from pathlib import Path
from typing import Iterable

from .hardware_contracts import MeasurementBackend


@dataclass(frozen=True)
class CalibrationResult:
    """Result object returned for a calibration run."""

    procedure_id: str
    procedure_name: str
    mode: str
    timestamp_iso: str
    expected: float
    sample_count: int
    mean: float
    stdev: float
    max_abs_residual: float
    tolerance: float
    passed: bool
    samples: tuple[float, ...]
    notes: str = ""

    @property
    def status(self) -> str:
        return "PASS" if self.passed else "FAIL"


@dataclass(frozen=True)
class IrmVoltageCalibrationPoint:
    """One measured IRM field/drive-voltage calibration point."""

    field_mT: float
    voltage_v: float

    def __post_init__(self) -> None:
        if not math.isfinite(self.field_mT) or not math.isfinite(self.voltage_v):
            raise ValueError("IRM voltage calibration points must be finite")
        if self.field_mT < 0.0:
            raise ValueError("IRM calibration field must be non-negative")
        if self.voltage_v < 0.0:
            raise ValueError("IRM calibration voltage must be non-negative")


@dataclass(frozen=True)
class IrmVoltageRequest:
    """Validated voltage request for an IRM field target."""

    field_mT: float
    voltage_v: float
    max_voltage_v: float
    within_limits: bool


@dataclass(frozen=True)
class IrmVoltageCalibrationFit:
    """Linear field-to-voltage mapping for IRM pulse setup."""

    points: tuple[IrmVoltageCalibrationPoint, ...]
    slope_v_per_mT: float
    intercept_v: float
    r_squared: float
    max_voltage_v: float
    force_zero_intercept: bool = False

    def voltage_for_field(self, field_mT: float) -> float:
        if not math.isfinite(field_mT):
            raise ValueError("IRM field request must be finite")
        if field_mT < 0.0:
            raise ValueError("IRM field request must be non-negative")
        return self.intercept_v + self.slope_v_per_mT * field_mT

    def request_for_field(self, field_mT: float) -> IrmVoltageRequest:
        voltage = self.voltage_for_field(field_mT)
        return IrmVoltageRequest(
            field_mT=float(field_mT),
            voltage_v=voltage,
            max_voltage_v=self.max_voltage_v,
            within_limits=0.0 <= voltage <= self.max_voltage_v,
        )


def fit_irm_voltage_calibration(
    points: Iterable[IrmVoltageCalibrationPoint | tuple[float, float]],
    *,
    max_voltage_v: float = 10.0,
    force_zero_intercept: bool = False,
) -> IrmVoltageCalibrationFit:
    """Fit a linear IRM calibration from measured field/voltage points.

    The fit maps field in mT to requested drive voltage. It is intentionally
    pure and hardware-free so the calibration panel, settings import, and
    hardware acceptance workflow can share the same validation semantics.
    """

    normalized = tuple(_coerce_irm_voltage_point(point) for point in points)
    if len(normalized) < 2:
        raise ValueError("IRM voltage calibration requires at least two points")
    if not math.isfinite(max_voltage_v) or max_voltage_v <= 0.0:
        raise ValueError("max_voltage_v must be positive")

    fields = [point.field_mT for point in normalized]
    voltages = [point.voltage_v for point in normalized]
    if len(set(fields)) < 2:
        raise ValueError("IRM voltage calibration requires at least two distinct fields")

    if force_zero_intercept:
        denominator = sum(field * field for field in fields)
        if denominator <= 0.0:
            raise ValueError("zero-intercept IRM calibration requires a non-zero field")
        slope = sum(field * voltage for field, voltage in zip(fields, voltages)) / denominator
        intercept = 0.0
    else:
        field_mean = sum(fields) / len(fields)
        voltage_mean = sum(voltages) / len(voltages)
        denominator = sum((field - field_mean) ** 2 for field in fields)
        if denominator <= 0.0:
            raise ValueError("IRM voltage calibration requires field variance")
        slope = sum(
            (field - field_mean) * (voltage - voltage_mean)
            for field, voltage in zip(fields, voltages)
        ) / denominator
        intercept = voltage_mean - slope * field_mean

    predictions = [intercept + slope * field for field in fields]
    r_squared = _r_squared(voltages, predictions)
    return IrmVoltageCalibrationFit(
        points=normalized,
        slope_v_per_mT=slope,
        intercept_v=intercept,
        r_squared=r_squared,
        max_voltage_v=float(max_voltage_v),
        force_zero_intercept=force_zero_intercept,
    )


def irm_voltage_calibration_artifact(
    fit: IrmVoltageCalibrationFit,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
) -> dict[str, object]:
    """Return a deterministic JSON-ready artifact for an IRM voltage fit."""
    return {
        "procedure_id": "calibration/irm-voltage",
        "procedure_name": "IRM Voltage Calibration",
        "timestamp_iso": _now_iso(),
        "run_context": str(run_context),
        "operator": str(operator),
        "notes": str(notes),
        "slope_v_per_mT": fit.slope_v_per_mT,
        "intercept_v": fit.intercept_v,
        "r_squared": fit.r_squared,
        "max_voltage_v": fit.max_voltage_v,
        "force_zero_intercept": fit.force_zero_intercept,
        "points": [
            {"field_mT": point.field_mT, "voltage_v": point.voltage_v}
            for point in fit.points
        ],
    }


def write_irm_voltage_calibration_artifact(
    path: str | Path,
    fit: IrmVoltageCalibrationFit,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
) -> Path:
    """Write an IRM voltage calibration fit as a JSON acceptance artifact."""
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = irm_voltage_calibration_artifact(
        fit,
        run_context=run_context,
        operator=operator,
        notes=notes,
    )
    output.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return output


def calibration_result_artifact(
    result: CalibrationResult,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
) -> dict[str, object]:
    """Return a JSON-ready artifact for a generic calibration result."""

    return {
        "procedure_id": result.procedure_id,
        "procedure_name": result.procedure_name,
        "status": result.status,
        "mode": result.mode,
        "timestamp_iso": result.timestamp_iso,
        "run_context": str(run_context),
        "operator": str(operator),
        "notes": str(notes or result.notes),
        "expected": result.expected,
        "sample_count": result.sample_count,
        "mean": result.mean,
        "stdev": result.stdev,
        "max_abs_residual": result.max_abs_residual,
        "tolerance": result.tolerance,
        "passed": result.passed,
        "samples": list(result.samples),
        "hardware_validation_required": True,
        "hardware_validation_statement": (
            "Software calibration artifact only; live meter/coil/bridge "
            "acceptance must be confirmed against lab hardware and reference standards."
        ),
    }


def write_calibration_result_artifact(
    path: str | Path,
    result: CalibrationResult,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
) -> Path:
    """Write a generic calibration result as a JSON acceptance artifact."""

    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = calibration_result_artifact(
        result,
        run_context=run_context,
        operator=operator,
        notes=notes,
    )
    output.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return output


def _sample_magnitude_from_quiet(
    raw: tuple[float, float, float],
) -> float:
    """Return magnitude of a three-axis reading."""

    x, y, z = raw
    return math.sqrt(x * x + y * y + z * z)


def _stats(values: Iterable[float]) -> tuple[float, float]:
    vals = list(values)
    if not vals:
        raise ValueError("calibration requires at least one sample")

    mean = sum(vals) / len(vals)
    if len(vals) == 1:
        return mean, 0.0

    variance = sum((v - mean) ** 2 for v in vals) / len(vals)
    return mean, math.sqrt(variance)


def _coerce_irm_voltage_point(
    point: IrmVoltageCalibrationPoint | tuple[float, float],
) -> IrmVoltageCalibrationPoint:
    if isinstance(point, IrmVoltageCalibrationPoint):
        return point
    field_mT, voltage_v = point
    return IrmVoltageCalibrationPoint(float(field_mT), float(voltage_v))


def _r_squared(observed: list[float], predicted: list[float]) -> float:
    mean_observed = sum(observed) / len(observed)
    total = sum((value - mean_observed) ** 2 for value in observed)
    if total == 0.0:
        return 1.0 if all(value == predicted[0] for value in predicted) else 0.0
    residual = sum((value - fit) ** 2 for value, fit in zip(observed, predicted))
    return max(0.0, min(1.0, 1.0 - residual / total))


def run_automated_squid_baseline_calibration(
    backend: MeasurementBackend,
    *,
    procedure_id: str = "gaussmeter/squid-baseline",
    expected: float = 1.0,
    samples: int = 5,
    tolerance: float = 0.05,
    procedure_name: str = "SQUID / Gaussmeter baseline check",
) -> CalibrationResult:
    """Collect baseline samples from the active backend and evaluate residuals."""

    if samples <= 0:
        raise ValueError("samples must be positive")

    raw_samples: list[float] = []
    for _ in range(samples):
        raw_samples.append(_sample_magnitude_from_quiet(backend.read_squid()))

    mean, stdev = _stats(raw_samples)
    residual = abs(mean - expected)
    max_abs_residual = max(abs(v - expected) for v in raw_samples)
    passed = residual <= tolerance and max_abs_residual <= tolerance

    return CalibrationResult(
        procedure_id=procedure_id,
        procedure_name=procedure_name,
        mode="automated",
        timestamp_iso=_now_iso(),
        expected=expected,
        sample_count=len(raw_samples),
        mean=mean,
        stdev=stdev,
        max_abs_residual=max_abs_residual,
        tolerance=tolerance,
        passed=passed,
        samples=tuple(raw_samples),
        notes="Automated run from measurement backend SQUID path.",
    )


def run_manual_calibration(
    measured: float,
    *,
    procedure_id: str = "gaussmeter/squid-baseline",
    procedure_name: str = "SQUID / Gaussmeter baseline check",
    expected: float = 1.0,
    tolerance: float = 0.05,
    notes: str = "Manual operator entry.",
) -> CalibrationResult:
    """Record a manual entry calibration result."""

    residual = abs(measured - expected)
    return CalibrationResult(
        procedure_id=procedure_id,
        procedure_name=procedure_name,
        mode="manual",
        timestamp_iso=_now_iso(),
        expected=expected,
        sample_count=1,
        mean=measured,
        stdev=0.0,
        max_abs_residual=residual,
        tolerance=tolerance,
        passed=residual <= tolerance,
        samples=(float(measured),),
        notes=notes,
    )


def format_calibration_result(result: CalibrationResult) -> str:
    """Format a human-friendly summary for status panels."""

    return (
        f"{result.procedure_name} [{result.status}]\\n"
        f"Mode: {result.mode}\\n"
        f"Time: {result.timestamp_iso}\\n"
        f"Expected: {result.expected:.6g} "
        f"(tol ±{result.tolerance:.6g})\\n"
        f"Mean: {result.mean:.6g}\\n"
        f"Stdev: {result.stdev:.6g}\\n"
        f"Max residual: {result.max_abs_residual:.6g}\\n"
        f"Notes: {result.notes}"
    )


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()
