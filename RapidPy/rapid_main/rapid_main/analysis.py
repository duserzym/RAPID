"""Deterministic analysis helpers for measurement review.

The legacy VB6 ``modDataAnalysis`` routines mixed plotting, vector extraction,
and statistics inside form workflows. This module keeps the RapidPy equivalent
small and reusable: it consumes existing :class:`MeasurementStep` records and
returns plain dataclass evidence suitable for measurement output, plots, and
tests.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from statistics import mean, median, pstdev
from typing import TYPE_CHECKING, Iterable, Sequence

from rapid_main.data_model import MeasurementStep

if TYPE_CHECKING:
    import numpy as np


Vector3 = tuple[float, float, float]


@dataclass(frozen=True)
class MeasurementVector:
    """Cartesian vector extracted from one measurement step."""

    label: str
    vector: Vector3
    moment: float

    @property
    def magnitude(self) -> float:
        x, y, z = self.vector
        return math.sqrt(x * x + y * y + z * z)


@dataclass(frozen=True)
class PrincipalAxisFit:
    """Best-fit line through measurement vectors with population variances."""

    count: int
    centroid: Vector3
    axis: Vector3
    declination_deg: float
    inclination_deg: float
    primary_variance: float
    total_variance: float
    variance_fraction: float
    rms_perpendicular: float
    mad_deg: float | None = None
    dang_deg: float | None = None

    @property
    def is_degenerate(self) -> bool:
        return self.count < 2 or self.variance_fraction <= 0.0


@dataclass(frozen=True)
class CartesianComponent:
    """Principal component in the same Cartesian frame as its inputs."""

    center: np.ndarray
    direction: np.ndarray
    eigenvalues: np.ndarray
    variance_fraction: float
    rms_perpendicular: float
    mad_deg: float
    dang_deg: float


@dataclass(frozen=True)
class MomentDecaySummary:
    """Evidence for monotonic moment loss across a treatment sequence."""

    initial: float | None
    final: float | None
    final_ratio: float | None
    percent_loss: float | None
    per_step_ratios: tuple[float | None, ...]
    monotonic_nonincreasing: bool
    log_decay_slope: float | None


@dataclass(frozen=True)
class MomentStatistics:
    """Summary statistics for measured moments."""

    count: int
    minimum: float | None
    maximum: float | None
    mean: float | None
    median: float | None
    population_stdev: float | None
    decay: MomentDecaySummary = field(default_factory=lambda: moment_decay_summary([]))


@dataclass(frozen=True)
class ReadingCycleStatistics:
    """Quality evidence calculated from one repeated SQUID reading cycle."""

    count: int
    mean_vector: Vector3
    axis_ranges: Vector3
    mean_magnitude: float
    rms_spread: float
    directional_spread_deg: float | None
    signal_to_drift: float | None


def extract_measurement_vectors(steps: Iterable[MeasurementStep]) -> list[MeasurementVector]:
    """Return specimen-frame vectors from ``MeasurementStep.sdx/sdy/sdz``.

    The function intentionally trusts the existing measurement record fields
    instead of recalculating vectors from angles. That preserves the exact
    Cartesian values written by measurement output paths.
    """

    return [
        MeasurementVector(
            label=str(step.demag_label),
            vector=(float(step.sdx), float(step.sdy), float(step.sdz)),
            moment=float(step.moment),
        )
        for step in steps
    ]


def vector_centroid(vectors: Iterable[MeasurementVector | Vector3]) -> Vector3:
    """Return the arithmetic centroid for vectors or extracted records."""

    points = [_coerce_vector(vector) for vector in vectors]
    if not points:
        return (0.0, 0.0, 0.0)
    return (
        sum(point[0] for point in points) / len(points),
        sum(point[1] for point in points) / len(points),
        sum(point[2] for point in points) / len(points),
    )


def fit_cartesian_component(
    vectors: Iterable[Sequence[float]], *, include_origin: bool = False,
) -> CartesianComponent | None:
    """Fit a line through the centroid or constrained to pass through the origin."""
    samples = list(vectors)
    if len(samples) < 2:
        return None
    import numpy as np

    points = np.array(samples, dtype=float)
    if not np.isfinite(points).all():
        return None
    scale = float(np.max(np.abs(points)))
    if scale == 0.0 or (not include_origin and np.all(points == points[0])):
        return None
    # Fit on a unitless scale so the result is independent of moment units.
    points /= scale
    mean = points.mean(axis=0)
    center = np.zeros(3, dtype=float) if include_origin else mean
    centered = points - center
    eigenvalues, eigenvectors = np.linalg.eigh(centered.T @ centered)
    eigenvalues = np.maximum(eigenvalues[::-1], 0.0)
    primary, second, third = map(float, eigenvalues)
    if primary <= 0.0:
        return None
    direction = eigenvectors[:, -1]
    # Free fits point along the remanence removed between the endpoint steps.
    # Anchored fits point away from the origin toward the measured vectors.
    removed_component = points[0] - points[-1]
    reference = mean if include_origin else removed_component
    alignment = float(np.dot(direction, reference))
    if include_origin and alignment == 0.0:
        alignment = float(np.dot(direction, removed_component))
    if alignment < 0.0:
        direction = -direction

    perpendicular = second + third
    mad = math.degrees(math.atan(math.sqrt(perpendicular / primary)))
    center_norm = float(np.linalg.norm(center))
    dang = 0.0
    if center_norm > 0.0:
        cosine = float(np.dot(center, direction)) / center_norm
        dang = math.degrees(math.acos(max(-1.0, min(1.0, cosine))))
    residuals = centered - np.outer(centered @ direction, direction)
    return CartesianComponent(
        center=center * scale,
        direction=direction,
        eigenvalues=(eigenvalues * scale) * scale,
        variance_fraction=primary / (primary + perpendicular),
        rms_perpendicular=scale * math.sqrt(float(np.sum(residuals**2)) / len(points)),
        mad_deg=mad,
        dang_deg=dang,
    )


def fit_principal_axis(steps_or_vectors: Iterable[MeasurementStep | MeasurementVector | Vector3]) -> PrincipalAxisFit:
    """Report the free principal-component fit in North/East/Down coordinates."""

    points = [_coerce_step_or_vector(item) for item in steps_or_vectors]
    count = len(points)
    component = fit_cartesian_component(points)
    if component is None:
        centroid = vector_centroid(points)
        axis = _stable_axis((1.0, 0.0, 0.0), centroid, points)
        dec, inc = vector_orientation(axis)
        return PrincipalAxisFit(count, centroid, axis, dec, inc, 0.0, 0.0, 0.0, 0.0)

    centroid = tuple(float(value) for value in component.center)
    axis = tuple(float(value) for value in component.direction)
    variances = component.eigenvalues / count
    dec, inc = vector_orientation(axis)
    return PrincipalAxisFit(
        count=count,
        centroid=centroid,
        axis=axis,
        declination_deg=dec,
        inclination_deg=inc,
        primary_variance=float(variances[0]),
        total_variance=float(variances.sum()),
        variance_fraction=component.variance_fraction,
        rms_perpendicular=component.rms_perpendicular,
        mad_deg=component.mad_deg,
        dang_deg=component.dang_deg,
    )


def vector_orientation(vector: Vector3) -> tuple[float, float]:
    """Return declination and inclination for a North/East/Down vector."""

    x, y, z = vector
    horizontal = math.hypot(x, y)
    if horizontal == 0.0 and z == 0.0:
        return (0.0, 0.0)
    declination = math.degrees(math.atan2(y, x)) % 360.0
    inclination = math.degrees(math.atan2(z, horizontal))
    return (declination, inclination)


def moment_statistics(steps: Iterable[MeasurementStep]) -> MomentStatistics:
    """Return deterministic moment statistics and decay evidence."""

    return moment_value_statistics(float(step.moment) for step in steps)


def moment_value_statistics(moments: Iterable[float]) -> MomentStatistics:
    """Return the same statistics for already-extracted moment magnitudes."""

    moments = [float(moment) for moment in moments]
    if not moments:
        return MomentStatistics(0, None, None, None, None, None, moment_decay_summary([]))
    return MomentStatistics(
        count=len(moments),
        minimum=min(moments),
        maximum=max(moments),
        mean=mean(moments),
        median=median(moments),
        population_stdev=pstdev(moments),
        decay=moment_decay_summary(moments),
    )


def reading_cycle_statistics(readings: Iterable[Vector3]) -> ReadingCycleStatistics:
    """Summarize repeated calibrated SQUID readings for one measurement step.

    The values use the same emu vector convention as ``MeasurementStep``. The
    axis ranges correspond to the legacy per-axis delta displays, while the
    spread and signal-to-drift ratio provide the evidence needed to flag an
    unstable read cycle before the output bundle is written.
    """

    samples = [_coerce_vector(reading) for reading in readings]
    if not samples:
        raise ValueError("reading cycle requires at least one sample")
    if not all(math.isfinite(component) for sample in samples for component in sample):
        raise ValueError("reading cycle contains non-finite values")

    mean_vector = vector_centroid(samples)
    axis_ranges = tuple(
        max(sample[index] for sample in samples) - min(sample[index] for sample in samples)
        for index in range(3)
    )
    mean_magnitude = _norm(mean_vector)
    rms_spread = math.sqrt(
        sum(_norm(_sub(sample, mean_vector)) ** 2 for sample in samples) / len(samples)
    )
    signal_to_drift = (
        None
        if mean_magnitude == 0.0
        else math.inf
        if rms_spread == 0.0
        else mean_magnitude / rms_spread
    )
    directional_spread = _directional_spread_deg(samples, mean_vector)
    return ReadingCycleStatistics(
        count=len(samples),
        mean_vector=mean_vector,
        axis_ranges=(axis_ranges[0], axis_ranges[1], axis_ranges[2]),
        mean_magnitude=mean_magnitude,
        rms_spread=rms_spread,
        directional_spread_deg=directional_spread,
        signal_to_drift=signal_to_drift,
    )


def moment_decay_summary(moments_or_steps: Iterable[float | MeasurementStep]) -> MomentDecaySummary:
    """Summarize treatment-to-treatment moment decay.

    Ratios are ``current / previous``. A ratio is ``None`` when the previous
    moment is zero. The log-decay slope is a least-squares slope of
    ``ln(moment / initial)`` over step index and is only reported when all
    moments are positive.
    """

    moments = [
        float(item.moment) if isinstance(item, MeasurementStep) else float(item)
        for item in moments_or_steps
    ]
    if not moments:
        return MomentDecaySummary(None, None, None, None, (), True, None)

    initial = moments[0]
    final = moments[-1]
    final_ratio = None if initial == 0.0 else final / initial
    percent_loss = None if initial == 0.0 else (1.0 - final_ratio) * 100.0
    ratios = tuple(
        None if previous == 0.0 else current / previous
        for previous, current in zip(moments, moments[1:])
    )
    monotonic = all(current <= previous for previous, current in zip(moments, moments[1:]))
    return MomentDecaySummary(
        initial=initial,
        final=final,
        final_ratio=final_ratio,
        percent_loss=percent_loss,
        per_step_ratios=ratios,
        monotonic_nonincreasing=monotonic,
        log_decay_slope=_log_decay_slope(moments),
    )


def _coerce_step_or_vector(item: MeasurementStep | MeasurementVector | Vector3) -> Vector3:
    if isinstance(item, MeasurementStep):
        return (float(item.sdx), float(item.sdy), float(item.sdz))
    return _coerce_vector(item)


def _coerce_vector(item: MeasurementVector | Vector3) -> Vector3:
    if isinstance(item, MeasurementVector):
        return item.vector
    x, y, z = item
    return (float(x), float(y), float(z))


def _stable_axis(axis: Vector3, centroid: Vector3, points: Sequence[Vector3]) -> Vector3:
    reference = centroid if _norm(centroid) > 0.0 else next((point for point in points if _norm(point) > 0.0), (1.0, 0.0, 0.0))
    axis = _normalize(axis)
    return _scale(axis, -1.0) if _dot(axis, reference) < 0.0 else axis


def _directional_spread_deg(samples: Sequence[Vector3], mean_vector: Vector3) -> float | None:
    """Return the RMS angular spread around a cycle's mean direction."""

    mean_magnitude = _norm(mean_vector)
    if mean_magnitude == 0.0:
        return None

    angles: list[float] = []
    for sample in samples:
        sample_magnitude = _norm(sample)
        if sample_magnitude == 0.0:
            continue
        cosine = _dot(sample, mean_vector) / (sample_magnitude * mean_magnitude)
        angles.append(math.degrees(math.acos(max(-1.0, min(1.0, cosine)))))
    if not angles:
        return None
    return math.sqrt(sum(angle * angle for angle in angles) / len(angles))


def _log_decay_slope(moments: Sequence[float]) -> float | None:
    if len(moments) < 2 or any(moment <= 0.0 for moment in moments):
        return None
    initial = moments[0]
    values = [math.log(moment / initial) for moment in moments]
    indexes = list(range(len(values)))
    index_mean = mean(indexes)
    value_mean = mean(values)
    denominator = sum((index - index_mean) ** 2 for index in indexes)
    if denominator == 0.0:
        return None
    return sum((index - index_mean) * (value - value_mean) for index, value in zip(indexes, values)) / denominator


def _sub(left: Vector3, right: Vector3) -> Vector3:
    return (left[0] - right[0], left[1] - right[1], left[2] - right[2])


def _scale(vector: Vector3, factor: float) -> Vector3:
    return (vector[0] * factor, vector[1] * factor, vector[2] * factor)


def _dot(left: Vector3, right: Vector3) -> float:
    return left[0] * right[0] + left[1] * right[1] + left[2] * right[2]


def _norm(vector: Vector3) -> float:
    return math.sqrt(_dot(vector, vector))


def _normalize(vector: Vector3) -> Vector3:
    norm = _norm(vector)
    if norm == 0.0:
        return (1.0, 0.0, 0.0)
    return (vector[0] / norm, vector[1] / norm, vector[2] / norm)
