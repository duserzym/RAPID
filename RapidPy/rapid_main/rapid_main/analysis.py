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
from typing import Iterable, Sequence

from rapid_main.data_model import MeasurementStep


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
    """PCA-style best-fit line through measurement vectors."""

    count: int
    centroid: Vector3
    axis: Vector3
    declination_deg: float
    inclination_deg: float
    primary_variance: float
    total_variance: float
    variance_fraction: float
    rms_perpendicular: float

    @property
    def is_degenerate(self) -> bool:
        return self.count < 2 or self.total_variance <= 0.0


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


def fit_principal_axis(steps_or_vectors: Iterable[MeasurementStep | MeasurementVector | Vector3]) -> PrincipalAxisFit:
    """Fit a deterministic PCA-style line to measurement vectors.

    The returned axis is the dominant eigenvector of the centered covariance
    matrix. Its sign is stabilized toward the centroid, or the first non-zero
    vector when the centroid is zero, so repeated runs produce identical output.
    """

    points = [_coerce_step_or_vector(item) for item in steps_or_vectors]
    count = len(points)
    centroid = vector_centroid(points)
    if count < 2:
        axis = _stable_axis((1.0, 0.0, 0.0), centroid, points)
        dec, inc = vector_orientation(axis)
        return PrincipalAxisFit(count, centroid, axis, dec, inc, 0.0, 0.0, 0.0, 0.0)

    centered = [_sub(point, centroid) for point in points]
    covariance = _covariance(centered)
    total_variance = covariance[0][0] + covariance[1][1] + covariance[2][2]
    if total_variance <= 0.0:
        axis = _stable_axis((1.0, 0.0, 0.0), centroid, points)
        dec, inc = vector_orientation(axis)
        return PrincipalAxisFit(count, centroid, axis, dec, inc, 0.0, 0.0, 0.0, 0.0)

    axis = _dominant_eigenvector(covariance)
    axis = _stable_axis(axis, centroid, points)
    primary_variance = _rayleigh(covariance, axis)
    variance_fraction = max(0.0, min(1.0, primary_variance / total_variance))
    residuals = [_perpendicular_distance(point, centroid, axis) for point in points]
    rms_perpendicular = math.sqrt(sum(value * value for value in residuals) / count)
    dec, inc = vector_orientation(axis)
    return PrincipalAxisFit(
        count=count,
        centroid=centroid,
        axis=axis,
        declination_deg=dec,
        inclination_deg=inc,
        primary_variance=primary_variance,
        total_variance=total_variance,
        variance_fraction=variance_fraction,
        rms_perpendicular=rms_perpendicular,
    )


def vector_orientation(vector: Vector3) -> tuple[float, float]:
    """Return specimen-frame declination and inclination in degrees."""

    x, y, z = vector
    horizontal = math.hypot(x, y)
    if horizontal == 0.0 and z == 0.0:
        return (0.0, 0.0)
    declination = math.degrees(math.atan2(y, x)) % 360.0
    inclination = math.degrees(math.atan2(z, horizontal))
    return (declination, inclination)


def moment_statistics(steps: Iterable[MeasurementStep]) -> MomentStatistics:
    """Return deterministic moment statistics and decay evidence."""

    moments = [float(step.moment) for step in steps]
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


def _covariance(centered: Sequence[Vector3]) -> tuple[Vector3, Vector3, Vector3]:
    count = len(centered)
    return (
        (
            sum(point[0] * point[0] for point in centered) / count,
            sum(point[0] * point[1] for point in centered) / count,
            sum(point[0] * point[2] for point in centered) / count,
        ),
        (
            sum(point[1] * point[0] for point in centered) / count,
            sum(point[1] * point[1] for point in centered) / count,
            sum(point[1] * point[2] for point in centered) / count,
        ),
        (
            sum(point[2] * point[0] for point in centered) / count,
            sum(point[2] * point[1] for point in centered) / count,
            sum(point[2] * point[2] for point in centered) / count,
        ),
    )


def _dominant_eigenvector(matrix: tuple[Vector3, Vector3, Vector3]) -> Vector3:
    vector = _normalize((1.0, 1.0, 1.0))
    for _ in range(32):
        next_vector = _matvec(matrix, vector)
        if _norm(next_vector) == 0.0:
            return (1.0, 0.0, 0.0)
        vector = _normalize(next_vector)
    return vector


def _stable_axis(axis: Vector3, centroid: Vector3, points: Sequence[Vector3]) -> Vector3:
    reference = centroid if _norm(centroid) > 0.0 else next((point for point in points if _norm(point) > 0.0), (1.0, 0.0, 0.0))
    axis = _normalize(axis)
    return _scale(axis, -1.0) if _dot(axis, reference) < 0.0 else axis


def _rayleigh(matrix: tuple[Vector3, Vector3, Vector3], axis: Vector3) -> float:
    return _dot(axis, _matvec(matrix, axis))


def _perpendicular_distance(point: Vector3, centroid: Vector3, axis: Vector3) -> float:
    delta = _sub(point, centroid)
    parallel = _scale(axis, _dot(delta, axis))
    return _norm(_sub(delta, parallel))


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


def _matvec(matrix: tuple[Vector3, Vector3, Vector3], vector: Vector3) -> Vector3:
    return (
        _dot(matrix[0], vector),
        _dot(matrix[1], vector),
        _dot(matrix[2], vector),
    )


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
