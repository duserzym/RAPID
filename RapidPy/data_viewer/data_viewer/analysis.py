from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Iterable, Literal, Sequence

import numpy as np

from data_viewer.data_loading import MeasurementStep, PaleointensityPoint, ViewerSpecimen
from rapid_main.analysis import fit_cartesian_component
from rapid_main.geometry import Angular3D, angular3d_to_viewer_cartesian3d


CoordinateSystem = Literal["specimen", "geographic", "tilt-corrected"]


@dataclass
class FitResult:
    center: np.ndarray
    direction: np.ndarray
    eigenvalues: np.ndarray
    dec: float
    inc: float
    mad: float
    dang: float


@dataclass
class Suggestion:
    title: str
    suggested_step: str
    reasons: list[str]
    confidence: str


@dataclass
class PaleointensitySummary:
    slope: float | None
    intercept: float | None
    point_count: int
    temperature_span: tuple[float, float] | None


def vector_for_step(step: MeasurementStep, coordinate_system: CoordinateSystem = "specimen") -> np.ndarray:
    if coordinate_system == "specimen":
        return _dir_to_cartesian(step.crdec, step.crinc, step.moment)
    if coordinate_system == "geographic":
        return _dir_to_cartesian(step.gdec, step.ginc, step.moment)
    if coordinate_system == "tilt-corrected":
        return _dir_to_cartesian(step.sdec, step.sinc, step.moment)
    raise ValueError(f"Unsupported coordinate system: {coordinate_system}")


def principal_component_fit(
    steps: Iterable[MeasurementStep],
    coordinate_system: CoordinateSystem = "specimen",
    include_origin: bool = False,
) -> FitResult | None:
    component = fit_cartesian_component(
        (vector_for_step(step, coordinate_system) for step in steps), include_origin=include_origin,
    )
    if component is None:
        return None
    dec, inc = _cartesian_to_dir(component.direction)
    return FitResult(
        center=component.center, direction=component.direction, eigenvalues=component.eigenvalues,
        dec=dec, inc=inc, mad=component.mad_deg, dang=component.dang_deg,
    )


def principal_component_segment(
    steps: Sequence[MeasurementStep],
    coordinate_system: CoordinateSystem = "specimen",
    include_origin: bool = False,
) -> tuple[np.ndarray, np.ndarray] | None:
    """Return the projected line segment in first-to-last treatment order."""
    fit = principal_component_fit(steps, coordinate_system, include_origin)
    if fit is None:
        return None
    distances = [float(np.dot(vector_for_step(step, coordinate_system) - fit.center, fit.direction))
                 for step in steps]
    descending = distances[0] > distances[-1]
    if include_origin:
        distances.append(0.0)
    start, end = min(distances), max(distances)
    if descending:
        start, end = end, start
    return (
        fit.center + fit.direction * start,
        fit.center + fit.direction * end,
    )


def stereonet_csd_arcs(
    dec_deg: float, inc_deg: float, csd_deg: float,
) -> list[tuple[bool, list[tuple[float, float]]]]:
    """Project the angular CSD boundary in pieces split at the equator.

    Each piece returns its lower-hemisphere flag and equal-area east/north points.
    """
    if not (math.isfinite(dec_deg) and -90 <= inc_deg <= 90 and 0 < csd_deg < 180):
        return []

    dec, inc, angle = map(math.radians, (dec_deg, inc_deg, csd_deg))
    sin_dec, cos_dec = math.sin(dec), math.cos(dec)
    sin_inc, cos_inc = math.sin(inc), math.cos(inc)
    sin_angle, cos_angle = math.sin(angle), math.cos(angle)
    depth_offset = sin_inc * cos_angle
    depth_amplitude = cos_inc * sin_angle
    breaks = [0.0, math.tau]
    if abs(depth_offset) < depth_amplitude:
        crossing = math.asin(-depth_offset / depth_amplitude)
        breaks = sorted({*breaks, crossing % math.tau, (math.pi - crossing) % math.tau})

    arcs = []
    for begin, end in zip(breaks, breaks[1:]):
        lower = depth_offset + depth_amplitude * math.sin((begin + end) / 2) >= 0
        points = []
        intervals = math.ceil((end - begin) / math.radians(5))
        for index in range(intervals + 1):
            phase = begin + (end - begin) * index / intervals
            along = cos_inc * cos_angle - sin_inc * sin_angle * math.sin(phase)
            across = math.cos(phase) * sin_angle
            depth = depth_offset + depth_amplitude * math.sin(phase)
            scale = 1 / math.sqrt(1 + abs(depth))
            points.append(((sin_dec * along + cos_dec * across) * scale,
                           (cos_dec * along - sin_dec * across) * scale))
        arcs.append((lower, points))
    return arcs
def summarize_paleointensity(points: Iterable[PaleointensityPoint]) -> PaleointensitySummary:
    series = [point for point in points if point.step_kind in {"ZF", "IF"}]
    if len(series) < 2:
        return PaleointensitySummary(None, None, len(series), None)
    xs = np.array([point.ptrm_gained for point in series], dtype=float)
    ys = np.array([point.nrm_remaining for point in series], dtype=float)
    slope, intercept = np.polyfit(xs, ys, 1)
    temps = [point.temperature_c for point in series]
    return PaleointensitySummary(float(slope), float(intercept), len(series), (min(temps), max(temps)))


def next_step_suggestion(specimen: ViewerSpecimen) -> Suggestion:
    exp_type = specimen.experiment_type
    steps = specimen.steps
    if len(steps) < 2:
        return Suggestion(
            title="Acquire another measurement",
            suggested_step="Add at least one demagnetization step beyond NRM.",
            reasons=["A single step is not enough to estimate direction change or remaining magnetization."],
            confidence="low",
        )
    if exp_type == "IZZI Thellier":
        return _suggest_izzi(steps, specimen.paleointensity_points)
    if exp_type == "AF":
        return _suggest_af(steps)
    if exp_type == "Thermal":
        return _suggest_thermal(steps)
    if exp_type == "AF + thermal":
        last = steps[-1]
        if last.treat_temp_K() is not None:
            base = _suggest_thermal(steps)
            base.reasons.insert(0, "The sequence is already in its thermal segment, so the next guidance is based on the most recent heating trend.")
            return base
        base = _suggest_af(steps)
        base.reasons.append("Once the low-coercivity overprint is mostly removed, the next phase is usually a thermal segment rather than larger AF jumps.")
        return base
    return Suggestion(
        title="Review file labels",
        suggested_step="The experiment type could not be auto-detected from the current step labels.",
        reasons=["Use AFxx, TTxxx, ZFxxx, IFxxx, or PTRMxxx-style labels to enable automated guidance."],
        confidence="low",
    )


def _suggest_af(steps: list[MeasurementStep]) -> Suggestion:
    af_steps = [step for step in steps if step.treat_ac_field_mT() is not None]
    if not af_steps:
        return Suggestion("Start AF demagnetization", "Add an AF step such as 5-10 mT.", ["No AF treatment steps were found yet."], "medium")
    values = [step.treat_ac_field_mT() or 0.0 for step in af_steps]
    last_value = values[-1]
    spacing = _recent_spacing(values, default=5.0)
    remaining = steps[-1].moment / max(steps[0].moment, 1e-15)
    turn = _direction_change_deg(steps)
    if remaining > 0.7 and turn < 10.0:
        increment = max(spacing, 10.0)
        confidence = "high"
        reasons = [
            "Most of the NRM is still present, so you can move faster through the low-coercivity part of the sequence.",
            f"The last directional change was only about {turn:.1f} degrees, which suggests the component is still stable at the current AF spacing.",
        ]
    elif remaining > 0.3:
        increment = max(5.0, min(spacing, 10.0))
        confidence = "medium"
        reasons = [
            "The specimen is actively demagnetizing, so a moderate AF increment keeps enough resolution to catch curvature in the Zijderveld path.",
            f"About {remaining * 100:.0f}% of the starting moment remains, which is a good range for 5-10 mT steps.",
        ]
    elif remaining > 0.12:
        increment = 5.0
        confidence = "medium"
        reasons = [
            "The remanence is getting weak, so smaller AF increments reduce the chance of stepping past the end of the stable component.",
            f"The last directional change was {turn:.1f} degrees, so fine resolution is more useful than speed now.",
        ]
    else:
        return Suggestion(
            title="Consider stopping or taking one cleanup AF step",
            suggested_step=f"If the specimen is still measurable, one final AF step around {last_value + 5:.0f} mT is enough before stopping.",
            reasons=[
                f"Only about {remaining * 100:.0f}% of the starting moment remains.",
                "Further AF steps are less likely to improve the interpretation than to add noise.",
            ],
            confidence="medium",
        )
    return Suggestion(
        title="Recommended next AF step",
        suggested_step=f"Try the next AF step at about {last_value + increment:.0f} mT.",
        reasons=reasons,
        confidence=confidence,
    )


def _suggest_thermal(steps: list[MeasurementStep]) -> Suggestion:
    thermal_steps = [step for step in steps if step.treat_temp_K() is not None]
    if not thermal_steps:
        return Suggestion("Start thermal demagnetization", "Add a heating step such as 100-150 C.", ["No thermal treatment steps were found yet."], "medium")
    values = [(step.treat_temp_K() or 273.15) - 273.15 for step in thermal_steps]
    last_value = values[-1]
    spacing = _recent_spacing(values, default=25.0)
    remaining = steps[-1].moment / max(steps[0].moment, 1e-15)
    turn = _direction_change_deg(steps)
    if remaining > 0.75 and turn < 8.0:
        increment = max(25.0, min(50.0, spacing or 50.0))
        reasons = [
            "The specimen still carries most of its original remanence, so broader temperature spacing is efficient early in the run.",
            f"The last directional change was only {turn:.1f} degrees, which suggests you have not yet reached the main unblocking interval.",
        ]
        confidence = "high"
    elif remaining > 0.35:
        increment = max(20.0, min(35.0, spacing or 25.0))
        reasons = [
            "The specimen is in the middle of the thermal unblocking path, so moderate temperature spacing balances speed and directional resolution.",
            f"Roughly {remaining * 100:.0f}% of the original moment remains.",
        ]
        confidence = "medium"
    elif remaining > 0.12:
        increment = 15.0 if turn > 12.0 else 20.0
        reasons = [
            "The remaining remanence is weaker, so smaller thermal increments help avoid overshooting the high-temperature component.",
            f"The latest change in direction was {turn:.1f} degrees.",
        ]
        confidence = "medium"
    else:
        return Suggestion(
            title="High-temperature tail is nearly exhausted",
            suggested_step=f"Consider stopping here or taking one final step near {last_value + 15:.0f} C if the signal-to-noise ratio is still good.",
            reasons=[
                f"Only about {remaining * 100:.0f}% of the starting moment remains.",
                "Additional heating now is more likely to damage or destabilize the specimen than clarify the trend.",
            ],
            confidence="medium",
        )
    return Suggestion(
        title="Recommended next thermal step",
        suggested_step=f"Try the next heating step at about {last_value + increment:.0f} C.",
        reasons=reasons,
        confidence=confidence,
    )


def _suggest_izzi(steps: list[MeasurementStep], points: list[PaleointensityPoint]) -> Suggestion:
    summary = summarize_paleointensity(points)
    thermal_labels = [step.demag_label for step in steps if any(step.demag_label.upper().startswith(prefix) for prefix in ("ZF", "IF", "PTRM", "IZ", "ZI"))]
    last_temp = 0.0
    if thermal_labels:
        last_temp = max(_extract_temperature(label) for label in thermal_labels)
    remaining = steps[-1].moment / max(steps[0].moment, 1e-15)
    increment = 25.0 if remaining > 0.25 else 15.0
    reasons = [
        "IZZI/Thellier experiments are most informative when the next zero-field and in-field pair stays close enough to preserve Arai-plot resolution.",
        f"About {remaining * 100:.0f}% of the original remanence remains, which supports another controlled temperature pair.",
    ]
    if summary.slope is not None:
        reasons.append(
            f"The current Arai-fit slope is {summary.slope:.2f} over {summary.point_count} points, so another nearby pair will help confirm whether that segment stays linear."
        )
    reasons.append("Add a pTRM check every 2-3 temperature levels if alteration is a concern.")
    return Suggestion(
        title="Recommended next IZZI temperature pair",
        suggested_step=f"Try the next zero-field / in-field pair near {last_temp + increment:.0f} C, then reassess the Arai slope and alteration checks.",
        reasons=reasons,
        confidence="medium",
    )


def _recent_spacing(values: list[float], default: float) -> float:
    if len(values) < 2:
        return default
    diffs = [values[index] - values[index - 1] for index in range(1, len(values)) if values[index] > values[index - 1]]
    if not diffs:
        return default
    return diffs[-1]


def _direction_change_deg(steps: list[MeasurementStep]) -> float:
    if len(steps) < 2:
        return 0.0
    v1 = vector_for_step(steps[-2])
    v2 = vector_for_step(steps[-1])
    n1 = float(np.linalg.norm(v1))
    n2 = float(np.linalg.norm(v2))
    if n1 <= 0 or n2 <= 0:
        return 0.0
    dot = float(np.dot(v1, v2) / (n1 * n2))
    return math.degrees(math.acos(np.clip(dot, -1.0, 1.0)))


def _cartesian_to_dir(vector: np.ndarray) -> tuple[float, float]:
    norm = float(np.linalg.norm(vector))
    if norm <= 0:
        return 0.0, 0.0
    dec = math.degrees(math.atan2(float(vector[1]), float(vector[0]))) % 360.0
    inc = math.degrees(math.asin(-float(vector[2]) / norm))
    return dec, inc


def _dir_to_cartesian(dec_deg: float, inc_deg: float, moment: float) -> np.ndarray:
    vector = angular3d_to_viewer_cartesian3d(Angular3D(dec_deg, inc_deg, moment))
    return np.array([vector.x, vector.y, vector.z], dtype=float)


def _extract_temperature(label: str) -> float:
    digits = []
    started = False
    for char in label:
        if char.isdigit() or (char == "." and started):
            digits.append(char)
            started = True
        elif started:
            break
    if not digits:
        return 0.0
    return float("".join(digits))
