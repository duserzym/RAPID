from __future__ import annotations

import json
import math
import os
import tempfile
from pathlib import Path
from typing import Sequence

from rapid_main.analysis import fit_principal_axis, moment_value_statistics
from rapid_main.data_model import MeasurementStep


def _cart_to_inc_dec(x: float, y: float, z: float) -> tuple[float, float]:
    """Convert Cartesian (N, E, Up) to (Inc, Dec) in degrees."""
    h = math.hypot(x, y)
    inc = math.degrees(math.atan2(-z, h))  # positive = below horizontal
    dec = math.degrees(math.atan2(y, x)) % 360.0
    return inc, dec


def build_quicklook_summary(
    north: Sequence[float],
    east: Sequence[float],
    up: Sequence[float],
    labels: Sequence[str],
) -> dict[str, object]:
    """Build the backend-independent quicklook payload without creating a dialog."""

    n = list(north)
    e = list(east)
    u = list(up)
    step_labels = list(labels)
    lengths = {len(n), len(e), len(u), len(step_labels)}
    if len(lengths) != 1:
        raise ValueError(
            "Quicklook north/east/up vectors and labels must have equal lengths "
            f"(got {len(n)}, {len(e)}, {len(u)}, {len(step_labels)})."
        )
    for axis, values in (("north", n), ("east", e), ("up", u)):
        for index, value in enumerate(values, start=1):
            if not math.isfinite(float(value)):
                raise ValueError(f"Quicklook {axis} value {index} is not finite.")
    down = [-z for z in u]
    intensity = [math.sqrt(x**2 + y**2 + z**2) for x, y, z in zip(n, e, u)]
    inc: list[float] = []
    dec: list[float] = []
    for x, y, z in zip(n, e, u):
        i, d = _cart_to_inc_dec(x, y, z)
        inc.append(i)
        dec.append(d)
    principal = fit_principal_axis(zip(n, e, down))
    moment_stats = moment_value_statistics(intensity)
    decay = moment_stats.decay
    return {
        "step_count": len(step_labels),
        "labels": step_labels,
        "vectors": {
            "north": n,
            "east": e,
            "up": u,
            "down": down,
        },
        "intensity": intensity,
        "inclination": inc,
        "declination": dec,
        "analysis": {
            "principal_axis": {
                "coordinate_convention": "north_east_down",
                "count": principal.count,
                "centroid": list(principal.centroid),
                "axis": list(principal.axis),
                "declination_deg": principal.declination_deg,
                "inclination_deg": principal.inclination_deg,
                "primary_variance": principal.primary_variance,
                "total_variance": principal.total_variance,
                "variance_fraction": principal.variance_fraction,
                "rms_perpendicular": principal.rms_perpendicular,
                "degenerate": principal.is_degenerate,
            },
            "moment_statistics": {
                "count": moment_stats.count,
                "minimum": moment_stats.minimum,
                "maximum": moment_stats.maximum,
                "mean": moment_stats.mean,
                "median": moment_stats.median,
                "population_stdev": moment_stats.population_stdev,
            },
            "decay": {
                "initial": decay.initial,
                "final": decay.final,
                "final_ratio": decay.final_ratio,
                "percent_loss": decay.percent_loss,
                "per_step_ratios": list(decay.per_step_ratios),
                "monotonic_nonincreasing": decay.monotonic_nonincreasing,
                "log_decay_slope": decay.log_decay_slope,
            },
        },
    }


def write_quicklook_json(path: str | Path, summary: dict[str, object]) -> Path:
    """Write a quicklook summary as a deterministic JSON sidecar artifact."""

    target = Path(path)
    return _atomic_write_text(
        target,
        json.dumps(summary, indent=2, sort_keys=True) + "\n",
    )


def _atomic_write_text(target: Path, text: str) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="",
            prefix=f".{target.name}.",
            suffix=".tmp",
            dir=target.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, target)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
    return target


def write_measurement_quicklook(
    path: Path, steps: Sequence[MeasurementStep], *, simulated: bool
) -> Path:
    """Write the current run's plot data with its measurement provenance."""
    summary = build_quicklook_summary(
        [step.sdx for step in steps],
        [step.sdy for step in steps],
        [step.sdz for step in steps],
        [step.demag_label for step in steps],
    )
    summary["provenance"] = {
        "kind": "simulated" if simulated else "measurement",
        "simulated": simulated,
        "statement": (
            "Current run used a simulated backend; not hardware evidence."
            if simulated
            else "Measurement quicklook data."
        ),
    }
    return write_quicklook_json(path, summary)
