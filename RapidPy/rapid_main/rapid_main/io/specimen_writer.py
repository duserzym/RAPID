"""
specimen_writer.py - Write / append VB6-format specimen data files.
"""
from __future__ import annotations
from pathlib import Path
from rapid_main.data_model import MeasurementStep, SpecimenMeta


def _fmt5(value: float) -> str:
    return f"{value:5.1f}"


def write_header(path, meta: SpecimenMeta) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    comment = (meta.comment or "")[:79]
    line2 = " " * 8
    line2 += _fmt5(meta.core_plate_strike) if meta.has_core_orientation else " " * 5
    line2 += " "
    line2 += _fmt5(meta.core_plate_dip) if meta.has_core_orientation else " " * 5
    line2 += " "
    line2 += _fmt5(meta.bedding_strike)
    line2 += " "
    line2 += _fmt5(meta.bedding_dip)
    line2 += " "
    line2 += _fmt5(meta.volume)
    if meta.fold_axis is not None:
        line2 += " "
        line2 += _fmt5(meta.fold_axis)
        if meta.fold_plunge is not None:
            line2 += " "
            line2 += _fmt5(meta.fold_plunge)
    with path.open("w", encoding="latin-1", newline="\r\n") as fh:
        fh.write(comment + "\n")
        fh.write(line2 + "\n")


def append_step(path, step: MeasurementStep) -> None:
    path = Path(path)
    ts_str = step.timestamp.strftime("%Y-%m-%d %H:%M:%S")
    operator = (step.operator or "")[:8].ljust(8)
    intensity = step.intensity_emu_per_cm3 if step.intensity_emu_per_cm3 is not None else step.moment
    moment_str = f"{intensity:.2E}"
    sx, sy, sz = step.axis_sd_raw if step.axis_sd_raw is not None else (step.sdx, step.sdy, step.sdz)
    line = (
        f"{step.demag_label:<7} "
        f"{step.gdec:7.1f} "
        f"{step.ginc:6.1f} "
        f"{step.sdec:7.1f} "
        f"{step.sinc:6.1f} "
        f"{moment_str:>12} "
        f"{step.error_angle:6.1f} "
        f"{step.crdec:7.1f} "
        f"{step.crinc:6.1f} "
        f"{sx:.4E} "
        f"{sy:.4E} "
        f"{sz:.4E} "
        f"{operator} "
        f"{ts_str}"
    )
    with path.open("a", encoding="latin-1", newline="\r\n") as fh:
        fh.write(line + "\n")
