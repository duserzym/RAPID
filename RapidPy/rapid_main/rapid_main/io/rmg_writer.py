"""
rmg_writer.py — Append entries to a VB6-format .rmg sidecar file.

Format: comma-delimited, 9+ columns per line:
  step_type,step_value,,,,,,,susceptibility
"""
from __future__ import annotations

from pathlib import Path

from rapid_main.data_model import MeasurementStep, RmgRecord


def _step_type_from_label(label: str) -> str:
    """Extract the 2-5 char step-type prefix from a demag label."""
    label_upper = label.upper()
    for prefix in ("AFMAX", "AFZ", "ARM", "IRM", "AF", "TT", "TH"):
        if label_upper.startswith(prefix):
            return prefix
    return label[:5]  # fallback: first 5 chars


def append_rmg_record(
    path: str | Path,
    step: MeasurementStep,
    susceptibility: float = 0.0,
) -> None:
    """
    Append one step + susceptibility record to a .rmg sidecar file.

    Parameters
    ----------
    path:
        Path to the .rmg file (created if it doesn't exist).
    step:
        The measurement step (provides step type and value).
    susceptibility:
        Susceptibility reading in emu/Oe at this step (0 if not measured).
    """
    path = Path(path)
    import re
    lbl = step.demag_label.upper()
    step_type = _step_type_from_label(lbl)
    # Extract numeric value from label (e.g. "AF20" → 20.0, "TT400" → 400.0)
    match = re.search(r"[\d.]+", lbl)
    step_value = float(match.group()) if match else 0.0

    # Build 9-column comma line (columns 2-7 are empty for future use)
    columns = [step_type, str(step_value), "", "", "", "", "", "", str(susceptibility)]
    line = ",".join(columns)

    with path.open("a", encoding="latin-1", newline="\r\n") as fh:
        fh.write(line + "\n")


def write_rmg_records(path: str | Path, records: list[RmgRecord]) -> None:
    """Write (overwrite) a .rmg file from a list of RmgRecord objects."""
    path = Path(path)
    with path.open("w", encoding="latin-1", newline="\r\n") as fh:
        for rec in records:
            if rec.raw_fields:
                fh.write(",".join(rec.raw_fields) + "\n")
            else:
                cols = [rec.step_type, str(rec.step_value), "", "", "", "", "", "", str(rec.susceptibility)]
                fh.write(",".join(cols) + "\n")
