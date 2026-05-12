"""
rmg_reader.py — Parse VB6-format .rmg sidecar files.

.rmg files store comma-delimited susceptibility + step data alongside
each specimen file.  Format discovered from frmMeasure.ImportZijRoutine:

  Column 0:  step-type prefix  (AF, TT, IRM, ARM, AFz, AFmax)
  Column 1:  step value        (field in Oe, temperature in °C)
  Columns 2-7: (reserved / unused in known files)
  Column 8:  susceptibility reading (emu/Oe)

One line per step, ordered oldest-first (same order as specimen file).
"""
from __future__ import annotations

from pathlib import Path

from rapid_main.data_model import RmgRecord


def read_rmg(path: str | Path) -> list[RmgRecord]:
    """
    Parse a .rmg file and return a list of RmgRecord objects.
    Lines that cannot be parsed are silently skipped.
    """
    path = Path(path)
    records: list[RmgRecord] = []
    for raw in path.read_text(encoding="latin-1", errors="replace").splitlines():
        line = raw.strip()
        if not line:
            continue
        parts = line.split(",")
        try:
            step_type = parts[0].strip()
            step_value = float(parts[1]) if len(parts) > 1 else 0.0
            susceptibility = float(parts[8]) if len(parts) > 8 else 0.0
            records.append(RmgRecord(
                step_type=step_type,
                step_value=step_value,
                susceptibility=susceptibility,
                raw_fields=parts,
            ))
        except (ValueError, IndexError):
            continue
    return records
