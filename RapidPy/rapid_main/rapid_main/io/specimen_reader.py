"""
specimen_reader.py — Parse VB6-format specimen data files.

Specimen file layout (from VB6 Sample.cls ReadSpec / ReadUpMeasurements):

  Line 1:  Free-text comment
  Line 2:  Fixed-width orientation params (space-padded, 1-based columns):
             cols  9-13  CorePlateStrike
             cols 15-19  CorePlateDip
             cols 21-25  BeddingStrike
             cols 27-31  BeddingDip
             cols 33-37  Volume (cm³)
             cols 39-43  FoldAxis    (optional)
             cols 45-49  FoldPlunge  (optional)
  Lines 3+: Space-delimited per-step records written by WriteData():
             demag gdec ginc sdec sinc moment errangle crdec crinc sdx sdy sdz operator datetime
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional

from rapid_main.data_model import MeasurementStep, SpecimenMeta


# ---------------------------------------------------------------------------
# Header parser (line 2 of specimen file)
# ---------------------------------------------------------------------------

def _parse_header_line(line: str) -> dict:
    """Extract fixed-width orientation fields from header line (line 2)."""
    def _col(s: str, start: int, end: int) -> Optional[float]:
        """1-based, inclusive column slice → float or None."""
        try:
            return float(s[start - 1:end].strip())
        except (ValueError, IndexError):
            return None

    return {
        "core_plate_strike": _col(line, 9, 13) or 0.0,
        "core_plate_dip":    _col(line, 15, 19) or 0.0,
        "bedding_strike":    _col(line, 21, 25) or 0.0,
        "bedding_dip":       _col(line, 27, 31) or 0.0,
        "volume":            _col(line, 33, 37) or 1.0,
        "fold_axis":         _col(line, 39, 43),
        "fold_plunge":       _col(line, 45, 49),
    }


# ---------------------------------------------------------------------------
# Step-line parser (lines 3+ of specimen file)
# ---------------------------------------------------------------------------

def _parse_step_line(line: str) -> Optional[MeasurementStep]:
    """Parse one space-delimited step record.  Returns None on failure."""
    parts = line.split()
    if len(parts) < 12:
        return None
    try:
        demag    = parts[0]
        gdec     = float(parts[1])
        ginc     = float(parts[2])
        sdec     = float(parts[3])
        sinc     = float(parts[4])
        moment   = float(parts[5])
        errangle = float(parts[6])
        crdec    = float(parts[7])
        crinc    = float(parts[8])
        sdx      = float(parts[9])
        sdy      = float(parts[10])
        sdz      = float(parts[11])
        operator = parts[12] if len(parts) > 12 else ""
        ts_str   = " ".join(parts[13:15]) if len(parts) > 14 else ""
        try:
            ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            ts = datetime.now()
        return MeasurementStep(
            demag_label=demag, gdec=gdec, ginc=ginc,
            sdec=sdec, sinc=sinc, moment=moment,
            error_angle=errangle, crdec=crdec, crinc=crinc,
            sdx=sdx, sdy=sdy, sdz=sdz,
            operator=operator, timestamp=ts,
        )
    except (ValueError, IndexError):
        return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_specimen(
    path: str | Path,
    specimen_name: Optional[str] = None,
) -> tuple[SpecimenMeta, list[MeasurementStep]]:
    """
    Read a VB6 specimen file and return (SpecimenMeta, [MeasurementStep, ...]).

    Parameters
    ----------
    path:
        Full path to the specimen file (no extension).
    specimen_name:
        Override for the specimen name; defaults to ``path.name``.
    """
    path = Path(path)
    name = specimen_name or path.name
    lines = path.read_text(encoding="latin-1", errors="replace").splitlines()

    comment = lines[0].strip() if len(lines) > 0 else ""

    meta_fields: dict = {}
    if len(lines) > 1:
        meta_fields = _parse_header_line(lines[1])

    meta = SpecimenMeta(
        name=name,
        comment=comment,
        **meta_fields,
    )

    steps: list[MeasurementStep] = []
    for line in lines[2:]:
        stripped = line.strip()
        if not stripped:
            continue
        step = _parse_step_line(stripped)
        if step is not None:
            steps.append(step)

    return meta, steps
