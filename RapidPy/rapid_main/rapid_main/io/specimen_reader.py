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

import math
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.geometry import Angular3D, angular3d_to_viewer_cartesian3d


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

    strike = _col(line, 9, 13)
    dip = _col(line, 15, 19)
    return {
        "core_plate_strike": strike or 0.0,
        "core_plate_dip":    dip or 0.0,
        "core_orientation_available": strike is not None and dip is not None,
        "bedding_strike":    _col(line, 21, 25) or 0.0,
        "bedding_dip":       _col(line, 27, 31) or 0.0,
        "volume":            _col(line, 33, 37) or 1.0,
        "fold_axis":         _col(line, 39, 43),
        "fold_plunge":       _col(line, 45, 49),
    }


# ---------------------------------------------------------------------------
# Step-line parser (lines 3+ of specimen file)
# ---------------------------------------------------------------------------

_SPLIT_LABEL_PREFIXES = ("TT", "AF", "TH", "TEMP", "IRM", "ARM", "ZF", "IF", "PTRM", "IZ", "ZI")
_STEP_VALUE_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?(?:_[+-]?\d+(?:\.\d+)?)?$")

def _normalize_demag_tokens(parts: list[str]) -> tuple[str, int]:
    """Normalize split CIT labels such as `TT 100` to `TT100`."""
    if len(parts) >= 2 and parts[0].upper() in _SPLIT_LABEL_PREFIXES and _STEP_VALUE_PATTERN.fullmatch(parts[1]):
        return f"{parts[0]}{parts[1]}", 1
    return parts[0], 0

def _parse_step_line(line: str) -> Optional[MeasurementStep]:
    """Parse one space-delimited step record.  Returns None on failure."""
    parts = line.split()
    if len(parts) < 12:
        return None
    try:
        demag, offset = _normalize_demag_tokens(parts)
        base = 1 + offset
        gdec     = float(parts[base + 0])
        ginc     = float(parts[base + 1])
        sdec     = float(parts[base + 2])
        sinc     = float(parts[base + 3])
        moment   = float(parts[base + 4])
        errangle = float(parts[base + 5])
        crdec    = float(parts[base + 6])
        crinc    = float(parts[base + 7])
        cart = angular3d_to_viewer_cartesian3d(Angular3D(dec=sdec, inc=sinc, mag=moment))
        sdx, sdy, sdz = cart.x, cart.y, cart.z
        trailing = parts[base + 11:]
        operator = trailing[0] if trailing else ""
        try:
            ts = datetime.strptime(" ".join(trailing[-2:]), "%Y-%m-%d %H:%M:%S")
            operator = " ".join(trailing[:-2])
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
    *,
    validate_format: bool = False,
) -> tuple[SpecimenMeta, list[MeasurementStep]]:
    """
    Read a VB6 specimen file and return (SpecimenMeta, [MeasurementStep, ...]).

    Parameters
    ----------
    path:
        Full path to the specimen file (no extension).
    specimen_name:
        Override for the specimen name; defaults to ``path.name``.
    validate_format:
        Require identifiable CIT content when opening an arbitrary file.
    """
    path = Path(path)
    name = specimen_name or path.name
    return parse_specimen(path.read_bytes(), name, validate_format=validate_format)


def parse_specimen(data: bytes, name: str, *, validate_format: bool = False) -> tuple[SpecimenMeta, list[MeasurementStep]]:
    """Parse the same CIT bytes from an open file or a verified archive member."""
    lines = data.decode("latin-1").splitlines()

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
            if validate_format and not all(math.isfinite(value) for value in (
                step.gdec, step.ginc, step.sdec, step.sinc, step.crdec,
                step.crinc, step.moment, step.error_angle,
            )):
                raise ValueError(f"{name} contains a nonfinite CIT measurement.")
            steps.append(step)

    if validate_format and not steps:
        # A new specimen can have a valid header before the first reading.
        try:
            volume = float(lines[1][32:37].strip())
        except (IndexError, ValueError):
            volume = 0.0
        if not math.isfinite(volume) or volume <= 0 or any(line.strip() for line in lines[2:]):
            raise ValueError(f"{name} has no readable CIT specimen header or measurements.")
    return meta, steps
