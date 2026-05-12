"""
data_model.py — Core dataclasses for RAPID v4 specimen data.

Mirrors the VB6 Sample.cls data structures with added MagIC-derived fields.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


# ---------------------------------------------------------------------------
# Specimen metadata (header lines of the specimen file)
# ---------------------------------------------------------------------------

@dataclass
class SpecimenMeta:
    """Orientation and identification info read from the specimen file header."""
    name: str                          # specimen name (== filename)
    comment: str = ""                  # line 1 of specimen file (free text)
    core_plate_strike: float = 0.0     # °
    core_plate_dip: float = 0.0        # °
    bedding_strike: float = 0.0        # °
    bedding_dip: float = 0.0           # °
    volume: float = 1.0                # cm³
    fold_axis: Optional[float] = None  # ° (optional, col 39-43)
    fold_plunge: Optional[float] = None  # ° (optional, col 45-49)
    # Hierarchical MagIC names (filled in by caller if known)
    sample: str = ""
    site: str = ""
    location: str = ""


# ---------------------------------------------------------------------------
# Per-step measurement record
# ---------------------------------------------------------------------------

@dataclass
class MeasurementStep:
    """One demagnetisation/measurement step, exactly as written by VB6 WriteData()."""
    demag_label: str          # e.g. "NRM", "AF20", "TT400", "IRM1000", "ARM"
    gdec: float               # geographic declination (°)
    ginc: float               # geographic inclination (°)
    sdec: float               # specimen declination (°)
    sinc: float               # specimen inclination (°)
    moment: float             # moment in emu
    error_angle: float        # CSD / error angle (°)
    crdec: float              # core declination (°)
    crinc: float              # core inclination (°)
    sdx: float                # specimen Cartesian X (emu)
    sdy: float                # specimen Cartesian Y (emu)
    sdz: float                # specimen Cartesian Z (emu)
    operator: str = ""        # operator name (≤ 8 chars)
    timestamp: datetime = field(default_factory=datetime.now)

    # -----------------------------------------------------------------------
    # Derived MagIC fields
    # -----------------------------------------------------------------------

    def magn_moment_Am2(self) -> float:
        """Moment in A·m²  (1 emu = 1e-3 A·m²)."""
        return self.moment * 1e-3

    def treat_ac_field_T(self) -> Optional[float]:
        """AF peak field in Tesla (1 Oe = 1e-4 T).  None if not an AF step."""
        lbl = self.demag_label.upper()
        match = re.match(r'^(?:AF|AFZ|AFMAX)(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) * 1e-4  # Oe → T
        return None

    def treat_dc_field_T(self) -> Optional[float]:
        """DC bias field in Tesla.  Encoded as 'ARM<field>_<bias>' convention."""
        lbl = self.demag_label.upper()
        # ARM steps may encode DC bias after underscore: e.g. "ARM100_1"
        match = re.match(r'^ARM\d+_(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) * 1e-4  # Oe → T
        return None

    def treat_temp_K(self) -> Optional[float]:
        """Treatment temperature in Kelvin.  None if not a thermal step."""
        lbl = self.demag_label.upper()
        match = re.match(r'^(?:TT|TH|TEMP)(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) + 273.15  # °C → K
        return None

    def treat_ac_field_mT(self) -> Optional[float]:
        """AF field in mT (for display).  1 Oe = 0.1 mT."""
        t = self.treat_ac_field_T()
        return t * 1e3 if t is not None else None

    def method_codes(self) -> str:
        """MagIC LP-* method code(s) for this step."""
        from rapid_main.io.magic_method_codes import label_to_method_codes  # lazy import
        return label_to_method_codes(self.demag_label)

    def measurement_name(self, specimen_name: str) -> str:
        """Unique MagIC measurement identifier."""
        return f"{specimen_name}-{self.demag_label}"


# ---------------------------------------------------------------------------
# RMG sidecar record (susceptibility vs. demagnetisation)
# ---------------------------------------------------------------------------

@dataclass
class RmgRecord:
    """One line from a `.rmg` sidecar file (comma-delimited, 9+ columns)."""
    step_type: str          # column 0: AF, TT, IRM, ARM, AFz, AFmax
    step_value: float       # column 1: field (Oe) or temperature (°C)
    susceptibility: float   # column 8: susceptibility (emu/Oe)
    raw_fields: list = field(default_factory=list)  # all raw columns for round-trip fidelity
