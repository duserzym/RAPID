"""
magic_writer.py — Write MagIC v3.0 measurements.txt files.

Produces tab-delimited MagIC format directly compatible with the PmagPy
data processing pipeline (https://earthref.org/PmagPy).

Reference: resources/MagIC Data Model v3.0 - February 26th, 2025.json
Guide:     resources/magic_data_entry_guide.md

measurements.txt layout:
  Row 1:  "tab delimited\\tmeasurements"
  Row 2:  Column header names (tab-delimited)
  Row 3+: Data rows (tab-delimited)

Unit conventions (MagIC v3.0):
  moment:        A·m²   (from emu: ×1e-3)
  AF field:      T      (from Oe: ×1e-4)
  DC bias field: T      (from Oe: ×1e-4)
  temperature:   K      (from °C: +273.15)
  declination/inclination: degrees (unchanged)
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.magic_method_codes import label_to_method_codes


# ---------------------------------------------------------------------------
# Column definitions (subset of MagIC measurements table used by RAPID)
# ---------------------------------------------------------------------------

MAGIC_COLUMNS = [
    "measurement",
    "specimen",
    "sample",
    "site",
    "location",
    "method_codes",
    "treat_ac_field",    # T
    "treat_dc_field",    # T
    "treat_temp",        # K
    "magn_moment",       # A·m²
    "magn_x",            # A·m²
    "magn_y",            # A·m²
    "magn_z",            # A·m²
    "dir_dec",           # °
    "dir_inc",           # °
    "meas_n_orient",
    "analysts",
    "timestamp",
]

_HEADER_LINE = "tab delimited\tmeasurements"
_EMPTY = ""


def _fmt(value: Optional[float]) -> str:
    """Format a float or return empty string for None."""
    if value is None:
        return _EMPTY
    return f"{value:.6g}"


def _step_to_row(
    step: MeasurementStep,
    meta: SpecimenMeta,
) -> list[str]:
    """Convert one MeasurementStep to a list of column values."""
    meas_name = f"{meta.name}-{step.demag_label}"
    method_code = label_to_method_codes(step.demag_label)
    magn_moment = step.magn_moment_Am2()
    # Cartesian in A·m² (emu → A·m² same factor ×1e-3)
    magn_x = step.sdx * 1e-3
    magn_y = step.sdy * 1e-3
    magn_z = step.sdz * 1e-3

    return [
        meas_name,                            # measurement
        meta.name,                            # specimen
        meta.sample,                          # sample
        meta.site,                            # site
        meta.location,                        # location
        method_code,                          # method_codes
        _fmt(step.treat_ac_field_T()),        # treat_ac_field (T)
        _fmt(step.treat_dc_field_T()),        # treat_dc_field (T)
        _fmt(step.treat_temp_K()),            # treat_temp (K)
        _fmt(magn_moment),                    # magn_moment (A·m²)
        _fmt(magn_x),                         # magn_x
        _fmt(magn_y),                         # magn_y
        _fmt(magn_z),                         # magn_z
        _fmt(step.gdec),                      # dir_dec (geographic)
        _fmt(step.ginc),                      # dir_inc (geographic)
        "4",                                  # meas_n_orient (4-position standard)
        step.operator,                        # analysts
        step.timestamp.strftime("%Y-%m-%dT%H:%M:%S"),  # timestamp
    ]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def write_measurements(
    path: str | Path,
    meta: SpecimenMeta,
    steps: list[MeasurementStep],
) -> None:
    """
    Write a complete MagIC measurements.txt for one specimen.
    Overwrites any existing file at *path*.

    Parameters
    ----------
    path:
        Output file path (typically ``<outdir>/measurements.txt``).
    meta:
        SpecimenMeta for this specimen (provides hierarchy labels).
    steps:
        Ordered list of MeasurementStep objects to write.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as fh:
        fh.write(_HEADER_LINE + "\n")
        fh.write("\t".join(MAGIC_COLUMNS) + "\n")
        for step in steps:
            row = _step_to_row(step, meta)
            fh.write("\t".join(row) + "\n")


def append_measurement(
    path: str | Path,
    meta: SpecimenMeta,
    step: MeasurementStep,
) -> None:
    """
    Append a single measurement step to an existing MagIC measurements.txt.
    Creates the file (with header) if it does not yet exist.

    Parameters
    ----------
    path:
        Output file path.
    meta:
        SpecimenMeta for this specimen.
    step:
        The MeasurementStep to append.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    write_header = not path.exists()
    with path.open("a", encoding="utf-8", newline="\n") as fh:
        if write_header:
            fh.write(_HEADER_LINE + "\n")
            fh.write("\t".join(MAGIC_COLUMNS) + "\n")
        row = _step_to_row(step, meta)
        fh.write("\t".join(row) + "\n")
