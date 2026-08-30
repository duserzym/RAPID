from __future__ import annotations
"""Utilities for VB6-style sample index files (.sam) and queue registry rows."""
from pathlib import Path
from typing import Iterable

from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations
from rapid_main.io.sam_reader import _looks_like_coordinate_header


def _build_rows(raw_lines: Iterable[str]) -> list[str]:
    """Drop comments/blank lines and normalize whitespace."""
    rows: list[str] = []
    for raw in raw_lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        rows.append(line)
    return rows


def read_sample_index_registrations(
    sam_path: str | Path,
) -> SampleIndexRegistrations:
    """Read a ``.sam`` registry file into typed registration objects.

    VB6 registry files may start with a two-line header
    (set name + coordinate pair). That header is preserved as:

    - ``sample_set``: first line
    - ``location``/``formation``: second line when present

    The remaining lines are treated as specimen names and are stored in order.
    """
    sam_path = Path(sam_path)
    with sam_path.open("r", encoding="latin-1", errors="replace") as fh:
        rows = _build_rows(fh)

    if not rows:
        return SampleIndexRegistrations([])

    data_start = 0
    sample_set = sam_path.stem
    formation = ""
    location = str(sam_path.parent)

    if len(rows) >= 2 and _looks_like_coordinate_header(rows[1]):
        sample_set = rows[0]
        location = rows[1]
        data_start = 2
    
    registrations: list[SampleIndexRegistration] = []
    for idx, line in enumerate(rows[data_start:], start=1):
        columns = [part.strip() for part in line.split()]
        specimen = columns[0] if columns else ""
        if not specimen:
            continue

        depth_cm = columns[1] if len(columns) >= 2 else ""
        formation = columns[2] if len(columns) >= 3 else ""
        location_field = columns[3] if len(columns) >= 4 else location

        registrations.append(
            SampleIndexRegistration(
                specimen_name=specimen,
                sample_set=sample_set,
                location=location_field,
                formation=formation,
                depth_cm=depth_cm,
                order=idx,
            )
        )

    return SampleIndexRegistrations(registrations)


def registrations_to_samples(registrations: SampleIndexRegistrations) -> list[str]:
    """Extract specimen names from a registry collection."""
    return registrations.names
