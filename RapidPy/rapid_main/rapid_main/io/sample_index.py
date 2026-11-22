from __future__ import annotations
"""Utilities for VB6-style sample index files (.sam) and queue registry rows."""
from pathlib import Path
import csv
import math
from typing import Iterable

from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations
from rapid_main.io.sam_reader import locality_header_index


def _build_rows(raw_lines: Iterable[str]) -> list[str]:
    """Drop comments/blank lines and normalize whitespace."""
    rows: list[str] = []
    for raw in raw_lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        rows.append(raw.rstrip('\r\n'))
    return rows


def _site_declination(line: str) -> float | None:
    # CIT has a six-character stride. Preserve an empty third field even
    # when optional fold and bedding fields follow it.
    fixed = (len(line) >= 17 and line[5].isspace() and line[11].isspace()
             and (len(line) == 17 or line[17].isspace()))
    if fixed:
        try:
            float(line[:5])
            float(line[6:11])
        except ValueError:
            fixed = False
    fields = line.split()
    text = line[12:17].strip() if fixed else (fields[2] if len(fields) >= 3 else '')
    if not text:
        return None
    declination = float(text)
    if not math.isfinite(declination):
        raise ValueError('Sample index magnetic declination must be finite.')
    return declination


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
    if sam_path.suffix.lower() == '.csv':
        with sam_path.open('r', encoding='utf-8-sig', newline='') as handle:
            rows = list(csv.reader(handle))
        if rows and rows[0] and rows[0][0].strip().lower() in {'sample name', 'name', 'specimen'}:
            rows = rows[1:]
        entries = []
        for row in rows:
            if not row or not row[0].strip():
                continue
            values = [row[index].strip() if index < len(row) else '' for index in range(4)]
            entries.append(SampleIndexRegistration(specimen_name=values[0], sample_set=sam_path.stem,
                depth_cm=values[1], formation=values[2], location=values[3], order=len(entries) + 1,
                source_file=str(sam_path.resolve())))
        return SampleIndexRegistrations(entries)
    with sam_path.open("r", encoding="latin-1", errors="replace") as fh:
        rows = _build_rows(fh)

    if not rows:
        return SampleIndexRegistrations([])

    data_start = 0
    sample_set = sam_path.stem
    formation = ""
    location = str(sam_path.parent)

    magnetic_declination = None
    header = locality_header_index(rows)
    if header is not None:
        sample_set = rows[header - 1].strip()
        location = rows[header].strip()
        data_start = header + 1
        magnetic_declination = _site_declination(rows[header])
    
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
                source_file=str(sam_path.resolve()),
                magnetic_declination_e=magnetic_declination,
            )
        )

    return SampleIndexRegistrations(registrations)


def registrations_to_samples(registrations: SampleIndexRegistrations) -> list[str]:
    """Extract specimen names from a registry collection."""
    return registrations.names
