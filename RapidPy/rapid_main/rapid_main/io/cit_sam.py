from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class CitSamHeader:
    format_id: str = "CIT"
    comment: str = ""
    latitude_n: float = 0.0
    longitude_e: float = 0.0
    magnetic_declination_e: float = 0.0
    fold_axis_azimuth: float | None = None
    fold_axis_plunge: float | None = None
    bedding_strike: float | None = None
    bedding_dip: float | None = None


@dataclass(slots=True)
class CitSamEntry:
    specimen_name: str
    strat_level: float | None = None
    site_id: str = ""


def _fmt5(value: float | None) -> str:
    if value is None:
        return " " * 5
    return f"{value:5.1f}"[-5:]


def _slice_float(line: str, start: int, end: int) -> float | None:
    try:
        text = line[start:end].strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def read_cit_sam(path: str | Path) -> tuple[CitSamHeader, list[CitSamEntry]]:
    """Read a CIT-style .sam locality file.

    Supports legacy files with or without an explicit format-id first line.
    """
    path = Path(path)
    lines = path.read_text(encoding="latin-1", errors="replace").splitlines()
    if not lines:
        return CitSamHeader(), []

    idx = 0
    first = lines[0].strip().upper()
    format_id = "CIT"
    if first in {"CIT", "2G", "APP", "JRA"}:
        format_id = first
        idx = 1

    comment = lines[idx].rstrip() if idx < len(lines) else ""
    idx += 1

    locality = lines[idx] if idx < len(lines) else ""
    idx += 1

    header = CitSamHeader(
        format_id=format_id,
        comment=comment,
        latitude_n=_slice_float(locality, 0, 5) or 0.0,
        longitude_e=_slice_float(locality, 6, 11) or 0.0,
        magnetic_declination_e=_slice_float(locality, 12, 17) or 0.0,
        fold_axis_azimuth=_slice_float(locality, 18, 23),
        fold_axis_plunge=_slice_float(locality, 24, 29),
        bedding_strike=_slice_float(locality, 30, 35),
        bedding_dip=_slice_float(locality, 36, 41),
    )

    entries: list[CitSamEntry] = []
    for raw in lines[idx:]:
        line = raw.rstrip()
        if not line or line.lstrip().startswith("#!"):
            continue
        specimen = line[:22].strip()
        if not specimen:
            continue
        level = _slice_float(line, 22, 30)
        site = line[30:32].strip() if len(line) >= 32 else ""
        entries.append(CitSamEntry(specimen_name=specimen, strat_level=level, site_id=site))

    return header, entries


def write_cit_sam(
    path: str | Path,
    header: CitSamHeader,
    entries: list[CitSamEntry],
    *,
    include_format_line: bool = True,
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    locality_line = (
        f"{_fmt5(header.latitude_n)} {_fmt5(header.longitude_e)} {_fmt5(header.magnetic_declination_e)}"
        f" {_fmt5(header.fold_axis_azimuth)} {_fmt5(header.fold_axis_plunge)}"
        f" {_fmt5(header.bedding_strike)} {_fmt5(header.bedding_dip)}"
    ).rstrip()

    with path.open("w", encoding="latin-1", newline="\r\n") as fh:
        if include_format_line:
            fh.write((header.format_id or "CIT").upper() + "\n")
        fh.write((header.comment or "")[:255] + "\n")
        fh.write(locality_line + "\n")
        for item in entries:
            specimen = (item.specimen_name or "")[:22].ljust(22)
            level = " " * 8 if item.strat_level is None else f"{item.strat_level:8.2f}"[:8]
            site = (item.site_id or "")[:2].ljust(2)
            fh.write(f"{specimen}{level}{site}".rstrip() + "\n")
