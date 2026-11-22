"""
sam_reader.py — Parse VB6-format .sam sample index files.

Some CIT `.sam` files are just specimen-name lists. Others start with a
two-line header containing the set name and site coordinates, followed by the
specimen names. The reader normalises both forms to the ordered specimen list.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional


def _looks_like_coordinate_header(line: str) -> bool:
    """Return True for the locality coordinates and optional orientation fields."""
    parts = line.split()
    if not 2 <= len(parts) <= 7:
        return False
    try:
        [float(part) for part in parts]
    except ValueError:
        return False
    return True


def locality_header_index(rows: list[str]) -> Optional[int]:
    index = 2 if rows and rows[0].strip().upper() in {"CIT", "2G", "APP", "JRA"} else 1
    return index if len(rows) > index and _looks_like_coordinate_header(rows[index]) else None


def read_sam(sam_path: str | Path) -> list[str]:
    """
    Parse a .sam index file and return the ordered list of specimen names.

    Lines starting with '#' and blank lines are skipped (treated as comments).
    If the first non-comment lines are a set name followed by a coordinate
    header, both header lines are ignored.
    """
    sam_path = Path(sam_path)
    entries: list[str] = []
    with sam_path.open("r", encoding="latin-1", errors="replace") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            entries.append(line)

    header = locality_header_index(entries)
    if header is not None:
        return entries[header + 1:]
    return entries


def specimen_path(sam_path: str | Path, specimen_name: str) -> Path:
    """
    Resolve the specimen data file path from a .sam file location.

    Convention (VB6): specimen file = sam_path.parent / specimen_name
    """
    return Path(sam_path).parent / specimen_name


def find_sam_files(search_dir: str | Path) -> list[Path]:
    """Recursively find all .sam files under *search_dir*."""
    return sorted(Path(search_dir).rglob("*.sam"))
