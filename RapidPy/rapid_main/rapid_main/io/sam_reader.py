"""
sam_reader.py — Parse VB6-format .sam sample index files.

A .sam file lives at:  <filedir>/<SampleCode>/<SampleCode>.sam
Each non-blank, non-comment line is a specimen name that belongs to that
sample set.  The matching specimen file lives at:
  <filedir>/<SampleCode>/<SpecimenName>   (no extension)
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional


def read_sam(sam_path: str | Path) -> list[str]:
    """
    Parse a .sam index file and return the ordered list of specimen names.

    Lines starting with '#' and blank lines are skipped (treated as comments).
    """
    sam_path = Path(sam_path)
    specimens: list[str] = []
    with sam_path.open("r", encoding="latin-1", errors="replace") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            specimens.append(line)
    return specimens


def specimen_path(sam_path: str | Path, specimen_name: str) -> Path:
    """
    Resolve the specimen data file path from a .sam file location.

    Convention (VB6): specimen file = sam_path.parent / specimen_name
    """
    return Path(sam_path).parent / specimen_name


def find_sam_files(search_dir: str | Path) -> list[Path]:
    """Recursively find all .sam files under *search_dir*."""
    return sorted(Path(search_dir).rglob("*.sam"))
