"""
sequence_io.py — Save and load RAPID sequence step lists.

Sequences are stored as JSON (list of step-label strings).
A plain-text one-per-line format is also supported for compatibility.
"""
from __future__ import annotations

import json
from pathlib import Path


def save_sequence_json(path: Path | str, labels: list[str]) -> None:
    """Write a sequence step list to a JSON file."""
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"steps": labels}, indent=2), encoding="utf-8")


def load_sequence_json(path: Path | str) -> list[str]:
    """Read a sequence from a JSON file.  Returns empty list on error."""
    p = Path(path)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        steps = data.get("steps", [])
        return [str(s) for s in steps]
    except Exception:
        return []


def save_sequence_txt(path: Path | str, labels: list[str]) -> None:
    """Write a sequence as plain text (one label per line)."""
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("\n".join(labels), encoding="utf-8")


def load_sequence_txt(path: Path | str) -> list[str]:
    """Read a sequence from a plain-text file.  Returns empty list on error."""
    p = Path(path)
    try:
        lines = p.read_text(encoding="utf-8").splitlines()
        return [ln.strip() for ln in lines if ln.strip() and not ln.startswith("#")]
    except Exception:
        return []


def load_sequence(path: Path | str) -> list[str]:
    """Auto-detect JSON vs. plain text based on file extension."""
    p = Path(path)
    if p.suffix.lower() == ".json":
        return load_sequence_json(p)
    return load_sequence_txt(p)
