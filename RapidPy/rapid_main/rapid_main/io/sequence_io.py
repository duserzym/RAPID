"""
sequence_io.py — Save and load RAPID sequence step lists.

Sequences are stored as JSON (list of step-label strings).
A plain-text one-per-line format is also supported for compatibility.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
import tempfile


class SequenceFormatError(ValueError):
    """Raised when a sequence file is readable but does not contain valid steps."""


def _validated_labels(labels: object) -> list[str]:
    if not isinstance(labels, list):
        raise SequenceFormatError("Sequence steps must be a JSON list.")
    normalized: list[str] = []
    for index, value in enumerate(labels, start=1):
        if not isinstance(value, str):
            raise SequenceFormatError(f"Sequence step {index} must be text.")
        label = value.strip()
        if not label:
            raise SequenceFormatError(f"Sequence step {index} is blank.")
        if any(char in label for char in "\r\n\0"):
            raise SequenceFormatError(f"Sequence step {index} contains invalid control characters.")
        normalized.append(label)
    if not normalized:
        raise SequenceFormatError("Sequence contains no steps.")
    return normalized


def _atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        newline="\n",
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
        delete=False,
    )
    temp_path = Path(handle.name)
    try:
        with handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def save_sequence_json(path: Path | str, labels: list[str]) -> None:
    """Write a sequence step list to a JSON file."""
    p = Path(path)
    clean = _validated_labels(labels)
    _atomic_write_text(p, json.dumps({"schema": "rapidpy.sequence.v1", "steps": clean}, indent=2) + "\n")


def load_sequence_json(path: Path | str) -> list[str]:
    """Read a sequence from a JSON file.  Returns empty list on error."""
    p = Path(path)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return []
        return _validated_labels(data.get("steps", []))
    except Exception:
        return []


def save_sequence_txt(path: Path | str, labels: list[str]) -> None:
    """Write a sequence as plain text (one label per line)."""
    p = Path(path)
    clean = _validated_labels(labels)
    _atomic_write_text(p, "\n".join(clean) + "\n")


def load_sequence_txt(path: Path | str) -> list[str]:
    """Read a sequence from a plain-text file.  Returns empty list on error."""
    p = Path(path)
    try:
        lines = p.read_text(encoding="utf-8").splitlines()
        return _validated_labels(
            [ln.strip() for ln in lines if ln.strip() and not ln.lstrip().startswith("#")]
        )
    except Exception:
        return []


def load_sequence(path: Path | str) -> list[str]:
    """Auto-detect JSON vs. plain text based on file extension."""
    p = Path(path)
    if p.suffix.lower() == ".json":
        return load_sequence_json(p)
    return load_sequence_txt(p)


def load_sequence_strict(path: Path | str) -> list[str]:
    """Load a sequence and preserve a useful error for the operator UI."""
    p = Path(path)
    try:
        if p.suffix.lower() == ".json":
            payload = json.loads(p.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                raise SequenceFormatError("JSON sequence must be an object with a 'steps' list.")
            return _validated_labels(payload.get("steps"))
        lines = p.read_text(encoding="utf-8").splitlines()
        return _validated_labels(
            [line.strip() for line in lines if line.strip() and not line.lstrip().startswith("#")]
        )
    except SequenceFormatError:
        raise
    except json.JSONDecodeError as exc:
        raise SequenceFormatError(f"Invalid JSON at line {exc.lineno}, column {exc.colno}.") from exc
    except OSError as exc:
        raise SequenceFormatError(f"Unable to read sequence file: {exc}") from exc
