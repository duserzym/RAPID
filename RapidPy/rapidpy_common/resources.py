"""Package-aware resource lookup with a source-checkout fallback."""
from __future__ import annotations

from importlib import resources
from pathlib import Path


def asset_directory(package: str, source_fallback: str | Path) -> Path:
    """Return an installed asset package directory or a development fallback."""

    try:
        candidate = resources.files(package)
        path = Path(str(candidate))
        if path.is_dir():
            return path
    except (ImportError, ModuleNotFoundError, TypeError):
        pass
    return Path(source_fallback).resolve()

