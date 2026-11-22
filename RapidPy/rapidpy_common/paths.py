"""Filesystem paths for data and evidence written by RapidPy tools."""
import os
from pathlib import Path


def absolute_io_path(path: str | Path) -> Path:
    """Return an absolute path with Windows extended-length support."""
    absolute = os.path.abspath(path)
    if os.name != "nt" or absolute.startswith(("\\\\?\\", "\\\\.\\")):
        return Path(absolute)
    if absolute.startswith("\\\\"):
        return Path("\\\\?\\UNC\\" + absolute[2:])
    return Path("\\\\?\\" + absolute)
