"""rapid_main package."""
from __future__ import annotations

import os

__all__ = ["__version__", "software_version"]

__version__ = "4.0.0"


def software_version() -> str:
    """Version string stamped into audit records and evidence bundles.

    Set ``RAPID_BUILD_COMMIT`` at build/deploy time so hardware evidence can
    name the exact commit that produced it.
    """

    commit = os.environ.get("RAPID_BUILD_COMMIT", "").strip()
    return f"{__version__}+{commit}" if commit else __version__
