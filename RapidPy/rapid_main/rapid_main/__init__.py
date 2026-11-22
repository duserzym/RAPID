"""rapid_main package."""
from __future__ import annotations

from functools import lru_cache
import json
import os
from pathlib import Path
import re
import sys

__all__ = ["__version__", "software_version"]

__version__ = "4.0.0"


@lru_cache(maxsize=1)
def _packaged_commit(executable: str) -> str:
    try:
        manifest = json.loads(Path(executable).with_name("build_manifest.json").read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return ""
    if not isinstance(manifest, dict) or manifest.get("schema") != "rapidpy.portable_build.v2":
        return ""
    commit = manifest.get("source_revision")
    if isinstance(commit, str) and re.fullmatch(r"(?:[0-9a-f]{40}|[0-9a-f]{64})", commit):
        return commit
    return ""


def software_version() -> str:
    """Version string stamped into audit records and evidence bundles.

    Portable releases use the adjacent build manifest. Source runs can set
    ``RAPID_BUILD_COMMIT`` to identify the checked-out commit.
    """

    commit = _packaged_commit(sys.executable) if getattr(sys, "frozen", False) else ""
    commit = commit or os.environ.get("RAPID_BUILD_COMMIT", "").strip()
    return f"{__version__}+{commit}" if commit else __version__
