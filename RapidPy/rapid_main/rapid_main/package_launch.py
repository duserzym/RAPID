"""Resolve helper applications in source checkouts and installed distributions."""
from __future__ import annotations

from dataclasses import dataclass
from importlib.util import find_spec
from pathlib import Path
import sys

HELPER_MODULES = (
    "af_tuner", "data_viewer", "gaussmeter_control", "updown_control",
    "vrm_logger", "webcam_viewer",
)


class ToolUnavailableError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class ToolLaunch:
    command: tuple[str, ...]
    cwd: Path | None
    source: str


def resolve_tool_launch(
    *,
    module: str,
    source_root: str | Path,
    source_relative: str | Path,
    python_executable: str | None = None,
) -> ToolLaunch:
    """Prefer a checkout script, then fall back to an installed ``-m`` entry."""

    executable = python_executable or sys.executable
    if getattr(sys, "frozen", False):
        if module not in HELPER_MODULES:
            raise ToolUnavailableError(f"Tool {module!r} is not bundled with RAPID.")
        return ToolLaunch((executable, "--tool", module), None, "bundled-tool")
    root = Path(source_root).resolve()
    script = root / source_relative
    if script.is_file():
        return ToolLaunch((executable, str(script)), root, "source-checkout")
    try:
        available = find_spec(module) is not None
    except (ImportError, ModuleNotFoundError, ValueError):
        available = False
    if available:
        return ToolLaunch((executable, "-m", module), None, "installed-package")
    raise ToolUnavailableError(
        f"Neither source entry point {script} nor installed module {module!r} is available."
    )

