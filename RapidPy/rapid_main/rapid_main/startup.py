"""Read-only startup and packaging diagnostics for the main application."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from importlib import metadata
from importlib import resources
from importlib.util import find_spec
import json
import os
from pathlib import Path
import sys
from typing import Iterable

REQUIRED_MODULES = ("PySide6", "rapidpy_common")
OPTIONAL_MODULES = (
    "numpy",
    "pyqtgraph",
    "serial",
    "af_tuner",
    "data_viewer",
    "gaussmeter_control",
    "updown_control",
    "vrm_logger",
    "webcam_viewer",
)
ICON_CANDIDATES = (
    "rapid_main_icon.ico",
    "rapid_main_window_icon.ico",
    "rapid_main_window_icon.png",
    "rapid_main_icon.png",
    "rapid_icon.ico",
    "rapid_icon.png",
)


@dataclass(frozen=True, slots=True)
class StartupEnvironmentReport:
    python_executable: str
    python_version: str
    packaged: bool
    required_modules: dict[str, bool]
    optional_modules: dict[str, bool]
    assets_dir: str
    icon_path: str
    config_path: str
    config_exists: bool
    config_valid: bool
    blockers: tuple[str, ...]
    warnings: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.blockers

    def to_payload(self) -> dict[str, object]:
        payload = asdict(self)
        payload["schema"] = "rapidpy.startup_environment.v1"
        payload["ok"] = self.ok
        return payload

    def to_json(self) -> str:
        return json.dumps(self.to_payload(), indent=2, sort_keys=True)


def main_assets_dir() -> Path:
    try:
        packaged = Path(str(resources.files("rapid_main_assets")))
        if packaged.is_dir():
            return packaged
    except (ImportError, ModuleNotFoundError, TypeError):
        pass
    return (Path(__file__).resolve().parent.parent / "assets").resolve()


def select_main_icon(assets_dir: Path | None = None) -> tuple[str, Path]:
    directory = assets_dir or main_assets_dir()
    for name in ICON_CANDIDATES:
        if (directory / name).is_file():
            return name, directory / name
    return "", directory


def collect_startup_environment(
    *,
    required_modules: Iterable[str] = REQUIRED_MODULES,
    optional_modules: Iterable[str] = OPTIONAL_MODULES,
) -> StartupEnvironmentReport:
    required = {name: find_spec(name) is not None for name in required_modules}
    optional = {name: find_spec(name) is not None for name in optional_modules}
    assets = main_assets_dir()
    _icon_name, icon_path = select_main_icon(assets)
    config_path = Path(
        os.environ.get("RAPID_CONFIG", Path.home() / ".rapid" / "config.json")
    ).expanduser().resolve()
    config_exists = config_path.exists()
    config_valid = True
    if config_exists:
        try:
            config_payload = json.loads(config_path.read_text(encoding="utf-8"))
            config_valid = isinstance(config_payload, dict)
        except (OSError, ValueError):
            config_valid = False
    blockers = [f"Required Python module is unavailable: {name}" for name, ok in required.items() if not ok]
    if not icon_path.is_file():
        blockers.append(f"No RAPID main-window icon found in {assets}")
    if config_exists and not config_valid:
        blockers.append(f"Existing RAPID configuration is not a valid JSON object: {config_path}")
    warnings = [
        f"Optional component is unavailable: {name}"
        for name, ok in optional.items()
        if not ok
    ]
    if not config_exists:
        warnings.append(f"No configuration file exists yet; RapidPy will create {config_path}")
    try:
        metadata.version("berkeley-rapidpy")
        packaged = True
    except metadata.PackageNotFoundError:
        packaged = bool(getattr(sys, "frozen", False))
    return StartupEnvironmentReport(
        python_executable=sys.executable,
        python_version=sys.version.split()[0],
        packaged=packaged,
        required_modules=required,
        optional_modules=optional,
        assets_dir=str(assets),
        icon_path=str(icon_path) if icon_path.is_file() else "",
        config_path=str(config_path),
        config_exists=config_exists,
        config_valid=config_valid,
        blockers=tuple(blockers),
        warnings=tuple(warnings),
    )

