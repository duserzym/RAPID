from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from .models import CalibrationFactors


CONFIG_FILE = Path.home() / ".rapidpy_vrm_config.json"


@dataclass(slots=True)
class AppConfig:
    port: str = ""
    interval_s: float = 1.0
    spacing_mode: str = "Linear"
    output_file: str = "vrm_log.csv"
    display_unit: str = "Moment"
    calibration_x: float = -3.410
    calibration_y: float = -3.470
    calibration_z: float = -2.516
    range_fact: float = 1e-5
    ini_path: str = ""
    window_geometry: str = ""

    @property
    def factors(self) -> CalibrationFactors:
        return CalibrationFactors(
            x=self.calibration_x,
            y=self.calibration_y,
            z=self.calibration_z,
        )


def load_config() -> AppConfig:
    if not CONFIG_FILE.exists():
        return AppConfig()

    try:
        payload = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return AppConfig()

    defaults = AppConfig()
    merged = asdict(defaults)
    for key in merged:
        if key in payload:
            merged[key] = payload[key]

    return AppConfig(**merged)


def save_config(config: AppConfig) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(asdict(config), indent=2, sort_keys=True),
            encoding="utf-8",
        )
    except OSError:
        # Config persistence failure should not stop the run.
        return


# ---------------------------------------------------------------------------
# INI calibration helpers
# ---------------------------------------------------------------------------

def _auto_find_ini() -> Path | None:
    """Search for Paleomag_v3.INI relative to this package."""
    # Look in VB6/ sibling of the repo root (4 levels up from this file:
    # vrm_logger/vrm_logger/config.py → vrm_logger → RapidPy → repo root)
    here = Path(__file__).resolve()
    for depth in range(2, 6):
        candidate = here.parents[depth] / "VB6" / "Paleomag_v3.INI"
        if candidate.exists():
            return candidate
    return None


def read_calibration_from_ini(
    ini_path: Path,
) -> tuple[float, float, float, float] | None:
    """Return ``(xcal, ycal, zcal, range_fact)`` from *ini_path*, or ``None``."""
    try:
        text = ini_path.read_text(encoding="latin-1")
    except OSError:
        return None

    in_section = False
    vals: dict[str, str] = {}
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.lower() == "[magnetometercalibration]":
            in_section = True
            continue
        if in_section:
            if stripped.startswith("["):
                break
            if "=" in stripped:
                k, _, v = stripped.partition("=")
                vals[k.strip().lower()] = v.strip()

    try:
        xcal = float(vals.get("xcal", "-3.410"))
        ycal = float(vals.get("ycal", "-3.470"))
        zcal = float(vals.get("zcal", "-2.516"))
        rfact = float(vals.get("rangefact", "0.00001"))
        return xcal, ycal, zcal, rfact
    except ValueError:
        return None
