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
    display_unit: str = "Volts"
    calibration_x: float = -3.410
    calibration_y: float = -3.470
    calibration_z: float = -2.516
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
