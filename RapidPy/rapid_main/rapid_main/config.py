"""
config.py — Application configuration for RAPID v4.

``AppConfig`` is a dataclass tree that mirrors all settings in the
Settings panel.  It serialises to / deserialises from a single JSON
file so every user preference survives across restarts.

Default path: ``~/.rapid/config.json``  (overridable via env var
``RAPID_CONFIG``).

Usage::

    from rapid_main.config import AppConfig

    cfg = AppConfig.load()        # load or create defaults
    cfg.general.operator = "JSmith"
    cfg.save()                    # write back to JSON
"""
from __future__ import annotations

import dataclasses
import json
import os
from dataclasses import dataclass, field, asdict
from datetime import date
from pathlib import Path
from typing import Any


# ─────────────────────────────────────────────────────────────────────────────
# Sub-configs (one per settings tab)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class GeneralConfig:
    data_dir:     str  = ""
    sample_dir:   str  = ""
    backup_dir:   str  = ""
    operator:     str  = ""
    lab_name:     str  = "Paleomagnetism Laboratory"
    nocomm:       bool = False
    auto_save:    bool = True


@dataclass
class SquidConfig:
    port:             str   = "COM1"
    baud:             int   = 9600
    range_label:      str   = "1×"
    samples_per_pos:  int   = 4
    settle_time:      float = 1.5


@dataclass
class IrmArmConfig:
    irm_max_field:  float = 1000.0   # mT
    irm_axis:       str   = "Z (up-axis)"
    irm_ramp:       str   = "Slow (60 s)"
    irm_steps:      int   = 10
    arm_peak_af:    float = 100.0    # mT
    arm_bias:       float = 0.05     # mT


@dataclass
class AfDemagConfig:
    board:        int   = 0
    peak:         float = 180.0   # mT
    ramp_speed:   str   = "Medium (8 Hz)"
    settle:       float = 1.5     # s
    tumble:       bool  = False
    tumble_pause: float = 0.5     # s


@dataclass
class VacuumConfig:
    port:            str   = "COM1"
    baud:            int   = 9600
    target_pressure: float = 5.0    # mTorr
    warn_threshold:  float = 20.0   # mTorr
    auto_pump:       bool  = False
    poll_interval:   float = 2.0    # s


@dataclass
class DataFilesConfig:
    format:       str  = "CSV (comma-separated)"
    naming:       str  = "SampleName_Date"
    decimals:     int  = 6
    auto_save:    bool = True
    auto_append:  bool = False
    backup:       bool = True
    write_header: bool = True
    write_meta:   bool = True


@dataclass
class ChangerConfig:
    port:         str   = "COM3"
    baud:         int   = 9600
    speed_xy:     float = 20.0   # %
    speed_z:      float = 15.0   # %
    home_x:       float = 0.0    # mm
    home_y:       float = 0.0    # mm
    home_z:       float = 0.0    # mm
    soft_limits:  bool  = True
    limit_stop:   bool  = True


@dataclass
class CalibrationConfig:
    cal_rod_moment: float = 1.234e-5   # A·m²
    cal_date_iso:   str   = ""         # ISO date string, e.g. "2026-05-12"
    cal_x:          float = 1.0
    cal_y:          float = 1.0
    cal_z:          float = 1.0
    range_factor:   float = 1.0e-5
    bg_x:           float = 0.0
    bg_y:           float = 0.0
    bg_z:           float = 0.0
    bg_subtract:    bool  = False

    def cal_date(self) -> date | None:
        try:
            return date.fromisoformat(self.cal_date_iso) if self.cal_date_iso else None
        except ValueError:
            return None


@dataclass
class SequenceTimesConfig:
    """Per-step-type time estimates in seconds (used by RuntimeEstimator)."""
    NRM:     int = 120
    AF:      int = 180
    TT:      int = 600
    TH:      int = 600
    IRM:     int = 300
    ARM:     int = 240
    RRM:     int = 200
    PTRM:    int = 600
    ZF:      int = 120
    IF:      int = 120
    default: int = 120   # fallback for unknown types

    def as_estimator_dict(self) -> dict[str, int]:
        """Convert to the dict format expected by RuntimeEstimator."""
        return {
            "NRM":     self.NRM,
            "AF":      self.AF,
            "AFMAX":   self.AF,
            "AFZ":     self.AF,
            "TT":      self.TT,
            "TH":      self.TH,
            "TEMP":    self.TT,
            "IRM":     self.IRM,
            "ARM":     self.ARM,
            "RRM":     self.RRM,
            "PTRM":    self.PTRM,
            "ZF":      self.ZF,
            "IF":      self.IF,
            "_default": self.default,
        }


# ─────────────────────────────────────────────────────────────────────────────
# Root config
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class AppConfig:
    general:    GeneralConfig      = field(default_factory=GeneralConfig)
    squid:      SquidConfig        = field(default_factory=SquidConfig)
    irm_arm:    IrmArmConfig       = field(default_factory=IrmArmConfig)
    af_demag:   AfDemagConfig      = field(default_factory=AfDemagConfig)
    vacuum:     VacuumConfig       = field(default_factory=VacuumConfig)
    data_files: DataFilesConfig    = field(default_factory=DataFilesConfig)
    changer:    ChangerConfig      = field(default_factory=ChangerConfig)
    calibration: CalibrationConfig = field(default_factory=CalibrationConfig)
    sequence:   SequenceTimesConfig = field(default_factory=SequenceTimesConfig)

    # ── Persistence ───────────────────────────────────────────────────────────

    @staticmethod
    def default_path() -> Path:
        env = os.environ.get("RAPID_CONFIG")
        if env:
            return Path(env)
        return Path.home() / ".rapid" / "config.json"

    @classmethod
    def load(cls, path: Path | str | None = None) -> "AppConfig":
        """Load config from *path* (or ``default_path()``).

        Returns a fully-defaulted ``AppConfig`` if the file does not exist or
        is corrupt.
        """
        p = Path(path) if path else cls.default_path()
        if not p.exists():
            return cls()
        try:
            raw: dict[str, Any] = json.loads(p.read_text(encoding="utf-8"))
            return cls._from_dict(raw)
        except Exception:
            return cls()

    def save(self, path: Path | str | None = None) -> None:
        """Save config to *path* (or ``default_path()``)."""
        p = Path(path) if path else self.default_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    # ── Internal helpers ──────────────────────────────────────────────────────

    @classmethod
    def _from_dict(cls, d: dict[str, Any]) -> "AppConfig":
        def _merge(dc_class: type, src: dict) -> Any:
            """Create a dataclass instance, ignoring unknown keys."""
            known = {f.name for f in dataclasses.fields(dc_class)}
            filtered = {k: v for k, v in src.items() if k in known}
            return dc_class(**filtered)

        return cls(
            general=    _merge(GeneralConfig,       d.get("general",     {})),
            squid=      _merge(SquidConfig,         d.get("squid",       {})),
            irm_arm=    _merge(IrmArmConfig,        d.get("irm_arm",     {})),
            af_demag=   _merge(AfDemagConfig,       d.get("af_demag",    {})),
            vacuum=     _merge(VacuumConfig,        d.get("vacuum",      {})),
            data_files= _merge(DataFilesConfig,     d.get("data_files",  {})),
            changer=    _merge(ChangerConfig,       d.get("changer",     {})),
            calibration=_merge(CalibrationConfig,   d.get("calibration", {})),
            sequence=   _merge(SequenceTimesConfig, d.get("sequence",    {})),
        )
