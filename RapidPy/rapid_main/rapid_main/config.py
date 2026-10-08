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
import tempfile
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any


CONFIG_BACKUP_SCHEMA = "rapidpy.config.backup.v1"


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    """Durably publish JSON without exposing a partially written config."""

    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            json.dump(payload, handle, indent=2, ensure_ascii=False, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except Exception:
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Sub-configs (one per settings tab)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class GeneralConfig:
    data_dir:     str  = ""
    sample_dir:   str  = ""
    backup_dir:   str  = ""
    operator:     str  = ""
    operator_email: str = ""  # VB6 LoginEmail: receives this session's notices
    lab_name:     str  = "Paleomagnetism Laboratory"
    nocomm:       bool = False
    auto_save:    bool = True


#: VB6 ``frmSQUID.Connect`` hardcodes ``MSCommSquid.Settings = "1200,N,8,1"``.
VB6_SQUID_BAUD = 1200
SQUID_RANGE_LABELS = ("1×", "10×", "100×", "1000×")


def normalize_squid_range_label(label: object) -> str:
    """Return the canonical range wording, repairing damaged encodings.

    Older config files stored ``"1×"`` through a lossy encoding (``"1\ufffd"``
    or ``"1Ã—"``).  The leading digits still identify the 2G control rate.
    """

    text = str(label or "").strip()
    if "flux" in text.lower():
        return text
    digits = "".join(ch for ch in text if ch.isdigit())
    for canonical in reversed(SQUID_RANGE_LABELS):  # longest prefix first
        if digits.startswith(canonical[:-1]):
            return canonical
    return SQUID_RANGE_LABELS[0]


@dataclass
class SquidConfig:
    # Defaults follow the VB6 station: COMPortSquids=1, 1200,N,8,1 (hardcoded in
    # frmSQUID), ReadDelay=1 s, and one bracketed block per position (AvgSteps=1).
    port:             str   = "COM1"
    baud:             int   = VB6_SQUID_BAUD
    range_label:      str   = "1×"
    samples_per_pos:  int   = 1
    settle_time:      float = 1.0
    # Continuous read rate (live stream and motion capture).  The latch holds
    # default to the VB6 LatchCount/LatchData pauses; bracketed static reads
    # always keep the VB6 values regardless of these settings.
    stream_interval_ms:         int  = 500
    stream_axes:                str  = "XYZ"
    stream_counts_every:        int  = 1
    stream_latch_count_hold_ms: int  = 100
    stream_latch_data_hold_ms:  int  = 120
    # Opt-in continuous capture while the specimen descends, turns and ascends.
    motion_capture_enabled:     bool = False
    motion_capture_descent:     bool = True
    motion_capture_turns:       bool = True
    motion_capture_ascent:      bool = True
    motion_capture_dir:         str  = ""

    def __post_init__(self) -> None:
        self.range_label = normalize_squid_range_label(self.range_label)


@dataclass
class IrmArmConfig:
    irm_max_field:  float = 1000.0   # mT
    irm_axis:       str   = "Z (up-axis)"
    irm_ramp:       str   = "Slow (60 s)"
    irm_steps:      int   = 10
    arm_peak_af:    float = 100.0    # mT
    arm_bias:       float = 0.05     # mT
    arm_enabled: bool = False
    arm_bias_max_mT: float = 0.0
    arm_voltage_per_mT: float = 0.0
    arm_voltage_max: float = 0.0
    arm_board: int = -1
    arm_dac_channel: int = -1
    arm_gate_bit: int = -1
    arm_digital_port: int = 1  # MCC AUXPORT
    arm_voltage_range: int = 100  # MCC UNI10VOLTS
    arm_calibration_source: str = ""
    irm_voltage_slope: float = 0.01  # V / mT
    irm_voltage_intercept: float = 0.0  # V
    irm_max_voltage: float = 10.0  # V


@dataclass
class PulseIrmConfig:
    system: str = ""
    calibration_source: str = ""
    axial_enabled: bool = False
    transverse_enabled: bool = False
    backfield_enabled: bool = False
    coil_position: int = 0
    axial_calibrated: bool = False
    transverse_calibrated: bool = False
    axial_calibration: list[list[float]] = field(default_factory=list)  # capacitor V, mT
    transverse_calibration: list[list[float]] = field(default_factory=list)
    axial_min_mT: float = 0
    axial_max_mT: float = 0
    transverse_min_mT: float = 0
    transverse_max_mT: float = 0
    axial_capacitor_max_v: float = 0
    transverse_capacitor_max_v: float = 0
    control_v_per_capacitor_v: float = 0
    feedback_v_per_capacitor_v: float = 0
    control_max_v: float = 0
    asc_boost_at_min: float = 1.5
    asc_boost_at_max: float = 1.13
    trim_on_high: bool = False
    board: int = -1
    dac_channel: int = -1
    capacitor_adc_channel: int = -1
    voltage_range: int = 100
    digital_port: int = 1
    fire_bit: int = -1
    trim_bit: int = -1
    relay_board: int = -1
    irm_relay_bit: int = -1
    axial_relay_bit: int = -1
    transverse_relay_bit: int = -1
    charge_timeout_s: float = 90
    discharge_timeout_s: float = 30
    discharged_max_v: float = 10  # legacy zero-field charge threshold
    poll_s: float = .1
    temperature_channels: list[int] = field(default_factory=list)
    temperature_slope: float = 0
    temperature_offset: float = 0
    temperature_hot_c: float = 0


@dataclass
class AfDemagConfig:
    board:        int   = 0
    peak:         float = 180.0   # mT
    ramp_speed:   str   = "Medium (8 Hz)"
    settle:       float = 1.5     # s
    tumble:       bool  = False
    tumble_pause: float = 0.5     # s
    enabled: bool = False
    system: str = "ADWIN"
    coil_position: int = 0
    calibration_source: str = ""
    axial_calibration: list[list[float]] = field(default_factory=list)  # monitor V, field mT
    transverse_calibration: list[list[float]] = field(default_factory=list)
    axial_calibrated: bool = False
    transverse_calibrated: bool = False
    axial_min_mT: float = 0.0
    axial_max_mT: float = 0.0
    transverse_min_mT: float = 0.0
    transverse_max_mT: float = 0.0
    axial_frequency_hz: float = 0.0
    transverse_frequency_hz: float = 0.0
    axial_ramp_max_v: float = 0.0
    transverse_ramp_max_v: float = 0.0
    axial_monitor_max_v: float = 0.0
    transverse_monitor_max_v: float = 0.0
    axial_ramp_up_vps: float = 0.0
    transverse_ramp_up_vps: float = 0.0
    ramp_up_min_ms: float = 0.0
    ramp_up_max_ms: float = 0.0
    ramp_down_min_periods: int = 0
    ramp_down_max_periods: int = 0
    ramp_down_periods_per_v: float = 0.0
    hold_peak_periods: int = 0
    io_rate_hz: float = 0.0
    axial_relay_bit: int = -1
    transverse_relay_bit: int = -1
    bin_folder: str = ""
    boot_file: str = ""
    process_file: str = ""


@dataclass
class VacuumConfig:
    port:            str   = "COM1"
    baud:            int   = 9600
    target_pressure: float = 5.0    # mTorr
    warn_threshold:  float = 20.0   # mTorr
    auto_pump:       bool  = False
    poll_interval:   float = 2.0    # s
    dropoff_delay_s:  float = 1.0    # VB6 DropoffVacuumDelay


@dataclass
class SusceptibilityConfig:
    """Bartington bridge settings mapped from the legacy RAPID INI."""

    enabled:           bool  = False
    port:              str   = ""
    baud:              int   = 1200
    parity:            str   = "N"
    bytesize:          int   = 8
    stopbits:          float = 2.0
    response_timeout:  float = 35.0
    scale_factor:      float = 1.0
    moment_factor_cgs: float = 1.0e-5
    coil_position:     int   = 0


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
class MotorStationConfig:
    """Native Quicksilver calibration and per-axis serial wiring from the station."""

    calibration_source: str = ""
    ports: dict[str, str] = field(default_factory=dict)
    baud: int = 57600  # VB6 frmDCMotors: 57600,N,8,2
    addresses: dict[str, int] = field(default_factory=dict)
    controller: dict[str, float] = field(default_factory=dict)
    hole_slot: int = 0
    use_xy_table: bool | None = None
    xy_positions: dict[str, list[int]] = field(default_factory=dict)
    xy_home: list[int] = field(default_factory=list)


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
class MotionPositionsConfig:
    """Lift positions used by the bracketed SQUID measurement block.

    These map directly onto the legacy ``[SteppingMotor]`` INI keys. VB6
    measures at ``ZeroPos + SampleHeight / 2`` and ``MeasPos + SampleHeight / 2``
    where ``SampleHeight = SampleTop - SampleBottom``.

    They default to ``0`` (unconfigured) on purpose: RapidPy must not invent
    lift positions for a physical instrument. Until an operator imports the
    legacy INI or enters real values, hardware-mode preflight blocks.
    """

    zero_pos:      int = 0    # motor steps
    meas_pos:      int = 0    # motor steps
    sample_top:    int = 0    # motor steps
    sample_bottom: int = 0    # motor steps

    @property
    def sample_height(self) -> int:
        """VB6 ``SampleHeight = SampleTop - SampleBottom``."""
        return int(self.sample_top) - int(self.sample_bottom)

    def zero_position(self) -> int:
        """VB6 ``Int(ZeroPos + SampleHeight / 2)``."""
        import math
        return math.floor(self.zero_pos + self.sample_height / 2)

    def measurement_position(self) -> int:
        """VB6 ``Int(MeasPos + SampleHeight / 2)``."""
        import math
        return math.floor(self.meas_pos + self.sample_height / 2)

    @property
    def configured(self) -> bool:
        """True only when both lift positions are set and distinct."""
        return (
            int(self.zero_pos) != 0
            and int(self.meas_pos) != 0
            and self.zero_position() != self.measurement_position()
        )

    def unconfigured_reason(self) -> str:
        if int(self.zero_pos) == 0 or int(self.meas_pos) == 0:
            return (
                "Zero/measurement lift positions are not configured. Import the "
                "legacy INI ([SteppingMotor] ZeroPos/MeasPos/SampleTop/SampleBottom) "
                "or enter the measured positions before running on hardware."
            )
        if self.zero_position() == self.measurement_position():
            return "Zero and measurement lift positions are identical."
        return ""


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
    pulse_irm:  PulseIrmConfig     = field(default_factory=PulseIrmConfig)
    af_demag:   AfDemagConfig      = field(default_factory=AfDemagConfig)
    vacuum:     VacuumConfig       = field(default_factory=VacuumConfig)
    susceptibility: SusceptibilityConfig = field(default_factory=SusceptibilityConfig)
    data_files: DataFilesConfig    = field(default_factory=DataFilesConfig)
    changer:    ChangerConfig      = field(default_factory=ChangerConfig)
    motor_station: MotorStationConfig = field(default_factory=MotorStationConfig)
    calibration: CalibrationConfig = field(default_factory=CalibrationConfig)
    motion:     MotionPositionsConfig = field(default_factory=MotionPositionsConfig)
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
        _atomic_write_json(p, asdict(self))

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
            pulse_irm=  _merge(PulseIrmConfig,      d.get("pulse_irm",   {})),
            af_demag=   _merge(AfDemagConfig,       d.get("af_demag",    {})),
            vacuum=     _merge(VacuumConfig,        d.get("vacuum",      {})),
            susceptibility=_merge(SusceptibilityConfig, d.get("susceptibility", {})),
            data_files= _merge(DataFilesConfig,     d.get("data_files",  {})),
            changer=    _merge(ChangerConfig,       d.get("changer",     {})),
            motor_station=_merge(MotorStationConfig, d.get("motor_station", {})),
            calibration=_merge(CalibrationConfig,   d.get("calibration", {})),
            motion=     _merge(MotionPositionsConfig, d.get("motion",    {})),
            sequence=   _merge(SequenceTimesConfig, d.get("sequence",    {})),
        )


def write_config_backup(config: AppConfig, path: Path | str) -> Path:
    """Write a versioned, human-readable snapshot of every application setting."""

    target = Path(path)
    payload: dict[str, Any] = {
        "schema": CONFIG_BACKUP_SCHEMA,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "config": asdict(config),
    }
    _atomic_write_json(target, payload)
    return target


def read_config_backup(path: Path | str) -> AppConfig:
    """Strictly validate and decode a settings backup before any live mutation."""

    source = Path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read settings backup: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("Settings backup root must be a JSON object.")
    if payload.get("schema") != CONFIG_BACKUP_SCHEMA:
        raise ValueError(
            f"Unsupported settings backup schema: {payload.get('schema')!r}."
        )
    config_payload = payload.get("config")
    if not isinstance(config_payload, dict):
        raise ValueError("Settings backup does not contain a configuration object.")
    required_sections = {field.name for field in dataclasses.fields(AppConfig)}
    missing = sorted(required_sections - set(config_payload))
    if missing:
        raise ValueError("Settings backup is missing sections: " + ", ".join(missing))
    try:
        return AppConfig._from_dict(config_payload)
    except (AttributeError, TypeError, ValueError) as exc:
        raise ValueError(f"Settings backup contains invalid values: {exc}") from exc
