"""Utilities to import legacy VB6 INI settings into rapid_main AppConfig.

This module intentionally performs a safe, conservative mapping:

* map high-confidence keys into existing rapid_main settings,
* ignore unknown/obsolete keys without failing,
* and return a report so the caller can show user-facing import feedback.
"""

from __future__ import annotations

import configparser
from dataclasses import dataclass
from pathlib import Path

from .config import AppConfig


@dataclass
class LegacyIniImportReport:
    """Summary of a legacy INI import attempt."""

    source: Path
    mapped_fields: list[str]
    unmapped_fields: list[str]
    warnings: list[str]


def _read_legacy_ini(path: Path) -> configparser.ConfigParser:
    if not path.exists():
        raise FileNotFoundError(f"Legacy INI file not found: {path}")
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.optionxform = str
    cfg.read(path, encoding="utf-8")
    if not cfg.sections():
        raise ValueError(f"No sections found in legacy INI: {path}")
    return cfg


def _value(cfg: configparser.ConfigParser, section: str, key: str) -> str | None:
    if not cfg.has_section(section):
        return None
    raw = cfg.get(section, key, fallback=None)
    return None if raw is None else str(raw).strip()


def _parse_bool(raw: str | None, *, default: bool, section_key: str, field: str, warnings: list[str]) -> bool:
    if raw is None:
        return default
    text = raw.strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    warnings.append(f"{section_key}.{field}: '{raw}' is not a recognized boolean, using default {default}")
    return default


def _parse_int(raw: str | None, *, default: int, section_key: str, field: str, warnings: list[str]) -> int:
    if raw is None:
        return default
    try:
        value = float(raw.strip())
        return int(round(value))
    except ValueError:
        warnings.append(f"{section_key}.{field}: '{raw}' is not numeric, using default {default}")
        return default


def _parse_float(
    raw: str | None,
    *,
    default: float | None,
    section_key: str,
    field: str,
    warnings: list[str],
) -> float | None:
    if raw is None:
        return default
    try:
        return float(raw.strip())
    except ValueError:
        warnings.append(f"{section_key}.{field}: '{raw}' is not numeric, using default {default}")
        return default


def _normalize_com_port(raw: str | None, *, field: str, warnings: list[str]) -> str | None:
    if raw is None:
        return None
    text = raw.strip()
    if not text:
        return None
    if text.upper().startswith("COM"):
        return text.upper()
    try:
        port = int(float(text))
    except ValueError:
        warnings.append(f"{field}: '{raw}' is not a valid COM port value")
        return None
    if port <= 0:
        warnings.append(f"{field}: '{raw}' is not a valid COM port value")
        return None
    return f"COM{port}"


def _ramp_speed_label(raw_rate: float) -> str:
    if raw_rate <= 3.5:
        return "Slow (3 Hz)"
    if raw_rate <= 9:
        return "Medium (8 Hz)"
    return "Fast (15 Hz)"


def import_vb6_ini(config: AppConfig, path: str | Path) -> LegacyIniImportReport:
    """Import a VB6 INI file into an existing :class:`AppConfig` in place.

    Returns a mapping/warning report to allow callers to surface transparent
    operator feedback.
    """

    ini_path = Path(path)
    parser = _read_legacy_ini(ini_path)
    mapped_fields: list[str] = []
    unmapped_fields: list[str] = []
    warnings: list[str] = []
    mapped_keys: set[tuple[str, str]] = set()

    def mark_mapped(section: str, key: str, destination: str) -> None:
        mapped_fields.append(f"{section}.{key} -> {destination}")
        mapped_keys.add((section, key))

    # Program section (legacy general settings)
    program_section = "Program"
    raw = _value(parser, program_section, "NoCommMode")
    if raw is not None:
        config.general.nocomm = _parse_bool(raw, default=config.general.nocomm,
                                           section_key=program_section, field="NoCommMode", warnings=warnings)
        mark_mapped(program_section, "NoCommMode", "general.nocomm")

    raw = _value(parser, program_section, "LastLogin")
    if raw:
        config.general.operator = raw
        mark_mapped(program_section, "LastLogin", "general.operator")

    raw = _value(parser, program_section, "DefaultPath")
    if raw:
        config.general.data_dir = raw
        mark_mapped(program_section, "DefaultPath", "general.data_dir")

    raw = _value(parser, program_section, "DefaultBackupDrive")
    if raw:
        # VB6 often uses a drive letter + colon; keep as-is.
        config.general.backup_dir = raw
        mark_mapped(program_section, "DefaultBackupDrive", "general.backup_dir")

    raw = _value(parser, program_section, "UsageFile")
    if raw:
        sample_file = Path(raw)
        config.general.sample_dir = str(sample_file.parent)
        mark_mapped(program_section, "UsageFile", "general.sample_dir")

    # COM ports (legacy communication mapping)
    com_section = "COMPorts"
    raw = _value(parser, com_section, "COMPortSquids")
    mapped = _normalize_com_port(raw, field="COMPorts.COMPortSquids", warnings=warnings)
    if mapped:
        config.squid.port = mapped
        mark_mapped(com_section, "COMPortSquids", "squid.port")

    raw = _value(parser, com_section, "COMPortVacuum")
    mapped = _normalize_com_port(raw, field="COMPorts.COMPortVacuum", warnings=warnings)
    if mapped:
        config.vacuum.port = mapped
        mark_mapped(com_section, "COMPortVacuum", "vacuum.port")

    raw = _value(parser, com_section, "COMPortChanger")
    mapped = _normalize_com_port(raw, field="COMPorts.COMPortChanger", warnings=warnings)
    if mapped:
        config.changer.port = mapped
        mark_mapped(com_section, "COMPortChanger", "changer.port")

    # Magnetometer calibration
    mag_section = "MagnetometerCalibration"
    if _value(parser, mag_section, "XCal") is not None:
        config.calibration.cal_x = _parse_float(
            _value(parser, mag_section, "XCal"),
            default=config.calibration.cal_x,
            section_key=mag_section,
            field="XCal",
            warnings=warnings,
        )
        mark_mapped(mag_section, "XCal", "calibration.cal_x")
    if _value(parser, mag_section, "YCal") is not None:
        config.calibration.cal_y = _parse_float(
            _value(parser, mag_section, "YCal"),
            default=config.calibration.cal_y,
            section_key=mag_section,
            field="YCal",
            warnings=warnings,
        )
        mark_mapped(mag_section, "YCal", "calibration.cal_y")
    if _value(parser, mag_section, "ZCal") is not None:
        config.calibration.cal_z = _parse_float(
            _value(parser, mag_section, "ZCal"),
            default=config.calibration.cal_z,
            section_key=mag_section,
            field="ZCal",
            warnings=warnings,
        )
        mark_mapped(mag_section, "ZCal", "calibration.cal_z")
    if _value(parser, mag_section, "RangeFact") is not None:
        config.calibration.range_factor = _parse_float(
            _value(parser, mag_section, "RangeFact"),
            default=config.calibration.range_factor,
            section_key=mag_section,
            field="RangeFact",
            warnings=warnings,
        )
        mark_mapped(mag_section, "RangeFact", "calibration.range_factor")

    raw = _value(parser, mag_section, "ReadDelay")
    if raw is not None:
        config.squid.settle_time = _parse_float(
            raw,
            default=config.squid.settle_time,
            section_key=mag_section,
            field="ReadDelay",
            warnings=warnings,
        )
        mark_mapped(mag_section, "ReadDelay", "squid.settle_time")

    # AF controls
    af_section = "AF"
    raw = _value(parser, af_section, "AFWait")
    if raw is not None:
        config.af_demag.settle = _parse_float(
            raw,
            default=config.af_demag.settle,
            section_key=af_section,
            field="AFWait",
            warnings=warnings,
        )
        mark_mapped(af_section, "AFWait", "af_demag.settle")

    raw = _value(parser, af_section, "AFRampRate")
    if raw is not None:
        rate = _parse_float(
            raw,
            default=0.0,
            section_key=af_section,
            field="AFRampRate",
            warnings=warnings,
        )
        if rate > 0:
            config.af_demag.ramp_speed = _ramp_speed_label(rate)
            mark_mapped(af_section, "AFRampRate", "af_demag.ramp_speed")

    # IRM / ARM controls
    arm_section = "ARM"
    raw = _value(parser, arm_section, "ARMMax")
    if raw is not None:
        config.irm_arm.arm_peak_af = _parse_float(
            raw,
            default=config.irm_arm.arm_peak_af,
            section_key=arm_section,
            field="ARMMax",
            warnings=warnings,
        )
        mark_mapped(arm_section, "ARMMax", "irm_arm.arm_peak_af")

    raw = _value(parser, arm_section, "ARMVoltGauss")
    if raw is not None:
        config.irm_arm.arm_bias = _parse_float(
            raw,
            default=config.irm_arm.arm_bias,
            section_key=arm_section,
            field="ARMVoltGauss",
            warnings=warnings,
        )
        mark_mapped(arm_section, "ARMVoltGauss", "irm_arm.arm_bias")

    irm_axial = _parse_float(
        _value(parser, "IRMAxial", "IRMAxialVoltMax"),
        default=None,
        section_key="IRMAxial",
        field="IRMAxialVoltMax",
        warnings=warnings,
    )
    if irm_axial is not None:
        config.irm_arm.irm_max_field = irm_axial
        mark_mapped("IRMAxial", "IRMAxialVoltMax", "irm_arm.irm_max_field")

    irm_trans = _parse_float(
        _value(parser, "IRMTrans", "IRMTransVoltMax"),
        default=None,
        section_key="IRMTrans",
        field="IRMTransVoltMax",
        warnings=warnings,
    )
    if irm_trans is not None:
        if config.irm_arm.irm_max_field < irm_trans:
            config.irm_arm.irm_max_field = irm_trans
        mark_mapped("IRMTrans", "IRMTransVoltMax", "irm_arm.irm_max_field")

    raw = _value(parser, "IRMPulse", "IRMAxis")
    if raw is not None and raw.strip():
        config.irm_arm.irm_axis = f"{raw.strip().upper()} axis" if raw.strip().upper() in {"X", "Y", "Z"} else raw.strip()
        mark_mapped("IRMPulse", "IRMAxis", "irm_arm.irm_axis")

    # Vacuum
    vac_section = "Vacuum"
    raw = _value(parser, vac_section, "DoVacuumReset")
    if raw is not None:
        config.vacuum.auto_pump = _parse_bool(
            raw,
            default=config.vacuum.auto_pump,
            section_key=vac_section,
            field="DoVacuumReset",
            warnings=warnings,
        )
        mark_mapped(vac_section, "DoVacuumReset", "vacuum.auto_pump")

    # Optional safety / quality mappings (if present)
    raw = _value(parser, "SampleChanger", "HoleSlotNum")
    if raw is not None:
        config.changer.speed_z = _parse_int(
            raw,
            default=int(config.changer.speed_z),
            section_key="SampleChanger",
            field="HoleSlotNum",
            warnings=warnings,
        )
        mark_mapped("SampleChanger", "HoleSlotNum", "changer.speed_z")
        warnings.append(
            "SampleChanger.HoleSlotNum is migrated to changer.speed_z as a best-effort placeholder "
            "because no dedicated rapid_main field exists."
        )

    # Compositional summary of unmapped keys
    for sec_name in parser.sections():
        if not parser.has_section(sec_name):
            continue
        for key, _raw in parser.items(sec_name):
            if (sec_name, key) not in mapped_keys:
                unmapped_fields.append(f"{sec_name}.{key}")

    if unmapped_fields:
        if len(unmapped_fields) > 40:
            preview = ", ".join(unmapped_fields[:40]) + ", ..."
        else:
            preview = ", ".join(unmapped_fields)
        warnings.append(f"Unmapped legacy entries: {preview}")

    return LegacyIniImportReport(
        source=ini_path,
        mapped_fields=mapped_fields,
        unmapped_fields=unmapped_fields,
        warnings=warnings,
    )
