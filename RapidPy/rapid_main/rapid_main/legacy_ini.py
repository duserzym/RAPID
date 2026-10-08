"""Utilities to import legacy VB6 INI settings into rapid_main AppConfig.

This module intentionally performs a safe, conservative mapping:

* map high-confidence keys into existing rapid_main settings,
* ignore unknown/obsolete keys without failing,
* and return a report so the caller can show user-facing import feedback.
"""

from __future__ import annotations

import configparser
import math
import re
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

    # SteppingMotor section: lift positions used by the measurement block.
    motor_section = "SteppingMotor"
    for key, field_name in (
        ("ZeroPos", "zero_pos"),
        ("MeasPos", "meas_pos"),
        ("SampleTop", "sample_top"),
        ("SampleBottom", "sample_bottom"),
    ):
        raw = _value(parser, motor_section, key)
        if raw is not None:
            setattr(
                config.motion,
                field_name,
                _parse_int(
                    raw,
                    default=getattr(config.motion, field_name),
                    section_key=motor_section,
                    field=key,
                    warnings=warnings,
                ),
            )
            mark_mapped(motor_section, key, f"motion.{field_name}")

    raw = _value(parser, motor_section, "SCoilPos")
    if raw is not None:
        config.susceptibility.coil_position = _parse_int(
            raw,
            default=config.susceptibility.coil_position,
            section_key=motor_section,
            field="SCoilPos",
            warnings=warnings,
        )
        mark_mapped(motor_section, "SCoilPos", "susceptibility.coil_position")

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
        # VB6 never stored SQUID framing: frmSQUID.Connect hardcodes 1200,N,8,1.
        config.squid.baud = 1200
        mapped_fields.append("frmSQUID.Connect 1200,N,8,1 (hardcoded) -> squid.baud")

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

    raw = _value(parser, com_section, "COMPortSusceptibility")
    mapped = _normalize_com_port(
        raw, field="COMPorts.COMPortSusceptibility", warnings=warnings
    )
    if mapped:
        config.susceptibility.port = mapped
        mark_mapped(com_section, "COMPortSusceptibility", "susceptibility.port")

    raw = _value(parser, com_section, "SusceptibilitySettings")
    if raw is not None:
        parts = [part.strip() for part in raw.split(",")]
        if len(parts) == 4:
            config.susceptibility.baud = _parse_int(
                parts[0],
                default=config.susceptibility.baud,
                section_key=com_section,
                field="SusceptibilitySettings.baud",
                warnings=warnings,
            )
            config.susceptibility.parity = parts[1].upper()
            config.susceptibility.bytesize = _parse_int(
                parts[2],
                default=config.susceptibility.bytesize,
                section_key=com_section,
                field="SusceptibilitySettings.bytesize",
                warnings=warnings,
            )
            parsed_stopbits = _parse_float(
                parts[3],
                default=config.susceptibility.stopbits,
                section_key=com_section,
                field="SusceptibilitySettings.stopbits",
                warnings=warnings,
            )
            config.susceptibility.stopbits = float(parsed_stopbits)
            mark_mapped(
                com_section,
                "SusceptibilitySettings",
                "susceptibility serial framing",
            )
        else:
            warnings.append(
                "COMPorts.SusceptibilitySettings must use baud,parity,data,stop format"
            )

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
    af = config.af_demag
    units = (_value(parser, "AF", "AFUnits") or "").upper()
    field_factor = {"G": 0.1, "GAUSS": 0.1, "OE": 0.1, "OERSTED": 0.1, "MT": 1.0}.get(units)
    if field_factor is None:
        warnings.append("AFUnits is missing or unsupported; live AF calibration was not imported.")
    else:
        af.calibration_source = str(ini_path.resolve())
        for coil, section, prefix in (("axial", "AFAxial", "AFAxial"), ("transverse", "AFTrans", "AFTrans")):
            raw_count = _value(parser, section, prefix + "Count")
            if raw_count is not None:
                try:
                    count = int(raw_count)
                    if not 2 <= count <= 10000:
                        raise ValueError("at least two calibration points required")
                    points = [[float(_value(parser, section, f"{prefix}X{i}")),
                               float(_value(parser, section, f"{prefix}Y{i}")) * field_factor]
                              for i in range(1, count + 1)]
                    if any(not math.isfinite(x) or not math.isfinite(y) for x, y in points):
                        raise ValueError("non-finite calibration")
                    setattr(af, coil + "_calibration", points)
                    setattr(af, coil + "_calibrated", _parse_bool(_value(parser, section, prefix + "CalDone"), default=False,
                           section_key=section, field=prefix + "CalDone", warnings=warnings))
                    for i in range(1, count + 1):
                        mark_mapped(section, f"{prefix}X{i}", f"af_demag.{coil}_calibration")
                        mark_mapped(section, f"{prefix}Y{i}", f"af_demag.{coil}_calibration (mT)")
                    mark_mapped(section, prefix + "Count", f"af_demag.{coil}_calibration")
                    mark_mapped(section, prefix + "CalDone", f"af_demag.{coil}_calibrated")
                except (TypeError, ValueError) as exc:
                    setattr(af, coil + "_calibrated", False)
                    warnings.append(f"{section} calibration not imported: {exc}")
            for suffix, destination, factor in (("Min", "min_mT", field_factor), ("Max", "max_mT", field_factor),
                    ("ResFreq", "frequency_hz", 1), ("RampMax", "ramp_max_v", 1), ("MonMax", "monitor_max_v", 1)):
                raw_value = _value(parser, section, prefix + suffix)
                if raw_value is not None:
                    setattr(af, coil + "_" + destination, _parse_float(raw_value, default=0,
                            section_key=section, field=prefix + suffix, warnings=warnings) * factor)
                    mark_mapped(section, prefix + suffix, f"af_demag.{coil}_{destination}")
        af.peak = max(af.axial_max_mT, af.transverse_max_mT) or af.peak
        mark_mapped("AF", "AFUnits", "af_demag calibration fields converted to mT")
    for section, key, name in (("SteppingMotor", "AFPos", "coil_position"),
            ("AF", "AxialRampUpVoltsPerSec", "axial_ramp_up_vps"),
            ("AF", "TransRampUpVoltsPerSec", "transverse_ramp_up_vps"),
            ("AF", "MinRampUpTime_ms", "ramp_up_min_ms"), ("AF", "MaxRampUpTime_ms", "ramp_up_max_ms"),
            ("AF", "MinRampDown_NumPeriods", "ramp_down_min_periods"),
            ("AF", "MaxRampDown_NumPeriods", "ramp_down_max_periods"),
            ("AF", "RampDownNumPeriodsPerVolt", "ramp_down_periods_per_v"),
            ("AF", "HoldAtPeakField_NumPeriods", "hold_peak_periods")):
        raw_value = _value(parser, section, key)
        if raw_value is not None:
            current = getattr(af, name)
            convert = _parse_int if isinstance(current, int) else _parse_float
            setattr(af, name, convert(raw_value, default=current, section_key=section, field=key, warnings=warnings))
            mark_mapped(section, key, f"af_demag.{name}")
    for key, name in (("AFSystem", "system"), ("ADWINBinFolderPath", "bin_folder"),
                      ("ADWINBootFile", "boot_file"), ("ADWINRampProgFile", "process_file")):
        raw_value = _value(parser, "AF", key)
        if raw_value is not None:
            setattr(af, name, raw_value)
            mark_mapped("AF", key, f"af_demag.{name}")
    enabled = _value(parser, "Modules", "EnableAF")
    if enabled is not None:
        af.enabled = _parse_bool(enabled, default=False, section_key="Modules", field="EnableAF", warnings=warnings)
        mark_mapped("Modules", "EnableAF", "af_demag.enabled")
    # Resolve the named legacy relay through its board channel entry, not the
    # ordinal in the symbolic DO reference (the two can differ).
    relay_board = None
    for key, name in (("AFAxialRelay", "axial_relay_bit"), ("AFTransRelay", "transverse_relay_bit")):
        reference = _value(parser, "Channels", key)
        if reference:
            match = re.fullmatch(r"DO-(\d+)-CH(\d+)", reference)
            mapping = _value(parser, "Boards", reference)
            try:
                if not match or not mapping:
                    raise ValueError("relay channel has no board mapping")
                board_index = match.group(1)
                protocol = int(_value(parser, "Boards", "CommProtocol" + board_index))
                if protocol != 2:
                    raise ValueError("AF relay does not belong to an ADwin board")
                board_number = int(_value(parser, "Boards", "BoardNum" + board_index))
                if relay_board is not None and relay_board != board_number:
                    raise ValueError("AF coil relays must belong to the same ADwin board")
                relay_board = board_number
                af.board = board_number
                setattr(af, name, int(mapping.split(",")[1]))
                mark_mapped("Channels", key, f"af_demag.{name}")
            except (TypeError, ValueError, IndexError) as exc:
                setattr(af, name, -1)
                warnings.append(f"{key}: {exc}")
    wave_count = _parse_int(_value(parser, "WaveForms", "WaveFormCount"), default=0,
                            section_key="WaveForms", field="WaveFormCount", warnings=warnings)
    for index in range(max(0, min(wave_count, 10000))):
        if (_value(parser, "WaveForms", f"WaveName{index}") or "").upper() == "AFRAMPUP":
            af.io_rate_hz = _parse_float(_value(parser, "WaveForms", f"IORate{index}"), default=0,
                                        section_key="WaveForms", field=f"IORate{index}", warnings=warnings)
            mark_mapped("WaveForms", f"IORate{index}", "af_demag.io_rate_hz")
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
        config.irm_arm.arm_bias_max_mT = .1 * _parse_float(
            raw,
            default=config.irm_arm.arm_bias_max_mT * 10,
            section_key=arm_section,
            field="ARMMax",
            warnings=warnings,
        )
        mark_mapped(arm_section, "ARMMax", "irm_arm.arm_bias_max_mT (G converted to mT)")

    raw = _value(parser, arm_section, "ARMVoltGauss")
    if raw is not None:
        config.irm_arm.arm_voltage_per_mT = 10 * _parse_float(
            raw,
            default=config.irm_arm.arm_voltage_per_mT / 10,
            section_key=arm_section,
            field="ARMVoltGauss",
            warnings=warnings,
        )
        mark_mapped(arm_section, "ARMVoltGauss", "irm_arm.arm_voltage_per_mT (V/G converted to V/mT)")

    arm = config.irm_arm
    arm.arm_calibration_source = str(Path(path).resolve())
    raw = _value(parser, "Modules", "EnableARM")
    if raw is not None:
        arm.arm_enabled = _parse_bool(raw, default=False, section_key="Modules", field="EnableARM", warnings=warnings)
        mark_mapped("Modules", "EnableARM", "irm_arm.arm_enabled")
    raw = _value(parser, "ARM", "ARMVoltMax")
    if raw is not None:
        arm.arm_voltage_max = _parse_float(raw, default=0, section_key="ARM", field="ARMVoltMax", warnings=warnings)
        mark_mapped("ARM", "ARMVoltMax", "irm_arm.arm_voltage_max")
    arm_board = None
    for key, kind, name in (("ARMVoltageOut", "AO", "arm_dac_channel"), ("ARMSet", "DO", "arm_gate_bit")):
        reference = _value(parser, "Channels", key)
        if reference:
            try:
                match = re.fullmatch(kind + r"-(\d+)-CH(\d+)", reference)
                if not match:
                    raise ValueError("invalid ARM channel reference")
                index = match.group(1)
                if int(_value(parser, "Boards", "CommProtocol" + index)) != 1:
                    raise ValueError("ARM bias channel requires an MCC board")
                number = int(_value(parser, "Boards", "BoardNum" + index))
                if arm_board is not None and arm_board != number:
                    raise ValueError("ARM bias outputs must use the same MCC board")
                arm_board = number
                arm.arm_board = number
                setattr(arm, name, int(_value(parser, "Boards", reference).split(",")[1]))
                arm.arm_voltage_range = int(_value(parser, "Boards", "RangeType" + index))
                arm.arm_digital_port = int(_value(parser, "Boards", "DOutPortType" + index))
                mark_mapped("Channels", key, "irm_arm." + name)
            except (AttributeError, TypeError, ValueError, IndexError) as exc:
                setattr(arm, name, -1)
                warnings.append(f"{key}: {exc}")

    pulse = config.pulse_irm
    pulse.calibration_source = str(Path(path).resolve())
    for section,key,name in (("IRMPulse","IRMSystem","system"),("IRMPulse","PulseMCCVoltConversion","control_v_per_capacitor_v"),
                             ("IRMPulse","PulseReturnMCCVoltConversion","feedback_v_per_capacitor_v"),
                             ("IRMPulse","PulseVoltMax","control_max_v"),("IRMPulse","AscSetVoltageMinBoostMultiplier","asc_boost_at_min"),
                             ("IRMPulse","AscSetVoltageMaxBoostMultiplier","asc_boost_at_max"),("IRMPulse","TrimOnTrue","trim_on_high"),
                             ("Modules","EnableAxialIRM","axial_enabled"),("Modules","EnableTransIRM","transverse_enabled"),
                             ("Modules","EnableIRMBackfield","backfield_enabled"),("SteppingMotor","IRMPos","coil_position")):
        raw = _value(parser,section,key)
        if raw is None:
            continue
        old = getattr(pulse,name)
        if isinstance(old,bool):
            value = _parse_bool(raw,default=old,section_key=section,field=key,warnings=warnings)
        elif isinstance(old,str):
            value = raw.strip()
        elif name=="coil_position":
            value = _parse_int(raw,default=int(old),section_key=section,field=key,warnings=warnings)
        else:
            value = _parse_float(raw,default=old,section_key=section,field=key,warnings=warnings)
        setattr(pulse,name,value)
        mark_mapped(section,key,"pulse_irm."+name)
    for coil,prefix in (("axial","Axial"),("transverse","Trans")):
        section = "IRM"+prefix
        count = _parse_int(_value(parser,section,"Pulse"+prefix+"Count"),default=0,section_key=section,field="Count",warnings=warnings)
        setattr(pulse,coil+"_calibrated",False)
        if count >= 2 and count <= 10000:
            try:
                points = [[float(_value(parser,section,f"Pulse{prefix}X{i}")),.1*float(_value(parser,section,f"Pulse{prefix}Y{i}"))] for i in range(1,count+1)]
                setattr(pulse,coil+"_calibration",points)
                done = _parse_bool(_value(parser,section,f"IRM{prefix}CalDone"),default=False,section_key=section,field="CalDone",warnings=warnings)
                setattr(pulse,coil+"_calibrated",done)
                for i in range(1,count+1):
                    for axis in ("X","Y"):
                        mark_mapped(section,f"Pulse{prefix}{axis}{i}",f"pulse_irm.{coil}_calibration")
            except (TypeError,ValueError) as exc:
                warnings.append(f"{section} pulse calibration: {exc}")
        for suffix,name,scale in (("Min","min_mT",.1),("Max","max_mT",.1)):
            raw = _value(parser,section,"Pulse"+prefix+suffix)
            if raw is not None:
                setattr(pulse,coil+"_"+name,scale*_parse_float(raw,default=0,section_key=section,field=suffix,warnings=warnings))
                mark_mapped(section,"Pulse"+prefix+suffix,f"pulse_irm.{coil}_{name}")
        raw = _value(parser,section,"IRM"+prefix+"VoltMax")
        if raw is not None:
            setattr(pulse,coil+"_capacitor_max_v",_parse_float(raw,default=0,section_key=section,field="VoltMax",warnings=warnings))
            mark_mapped(section,"IRM"+prefix+"VoltMax",f"pulse_irm.{coil}_capacitor_max_v")

    mcc_board = None
    for key,kind,name in (("IRMVoltageOut","AO","dac_channel"),("IRMCapacitorVoltageIn","AI","capacitor_adc_channel"),
                          ("IRMFire","DO","fire_bit"),("IRMTrim","DO","trim_bit")):
        reference = _value(parser,"Channels",key)
        if reference is None:
            continue
        try:
            match = re.fullmatch(kind+r"-(\d+)-CH(\d+)",reference)
            if not match:
                raise ValueError("invalid pulse IRM channel reference")
            index = match.group(1)
            if int(_value(parser,"Boards","CommProtocol"+index)) != 1:
                raise ValueError("pulse charge/readback/fire/trim requires MCC channels")
            number = int(_value(parser,"Boards","BoardNum"+index))
            if mcc_board is not None and mcc_board != number:
                raise ValueError("pulse circuit channels must use one MCC board")
            mcc_board = number
            pulse.board = number
            setattr(pulse,name,int(_value(parser,"Boards",reference).split(",")[1]))
            pulse.voltage_range = int(_value(parser,"Boards","RangeType"+index))
            pulse.digital_port = int(_value(parser,"Boards","DOutPortType"+index))
            mark_mapped("Channels",key,"pulse_irm."+name)
        except (AttributeError,TypeError,ValueError,IndexError) as exc:
            setattr(pulse,name,-1)
            warnings.append(f"{key}: {exc}")
    for key,name in (("IRMRelay","irm_relay_bit"),("AFAxialRelay","axial_relay_bit"),("AFTransRelay","transverse_relay_bit")):
        reference = _value(parser,"Channels",key)
        if reference is None:
            continue
        try:
            match = re.fullmatch(r"DO-(\d+)-CH(\d+)",reference)
            if not match:
                raise ValueError("invalid pulse coil relay reference")
            index = match.group(1)
            if int(_value(parser,"Boards","CommProtocol"+index)) != 2:
                raise ValueError("pulse coil selection requires ADwin relay channels")
            number = int(_value(parser,"Boards","BoardNum"+index))
            if pulse.relay_board >= 0 and pulse.relay_board != number:
                raise ValueError("pulse coil relays must use one ADwin board")
            pulse.relay_board = number
            setattr(pulse,name,int(_value(parser,"Boards",reference).split(",")[1]))
            mark_mapped("Channels",key,"pulse_irm."+name)
        except (AttributeError,TypeError,ValueError,IndexError) as exc:
            setattr(pulse,name,-1)
            warnings.append(f"{key}: {exc}")

    pulse.temperature_channels = []
    for index in (1,2):
        enabled = _parse_bool(_value(parser,"Modules",f"EnableT{index}"),default=False,section_key="Modules",field=f"EnableT{index}",warnings=warnings)
        if not enabled:
            continue
        try:
            reference = _value(parser,"Channels",f"AnalogT{index}")
            match = re.fullmatch(r"AI-(\d+)-CH(\d+)",reference or "")
            if not match or int(_value(parser,"Boards","BoardNum"+match.group(1)))!=pulse.board:
                raise ValueError("pulse temperature sensor must use its configured MCC board")
            channel = int(_value(parser,"Boards",reference).split(",")[1])
            pulse.temperature_channels.append(channel)
            mark_mapped("Channels",f"AnalogT{index}","pulse_irm.temperature_channels")
        except (AttributeError,TypeError,ValueError,IndexError) as exc:
            pulse.temperature_channels.append(-1)
            warnings.append(f"AnalogT{index}: {exc}")
    for key,name in (("TSlope","temperature_slope"),("Toffset","temperature_offset"),("Thot","temperature_hot_c")):
        raw = _value(parser,"AF",key)
        if raw is not None:
            setattr(pulse,name,_parse_float(raw,default=0,section_key="AF",field=key,warnings=warnings))
            mark_mapped("AF",key,"pulse_irm."+name)

    calibrated_limits = [getattr(pulse,coil+"_max_mT") for coil in ("axial","transverse")
                         if getattr(pulse,coil+"_calibrated") and math.isfinite(getattr(pulse,coil+"_max_mT")) and getattr(pulse,coil+"_max_mT")>0]
    if calibrated_limits:
        config.irm_arm.irm_max_field = max(calibrated_limits)

    raw = _value(parser, "IRMPulse", "IRMAxis")
    if raw is not None and raw.strip():
        config.irm_arm.irm_axis = {"X":"X","Y":"Y","Z":"Z (up-axis)"}.get(raw.strip().upper(),raw.strip())
        mark_mapped("IRMPulse", "IRMAxis", "irm_arm.irm_axis")

    # Vacuum
    vac_section = "Vacuum"
    raw = _value(parser, vac_section, 'DropoffVacuumDelay')
    delay = 1.0 if raw is None else _parse_float(raw, default=-1.0,
        section_key=vac_section, field='DropoffVacuumDelay', warnings=warnings)
    if not math.isfinite(delay) or delay < 0:
        warnings.append('Vacuum.DropoffVacuumDelay: finite nonnegative delay required; transfer delay remains unaccepted')
        delay = -1.0
    config.vacuum.dropoff_delay_s = delay
    if raw is not None:
        mark_mapped(vac_section, 'DropoffVacuumDelay', 'vacuum.dropoff_delay_s')
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

    susc_cal_section = "SusceptibilityCalibration"
    for key, field_name in (
        ("SusceptibilityMomentFactorCGS", "moment_factor_cgs"),
        ("SusceptibilityScaleFactor", "scale_factor"),
    ):
        raw = _value(parser, susc_cal_section, key)
        if raw is not None:
            parsed = _parse_float(
                raw,
                default=getattr(config.susceptibility, field_name),
                section_key=susc_cal_section,
                field=key,
                warnings=warnings,
            )
            setattr(config.susceptibility, field_name, float(parsed))
            mark_mapped(susc_cal_section, key, f"susceptibility.{field_name}")

    raw = _value(parser, "Modules", "EnableSusceptibility")
    if raw is not None:
        config.susceptibility.enabled = _parse_bool(
            raw,
            default=config.susceptibility.enabled,
            section_key="Modules",
            field="EnableSusceptibility",
            warnings=warnings,
        )
        mark_mapped("Modules", "EnableSusceptibility", "susceptibility.enabled")

    # Native motion calibration must not be confused with UI speed percentages.
    station = config.motor_station
    # A new station file is one calibration provenance. Never combine omitted
    # wiring/geometry with values accepted from an earlier file.
    station.calibration_source = ''
    station.ports = {}
    station.addresses = {}
    station.controller = {}
    station.hole_slot = 0
    station.xy_positions = {}
    station.xy_home = []
    station.use_xy_table = None
    for axis, port_key, address_key in (
        ("changer_x", "COMPortChanger", "MotorIDChanger"),
        ("changer_y", "COMPortChangerY", "MotorIDChangerY"),
        ("updown", "COMPortUpDown", "MotorIDUpDown"),
        ("turning", "COMPortTurning", "MotorIDTurning"),
    ):
        port = _normalize_com_port(_value(parser, "COMPorts", port_key), field=port_key, warnings=warnings)
        address = _value(parser, "MotorPrograms", address_key)
        if port:
            station.ports[axis] = port
            mark_mapped("COMPorts", port_key, f"motor_station.ports.{axis}")
        if address is not None:
            station.addresses[axis] = _parse_int(address, default=0, section_key="MotorPrograms", field=address_key, warnings=warnings)
            mark_mapped("MotorPrograms", address_key, f"motor_station.addresses.{axis}")
    for section, key, destination in (
        ("SampleChanger", "SlotMin", "slot_min"),
        ("SampleChanger", "SlotMax", "slot_max"),
        ("SampleChanger", "OneStep", "one_step"),
        ("SteppingMotor", "SampleHoleAlignmentOffset", "sample_hole_alignment_offset"),
        ("SteppingMotor", "ChangerSpeed", "changer_speed"),
        ("SteppingMotor", "TurnerSpeed", "turner_speed"),
        ("SteppingMotor", "TurningMotorFullRotation", "turning_motor_full_rotation"),
        ("SteppingMotor", "TurningMotor1rps", "turning_motor_1rps"),
        ("SteppingMotor", "LiftSpeedSlow", "lift_speed_slow"),
        ("SteppingMotor", "LiftSpeedNormal", "lift_speed_normal"),
        ("SteppingMotor", "LiftSpeedFast", "lift_speed_fast"),
        ("SteppingMotor", "LiftAcceleration", "lift_acceleration"),
        ("SteppingMotor", "MeasPos", "meas_pos"),
        ("SteppingMotor", "SampleBottom", "sample_bottom"),
        ("SteppingMotor", "UpDownTorqueFactor", "updown_torque_factor"),
        ("SteppingMotor", "UpDownMaxTorque", "updown_max_torque"),
        ("SteppingMotor", "PickupTorqueThrottle", "pickup_torque_throttle"),
    ):
        raw = _value(parser, section, key)
        if raw is not None:
            value = _parse_float(raw, default=None, section_key=section, field=key, warnings=warnings)
            if value is not None and math.isfinite(value):
                station.controller[destination] = value
                mark_mapped(section, key, f"motor_station.controller.{destination}")
    if station.ports or station.addresses or station.controller:
        station.calibration_source = str(ini_path.resolve())
        station.controller["sample_height"] = config.motion.sample_height

    raw = _value(parser, "SampleChanger", "HoleSlotNum")
    if raw is not None:
        station.hole_slot = _parse_int(
            raw,
            default=station.hole_slot,
            section_key="SampleChanger",
            field="HoleSlotNum",
            warnings=warnings,
        )
        mark_mapped("SampleChanger", "HoleSlotNum", "motor_station.hole_slot")

    # XY positions are controller counts, not chain spacing or UI millimetres.
    # A partial reimport must not reuse coordinates from an earlier station.
    if parser.has_section('XYTable'):
        station.xy_positions = {}
        station.xy_home = []
        station.use_xy_table = None
        raw = _value(parser, 'XYTable', 'UseXYTableAPS')
        if raw is not None:
            if raw.lower() in {'true', '1', 'yes', 'y', 'on'}:
                station.use_xy_table = True
            elif raw.lower() in {'false', '0', 'no', 'n', 'off'}:
                station.use_xy_table = False
            else:
                warnings.append('XYTable.UseXYTableAPS: invalid table mode; station mode remains unaccepted')
            mark_mapped('XYTable', 'UseXYTableAPS', 'motor_station.use_xy_table')

        def xy_pair(x_key, y_key, destination='motor_station.xy_positions'):
            raw_values = [_value(parser, 'XYTable', key) for key in (x_key, y_key)]
            if raw_values == [None, None]:
                return None
            for key, raw_value in zip((x_key, y_key), raw_values):
                if raw_value is not None:
                    mark_mapped('XYTable', key, destination)
            try:
                values = [float(value) for value in raw_values]
                if any(not math.isfinite(value) or not value.is_integer()
                       or not -(2**31) <= value < 2**31 for value in values):
                    raise ValueError('invalid controller count')
                return [int(value) for value in values]
            except (TypeError, ValueError):
                warnings.append(f'XYTable.{x_key}/{y_key}: complete signed 32-bit integer coordinate pair required; omitted')
                return None

        home = xy_pair('XYHomeX', 'XYHomeY', 'motor_station.xy_home')
        if home is not None:
            station.xy_home = home
        slots = {int(match.group(1)) for key, _ in parser.items('XYTable')
                 if (match := re.fullmatch(r'XY([1-9][0-9]*)[XY]', key))}
        for slot in sorted(slots):
            coordinates = xy_pair(f'XY{slot}X', f'XY{slot}Y')
            if coordinates is not None:
                station.xy_positions[str(slot)] = coordinates
        station.calibration_source = str(ini_path.resolve())

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
