from __future__ import annotations

from dataclasses import dataclass, field
import re
import time
from typing import Callable, Iterable, Sequence

import serial
from serial.tools import list_ports

from rapidpy_common.gaussmeter import probe_gaussmeter_port
from rapidpy_common.hardware import HardwareError, MotorSerialClient


FLOAT_RE = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
COM_PORT_RE = re.compile(r"COM(\d+)$", re.IGNORECASE)
PCI_KEYWORDS = (
    "enhanced",
    "pci",
    "pcie",
    "communications port",
    "serial port",
    "oxford",
    "sunix",
    "moschip",
)

LEGACY_HINTS = {
    3: [("Vacuum", "9600,N,8,1")],
    4: [("Up/Down motor", "57600,N,8,2")],
    5: [("Turning motor", "57600,N,8,2")],
    6: [("X / changer motor", "57600,N,8,2")],
    7: [("Y motor", "57600,N,8,2")],
    8: [("Susceptibility", "9600,N,8,2")],
    9: [("AF", "1200,N,8,1")],
    10: [("SQUID", "1200,N,8,1")],
}


@dataclass(slots=True)
class PortProbeResult:
    port: str
    description: str
    manufacturer: str
    hwid: str
    location: str
    adapter_family: str
    legacy_hints: list[str] = field(default_factory=list)
    detected_device: str = "Unidentified"
    confidence: str = "Low"
    protocol: str = ""
    notes: str = ""
    raw_response: str = ""


@dataclass(slots=True)
class ProtocolMatch:
    device: str
    confidence: str
    protocol: str
    notes: str
    raw_response: str


def list_serial_ports() -> list[list_ports.ListPortInfo]:
    ports = list(list_ports.comports())
    return sorted(ports, key=lambda port: _port_sort_key(port.device))


def sweep_ports(
    *,
    enhanced_only: bool,
    progress: Callable[[int, int, str], None] | None = None,
    stop_requested: Callable[[], bool] | None = None,
) -> list[PortProbeResult]:
    ports = list_serial_ports()
    total = len(ports)
    results: list[PortProbeResult] = []
    for index, info in enumerate(ports, start=1):
        if stop_requested is not None and stop_requested():
            break
        if progress is not None:
            progress(index, total, info.device)
        results.append(probe_port(info, enhanced_only=enhanced_only))
    return results


def probe_port(info: list_ports.ListPortInfo, *, enhanced_only: bool) -> PortProbeResult:
    result = PortProbeResult(
        port=info.device,
        description=(info.description or "").strip(),
        manufacturer=(info.manufacturer or "").strip(),
        hwid=(info.hwid or "").strip(),
        location=(getattr(info, "location", "") or "").strip(),
        adapter_family=classify_adapter(info),
        legacy_hints=_legacy_hints_for_port(info.device),
    )

    if enhanced_only and not is_enhanced_adapter(info):
        result.notes = "Skipped active probe because the adapter description does not look like a PCI/enhanced serial port."
        return result

    motor_match = _probe_motor_controller(info.device)
    if motor_match is not None:
        return _apply_match(result, motor_match)

    squid_match = _probe_squid(info.device)
    if squid_match is not None:
        return _apply_match(result, squid_match)

    gaussmeter_match = _probe_gaussmeter(info.device)
    if gaussmeter_match is not None:
        return _apply_match(result, gaussmeter_match)

    if result.legacy_hints:
        result.confidence = "Hint"
        result.notes = "No high-confidence protocol match. Compare the cable path against the legacy VB6 role hints."
    else:
        result.notes = "No RAPID protocol match found during the safe probe sweep."
    return result


def classify_adapter(info: list_ports.ListPortInfo) -> str:
    haystack = " ".join(
        filter(
            None,
            (
                info.device,
                info.description,
                getattr(info, "manufacturer", ""),
                info.hwid,
            ),
        )
    ).lower()
    if any(keyword in haystack for keyword in PCI_KEYWORDS):
        return "Enhanced / PCI serial"
    return "Other serial"


def is_enhanced_adapter(info: list_ports.ListPortInfo) -> bool:
    return classify_adapter(info) == "Enhanced / PCI serial"


def _apply_match(result: PortProbeResult, match: ProtocolMatch) -> PortProbeResult:
    result.detected_device = match.device
    result.confidence = match.confidence
    result.protocol = match.protocol
    result.notes = match.notes
    result.raw_response = match.raw_response
    return result


def _legacy_hints_for_port(port_name: str) -> list[str]:
    match = COM_PORT_RE.search(port_name)
    if not match:
        return []
    com_number = int(match.group(1))
    hints = LEGACY_HINTS.get(com_number, [])
    return [f"{role} ({settings})" for role, settings in hints]


def _port_sort_key(port_name: str) -> tuple[int, str]:
    match = COM_PORT_RE.search(port_name)
    if match:
        return (0, f"{int(match.group(1)):05d}")
    return (1, port_name.upper())


def _probe_motor_controller(port: str) -> ProtocolMatch | None:
    try:
        with serial.Serial(
            port=port,
            baudrate=57600,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_TWO,
            timeout=0.20,
            write_timeout=0.20,
        ) as ser:
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            _write_line(ser, "@255 173 416")
            time.sleep(0.03)
            ser.reset_input_buffer()

            responders: list[int] = []
            raw: list[str] = []
            for address in (1, 2, 3, 4):
                response = _query_line(ser, f"@{address} 20")
                if not response:
                    continue
                try:
                    MotorSerialClient._parse_status_word(response)
                except HardwareError:
                    continue
                responders.append(address)
                raw.append(f"@{address} 20 -> {response}")

            if not responders:
                return None

            return ProtocolMatch(
                device="Quicksilver motor controller",
                confidence="High",
                protocol="57600,N,8,2",
                notes=f"Motor protocol responded on address(es): {', '.join(str(address) for address in responders)}.",
                raw_response=" | ".join(raw),
            )
    except (OSError, serial.SerialException) as exc:
        if _is_busy_port_error(exc):
            return ProtocolMatch(
                device="Port busy",
                confidence="Blocked",
                protocol="",
                notes="The port is already open in another process, so the mapper could not probe it.",
                raw_response=str(exc),
            )
        return None


def _probe_squid(port: str) -> ProtocolMatch | None:
    try:
        with serial.Serial(
            port=port,
            baudrate=1200,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=0.25,
            write_timeout=0.25,
        ) as ser:
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            _write_squid_command(ser, "ALC")
            time.sleep(0.08)
            _write_squid_command(ser, "ALD")
            time.sleep(0.12)
            count_response = _query_squid(ser, "XSC")
            data_response = _query_squid(ser, "XSD")
            count = _extract_float(count_response)
            data = _extract_float(data_response)
            if count is None or data is None:
                return None

            return ProtocolMatch(
                device="SQUID magnetometer",
                confidence="High",
                protocol="1200,N,8,1",
                notes="Returned valid numeric SQUID count/data values for the X axis.",
                raw_response=f"XSC -> {count_response} | XSD -> {data_response}",
            )
    except (OSError, serial.SerialException):
        return None


def _probe_gaussmeter(port: str) -> ProtocolMatch | None:
    reading = probe_gaussmeter_port(
        port,
        connect_timeout_s=2.5,
        sample_timeout_s=1.0,
    )
    if reading is None:
        return None

    return ProtocolMatch(
        device="908A gaussmeter",
        confidence="High",
        protocol="gm0.dll driver",
        notes=(
            f"The gaussmeter driver connected and returned {reading.value:.6g} "
            f"{reading.units_label} in {reading.mode_label} mode."
        ),
        raw_response=(
            f"value={reading.raw_value}; converted={reading.value}; "
            f"units={reading.units_label}; mode={reading.mode_label}; range={reading.range_index}"
        ),
    )


def _write_line(ser: serial.Serial, command: str) -> None:
    ser.write(f"{command}\r\n".encode("ascii", errors="ignore"))
    ser.flush()


def _query_line(ser: serial.Serial, command: str) -> str:
    ser.reset_input_buffer()
    _write_line(ser, command)
    return ser.read_until(b"\r").decode("ascii", errors="ignore").strip()


def _write_squid_command(ser: serial.Serial, command: str) -> None:
    ser.write(f"\r{command}\r".encode("ascii", errors="ignore"))
    ser.flush()


def _query_squid(ser: serial.Serial, command: str, timeout_s: float = 0.35) -> str:
    ser.reset_input_buffer()
    _write_squid_command(ser, command)
    deadline = time.monotonic() + timeout_s
    chunks = bytearray()
    while time.monotonic() < deadline:
        byte = ser.read(1)
        if not byte:
            continue
        if byte == b"\r":
            if chunks:
                break
            continue
        chunks.extend(byte)
    return chunks.decode("ascii", errors="ignore").strip()


def _extract_float(response: str) -> float | None:
    match = FLOAT_RE.search(response)
    if match is None:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _is_busy_port_error(exc: BaseException) -> bool:
    message = str(exc).lower()
    busy_markers: Sequence[str] = (
        "permissionerror",
        "access is denied",
        "could not open port",
        "file not found",
        "device or resource busy",
    )
    return any(marker in message for marker in busy_markers)