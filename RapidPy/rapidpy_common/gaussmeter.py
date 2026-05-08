from __future__ import annotations

import ctypes
import os
import re
import subprocess
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence


AUTO_PORT = -1
DEFAULT_MODE = 1
DEFAULT_CONNECT_TIMEOUT_S = 30.0
DEFAULT_SAMPLE_TIMEOUT_S = 1.0

MODE_LABELS = (
    "DC",
    "DC Peak",
    "AC",
    "AC Max",
    "AC Peak",
)

BASE_UNIT_LABELS = (
    "T",
    "G",
    "A/m",
    "Oe",
)

UNIT_RANGE_LABELS = (
    ("T", "mT", "mT", "mT"),
    ("kG", "kG", "G", "G"),
    ("kA/m", "kA/m", "kA/m", "kA/m"),
    ("kOe", "kOe", "Oe", "Oe"),
)

UNIT_RANGE_SCALARS = (
    (1.0, 1000.0, 1000.0, 1000.0),
    (0.001, 0.001, 1.0, 1.0),
    (0.001, 0.001, 0.001, 0.001),
    (0.001, 0.001, 1.0, 1.0),
)

_PORT_PREFIX = "COM"
_FW_BELL_READING_RE = re.compile(r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s*([A-Za-z/]+)\s*$")


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def _is_transient_fwbell_path(path: Path) -> bool:
    try:
        resolved = path.resolve(strict=False)
    except OSError:
        resolved = path
    lowered = str(resolved).lower()
    temp_root = os.environ.get("TEMP", "").lower()
    if temp_root and lowered.startswith(temp_root):
        return True
    return "rapid_fwbell_5100" in lowered


class GaussmeterError(RuntimeError):
    """Base error for Hirst 908A / gm0.dll integration."""


class GaussmeterDriverError(GaussmeterError):
    """Raised when gm0.dll cannot be found or loaded."""


class GaussmeterConnectionError(GaussmeterError):
    """Raised when a gaussmeter connection cannot be established."""


class GaussmeterTimeoutError(GaussmeterError):
    """Raised when the driver does not report a state transition in time."""


class GmTime(ctypes.Structure):
    _fields_ = [
        ("sec", ctypes.c_ubyte),
        ("min", ctypes.c_ubyte),
        ("hour", ctypes.c_ubyte),
        ("day", ctypes.c_ubyte),
        ("month", ctypes.c_ubyte),
        ("year", ctypes.c_ubyte),
    ]


class GmStore(ctypes.Structure):
    _fields_ = [
        ("time", GmTime),
        ("range", ctypes.c_ubyte),
        ("Mode", ctypes.c_ubyte),
        ("Units", ctypes.c_ubyte),
        ("value", ctypes.c_float),
    ]


@dataclass(slots=True)
class GaussmeterReading:
    value: float
    raw_value: float
    units_index: int
    units_label: str
    base_units_label: str
    mode_index: int
    mode_label: str
    range_index: int
    timestamp: datetime | None


def normalize_range_index(range_index: int) -> int:
    if range_index > 3:
        return range_index - 4
    return range_index


def mode_label_for(mode_index: int) -> str:
    if 0 <= mode_index < len(MODE_LABELS):
        return MODE_LABELS[mode_index]
    return f"Mode {mode_index}"


def base_units_label_for(units_index: int) -> str:
    if 0 <= units_index < len(BASE_UNIT_LABELS):
        return BASE_UNIT_LABELS[units_index]
    return f"Units {units_index}"


def units_label_for(units_index: int, range_index: int) -> str:
    normalized_range = normalize_range_index(range_index)
    if 0 <= units_index < len(UNIT_RANGE_LABELS):
        labels = UNIT_RANGE_LABELS[units_index]
        if 0 <= normalized_range < len(labels):
            return labels[normalized_range]
    return base_units_label_for(units_index)


def make_actual_value(store: GmStore) -> float:
    normalized_range = normalize_range_index(int(store.range))
    units_index = int(store.Units)
    if 0 <= units_index < len(UNIT_RANGE_SCALARS):
        scalars = UNIT_RANGE_SCALARS[units_index]
        if 0 <= normalized_range < len(scalars):
            return float(store.value) * float(scalars[normalized_range])
    return float(store.value)


def gm_time_to_datetime(value: GmTime) -> datetime | None:
    if not any((value.year, value.month, value.day, value.hour, value.min, value.sec)):
        return None
    try:
        return datetime(
            year=2000 + int(value.year),
            month=int(value.month),
            day=int(value.day),
            hour=int(value.hour),
            minute=int(value.min),
            second=int(value.sec),
        )
    except ValueError:
        return None


def datetime_to_gm_time(value: datetime | None = None) -> GmTime:
    if value is None:
        value = datetime.now()
    return GmTime(
        sec=value.second,
        min=value.minute,
        hour=value.hour,
        day=value.day,
        month=value.month,
        year=value.year - 2000,
    )


def serial_port_name_to_number(port_name: str | int) -> int:
    if isinstance(port_name, int):
        return port_name
    text = port_name.strip().upper()
    if text.startswith(_PORT_PREFIX):
        text = text[len(_PORT_PREFIX) :]
    return int(text)


def _candidate_dll_paths(extra_paths: Sequence[str | Path] | None = None) -> list[Path]:
    candidates: list[Path] = []

    env_path = os.environ.get("RAPID_GM0_DLL")
    if env_path:
        candidates.append(Path(env_path))

    module_dir = Path(__file__).resolve().parent
    repo_root = module_dir.parent.parent
    candidates.extend(
        [
            module_dir / "gm0.dll",
            module_dir.parent / "gm0.dll",
            repo_root / "lib" / "gm0.dll",
            Path.cwd() / "gm0.dll",
        ]
    )

    if extra_paths:
        for raw_path in extra_paths:
            path = Path(raw_path)
            if path.suffix.lower() == ".dll":
                candidates.append(path)
            else:
                candidates.append(path / "gm0.dll")

    for raw_dir in os.environ.get("PATH", "").split(os.pathsep):
        if raw_dir:
            candidates.append(Path(raw_dir) / "gm0.dll")

    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            key = str(candidate.resolve(strict=False)).lower()
        except OSError:
            key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def find_gm0_dll(extra_paths: Sequence[str | Path] | None = None) -> Path | None:
    for candidate in _candidate_dll_paths(extra_paths):
        if candidate.is_file():
            return candidate
    return None


def gaussmeter_driver_status(extra_paths: Sequence[str | Path] | None = None) -> tuple[bool, str]:
    try:
        driver = load_gaussmeter_driver(extra_paths)
    except GaussmeterDriverError as gm0_exc:
        gm0_message = str(gm0_exc)
    else:
        return True, str(driver.dll_path)

    available, detail = fwbell_driver_status(extra_paths)
    if available:
        return True, detail
    return False, f"{gm0_message}; FW Bell: {detail}"


def _candidate_usb5100_dll_paths(extra_paths: Sequence[str | Path] | None = None) -> list[Path]:
    candidates: list[Path] = []

    for env_name in ("RAPID_USB5100_DLL", "RAPID_FW_BELL_DLL"):
        env_path = os.environ.get(env_name)
        if env_path:
            candidates.append(Path(env_path))

    repo_root = _repo_root()
    candidates.extend(
        [
            repo_root / "lib" / "usb5100.dll",
            repo_root / "tools" / "usb5100.dll",
            Path(r"C:/Program Files (x86)/FW Bell/PC5180/usb5100.dll"),
        ]
    )

    if extra_paths:
        for raw_path in extra_paths:
            path = Path(raw_path)
            if path.suffix.lower() == ".dll":
                candidates.append(path)
                candidates.append(path.parent / "usb5100.dll")
            else:
                candidates.append(path / "usb5100.dll")

    for raw_dir in os.environ.get("PATH", "").split(os.pathsep):
        if raw_dir:
            candidates.append(Path(raw_dir) / "usb5100.dll")

    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            key = str(candidate.resolve(strict=False)).lower()
        except OSError:
            key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def find_usb5100_dll(extra_paths: Sequence[str | Path] | None = None) -> Path | None:
    for candidate in _candidate_usb5100_dll_paths(extra_paths):
        if extra_paths is None and _is_transient_fwbell_path(candidate):
            continue
        if candidate.is_file():
            return candidate
    return None


def _candidate_usb5100_helper_paths() -> list[Path]:
    candidates = []

    env_path = os.environ.get("RAPID_USB5100_HELPER")
    if env_path:
        candidates.append(Path(env_path))

    repo_root = _repo_root()
    candidates.extend(
        [
            repo_root / "tools" / "usb5100_probe.exe",
            Path.cwd() / "usb5100_probe.exe",
        ]
    )

    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            key = str(candidate.resolve(strict=False)).lower()
        except OSError:
            key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def find_usb5100_helper() -> Path | None:
    for candidate in _candidate_usb5100_helper_paths():
        if candidate.is_file():
            return candidate
    return None


def _parse_fwbell_helper_output(output: str) -> dict[str, str]:
    parsed: dict[str, str] = {}
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key in {"status", "dll", "command", "response", "dll_dir", "loaded"}:
            parsed[key] = value.strip()
    return parsed


def _run_fwbell_helper(
    command_name: str,
    command_arg: str | None = None,
    *,
    extra_paths: Sequence[str | Path] | None = None,
    timeout_s: float = 10.0,
) -> dict[str, str]:
    helper_path = find_usb5100_helper()
    if helper_path is None:
        raise GaussmeterDriverError(
            "usb5100_probe.exe was not found. Build the x86 helper under tools to enable FW Bell gaussmeter access."
        )

    args = [str(helper_path)]
    dll_path = find_usb5100_dll(extra_paths)
    if dll_path is not None:
        args.extend(["--dll", str(dll_path)])
    args.append(command_name)
    if command_arg is not None:
        args.append(command_arg)

    try:
        completed = subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
    except FileNotFoundError as exc:
        raise GaussmeterDriverError(f"Unable to launch {helper_path}: {exc}") from exc
    except subprocess.TimeoutExpired as exc:
        raise GaussmeterTimeoutError(f"Timed out waiting for FW Bell helper command '{command_name}'.") from exc

    parsed = _parse_fwbell_helper_output(completed.stdout)
    if completed.returncode != 0:
        detail = completed.stderr.strip() or parsed.get("response") or completed.stdout.strip()
        raise GaussmeterConnectionError(detail or f"FW Bell helper failed with exit code {completed.returncode}.")
    if parsed.get("status") != "ok":
        raise GaussmeterConnectionError("FW Bell helper returned an unexpected response.")
    return parsed


def fwbell_driver_status(extra_paths: Sequence[str | Path] | None = None) -> tuple[bool, str]:
    helper_path = find_usb5100_helper()
    if helper_path is None:
        return False, "usb5100_probe.exe not found"

    dll_path = find_usb5100_dll(extra_paths)
    if dll_path is None:
        return (
            False,
            "usb5100.dll not found. Set RAPID_USB5100_DLL, install the FW Bell runtime, or browse to the vendor DLL.",
        )

    try:
        parsed = _run_fwbell_helper("status", extra_paths=extra_paths)
    except GaussmeterError as exc:
        return False, f"{dll_path}: {exc}"
    detail = parsed.get("dll") or str(dll_path)
    response = parsed.get("response")
    if response:
        detail = f"{detail} [{response}]"
    return True, detail


class _GaussmeterDriver:
    def __init__(self, dll_path: Path) -> None:
        try:
            dll = ctypes.WinDLL(str(dll_path))
        except OSError as exc:
            raise GaussmeterDriverError(f"Unable to load {dll_path}: {exc}") from exc

        self.dll_path = dll_path
        self.gm0_newgm = dll.gm0_newgm
        self.gm0_newgm.argtypes = [ctypes.c_long, ctypes.c_long]
        self.gm0_newgm.restype = ctypes.c_long

        self.gm0_startconnect = dll.gm0_startconnect
        self.gm0_startconnect.argtypes = [ctypes.c_long]
        self.gm0_startconnect.restype = ctypes.c_long

        self.gm0_killgm = dll.gm0_killgm
        self.gm0_killgm.argtypes = [ctypes.c_long]
        self.gm0_killgm.restype = ctypes.c_long

        self.gm0_getconnect = dll.gm0_getconnect
        self.gm0_getconnect.argtypes = [ctypes.c_long]
        self.gm0_getconnect.restype = ctypes.c_int

        self.gm0_setrange = dll.gm0_setrange
        self.gm0_setrange.argtypes = [ctypes.c_long, ctypes.c_ubyte]
        self.gm0_setrange.restype = ctypes.c_long

        self.gm0_setunits = dll.gm0_setunits
        self.gm0_setunits.argtypes = [ctypes.c_long, ctypes.c_ubyte]
        self.gm0_setunits.restype = ctypes.c_long

        self.gm0_setmode = dll.gm0_setmode
        self.gm0_setmode.argtypes = [ctypes.c_long, ctypes.c_ubyte]
        self.gm0_setmode.restype = ctypes.c_long

        self.gm0_isnewdata = dll.gm0_isnewdata
        self.gm0_isnewdata.argtypes = [ctypes.c_long]
        self.gm0_isnewdata.restype = ctypes.c_long

        self.gm0_getrange = dll.gm0_getrange
        self.gm0_getrange.argtypes = [ctypes.c_long]
        self.gm0_getrange.restype = ctypes.c_long

        self.gm0_getunits = dll.gm0_getunits
        self.gm0_getunits.argtypes = [ctypes.c_long]
        self.gm0_getunits.restype = ctypes.c_long

        self.gm0_getmode = dll.gm0_getmode
        self.gm0_getmode.argtypes = [ctypes.c_long]
        self.gm0_getmode.restype = ctypes.c_long

        self.gm0_getvalue = dll.gm0_getvalue
        self.gm0_getvalue.argtypes = [ctypes.c_long]
        self.gm0_getvalue.restype = ctypes.c_double

        self.gm0_donull = dll.gm0_donull
        self.gm0_donull.argtypes = [ctypes.c_long]
        self.gm0_donull.restype = ctypes.c_long

        self.gm0_doaz = dll.gm0_doaz
        self.gm0_doaz.argtypes = [ctypes.c_long]
        self.gm0_doaz.restype = ctypes.c_long

        self.gm0_resetnull = dll.gm0_resetnull
        self.gm0_resetnull.argtypes = [ctypes.c_long]
        self.gm0_resetnull.restype = ctypes.c_long

        self.gm0_resetpeak = dll.gm0_resetpeak
        self.gm0_resetpeak.argtypes = [ctypes.c_long]
        self.gm0_resetpeak.restype = ctypes.c_long

        self.gm0_sendtime = dll.gm0_sendtime
        self.gm0_sendtime.argtypes = [ctypes.c_long, ctypes.c_int]
        self.gm0_sendtime.restype = ctypes.c_long

        self.gm0_settime2 = dll.gm0_settime2
        self.gm0_settime2.argtypes = [ctypes.c_long, GmTime]
        self.gm0_settime2.restype = ctypes.c_long

        self.gm0_gettime = dll.gm0_gettime
        self.gm0_gettime.argtypes = [ctypes.c_long]
        self.gm0_gettime.restype = GmTime

        self.gm0_getstore = dll.gm0_getstore
        self.gm0_getstore.argtypes = [ctypes.c_long, ctypes.c_long]
        self.gm0_getstore.restype = GmStore

        self.gm0_startcmd = dll.gm0_startcmd
        self.gm0_startcmd.argtypes = [ctypes.c_long]
        self.gm0_startcmd.restype = ctypes.c_long

        self.gm0_endcmd = dll.gm0_endcmd
        self.gm0_endcmd.argtypes = [ctypes.c_long]
        self.gm0_endcmd.restype = ctypes.c_long


def load_gaussmeter_driver(extra_paths: Sequence[str | Path] | None = None) -> _GaussmeterDriver:
    dll_path = find_gm0_dll(extra_paths)
    if dll_path is None:
        raise GaussmeterDriverError(
            "gm0.dll was not found. Set RAPID_GM0_DLL or place the DLL on PATH to enable gaussmeter control."
        )
    return _GaussmeterDriver(dll_path)


class _Gm0GaussmeterBackend:
    def __init__(
        self,
        *,
        port: int = AUTO_PORT,
        mode: int = DEFAULT_MODE,
        dll_search_paths: Sequence[str | Path] | None = None,
    ) -> None:
        self.port = port
        self.mode = mode
        self._driver = load_gaussmeter_driver(dll_search_paths)
        self._handle = -1
        self.connected = False

    @property
    def dll_path(self) -> Path:
        return self._driver.dll_path

    @property
    def handle(self) -> int:
        return self._handle

    def connect(self, timeout_s: float = DEFAULT_CONNECT_TIMEOUT_S) -> None:
        if self.connected and self._handle >= 0:
            return

        handle = int(self._driver.gm0_newgm(self.port, self.mode))
        if handle < 0:
            raise GaussmeterConnectionError(
                f"gm0_newgm returned {handle} for port {self.port} in mode {self.mode}."
            )

        self._handle = handle
        self._driver.gm0_startconnect(self._handle)
        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline:
            if bool(self._driver.gm0_getconnect(self._handle)):
                self.connected = True
                return
            time.sleep(0.05)

        self.disconnect()
        raise GaussmeterTimeoutError(
            f"Timed out waiting for gaussmeter connection on port {self.port}."
        )

    def disconnect(self) -> None:
        if self._handle >= 0:
            try:
                self._driver.gm0_killgm(self._handle)
            finally:
                self._handle = -1
                self.connected = False

    def wait_for_data(self, timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> None:
        self._require_connection()
        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline:
            if int(self._driver.gm0_isnewdata(self._handle)):
                return
            time.sleep(0.05)
        raise GaussmeterTimeoutError("Timed out waiting for a new gaussmeter sample.")

    def read_store(self) -> GmStore:
        self._require_connection()
        time_value = self.get_time()
        range_index = normalize_range_index(int(self._driver.gm0_getrange(self._handle)))
        return GmStore(
            time=time_value,
            range=range_index,
            Mode=int(self._driver.gm0_getmode(self._handle)),
            Units=int(self._driver.gm0_getunits(self._handle)),
            value=float(self._driver.gm0_getvalue(self._handle)),
        )

    def read(self, *, wait_for_new_data: bool = False, sample_timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> GaussmeterReading:
        if wait_for_new_data:
            self.wait_for_data(sample_timeout_s)

        store = self.read_store()
        return GaussmeterReading(
            value=make_actual_value(store),
            raw_value=float(store.value),
            units_index=int(store.Units),
            units_label=units_label_for(int(store.Units), int(store.range)),
            base_units_label=base_units_label_for(int(store.Units)),
            mode_index=int(store.Mode),
            mode_label=mode_label_for(int(store.Mode)),
            range_index=int(store.range),
            timestamp=gm_time_to_datetime(store.time),
        )

    def set_range(self, range_index: int) -> None:
        self._require_connection()
        self._driver.gm0_setrange(self._handle, int(range_index))

    def set_units(self, units_index: int) -> None:
        self._require_connection()
        self._driver.gm0_setunits(self._handle, int(units_index))

    def set_mode(self, mode_index: int) -> None:
        self._require_connection()
        self._driver.gm0_setmode(self._handle, int(mode_index))

    def null(self) -> None:
        self._require_connection()
        self._driver.gm0_donull(self._handle)

    def auto_zero(self) -> None:
        self._require_connection()
        self._driver.gm0_doaz(self._handle)

    def reset_null(self) -> None:
        self._require_connection()
        self._driver.gm0_resetnull(self._handle)

    def reset_peak(self) -> None:
        self._require_connection()
        self._driver.gm0_resetpeak(self._handle)

    def enable_time_stream(self, enabled: bool) -> None:
        self._require_connection()
        self._driver.gm0_sendtime(self._handle, int(enabled))

    def get_time(self) -> GmTime:
        self._require_connection()
        return self._driver.gm0_gettime(self._handle)

    def set_time(self, value: datetime | None = None) -> None:
        self._require_connection()
        self._driver.gm0_settime2(self._handle, datetime_to_gm_time(value))

    def set_system_time(self) -> None:
        self.set_time(datetime.now())

    def start_command_mode(self) -> None:
        self._require_connection()
        self._driver.gm0_startcmd(self._handle)

    def end_command_mode(self) -> None:
        self._require_connection()
        self._driver.gm0_endcmd(self._handle)

    def _require_connection(self) -> None:
        if not self.connected or self._handle < 0:
            raise GaussmeterConnectionError("The gaussmeter is not connected.")

    def __enter__(self) -> GaussmeterClient:
        self.connect()
        return self

    def __exit__(self, exc_type, exc, exc_tb) -> None:
        self.disconnect()


def _normalize_fwbell_mode_index(mode_index: int) -> int:
    if mode_index in (1,):
        return 0
    if mode_index in (3, 4):
        return 2
    if mode_index in (0, 2):
        return mode_index
    raise GaussmeterError(f"FW Bell backend does not support mode index {mode_index}.")


def _fwbell_unit_token(mode_index: int, units_index: int) -> str:
    normalized_mode = _normalize_fwbell_mode_index(mode_index)
    if units_index == 0:
        units_token = "TESLA"
    elif units_index == 1:
        units_token = "GAUSS"
    elif units_index == 2:
        units_token = "AM"
    else:
        raise GaussmeterError(f"FW Bell backend does not support units index {units_index}.")

    mode_token = "DC" if normalized_mode == 0 else "AC"
    return f":UNIT:FLUX:{mode_token}:{units_token}"


def _parse_fwbell_reading(response: str) -> tuple[float, int, str, str]:
    match = _FW_BELL_READING_RE.match(response)
    if match is None:
        raise GaussmeterError(f"Unexpected FW Bell reading: {response!r}")

    value = float(match.group(1))
    units_label = match.group(2)
    normalized = units_label.upper()
    if normalized in {"T", "MT"}:
        return value, 0, units_label, "T"
    if normalized in {"G", "KG"}:
        return value, 1, units_label, "G"
    if normalized in {"A/M", "KA/M", "AM", "KAM"}:
        return value, 2, units_label, "A/m"
    if normalized in {"OE", "KOE"}:
        return value, 3, units_label, "Oe"
    return value, -1, units_label, units_label


class _FwBellGaussmeterBackend:
    def __init__(
        self,
        *,
        port: int = AUTO_PORT,
        mode: int = DEFAULT_MODE,
        dll_search_paths: Sequence[str | Path] | None = None,
    ) -> None:
        if port != AUTO_PORT:
            raise GaussmeterConnectionError("FW Bell USB access does not support manual COM port selection.")

        self.port = port
        self.mode = mode
        self._dll_search_paths = dll_search_paths
        self._helper_path = find_usb5100_helper()
        if self._helper_path is None:
            raise GaussmeterDriverError(
                "usb5100_probe.exe was not found. Build the x86 helper under tools to enable FW Bell gaussmeter access."
            )

        self._dll_path = find_usb5100_dll(dll_search_paths) or self._helper_path
        self._mode_index = 0
        self._units_index = 0
        self._range_index = 4
        self._handle = -1
        self.connected = False

    @property
    def dll_path(self) -> Path:
        return self._dll_path

    @property
    def handle(self) -> int:
        return self._handle

    def connect(self, timeout_s: float = DEFAULT_CONNECT_TIMEOUT_S) -> None:
        parsed = _run_fwbell_helper("status", extra_paths=self._dll_search_paths, timeout_s=max(timeout_s, 5.0))
        dll_text = parsed.get("dll")
        if dll_text:
            self._dll_path = Path(dll_text)
        self._handle = 1
        self.connected = True

    def disconnect(self) -> None:
        self._handle = -1
        self.connected = False

    def wait_for_data(self, timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> None:
        self._require_connection()

    def read(self, *, wait_for_new_data: bool = False, sample_timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> GaussmeterReading:
        self._require_connection()
        parsed = _run_fwbell_helper(
            "read",
            extra_paths=self._dll_search_paths,
            timeout_s=max(sample_timeout_s, 5.0),
        )
        response = parsed.get("response", "")
        value, units_index, units_label, base_units_label = _parse_fwbell_reading(response)
        mode_index = _normalize_fwbell_mode_index(self._mode_index)
        return GaussmeterReading(
            value=value,
            raw_value=value,
            units_index=units_index,
            units_label=units_label,
            base_units_label=base_units_label,
            mode_index=mode_index,
            mode_label=mode_label_for(mode_index),
            range_index=self._range_index,
            timestamp=None,
        )

    def set_range(self, range_index: int) -> None:
        self._require_connection()
        if range_index == 4:
            command = ":SENSE:FLUX:RANGE:AUTO"
        elif range_index in (0, 1, 2):
            command = f":SENSE:FLUX:RANGE {range_index}"
        else:
            raise GaussmeterError(f"FW Bell backend does not support range index {range_index}.")
        self._run_command(command)
        self._range_index = range_index

    def set_units(self, units_index: int) -> None:
        self._require_connection()
        self._run_command(_fwbell_unit_token(self._mode_index, units_index))
        self._units_index = units_index
        self._mode_index = _normalize_fwbell_mode_index(self._mode_index)

    def set_mode(self, mode_index: int) -> None:
        self._require_connection()
        normalized_mode = _normalize_fwbell_mode_index(mode_index)
        self._run_command(_fwbell_unit_token(normalized_mode, self._units_index))
        self._mode_index = normalized_mode

    def null(self) -> None:
        self.auto_zero()

    def auto_zero(self) -> None:
        self._require_connection()
        self._run_command(":SYSTEM:AZERO", timeout_s=10.0)
        time.sleep(6.0)

    def reset_null(self) -> None:
        raise GaussmeterError("FW Bell backend does not expose a reset-null command.")

    def reset_peak(self) -> None:
        raise GaussmeterError("FW Bell backend does not expose peak tracking commands.")

    def enable_time_stream(self, enabled: bool) -> None:
        if enabled:
            raise GaussmeterError("FW Bell backend does not support time streaming.")

    def get_time(self) -> GmTime:
        raise GaussmeterError("FW Bell backend does not expose instrument time.")

    def set_time(self, value: datetime | None = None) -> None:
        raise GaussmeterError("FW Bell backend does not expose instrument time.")

    def set_system_time(self) -> None:
        raise GaussmeterError("FW Bell backend does not expose instrument time.")

    def start_command_mode(self) -> None:
        self._require_connection()

    def end_command_mode(self) -> None:
        self._require_connection()

    def _run_command(self, command: str, *, timeout_s: float = 5.0) -> str:
        parsed = _run_fwbell_helper(
            "command",
            command,
            extra_paths=self._dll_search_paths,
            timeout_s=timeout_s,
        )
        dll_text = parsed.get("dll")
        if dll_text:
            self._dll_path = Path(dll_text)
        return parsed.get("response", "")

    def _require_connection(self) -> None:
        if not self.connected or self._handle < 0:
            raise GaussmeterConnectionError("The gaussmeter is not connected.")


class GaussmeterClient:
    def __init__(
        self,
        *,
        port: int = AUTO_PORT,
        mode: int = DEFAULT_MODE,
        dll_search_paths: Sequence[str | Path] | None = None,
    ) -> None:
        self.port = port
        self.mode = mode
        self._backend = self._create_backend(port=port, mode=mode, dll_search_paths=dll_search_paths)

    @staticmethod
    def _create_backend(
        *,
        port: int,
        mode: int,
        dll_search_paths: Sequence[str | Path] | None,
    ) -> _Gm0GaussmeterBackend | _FwBellGaussmeterBackend:
        if port != AUTO_PORT:
            return _Gm0GaussmeterBackend(port=port, mode=mode, dll_search_paths=dll_search_paths)

        try:
            return _Gm0GaussmeterBackend(port=port, mode=mode, dll_search_paths=dll_search_paths)
        except GaussmeterDriverError:
            return _FwBellGaussmeterBackend(port=port, mode=mode, dll_search_paths=dll_search_paths)

    @property
    def dll_path(self) -> Path:
        return self._backend.dll_path

    @property
    def handle(self) -> int:
        return self._backend.handle

    @property
    def connected(self) -> bool:
        return self._backend.connected

    def connect(self, timeout_s: float = DEFAULT_CONNECT_TIMEOUT_S) -> None:
        self._backend.connect(timeout_s=timeout_s)

    def disconnect(self) -> None:
        self._backend.disconnect()

    def wait_for_data(self, timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> None:
        self._backend.wait_for_data(timeout_s=timeout_s)

    def read(self, *, wait_for_new_data: bool = False, sample_timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S) -> GaussmeterReading:
        return self._backend.read(wait_for_new_data=wait_for_new_data, sample_timeout_s=sample_timeout_s)

    def set_range(self, range_index: int) -> None:
        self._backend.set_range(range_index)

    def set_units(self, units_index: int) -> None:
        self._backend.set_units(units_index)

    def set_mode(self, mode_index: int) -> None:
        self._backend.set_mode(mode_index)

    def null(self) -> None:
        self._backend.null()

    def auto_zero(self) -> None:
        self._backend.auto_zero()

    def reset_null(self) -> None:
        self._backend.reset_null()

    def reset_peak(self) -> None:
        self._backend.reset_peak()

    def enable_time_stream(self, enabled: bool) -> None:
        self._backend.enable_time_stream(enabled)

    def get_time(self) -> GmTime:
        return self._backend.get_time()

    def set_time(self, value: datetime | None = None) -> None:
        self._backend.set_time(value)

    def set_system_time(self) -> None:
        self._backend.set_system_time()

    def start_command_mode(self) -> None:
        self._backend.start_command_mode()

    def end_command_mode(self) -> None:
        self._backend.end_command_mode()

    def __enter__(self) -> GaussmeterClient:
        self.connect()
        return self

    def __exit__(self, exc_type, exc, exc_tb) -> None:
        self.disconnect()


def probe_gaussmeter_port(
    port: str | int,
    *,
    mode: int = DEFAULT_MODE,
    connect_timeout_s: float = 3.0,
    sample_timeout_s: float = DEFAULT_SAMPLE_TIMEOUT_S,
    dll_search_paths: Sequence[str | Path] | None = None,
) -> GaussmeterReading | None:
    try:
        port_number = serial_port_name_to_number(port)
    except ValueError:
        return None

    try:
        client = GaussmeterClient(port=port_number, mode=mode, dll_search_paths=dll_search_paths)
    except GaussmeterDriverError:
        return None

    try:
        client.connect(timeout_s=connect_timeout_s)
        return client.read(wait_for_new_data=True, sample_timeout_s=sample_timeout_s)
    except GaussmeterError:
        return None
    finally:
        client.disconnect()