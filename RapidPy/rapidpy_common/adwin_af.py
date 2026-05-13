from __future__ import annotations

import ctypes
import os
import platform
import struct
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional


def _find_adwin_btl_folder() -> str:
    """Try to locate the ADwin BTL folder from the Windows registry.

    Returns the folder path string, or "" if not found or not on Windows.
    """
    if platform.system() != "Windows":
        return ""
    try:
        import winreg
        # Jäger's installer writes the install directory here (32- and 64-bit paths)
        for hive_flag in (winreg.KEY_READ | winreg.KEY_WOW64_32KEY,
                          winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
                          winreg.KEY_READ):
            for subkey in (
                r"SOFTWARE\Jäger Meßtechnik GmbH\ADwin\Directory",
                r"SOFTWARE\ADwin\Directory",
            ):
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey,
                                       access=hive_flag) as key:
                        btl_dir, _ = winreg.QueryValueEx(key, "BTL")
                        if btl_dir and Path(btl_dir).is_dir():
                            return str(btl_dir)
                        # Try the install root directly
                        install_dir, _ = winreg.QueryValueEx(key, "")
                        candidate = Path(install_dir) / "BTL"
                        if candidate.is_dir():
                            return str(candidate)
                except (FileNotFoundError, OSError, ValueError):
                    continue
    except Exception:  # pragma: no cover
        pass
    # Last-resort hard-coded paths used by the default installer
    for guess in (r"C:\ADwin\BTL", r"C:\ADwin9\BTL", r"C:\ADwin"):
        if Path(guess).is_dir():
            return guess
    return ""


def _find_adwin_dll() -> str:
    """Return the full path to the correct ADwin DLL for this Python bitness.

    Searches (in order):
      1. C:\\Windows\\               — standard installer location
      2. C:\\Windows\\System32\\     — some installer variants
      3. C:\\Windows\\SysWOW64\\     — 32-bit variant on 64-bit Windows
      4. ADwin install root          — from registry
      5. System PATH                 — via shutil.which

    Returns the first path that exists, or "" if none found.
    """
    import shutil
    dll_name = "adwin32.dll" if struct.calcsize("P") == 4 else "adwin64.dll"
    candidates: list[Path] = [
        Path(r"C:\Windows") / dll_name,
        Path(r"C:\Windows\System32") / dll_name,
        Path(r"C:\Windows\SysWOW64") / dll_name,
    ]
    # Also look relative to the BTL folder (install root / parent)
    btl = _find_adwin_btl_folder()
    if btl:
        for rel in ("", "Bin", ".."):
            candidates.append(Path(btl) / rel / dll_name)
    # PATH
    in_path = shutil.which(dll_name)
    if in_path:
        candidates.append(Path(in_path))
    for c in candidates:
        try:
            if c.resolve().is_file():
                return str(c.resolve())
        except Exception:
            continue
    return ""


def find_btl_files(extra_folder: str = "") -> list[str]:
    """Return all .btl firmware files found in known ADwin directories.

    Searches the registry-detected BTL folder, common hard-coded paths,
    and *extra_folder* if given.  Returns a list of absolute path strings
    sorted so the most-common models come first:
      ADwin9.btl, ADwin11.btl, ADwin12.btl, then rest alphabetically.
    """
    # Gather candidate directories
    dirs: list[Path] = []
    auto = _find_adwin_btl_folder()
    if auto:
        dirs.append(Path(auto))
    for hc in (r"C:\ADwin\BTL", r"C:\ADwin9\BTL", r"C:\ADwin", r"C:\ADwin9"):
        p = Path(hc)
        if p.is_dir() and p not in dirs:
            dirs.append(p)
    if extra_folder:
        p = Path(extra_folder)
        if p.is_dir() and p not in dirs:
            dirs.append(p)

    seen: set[str] = set()
    results: list[str] = []
    for d in dirs:
        try:
            for f in sorted(d.glob("*.btl"), key=lambda x: x.name.lower()):
                key = f.name.lower()
                if key not in seen:
                    seen.add(key)
                    results.append(str(f))
        except OSError:
            continue

    # Sort: priority models first, then alphabetical
    _priority = ["adwin9.btl", "adwin11.btl", "adwin12.btl", "adwin10.btl"]

    def _sort_key(path: str) -> tuple[int, str]:
        name = Path(path).name.lower()
        try:
            return (_priority.index(name), name)
        except ValueError:
            return (len(_priority), name)

    results.sort(key=_sort_key)
    return results


class AdwinError(RuntimeError):
    """Raised when ADWIN board operations fail."""


@dataclass(slots=True)
class AdwinBoardConfig:
    board_num: int = 1
    bin_folder: str = ""
    boot_file: str = "ADwin9.btl"
    process_file: str = "sineout.T91"
    ramp_dac_chan: int = 1
    monitor_adc_chan: int = 1
    process_num: int = 1
    axial_relay_bit: int = 0
    trans_relay_bit: int = 1


@dataclass(slots=True)
class AdwinCoilLimits:
    axial_ramp_max: float = 10.0
    axial_monitor_max: float = 10.0
    trans_ramp_max: float = 10.0
    trans_monitor_max: float = 10.0


@dataclass(slots=True)
class AdwinRampRequest:
    slope_up: float
    slope_down: float
    peak_monitor_voltage: float
    sine_freq_hz: float
    ramp_peak_voltage: float
    active_coil: str  # "axial" | "transverse"
    ramp_mode: int = 3
    hold_ms: int = 0
    ramp_down_mode: int = 1
    io_rate_hz: float = 1000.0
    noise_level: int = 5


@dataclass(slots=True)
class AdwinRampResult:
    out_count: int
    in_count: int
    up_count: int
    down_start: int
    monitor_peak_v: float
    ramp_peak_v: float
    down_slope_vps: float
    timestep_s: float
    points_per_period: float


@dataclass(slots=True)
class AdwinDenseCaptureRequest:
    sine_freq_hz: float
    amplitude_v: float
    io_rate_hz: float
    duration_s: float
    dac_chan: int = 1
    adc_chan: int = 1
    noise_level: int = 5
    ramp_mode: int = 3  # CLIPTEST in the legacy ADbasic process
    ramp_up_slope_vps: float = 200.0
    ramp_down_slope_vps: float = 200.0
    ramp_down_periods: int = 2


@dataclass(slots=True)
class AdwinDenseCaptureResult:
    time_s: list[float]
    dac_v: list[float]
    adc_v: list[float]
    out_count: int
    in_count: int
    up_count: int
    down_start: int
    timestep_s: float
    points_per_period: float
    steady_start_idx: int
    steady_stop_idx: int


class _AdwinDll:
    def __init__(self) -> None:
        if platform.system() != "Windows":
            raise AdwinError("ADWIN backend requires Windows (adwin32.dll / adwin64.dll).")
        dll_name = "adwin32.dll" if struct.calcsize("P") == 4 else "adwin64.dll"
        dll_path_str = _find_adwin_dll()
        if not dll_path_str:
            raise AdwinError(
                f"{dll_name} not found.\n"
                f"Searched: C:\\Windows\\, C:\\Windows\\System32\\, ADwin install dir, PATH.\n"
                "Install the ADwin runtime from https://www.adwin.de "
                "or copy the DLL to C:\\Windows\\."
            )
        try:
            self._dll = ctypes.WinDLL(dll_path_str)
            self._dll_path = dll_path_str
        except OSError as exc:
            raise AdwinError(
                f"Unable to load {dll_path_str}.\n"
                f"OS error: {exc}\n"
                "Try running this application as Administrator, "
                "or reinstall ADwin."
            ) from exc
        self._bind()

    def _bind(self) -> None:
        self.ADboot = self._dll.ADboot
        self.ADboot.argtypes = [ctypes.c_char_p, ctypes.c_int, ctypes.c_long, ctypes.c_int]
        self.ADboot.restype = ctypes.c_long

        self.ADBload = self._dll.ADBload
        self.ADBload.argtypes = [ctypes.c_char_p, ctypes.c_int, ctypes.c_int]
        self.ADBload.restype = ctypes.c_int

        self.ADTest_Version = self._dll.ADTest_Version
        self.ADTest_Version.argtypes = [ctypes.c_int, ctypes.c_int]
        self.ADTest_Version.restype = ctypes.c_int

        self.ADB_Start = self._dll.ADB_Start
        self.ADB_Start.argtypes = [ctypes.c_int, ctypes.c_int]
        self.ADB_Start.restype = ctypes.c_int

        self.ADB_Stop = self._dll.ADB_Stop
        self.ADB_Stop.argtypes = [ctypes.c_int, ctypes.c_int]
        self.ADB_Stop.restype = ctypes.c_int

        self.Clear_Process = self._dll.Clear_Process
        self.Clear_Process.argtypes = [ctypes.c_long, ctypes.c_int]
        self.Clear_Process.restype = ctypes.c_int

        self.Set_ADBPar = self._dll.Set_ADBPar
        self.Set_ADBPar.argtypes = [ctypes.c_int, ctypes.c_long, ctypes.c_int]
        self.Set_ADBPar.restype = ctypes.c_int

        self.Set_ADBFPar = self._dll.Set_ADBFPar
        self.Set_ADBFPar.argtypes = [ctypes.c_int, ctypes.c_float, ctypes.c_int]
        self.Set_ADBFPar.restype = ctypes.c_int

        self.Get_ADBPar = self._dll.Get_ADBPar
        self.Get_ADBPar.argtypes = [ctypes.c_int, ctypes.c_int]
        self.Get_ADBPar.restype = ctypes.c_long

        self.Get_ADBFPar = self._dll.Get_ADBFPar
        self.Get_ADBFPar.argtypes = [ctypes.c_int, ctypes.c_int]
        self.Get_ADBFPar.restype = ctypes.c_float

        self.Set_Digout = self._dll.Set_Digout
        # VB6 source-of-truth declaration in Adwin.bas is:
        #   Set_Digout(value As Long, DeviceNo As Integer)
        self.Set_Digout.argtypes = [ctypes.c_long, ctypes.c_int]
        self.Set_Digout.restype = ctypes.c_int

        self.Get_Digout = self._dll.Get_Digout
        self.Get_Digout.argtypes = [ctypes.c_int]
        self.Get_Digout.restype = ctypes.c_long

        self.Get_ADC = self._dll.Get_ADC
        self.Get_ADC.argtypes = [ctypes.c_int, ctypes.c_int]
        self.Get_ADC.restype = ctypes.c_long

        self.Set_DAC = self._dll.Set_DAC
        self.Set_DAC.argtypes = [ctypes.c_int, ctypes.c_long, ctypes.c_int]
        self.Set_DAC.restype = ctypes.c_int

        self.Get_Data = self._dll.Get_Data
        self.Get_Data.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_long, ctypes.c_long, ctypes.c_int]
        self.Get_Data.restype = ctypes.c_int

        self.ADGetErrorCode = self._dll.ADGetErrorCode
        self.ADGetErrorCode.argtypes = []
        self.ADGetErrorCode.restype = ctypes.c_long

        self.ADGetErrorText = self._dll.ADGetErrorText
        self.ADGetErrorText.argtypes = [ctypes.c_long, ctypes.c_char_p, ctypes.c_long]
        self.ADGetErrorText.restype = ctypes.c_long


class AdwinAFController:
    # ±10 V range, 16-bit (0..65535, 32768 = 0 V)
    _COUNTS_PER_VOLT: float = 32768.0 / 10.0

    def __init__(self, board: Optional[AdwinBoardConfig] = None, limits: Optional[AdwinCoilLimits] = None) -> None:
        self.board = board or AdwinBoardConfig()
        self.limits = limits or AdwinCoilLimits()
        self._dll = _AdwinDll()
        self._last_digout_bit: int = -1

    @property
    def _dev(self) -> int:
        return int(self.board.board_num)

    def _resolve_file(self, filename: str) -> str:
        """Resolve *filename* relative to bin_folder (or auto-detected ADwin BTL folder)."""
        if self.board.bin_folder:
            root = Path(self.board.bin_folder).expanduser()
        else:
            auto = _find_adwin_btl_folder()
            root = Path(auto) if auto else Path.cwd()
        path = root / filename
        return str(path)

    def _last_error(self) -> tuple[int, str]:
        code = int(self._dll.ADGetErrorCode())
        if code == 0:
            return 0, ""
        buf = ctypes.create_string_buffer(512)
        self._dll.ADGetErrorText(code, buf, len(buf))
        text = buf.value.decode(errors="replace").strip()
        return code, text

    def _raise_if_error(self, action: str, raw_return: object | None = None) -> None:
        code, text = self._last_error()
        if code == 0:
            return
        detail = f" ADwin error {code}: {text}" if text else f" ADwin error {code}."
        if raw_return is not None:
            detail += f" Raw return={raw_return}."
        raise AdwinError(f"{action} failed.{detail}")

    def boot_board(self) -> None:
        """Boot the ADwin board using the configured BTL file."""
        boot_path = self._resolve_file(self.board.boot_file)
        if not Path(boot_path).is_file():
            raise AdwinError(
                f"BTL boot file not found: {boot_path}\n"
                "Set the 'bin folder' to the directory containing your .btl file "
                "(typically C:\\ADwin\\BTL\\)."
            )
        ret = int(self._dll.ADboot(os.fsencode(boot_path), self._dev, 0, 1))
        if ret <= 0:
            raise AdwinError(f"ADWIN boot failed (return code {ret}) using {boot_path}.")
        self._raise_if_error(f"ADboot(dev={self._dev}, file={Path(boot_path).name})", raw_return=ret)

    def clear_all_processes(self) -> None:
        for proc in range(1, 11):
            self._dll.ADB_Stop(proc, self._dev)
            self._dll.Clear_Process(proc, self._dev)

    def load_process(self) -> None:
        process_path = self._resolve_file(self.board.process_file)
        original_cwd = os.getcwd()
        try:
            os.chdir(str(Path(process_path).parent))
            ret = int(self._dll.ADBload(os.fsencode(process_path), self._dev, 1))
        finally:
            os.chdir(original_cwd)
        if ret != 1:
            raise AdwinError(f"Failed to load ADWIN process {process_path}; return={ret}.")

    def set_par(self, index: int, value: int) -> None:
        ret = int(self._dll.Set_ADBPar(int(index), int(value), self._dev))
        if ret != 0:
            raise AdwinError(f"Set_Par({index}, {value}) failed with code {ret}.")

    def set_fpar(self, index: int, value: float) -> None:
        ret = int(self._dll.Set_ADBFPar(int(index), float(value), self._dev))
        if ret != 0:
            raise AdwinError(f"Set_Fpar({index}, {value}) failed with code {ret}.")

    def get_par(self, index: int) -> int:
        value = int(self._dll.Get_ADBPar(int(index), self._dev))
        self._raise_if_error(f"Get_Par({index}, dev={self._dev})", raw_return=value)
        return value

    def get_fpar(self, index: int) -> float:
        value = float(self._dll.Get_ADBFPar(int(index), self._dev))
        self._raise_if_error(f"Get_FPar({index}, dev={self._dev})", raw_return=value)
        return value

    def set_digout(self, value: int) -> None:
        value = int(value) & 0x3F
        ret = int(self._dll.Set_Digout(value, self._dev))
        if ret != 0:
            code, text = self._last_error()
            if code != 0:
                raise AdwinError(
                    f"Set_Digout({value}, dev={self._dev}) failed with code {ret}. "
                    f"ADwin error {code}: {text}"
                )
            raise AdwinError(f"Set_Digout({value}, dev={self._dev}) failed with code {ret}.")
        self._raise_if_error(f"Set_Digout({value}, dev={self._dev})", raw_return=ret)
        self._last_digout_bit = int(value)

    def get_digout(self) -> int:
        value = int(self._dll.Get_Digout(self._dev))
        self._raise_if_error(f"Get_Digout(dev={self._dev})", raw_return=value)
        return value

    def test_version(self) -> int:
        """Return ADTest_Version result. 0 = not booted / unreachable; nonzero = OK."""
        value = int(self._dll.ADTest_Version(self._dev, 0))
        self._raise_if_error(f"ADTest_Version(dev={self._dev})", raw_return=value)
        return value

    def voltage_to_count(self, volts: float) -> int:
        """Convert ±10 V to a 16-bit DAC count (0..65535, 32768 = 0 V)."""
        return max(0, min(65535, int(volts * self._COUNTS_PER_VOLT) + 32768))

    def count_to_voltage(self, count: int) -> float:
        """Convert a 16-bit ADC count to ±10 V."""
        return (int(count) - 32768) / self._COUNTS_PER_VOLT

    def set_dac(self, channel: int, voltage: float) -> None:
        """Write *voltage* (±10 V) to a DAC output channel. Board must be booted."""
        count = self.voltage_to_count(voltage)
        ret = int(self._dll.Set_DAC(int(channel), int(count), self._dev))
        if ret != 0:
            raise AdwinError(f"Set_DAC(ch={channel}, v={voltage:.3f}V, count={count}) failed with code {ret}.")

    def get_adc(self, channel: int) -> float:
        """Read ADC voltage (±10 V) from *channel*. Board must be booted."""
        count = int(self._dll.Get_ADC(int(channel), self._dev))
        self._raise_if_error(f"Get_ADC(ch={channel}, dev={self._dev})", raw_return=count)
        return self.count_to_voltage(count)

    def get_data_long(self, data_no: int, start_index: int, count: int) -> list[int]:
        if count <= 0:
            return []
        buffer = (ctypes.c_long * count)()
        ret = int(self._dll.Get_Data(buffer, 2, int(data_no), int(start_index), int(count), self._dev))
        self._raise_if_error(
            f"Get_Data(data_no={data_no}, start={start_index}, count={count}, dev={self._dev})",
            raw_return=ret,
        )
        return [int(buffer[i]) for i in range(count)]

    def calc_digout_bit(self, chan_num: int, set_high: bool, one_chan_on: bool = True) -> int:
        if chan_num < 0 or chan_num > 5:
            raise AdwinError(f"Digital channel {chan_num} is out of supported range 0..5.")

        bit_value = (2 ** int(chan_num)) if set_high else 0
        if one_chan_on:
            return bit_value

        last_value = self._last_digout_bit
        if not (0 <= last_value <= 63):
            last_value = self.get_digout()

        for bit in range(6):
            if bit == int(chan_num):
                continue
            if ((last_value // (2**bit)) % 2) == 1:
                bit_value += 2**bit

        return bit_value

    def set_af_relays(self, active_coil: str, one_chan_on: bool = True) -> int:
        key = active_coil.strip().lower()
        if key == "axial":
            chan = int(self.board.axial_relay_bit)
        elif key == "transverse":
            chan = int(self.board.trans_relay_bit)
        elif key in {"none", "off", ""}:
            self.boot_board()
            self.set_digout(0)
            return 0
        else:
            raise AdwinError(f"Unsupported coil {active_coil!r}; expected axial/transverse.")

        self.boot_board()
        bit_value = self.calc_digout_bit(chan, set_high=True, one_chan_on=one_chan_on)
        self.set_digout(bit_value)
        time.sleep(1.0)
        return bit_value

    def _coil_limits(self, coil: str) -> tuple[float, float]:
        key = coil.strip().lower()
        if key == "axial":
            return self.limits.axial_ramp_max, self.limits.axial_monitor_max
        if key == "transverse":
            return self.limits.trans_ramp_max, self.limits.trans_monitor_max
        raise AdwinError(f"Unsupported coil {coil!r}; expected axial/transverse.")

    def run_ramp(self, request: AdwinRampRequest, timeout_s: float = 90.0) -> AdwinRampResult:
        self.boot_board()
        self.set_af_relays(request.active_coil, one_chan_on=True)
        self.clear_all_processes()
        self.load_process()

        max_ramp_v, max_mon_v = self._coil_limits(request.active_coil)

        self.set_fpar(31, request.slope_up)
        self.set_fpar(32, request.slope_down)
        self.set_fpar(33, request.peak_monitor_voltage)
        self.set_fpar(34, request.sine_freq_hz)
        self.set_fpar(35, request.ramp_peak_voltage)
        self.set_fpar(36, max_ramp_v)
        self.set_fpar(37, max_mon_v)

        self.set_par(31, request.ramp_mode)
        self.set_par(32, self.board.ramp_dac_chan)
        self.set_par(33, self.board.monitor_adc_chan)
        self.set_par(34, int(1_000_000.0 / request.io_rate_hz * 40.0))
        self.set_par(35, int(request.noise_level))
        self.set_par(36, int(request.hold_ms * request.sine_freq_hz / 1000.0))
        self.set_par(37, int(request.ramp_peak_voltage * request.sine_freq_hz / max(request.slope_down, 1e-9)))
        self.set_par(38, int(request.ramp_down_mode))

        ret = int(self._dll.ADB_Start(self.board.process_num, self._dev))
        if ret != 0:
            raise AdwinError(f"Start_Process({self.board.process_num}) failed with code {ret}.")

        start = time.monotonic()
        while True:
            if self.get_par(4) == 7:
                break
            if time.monotonic() - start > timeout_s:
                self._dll.ADB_Stop(self.board.process_num, self._dev)
                raise AdwinError("ADWIN AF ramp timeout while waiting for process completion.")
            time.sleep(0.2)

        result = AdwinRampResult(
            out_count=self.get_par(5),
            in_count=max(0, self.get_par(6) - 10),
            up_count=self.get_par(7),
            down_start=self.get_par(8),
            monitor_peak_v=self.get_fpar(4),
            ramp_peak_v=self.get_fpar(5),
            down_slope_vps=self.get_fpar(32),
            timestep_s=self.get_fpar(6),
            points_per_period=self.get_fpar(7),
        )

        self.clear_all_processes()
        return result

    def run_dense_loopback(
        self,
        request: AdwinDenseCaptureRequest,
        timeout_s: float = 30.0,
        should_stop: Callable[[], bool] | None = None,
    ) -> AdwinDenseCaptureResult:
        """Run a dense, board-timed sine capture using the legacy ADwin process.

        Unlike the host-polled loopback path, this loads the ADbasic process onto
        the board, configures its timer/process delay from ``io_rate_hz``, and then
        bulk-reads the captured arrays from DATA_31 / DATA_32 after completion.
        """
        self.boot_board()
        self.clear_all_processes()
        self.load_process()

        amplitude_v = max(0.0, min(10.0, float(request.amplitude_v)))
        sine_freq_hz = max(0.1, float(request.sine_freq_hz))
        io_rate_hz = max(1.0, float(request.io_rate_hz))
        duration_s = max(0.05, float(request.duration_s))

        self.set_fpar(31, float(request.ramp_up_slope_vps))
        self.set_fpar(32, float(request.ramp_down_slope_vps))
        self.set_fpar(33, amplitude_v)
        self.set_fpar(34, sine_freq_hz)
        self.set_fpar(35, amplitude_v)
        self.set_fpar(36, 10.0)
        self.set_fpar(37, 10.0)

        self.set_par(31, int(request.ramp_mode))
        self.set_par(32, int(request.dac_chan))
        self.set_par(33, int(request.adc_chan))
        self.set_par(34, int(1_000_000.0 / io_rate_hz * 40.0))
        self.set_par(35, int(request.noise_level))
        self.set_par(36, max(1, int(round(duration_s * sine_freq_hz))))
        self.set_par(37, max(1, int(request.ramp_down_periods)))
        self.set_par(38, 1)  # ramp down using number of periods

        ret = int(self._dll.ADB_Start(self.board.process_num, self._dev))
        if ret != 0:
            raise AdwinError(f"Start_Process({self.board.process_num}) failed with code {ret}.")

        start = time.monotonic()
        while True:
            if should_stop is not None and should_stop():
                self._dll.ADB_Stop(self.board.process_num, self._dev)
                self.clear_all_processes()
                raise AdwinError("Dense loopback capture stopped by user.")
            if self.get_par(4) == 7:
                break
            if time.monotonic() - start > timeout_s:
                self._dll.ADB_Stop(self.board.process_num, self._dev)
                self.clear_all_processes()
                raise AdwinError("ADWIN dense loopback timeout while waiting for process completion.")
            time.sleep(0.05)

        time.sleep(max(0.1, min(1.0, duration_s / 4.0)))

        out_count = max(0, self.get_par(5) - 10)
        in_count = max(0, self.get_par(6) - 10)
        up_count = max(0, self.get_par(7))
        down_start = max(0, self.get_par(8))
        timestep_s = float(self.get_fpar(6))
        points_per_period = float(self.get_fpar(7))

        capture_count = min(in_count, out_count) if out_count > 0 else in_count
        raw_adc = self.get_data_long(31, 1, capture_count)
        raw_dac = self.get_data_long(32, 1, capture_count)

        steady_start_idx = min(max(up_count - 1, 0), capture_count)
        steady_stop_idx = min(max(down_start - 1, steady_start_idx), capture_count)
        if steady_stop_idx - steady_start_idx >= max(10, int(points_per_period)):
            start_idx = steady_start_idx
            stop_idx = steady_stop_idx
        else:
            start_idx = 0
            stop_idx = capture_count

        dac_v = [self.count_to_voltage(value) for value in raw_dac[start_idx:stop_idx]]
        adc_v = [self.count_to_voltage(value) for value in raw_adc[start_idx:stop_idx]]
        time_s = [i * timestep_s for i in range(len(adc_v))]

        self.clear_all_processes()
        return AdwinDenseCaptureResult(
            time_s=time_s,
            dac_v=dac_v,
            adc_v=adc_v,
            out_count=out_count,
            in_count=in_count,
            up_count=up_count,
            down_start=down_start,
            timestep_s=timestep_s,
            points_per_period=points_per_period,
            steady_start_idx=start_idx,
            steady_stop_idx=stop_idx,
        )
