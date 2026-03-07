from __future__ import annotations

import ctypes
import os
import platform
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


class AdwinError(RuntimeError):
    """Raised when ADWIN board operations fail."""


@dataclass(slots=True)
class AdwinBoardConfig:
    board_num: int = 1
    bin_folder: str = ""
    boot_file: str = "ADwin9.btl"
    process_file: str = "AF_Ramp_System.abp"
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


class _AdwinDll:
    def __init__(self) -> None:
        if platform.system() != "Windows":
            raise AdwinError("ADWIN backend requires Windows (adwin32.dll).")
        try:
            self._dll = ctypes.WinDLL("adwin32.dll")
        except OSError as exc:
            raise AdwinError("Unable to load adwin32.dll. Install ADWIN runtime on this machine.") from exc

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
        self.Set_Digout.argtypes = [ctypes.c_long, ctypes.c_int]
        self.Set_Digout.restype = ctypes.c_int

        self.Get_Digout = self._dll.Get_Digout
        self.Get_Digout.argtypes = [ctypes.c_int]
        self.Get_Digout.restype = ctypes.c_long


class AdwinAFController:
    def __init__(self, board: Optional[AdwinBoardConfig] = None, limits: Optional[AdwinCoilLimits] = None) -> None:
        self.board = board or AdwinBoardConfig()
        self.limits = limits or AdwinCoilLimits()
        self._dll = _AdwinDll()
        self._last_digout_bit: int = -1

    @property
    def _dev(self) -> int:
        return int(self.board.board_num)

    def _resolve_file(self, filename: str) -> str:
        root = Path(self.board.bin_folder).expanduser() if self.board.bin_folder else Path.cwd()
        path = root / filename
        return str(path)

    def boot_board(self) -> None:
        ver = int(self._dll.ADTest_Version(self._dev, 0))
        if ver != 0:
            boot_path = self._resolve_file(self.board.boot_file)
            ret = int(self._dll.ADboot(os.fsencode(boot_path), self._dev, 0, 1))
            if ret != 8000:
                raise AdwinError(f"ADWIN boot failed with code {ret} using {boot_path}.")

    def clear_all_processes(self) -> None:
        for proc in range(1, 11):
            self._dll.ADB_Stop(proc, self._dev)
            self._dll.Clear_Process(proc, self._dev)

    def load_process(self) -> None:
        process_path = self._resolve_file(self.board.process_file)
        ret = int(self._dll.ADBload(os.fsencode(process_path), self._dev, 1))
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
        return int(self._dll.Get_ADBPar(int(index), self._dev))

    def get_fpar(self, index: int) -> float:
        return float(self._dll.Get_ADBFPar(int(index), self._dev))

    def set_digout(self, value: int) -> None:
        ret = int(self._dll.Set_Digout(int(value), self._dev))
        if ret != 0:
            raise AdwinError(f"Set_Digout({value}) failed with code {ret}.")
        self._last_digout_bit = int(value)

    def get_digout(self) -> int:
        return int(self._dll.Get_Digout(self._dev))

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
