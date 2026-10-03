"""Checked MCC Universal Library calls used by the legacy ARM bias circuit.

Signatures/constants are from VB6/CBW.BAS and frmMCC.frm. Loading/probing
never configures a port or writes an output. Drivers must be installed by the lab.
"""
import ctypes
import math
import os
import struct


class MccError(RuntimeError):
    pass


class MccDaq:
    simulated = False

    def __init__(self, board: int, *, dll=None):
        if not isinstance(board, int) or board < 0:
            raise ValueError("An explicit MCC board number is required.")
        self.board = board
        if dll is None:
            if os.name != "nt":
                raise MccError("MCC Universal Library requires Windows.")
            name = "cbw64.dll" if struct.calcsize("P") == 8 else "cbw32.dll"
            try:
                dll = ctypes.WinDLL(name)
            except OSError as exc:
                raise MccError(f"Cannot load {name}; install the matching MCC Universal Library driver.") from exc
            signatures = {
                "cbGetBoardName": [ctypes.c_int, ctypes.c_char_p],
                "cbFromEngUnits": [ctypes.c_int, ctypes.c_int, ctypes.c_float, ctypes.POINTER(ctypes.c_ushort)],
                "cbAOut": [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_ushort],
                "cbAIn": [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_ushort)],
                "cbToEngUnits": [ctypes.c_int, ctypes.c_int, ctypes.c_ushort, ctypes.POINTER(ctypes.c_float)],
                "cbDConfigBit": [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int],
                "cbDBitOut": [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_ushort],
            }
            for function, args in signatures.items():
                getattr(dll, function).argtypes = args
                getattr(dll, function).restype = ctypes.c_int
        self.dll = dll
        self._configured_bits = set()
        self.board_name = self.probe()

    @staticmethod
    def _checked(operation, result):
        if int(result) != 0:
            raise MccError(f"{operation} failed with MCC status {int(result)}.")

    def probe(self):
        name = ctypes.create_string_buffer(256)
        self._checked("cbGetBoardName", self.dll.cbGetBoardName(self.board, name))
        value = name.value.decode("ascii", errors="replace").strip()
        if not value:
            raise MccError("MCC board probe returned an empty board name.")
        return value

    def analog_output(self, channel: int, voltage: float, voltage_range: int):
        if channel < 0 or not math.isfinite(float(voltage)) or not 0 <= voltage <= 10:
            raise ValueError("MCC DAC channel and voltage must be configured within 0..10 V.")
        raw = ctypes.c_ushort()
        self._checked("cbFromEngUnits", self.dll.cbFromEngUnits(self.board, voltage_range, float(voltage), ctypes.byref(raw)))
        self._checked("cbAOut", self.dll.cbAOut(self.board, channel, voltage_range, raw.value))
        return raw.value

    def digital_output(self, port: int, bit: int, high: bool):
        if port < 0 or bit < 0:
            raise ValueError("MCC digital port and bit must be configured.")
        if (port, bit) not in self._configured_bits:
            self._checked("cbDConfigBit", self.dll.cbDConfigBit(self.board, port, bit, 1))
            self._configured_bits.add((port, bit))
        self._checked("cbDBitOut", self.dll.cbDBitOut(self.board, port, bit, int(bool(high))))

    def analog_input(self, channel: int, voltage_range: int):
        if channel < 0:
            raise ValueError("MCC ADC channel must be configured.")
        raw = ctypes.c_ushort()
        voltage = ctypes.c_float()
        self._checked("cbAIn", self.dll.cbAIn(self.board, channel, voltage_range, ctypes.byref(raw)))
        self._checked("cbToEngUnits", self.dll.cbToEngUnits(self.board, voltage_range, raw.value, ctypes.byref(voltage)))
        if not math.isfinite(voltage.value):
            raise MccError("MCC ADC conversion returned a non-finite voltage.")
        return float(voltage.value)
