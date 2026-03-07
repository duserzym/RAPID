from __future__ import annotations

import re
import time
from typing import Final

import serial


class SquidCommunicationError(RuntimeError):
    """Raised when the SQUID serial line fails or returns invalid data."""


FLOAT_RE: Final[re.Pattern[str]] = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")


class SquidSerialClient:
    def __init__(self) -> None:
        self._serial: serial.Serial | None = None

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    def connect(
        self,
        port: str,
        baudrate: int = 1200,
        bytesize: int = serial.EIGHTBITS,
        parity: str = serial.PARITY_NONE,
        stopbits: float = serial.STOPBITS_ONE,
        timeout: float = 1.0,
    ) -> None:
        if self.is_connected:
            self.disconnect()

        self._serial = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=bytesize,
            parity=parity,
            stopbits=stopbits,
            timeout=timeout,
            write_timeout=timeout,
        )
        self._serial.reset_input_buffer()
        self._serial.reset_output_buffer()

    def disconnect(self) -> None:
        if self._serial is not None:
            try:
                self._serial.close()
            finally:
                self._serial = None

    def _require_port(self) -> serial.Serial:
        if self._serial is None or not self._serial.is_open:
            raise SquidCommunicationError("SQUID serial port is not connected.")
        return self._serial

    def _send_command(self, command: str) -> None:
        ser = self._require_port()
        payload = f"\r{command}\r".encode("ascii", errors="ignore")
        ser.write(payload)
        ser.flush()

    def _read_response(self, timeout_s: float = 1.0) -> str:
        ser = self._require_port()
        deadline = time.monotonic() + timeout_s
        chunks = bytearray()

        while time.monotonic() < deadline:
            byte = ser.read(1)
            if not byte:
                continue
            if byte == b"\r":
                break
            chunks.extend(byte)

        if not chunks:
            raise SquidCommunicationError("Timed out waiting for SQUID response.")

        return chunks.decode("ascii", errors="ignore").strip()

    def _query_float(self, command: str) -> float:
        self._send_command(command)
        response = self._read_response()
        match = FLOAT_RE.search(response)
        if not match:
            raise SquidCommunicationError(
                f"No numeric value in SQUID response for {command!r}: {response!r}"
            )
        return float(match.group(0))

    def _latch_all_axes(self) -> None:
        self._send_command("ALC")
        time.sleep(0.10)
        self._send_command("ALD")
        time.sleep(0.12)

    def read_axis_volts(self, axis: str) -> float:
        axis = axis.upper()
        count = self._query_float(f"{axis}SC")
        data = self._query_float(f"{axis}SD")
        return -(data + count)

    def read_xyz_volts(self) -> tuple[float, float, float]:
        self._latch_all_axes()
        x = self.read_axis_volts("X")
        y = self.read_axis_volts("Y")
        z = self.read_axis_volts("Z")
        return x, y, z
