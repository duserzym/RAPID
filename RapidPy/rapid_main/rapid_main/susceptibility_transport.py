"""Fail-closed Bartington susceptibility bridge serial transport.

The legacy VB6 ``frmSusceptibilityMeter`` protocol sends ``Z`` (zero) or
``M`` (measure) followed by CRLF and waits for a CR-terminated ASCII reply.
This module preserves that narrow, evidenced boundary without claiming that
coil motion, holder subtraction, or physical calibration has been accepted.
"""
from __future__ import annotations

from dataclasses import dataclass
import math
import threading
import time
from typing import Callable

import serial

from rapid_main.communication_log import CommunicationEvent, CommunicationLogger


class SusceptibilityTransportError(RuntimeError):
    """Raised when bridge configuration, transport, or response data is invalid."""


@dataclass(frozen=True)
class SusceptibilityTransportConfig:
    port: str
    baud: int = 1200
    parity: str = "N"
    bytesize: int = 8
    stopbits: float = 2.0
    read_timeout_s: float = 0.1
    response_timeout_s: float = 35.0
    scale_factor: float = 1.0

    def validate(self) -> None:
        if not str(self.port).strip():
            raise SusceptibilityTransportError("Susceptibility bridge serial port is not configured.")
        if int(self.baud) <= 0:
            raise SusceptibilityTransportError("Susceptibility bridge baud rate must be positive.")
        if str(self.parity).upper() not in {"N", "E", "O", "M", "S"}:
            raise SusceptibilityTransportError(
                f"Unsupported susceptibility bridge parity: {self.parity!r}."
            )
        if int(self.bytesize) not in {5, 6, 7, 8}:
            raise SusceptibilityTransportError(
                f"Unsupported susceptibility bridge data bits: {self.bytesize!r}."
            )
        if float(self.stopbits) not in {1.0, 1.5, 2.0}:
            raise SusceptibilityTransportError(
                f"Unsupported susceptibility bridge stop bits: {self.stopbits!r}."
            )
        if float(self.read_timeout_s) <= 0 or float(self.response_timeout_s) <= 0:
            raise SusceptibilityTransportError("Susceptibility bridge timeouts must be positive.")
        if not math.isfinite(float(self.scale_factor)) or float(self.scale_factor) <= 0:
            raise SusceptibilityTransportError(
                "Susceptibility bridge scale factor must be finite and positive."
            )


class SusceptibilitySerialClient:
    """Thread-safe serial client for the VB6-evidenced zero/measure protocol."""

    def __init__(
        self,
        config: SusceptibilityTransportConfig,
        *,
        serial_factory: Callable[..., object] = serial.Serial,
        clock: Callable[[], float] = time.monotonic,
        logger: CommunicationLogger | None = None,
    ) -> None:
        config.validate()
        self.config = config
        self._serial_factory = serial_factory
        self._clock = clock
        self._serial: object | None = None
        self._lock = threading.RLock()
        self._logger = logger or CommunicationLogger(
            "SUSCEPTIBILITY", port=str(config.port), max_payload_chars=2048
        )

    @property
    def is_connected(self) -> bool:
        transport = self._serial
        return bool(transport is not None and getattr(transport, "is_open", False))

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return tuple(self._logger.transcript.events)

    def connect(self) -> None:
        with self._lock:
            if self.is_connected:
                return
            if self._serial is not None:
                self.close()  # Settle a retained original handle before replacement.
            try:
                transport = self._serial_factory(
                    port=str(self.config.port),
                    baudrate=int(self.config.baud),
                    bytesize=int(self.config.bytesize),
                    parity=str(self.config.parity).upper(),
                    stopbits=float(self.config.stopbits),
                    timeout=float(self.config.read_timeout_s),
                    write_timeout=float(self.config.read_timeout_s),
                )
            except Exception as exc:
                self._logger.error(f"connect failed: {exc}")
                raise SusceptibilityTransportError(
                    f"Unable to open susceptibility bridge {self.config.port}:{self.config.baud}: {exc}"
                ) from exc
            if not bool(getattr(transport, "is_open", False)):
                try:
                    close = getattr(transport, "close", None)
                    if callable(close):
                        close()
                finally:
                    self._logger.error("connect failed: serial transport did not report open")
                raise SusceptibilityTransportError(
                    "Susceptibility bridge serial transport did not report an open connection."
                )
            self._serial = transport
            self._logger.info(
                f"connected {self.config.port}:{self.config.baud} "
                f"{self.config.parity},{self.config.bytesize},{self.config.stopbits:g}"
            )

    def close(self) -> None:
        with self._lock:
            transport = self._serial
            if transport is None:
                return
            try:
                transport.close()
                if getattr(transport, 'is_open', True) is not False:
                    raise SusceptibilityTransportError('Original susceptibility serial handle remains open after close.')
            except Exception as exc:
                self._logger.error(f'close failed; original handle retained: {exc}')
                raise
            self._serial = None
            self._logger.info("disconnected")

    def zero(self) -> str:
        """Zero the bridge and require a complete acknowledgment."""

        return self._exchange("Z", detail="zero susceptibility bridge")

    def measure(self) -> float:
        """Return one finite, scale-adjusted bridge reading."""

        response = self._exchange("M", detail="measure susceptibility")
        text = response.rstrip("\r").strip()
        try:
            raw_value = float(text)
        except ValueError as exc:
            self._logger.error("measure response is not numeric", payload=response)
            raise SusceptibilityTransportError(
                f"Susceptibility bridge returned a non-numeric measurement: {text!r}."
            ) from exc
        value = raw_value * float(self.config.scale_factor)
        if not math.isfinite(value):
            self._logger.error("measure response is non-finite", payload=response)
            raise SusceptibilityTransportError(
                f"Susceptibility bridge returned a non-finite measurement: {text!r}."
            )
        return float(value)

    def _exchange(self, command: str, *, detail: str) -> str:
        with self._lock:
            transport = self._require_connected()
            payload = f"{command}\r\n".encode("ascii")
            try:
                for method_name in ("reset_output_buffer", "reset_input_buffer"):
                    method = getattr(transport, method_name, None)
                    if callable(method):
                        method()
                self._logger.sent(payload.decode("ascii"), detail=detail)
                written = getattr(transport, "write")(payload)
                if written is not None and int(written) != len(payload):
                    raise SusceptibilityTransportError(
                        f"short write ({written}/{len(payload)} bytes)"
                    )
                flush = getattr(transport, "flush", None)
                if callable(flush):
                    flush()
                response = self._read_cr_terminated(transport)
            except Exception as exc:
                if isinstance(exc, SusceptibilityTransportError):
                    wrapped = exc
                else:
                    wrapped = SusceptibilityTransportError(str(exc))
                self._logger.error(f"{detail} failed: {wrapped}", payload=payload.decode("ascii"))
                raise wrapped from exc
            self._logger.received(response, detail=detail)
            return response

    def _require_connected(self) -> object:
        if not self.is_connected or self._serial is None:
            raise SusceptibilityTransportError("Susceptibility bridge is not connected.")
        return self._serial

    def _read_cr_terminated(self, transport: object) -> str:
        deadline = self._clock() + float(self.config.response_timeout_s)
        raw = bytearray()
        while self._clock() < deadline:
            chunk = getattr(transport, "read")(1)
            if chunk:
                if not isinstance(chunk, (bytes, bytearray)):
                    raise SusceptibilityTransportError("serial read returned non-byte data")
                raw.extend(chunk)
                if raw.endswith(b"\r"):
                    break
        if not raw:
            raise SusceptibilityTransportError("Susceptibility bridge returned an empty response.")
        if not raw.endswith(b"\r"):
            raise SusceptibilityTransportError(
                f"Susceptibility bridge returned an unterminated partial response: {bytes(raw)!r}."
            )
        try:
            return bytes(raw).decode("ascii")
        except UnicodeDecodeError as exc:
            raise SusceptibilityTransportError(
                f"Susceptibility bridge returned non-ASCII bytes: {bytes(raw)!r}."
            ) from exc
