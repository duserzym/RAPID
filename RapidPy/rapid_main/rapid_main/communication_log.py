"""Communication transcript helpers mapped from legacy ``modListenAndLog``.

The VB6 listener/logger routines acted as an operator-facing trace around
serial command traffic. RapidPy keeps that behavior transport-agnostic here:
callers can record events directly or wrap a serial-like transport without
changing the transport's command semantics.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Iterable


Clock = Callable[[], datetime]


class CommunicationDirection(str, Enum):
    """Stable communication transcript event directions."""

    INFO = "INFO"
    TX = "TX"
    RX = "RX"
    ERROR = "ERROR"


@dataclass(frozen=True)
class CommunicationEvent:
    """One auditable communication event."""

    timestamp: datetime
    channel: str
    direction: CommunicationDirection
    payload: str = ""
    port: str = ""
    detail: str = ""

    def format_for_log(self) -> str:
        """Return a single-line tab-separated transcript entry."""

        return "\t".join(
            (
                _format_timestamp(self.timestamp),
                _clean_field(self.channel),
                self.direction.value,
                _clean_field(self.port),
                _clean_field(self.payload),
                _clean_field(self.detail),
            )
        )


@dataclass
class CommunicationTranscript:
    """In-memory transcript with deterministic file export."""

    events: list[CommunicationEvent] = field(default_factory=list)

    def append(self, event: CommunicationEvent) -> CommunicationEvent:
        self.events.append(event)
        return event

    def extend(self, events: Iterable[CommunicationEvent]) -> None:
        self.events.extend(events)

    def lines(self, *, include_header: bool = True) -> list[str]:
        body = [event.format_for_log() for event in self.events]
        if include_header:
            return ["timestamp\tchannel\tdirection\tport\tpayload\tdetail", *body]
        return body

    def write_text(self, path: str | Path, *, include_header: bool = True) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(self.lines(include_header=include_header)) + "\n", encoding="utf-8")
        return target


class CommunicationLogger:
    """Records normalized communication events for one hardware channel."""

    def __init__(
        self,
        channel: str,
        *,
        port: str = "",
        clock: Clock | None = None,
        max_payload_chars: int = 240,
    ) -> None:
        self.channel = str(channel)
        self.port = str(port)
        self.clock = clock or (lambda: datetime.now(timezone.utc))
        self.max_payload_chars = int(max_payload_chars)
        self.transcript = CommunicationTranscript()

    def record(
        self,
        direction: CommunicationDirection,
        payload: Any = "",
        *,
        port: str | None = None,
        detail: str = "",
    ) -> CommunicationEvent:
        event = CommunicationEvent(
            timestamp=self.clock(),
            channel=self.channel,
            direction=direction,
            payload=normalize_payload(payload, max_chars=self.max_payload_chars),
            port=self.port if port is None else str(port),
            detail=normalize_payload(detail, max_chars=self.max_payload_chars),
        )
        return self.transcript.append(event)

    def info(self, detail: Any, *, port: str | None = None) -> CommunicationEvent:
        return self.record(CommunicationDirection.INFO, "", port=port, detail=str(detail))

    def sent(self, payload: Any, *, port: str | None = None, detail: str = "") -> CommunicationEvent:
        return self.record(CommunicationDirection.TX, payload, port=port, detail=detail)

    def received(self, payload: Any, *, port: str | None = None, detail: str = "") -> CommunicationEvent:
        return self.record(CommunicationDirection.RX, payload, port=port, detail=detail)

    def error(self, detail: Any, *, payload: Any = "", port: str | None = None) -> CommunicationEvent:
        return self.record(CommunicationDirection.ERROR, payload, port=port, detail=str(detail))

    def write_text(self, path: str | Path, *, include_header: bool = True) -> Path:
        return self.transcript.write_text(path, include_header=include_header)


class LoggingTransportBridge:
    """Wrap a serial-like transport while preserving its behavior.

    The bridge logs common ``connect``, ``disconnect``, ``write``, ``read``,
    ``readline``, and ``query`` calls. Unknown attributes delegate to the
    wrapped transport so existing hardware adapters can adopt the bridge
    incrementally.
    """

    def __init__(self, transport: Any, logger: CommunicationLogger) -> None:
        self._transport = transport
        self.logger = logger

    @property
    def transcript(self) -> CommunicationTranscript:
        return self.logger.transcript

    def connect(self, *args: Any, **kwargs: Any) -> Any:
        port = _port_from_call(args, kwargs, self.logger.port)
        self.logger.info("connect", port=port)
        try:
            return self._transport.connect(*args, **kwargs)
        except Exception as exc:
            self.logger.error(f"connect failed: {exc}", port=port)
            raise

    def disconnect(self, *args: Any, **kwargs: Any) -> Any:
        self.logger.info("disconnect")
        try:
            return self._transport.disconnect(*args, **kwargs)
        except Exception as exc:
            self.logger.error(f"disconnect failed: {exc}")
            raise

    def write(self, payload: Any, *args: Any, **kwargs: Any) -> Any:
        self.logger.sent(payload)
        try:
            return self._transport.write(payload, *args, **kwargs)
        except Exception as exc:
            self.logger.error(f"write failed: {exc}", payload=payload)
            raise

    def read(self, *args: Any, **kwargs: Any) -> Any:
        try:
            payload = self._transport.read(*args, **kwargs)
        except Exception as exc:
            self.logger.error(f"read failed: {exc}")
            raise
        self.logger.received(payload)
        return payload

    def readline(self, *args: Any, **kwargs: Any) -> Any:
        try:
            payload = self._transport.readline(*args, **kwargs)
        except Exception as exc:
            self.logger.error(f"readline failed: {exc}")
            raise
        self.logger.received(payload, detail="readline")
        return payload

    def query(self, payload: Any, *args: Any, **kwargs: Any) -> Any:
        self.write(payload)
        return self.read(*args, **kwargs)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._transport, name)


def normalize_payload(payload: Any, *, max_chars: int = 240) -> str:
    """Return a transcript-safe single-line payload string."""

    if payload is None:
        text = ""
    elif isinstance(payload, bytes):
        text = payload.decode("utf-8", errors="replace")
    else:
        text = str(payload)
    text = text.replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
    if len(text) <= max_chars:
        return text
    omitted = len(text) - max_chars
    return f"{text[:max_chars]}...<truncated {omitted} chars>"


def _format_timestamp(timestamp: datetime) -> str:
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    return timestamp.astimezone(timezone.utc).isoformat(timespec="milliseconds")


def _clean_field(value: str) -> str:
    return normalize_payload(value, max_chars=10_000)


def _port_from_call(args: tuple[Any, ...], kwargs: dict[str, Any], fallback: str) -> str:
    if "port" in kwargs:
        return str(kwargs["port"])
    if args:
        return str(args[0])
    return fallback
