"""Auditable, append-only calibration approval and activation registry.

Calibration result files are evidence, not approved configuration by themselves.
This registry snapshots an artifact, records who approved it and for how long,
and uses immutable lifecycle events for activation, invalidation, and rollback.
No existing record or event is edited during normal operation.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any, Mapping
from uuid import uuid4


RECORD_SCHEMA = "rapidpy.calibration.record.v1"
EVENT_SCHEMA = "rapidpy.calibration.event.v1"


class CalibrationRegistryError(RuntimeError):
    """Raised when calibration lifecycle data is missing, invalid, or unsafe."""


@dataclass(frozen=True, slots=True)
class CalibrationRecord:
    record_id: str
    procedure_id: str
    procedure_name: str
    version: int
    artifact_sha256: str
    artifact_path: str
    source_artifact_path: str
    approved_by: str
    approved_at_iso: str
    valid_from_iso: str
    expires_at_iso: str
    notes: str = ""
    schema: str = RECORD_SCHEMA

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> "CalibrationRecord":
        if payload.get("schema") != RECORD_SCHEMA:
            raise CalibrationRegistryError("unsupported calibration record schema")
        try:
            record = cls(
                record_id=str(payload["record_id"]),
                procedure_id=str(payload["procedure_id"]),
                procedure_name=str(payload["procedure_name"]),
                version=int(payload["version"]),
                artifact_sha256=str(payload["artifact_sha256"]),
                artifact_path=str(payload["artifact_path"]),
                source_artifact_path=str(payload.get("source_artifact_path", "")),
                approved_by=str(payload["approved_by"]),
                approved_at_iso=str(payload["approved_at_iso"]),
                valid_from_iso=str(payload["valid_from_iso"]),
                expires_at_iso=str(payload.get("expires_at_iso", "")),
                notes=str(payload.get("notes", "")),
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise CalibrationRegistryError(f"invalid calibration record: {exc}") from exc
        if not record.record_id or not record.procedure_id or record.version < 1:
            raise CalibrationRegistryError("calibration record identity/version is invalid")
        if not re.fullmatch(r"[0-9a-f]{64}", record.artifact_sha256):
            raise CalibrationRegistryError("calibration artifact hash is invalid")
        _parse_iso(record.approved_at_iso, "approved_at_iso")
        _parse_iso(record.valid_from_iso, "valid_from_iso")
        if record.expires_at_iso:
            _parse_iso(record.expires_at_iso, "expires_at_iso")
        return record

    def to_payload(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class CalibrationEvent:
    event_id: str
    event_type: str
    record_id: str
    procedure_id: str
    operator: str
    timestamp_iso: str
    reason: str = ""
    schema: str = EVENT_SCHEMA

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> "CalibrationEvent":
        if payload.get("schema") != EVENT_SCHEMA:
            raise CalibrationRegistryError("unsupported calibration event schema")
        try:
            event = cls(
                event_id=str(payload["event_id"]),
                event_type=str(payload["event_type"]),
                record_id=str(payload["record_id"]),
                procedure_id=str(payload["procedure_id"]),
                operator=str(payload["operator"]),
                timestamp_iso=str(payload["timestamp_iso"]),
                reason=str(payload.get("reason", "")),
            )
        except (KeyError, TypeError) as exc:
            raise CalibrationRegistryError(f"invalid calibration event: {exc}") from exc
        if event.event_type not in {"approved", "activated", "invalidated"}:
            raise CalibrationRegistryError(f"unknown calibration event: {event.event_type}")
        if not event.event_id or not event.record_id or not event.procedure_id:
            raise CalibrationRegistryError("calibration event identity is invalid")
        _parse_iso(event.timestamp_iso, "timestamp_iso")
        return event

    def to_payload(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class CalibrationRecordState:
    record: CalibrationRecord
    status: str
    active: bool
    integrity_ok: bool
    detail: str = ""


class CalibrationRegistry:
    """Filesystem registry whose records and lifecycle events are append-only."""

    def __init__(self, root: str | Path) -> None:
        self.root = Path(root)
        self.records_dir = self.root / "records"
        self.events_dir = self.root / "events"
        self.artifacts_dir = self.root / "artifacts"

    @classmethod
    def default(cls) -> "CalibrationRegistry":
        config_path = Path(os.environ.get("RAPID_CONFIG", Path.home() / ".rapid" / "config.json"))
        return cls(config_path.expanduser().parent / "calibrations" / "registry")

    def approve(
        self,
        artifact_path: str | Path,
        *,
        approved_by: str,
        valid_days: int | None = 365,
        notes: str = "",
        activate: bool = True,
        now: datetime | None = None,
    ) -> CalibrationRecord:
        """Snapshot, approve, and optionally activate a calibration artifact."""

        operator = approved_by.strip()
        if not operator:
            raise CalibrationRegistryError("approver is required")
        if valid_days is not None and valid_days <= 0:
            raise CalibrationRegistryError("valid_days must be positive or None")
        source = Path(artifact_path).expanduser().resolve()
        try:
            artifact_bytes = source.read_bytes()
            artifact_payload = json.loads(artifact_bytes.decode("utf-8"))
        except (OSError, UnicodeError, ValueError) as exc:
            raise CalibrationRegistryError(f"cannot read calibration artifact: {exc}") from exc
        if not isinstance(artifact_payload, dict):
            raise CalibrationRegistryError("calibration artifact must contain a JSON object")
        procedure_id = str(artifact_payload.get("procedure_id", "")).strip()
        if not procedure_id:
            raise CalibrationRegistryError("calibration artifact has no procedure_id")
        if artifact_payload.get("passed") is False or artifact_payload.get("status") == "FAIL":
            raise CalibrationRegistryError("a failed calibration artifact cannot be approved")

        timestamp = _as_utc(now)
        digest = hashlib.sha256(artifact_bytes).hexdigest()
        version = 1 + max(
            (record.version for record in self.records() if record.procedure_id == procedure_id),
            default=0,
        )
        record_id = (
            f"{_slug(procedure_id)}-v{version:03d}-"
            f"{timestamp.strftime('%Y%m%dT%H%M%SZ')}-{uuid4().hex[:8]}"
        )
        suffix = source.suffix.lower() if source.suffix else ".json"
        snapshot_name = f"{record_id}{suffix}"
        expires = timestamp + timedelta(days=valid_days) if valid_days is not None else None
        record = CalibrationRecord(
            record_id=record_id,
            procedure_id=procedure_id,
            procedure_name=str(artifact_payload.get("procedure_name", procedure_id)),
            version=version,
            artifact_sha256=digest,
            artifact_path=f"artifacts/{snapshot_name}",
            source_artifact_path=str(source),
            approved_by=operator,
            approved_at_iso=_iso(timestamp),
            valid_from_iso=_iso(timestamp),
            expires_at_iso=_iso(expires) if expires is not None else "",
            notes=notes.strip(),
        )

        self._ensure_dirs()
        snapshot = self.root / record.artifact_path
        _write_bytes_exclusive(snapshot, artifact_bytes)
        try:
            _write_json_exclusive(self.records_dir / f"{record_id}.json", record.to_payload())
            self._append_event("approved", record, operator, notes, timestamp)
            if activate:
                self._append_event("activated", record, operator, "approved and activated", timestamp)
        except BaseException:
            # A snapshot without a record has no lifecycle meaning and can be
            # safely cleaned up; records/events are never removed here.
            if not (self.records_dir / f"{record_id}.json").exists():
                try:
                    snapshot.unlink()
                except OSError:
                    pass
            raise
        return record

    def activate(
        self,
        record_id: str,
        *,
        operator: str,
        reason: str,
        now: datetime | None = None,
    ) -> CalibrationRecord:
        """Activate a valid prior record; this is the non-destructive rollback path."""

        actor = operator.strip()
        if not actor:
            raise CalibrationRegistryError("operator is required")
        if not reason.strip():
            raise CalibrationRegistryError("activation/rollback reason is required")
        record = self.get(record_id)
        state = self.state_for(record_id, now=now)
        if state.status in {"invalidated", "expired", "integrity-error"}:
            raise CalibrationRegistryError(f"cannot activate {state.status} calibration")
        self._append_event("activated", record, actor, reason.strip(), _as_utc(now))
        return record

    def invalidate(
        self,
        record_id: str,
        *,
        operator: str,
        reason: str,
        now: datetime | None = None,
    ) -> CalibrationRecord:
        actor = operator.strip()
        if not actor:
            raise CalibrationRegistryError("operator is required")
        if not reason.strip():
            raise CalibrationRegistryError("invalidation reason is required")
        record = self.get(record_id)
        if self._is_invalidated(record_id):
            raise CalibrationRegistryError("calibration record is already invalidated")
        self._append_event("invalidated", record, actor, reason.strip(), _as_utc(now))
        return record

    def records(self) -> tuple[CalibrationRecord, ...]:
        if not self.records_dir.exists():
            return ()
        result = []
        for path in sorted(self.records_dir.glob("*.json")):
            result.append(CalibrationRecord.from_payload(_read_json_object(path)))
        return tuple(sorted(result, key=lambda item: (item.procedure_id, item.version)))

    def events(self) -> tuple[CalibrationEvent, ...]:
        if not self.events_dir.exists():
            return ()
        return tuple(
            CalibrationEvent.from_payload(_read_json_object(path))
            for path in sorted(self.events_dir.glob("*.json"))
        )

    def get(self, record_id: str) -> CalibrationRecord:
        path = self.records_dir / f"{record_id}.json"
        if not path.exists():
            raise CalibrationRegistryError(f"unknown calibration record: {record_id}")
        return CalibrationRecord.from_payload(_read_json_object(path))

    def states(self, *, now: datetime | None = None) -> tuple[CalibrationRecordState, ...]:
        return tuple(self.state_for(record.record_id, now=now) for record in self.records())

    def state_for(
        self,
        record_id: str,
        *,
        now: datetime | None = None,
    ) -> CalibrationRecordState:
        record = self.get(record_id)
        integrity_ok = self._artifact_integrity(record)
        invalidated = self._is_invalidated(record_id)
        timestamp = _as_utc(now)
        expired = bool(
            record.expires_at_iso
            and timestamp >= _parse_iso(record.expires_at_iso, "expires_at_iso")
        )
        active_ids = self._selected_record_ids()
        selected = active_ids.get(record.procedure_id) == record.record_id
        if not integrity_ok:
            return CalibrationRecordState(record, "integrity-error", False, False, "artifact hash mismatch")
        if invalidated:
            return CalibrationRecordState(record, "invalidated", False, True)
        if expired:
            return CalibrationRecordState(record, "expired", False, True)
        if selected:
            return CalibrationRecordState(record, "active", True, True)
        return CalibrationRecordState(record, "inactive", False, True)

    def active_records(self, *, now: datetime | None = None) -> tuple[CalibrationRecord, ...]:
        states = self.states(now=now)
        return tuple(state.record for state in states if state.active)

    def provenance_refs(self, *, now: datetime | None = None) -> list[dict[str, Any]]:
        """Return stable active-record references suitable for measurement JSON."""

        return [
            {
                "record_id": record.record_id,
                "procedure_id": record.procedure_id,
                "version": record.version,
                "artifact_sha256": record.artifact_sha256,
                "approved_by": record.approved_by,
                "approved_at_iso": record.approved_at_iso,
                "expires_at_iso": record.expires_at_iso,
            }
            for record in self.active_records(now=now)
        ]

    def _selected_record_ids(self) -> dict[str, str]:
        selected: dict[str, str] = {}
        for event in self.events():
            if event.event_type == "activated":
                selected[event.procedure_id] = event.record_id
        return selected

    def _is_invalidated(self, record_id: str) -> bool:
        return any(
            event.event_type == "invalidated" and event.record_id == record_id
            for event in self.events()
        )

    def _artifact_integrity(self, record: CalibrationRecord) -> bool:
        path = self.root / record.artifact_path
        try:
            return hashlib.sha256(path.read_bytes()).hexdigest() == record.artifact_sha256
        except OSError:
            return False

    def _append_event(
        self,
        event_type: str,
        record: CalibrationRecord,
        operator: str,
        reason: str,
        timestamp: datetime,
    ) -> CalibrationEvent:
        self._ensure_dirs()
        event = CalibrationEvent(
            event_id=uuid4().hex,
            event_type=event_type,
            record_id=record.record_id,
            procedure_id=record.procedure_id,
            operator=operator,
            timestamp_iso=_iso(timestamp),
            reason=reason,
        )
        ordering = {"approved": "00", "activated": "10", "invalidated": "20"}[event_type]
        name = (
            f"{timestamp.strftime('%Y%m%dT%H%M%S%fZ')}-"
            f"{ordering}-{event.event_id}.json"
        )
        _write_json_exclusive(self.events_dir / name, event.to_payload())
        return event

    def _ensure_dirs(self) -> None:
        self.records_dir.mkdir(parents=True, exist_ok=True)
        self.events_dir.mkdir(parents=True, exist_ok=True)
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)


def _read_json_object(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise CalibrationRegistryError(f"cannot read {path}: {exc}") from exc
    if not isinstance(payload, dict):
        raise CalibrationRegistryError(f"{path} does not contain a JSON object")
    return payload


def _write_json_exclusive(path: Path, payload: Mapping[str, Any]) -> None:
    data = (json.dumps(payload, indent=2, sort_keys=True) + "\n").encode("utf-8")
    _write_bytes_exclusive(path, data)


def _write_bytes_exclusive(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "xb") as handle:
        handle.write(data)
        handle.flush()
        os.fsync(handle.fileno())


def _as_utc(value: datetime | None) -> datetime:
    result = value or datetime.now(timezone.utc)
    if result.tzinfo is None:
        result = result.replace(tzinfo=timezone.utc)
    return result.astimezone(timezone.utc)


def _iso(value: datetime) -> str:
    return _as_utc(value).replace(microsecond=0).isoformat()


def _parse_iso(value: str, field: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise CalibrationRegistryError(f"{field} is not a valid ISO timestamp") from exc
    return _as_utc(parsed)


def _slug(value: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return normalized[:48] or "calibration"

