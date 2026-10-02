from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Mapping, Sequence


VRM_CONTEXT_ENV = "RAPID_VRM_CONTEXT_JSON"
VRM_CONTEXT_SCHEMA = "rapidpy.vrm.launch_context.v1"
VRM_SESSION_SCHEMA = "rapidpy.vrm.session_manifest.v2"
_OUTCOMES = {"operator_stopped", "window_closed", "error", "aborted"}


def _utc_timestamp_from_epoch(epoch: float) -> str:
    return datetime.fromtimestamp(epoch, tz=timezone.utc).isoformat()


def _atomic_write_json(path: Path, payload: Mapping[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            json.dump(dict(payload), handle, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except Exception:
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
        raise


def _safe_session_id(session_id: str) -> str:
    value = str(session_id).strip()
    if not value or not re.fullmatch(r"[A-Za-z0-9._-]+", value):
        raise ValueError(
            "VRM session_id must contain only letters, digits, dot, dash, or underscore"
        )
    return value


def vrm_manifest_path(output_csv: Path | str, session_id: str | None = None) -> Path:
    path = Path(output_csv)
    stem = path.stem if path.suffix else path.name
    if session_id is None:
        return path.with_name(f"{stem}.vrm.json")
    return path.with_name(f"{stem}.{_safe_session_id(session_id)}.vrm.json")


def _handoff_error(path: Path | None, message: str) -> dict[str, object]:
    payload: dict[str, object] = {
        "schema": "rapidpy.vrm.handoff_error.v1",
        "association_status": "invalid",
        "run_association_status": "unassociated",
        "error": message,
    }
    if path is not None:
        payload["context_path"] = str(path)
    return payload


def validate_handoff_context(
    payload: object, *, path: Path | None = None
) -> dict[str, object]:
    """Validate a main-app handoff and label its association explicitly."""

    if not isinstance(payload, dict):
        return _handoff_error(path, "handoff context was not a JSON object")
    if payload.get("schema") != VRM_CONTEXT_SCHEMA:
        return _handoff_error(
            path,
            f"unsupported handoff schema: {payload.get('schema')!r}",
        )
    for field in ("launch_id", "source_app", "launch_action", "launched_at_utc"):
        if not isinstance(payload.get(field), str) or not str(payload[field]).strip():
            return _handoff_error(path, f"handoff context requires {field}")
    for field in ("active_automation", "nocomm"):
        if not isinstance(payload.get(field), bool):
            return _handoff_error(path, f"handoff context requires boolean {field}")
    if payload["active_automation"]:
        return _handoff_error(path, "handoff was issued during active main-app automation")
    context = dict(payload)
    context["association_status"] = "validated_handoff"
    context["run_association_status"] = (
        "associated" if str(context.get("run_id", "")).strip() else "unassociated"
    )
    if path is not None:
        context["context_path"] = str(path)
    return context


def load_handoff_context(env: Mapping[str, str] | None = None) -> dict[str, object]:
    source = os.environ if env is None else env
    raw_path = source.get(VRM_CONTEXT_ENV, "").strip()
    if not raw_path:
        return {
            "schema": "rapidpy.vrm.standalone_context.v1",
            "association_status": "standalone",
            "run_association_status": "unassociated",
            "reason": "no main-app handoff context was provided",
        }
    path = Path(raw_path)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        return _handoff_error(path, str(exc))
    return validate_handoff_context(payload, path=path)


def file_sha256(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_vrm_output_manifest(
    *,
    output_csv: Path | str,
    session_start_epoch: float,
    session_end_epoch: float,
    session_id: str,
    outcome: str,
    row_count: int,
    write_mode: str,
    interval_s: float,
    spacing_mode: str,
    display_unit: str,
    baseline_volts: Sequence[float] | None,
    calibration: Mapping[str, float],
    handoff_context: Mapping[str, object] | None = None,
    device_port: str = "",
    provenance: str = "live_hardware",
    error: str = "",
    output_size_bytes: int = 0,
    output_sha256: str = "",
) -> dict[str, object]:
    session = _safe_session_id(session_id)
    if outcome not in _OUTCOMES:
        raise ValueError(f"unsupported VRM session outcome: {outcome!r}")
    if float(session_end_epoch) < float(session_start_epoch):
        raise ValueError("VRM session end precedes its start")
    if int(row_count) < 0:
        raise ValueError("VRM row_count cannot be negative")
    if write_mode not in {"new", "append", "overwrite"}:
        raise ValueError(f"unsupported VRM write mode: {write_mode!r}")
    if provenance not in {"live_hardware", "simulated"}:
        raise ValueError(f"unsupported VRM provenance: {provenance!r}")
    if outcome == "error" and not str(error).strip():
        raise ValueError("an error outcome requires an error message")
    baseline = list(baseline_volts) if baseline_volts is not None else None
    return {
        "schema": VRM_SESSION_SCHEMA,
        "session_id": session,
        "outcome": outcome,
        "error": str(error),
        "provenance": {
            "kind": provenance,
            "simulated": provenance == "simulated",
            "statement": (
                "Simulated VRM session; not hardware evidence."
                if provenance == "simulated"
                else "VRM values acquired from the configured serial SQUID client."
            ),
        },
        "device": {"port": str(device_port)},
        "output_csv": str(Path(output_csv)),
        "output_size_bytes": int(output_size_bytes),
        "output_sha256": str(output_sha256),
        "row_count": int(row_count),
        "write_mode": write_mode,
        "session_start_utc": _utc_timestamp_from_epoch(session_start_epoch),
        "session_end_utc": _utc_timestamp_from_epoch(session_end_epoch),
        "interval_s": float(interval_s),
        "spacing_mode": str(spacing_mode),
        "display_unit": str(display_unit),
        "baseline_volts": baseline,
        "calibration": dict(calibration),
        "handoff_context": dict(handoff_context or {}),
    }


def write_vrm_output_manifest(
    output_csv: Path | str,
    *,
    session_start_epoch: float,
    session_end_epoch: float,
    session_id: str | None = None,
    outcome: str,
    row_count: int,
    write_mode: str,
    interval_s: float,
    spacing_mode: str,
    display_unit: str,
    baseline_volts: Sequence[float] | None,
    calibration: Mapping[str, float],
    handoff_context: Mapping[str, object] | None = None,
    device_port: str = "",
    provenance: str = "live_hardware",
    error: str = "",
) -> Path:
    csv_path = Path(output_csv)
    if not csv_path.is_file():
        raise FileNotFoundError(f"VRM output CSV does not exist: {csv_path}")
    resolved_session_id = _safe_session_id(session_id or str(uuid.uuid4()))
    manifest_path = vrm_manifest_path(csv_path, resolved_session_id)
    if manifest_path.exists():
        raise FileExistsError(f"VRM session manifest already exists: {manifest_path}")
    payload = build_vrm_output_manifest(
        output_csv=csv_path,
        session_start_epoch=session_start_epoch,
        session_end_epoch=session_end_epoch,
        session_id=resolved_session_id,
        outcome=outcome,
        row_count=row_count,
        write_mode=write_mode,
        interval_s=interval_s,
        spacing_mode=spacing_mode,
        display_unit=display_unit,
        baseline_volts=baseline_volts,
        calibration=calibration,
        handoff_context=handoff_context,
        device_port=device_port,
        provenance=provenance,
        error=error,
        output_size_bytes=csv_path.stat().st_size,
        output_sha256=file_sha256(csv_path),
    )
    _atomic_write_json(manifest_path, payload)
    return manifest_path
