from __future__ import annotations

import json
import os
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Mapping


VRM_CONTEXT_ENV = "RAPID_VRM_CONTEXT_JSON"
VRM_CONTEXT_SCHEMA = "rapidpy.vrm.launch_context.v1"


class VrmLaunchContextError(ValueError):
    """Raised when a VRM handoff cannot prove its launch context."""


def _utc_timestamp() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def default_vrm_context_dir() -> Path:
    return Path.home() / "RapidPy" / "vrm_sessions"


def build_vrm_launch_context(
    *,
    source_app: str = "rapid_main",
    launch_action: str = "Diagnostics > VRM Data Collection",
    active_automation: bool = False,
    nocomm: bool = False,
    launch_id: str | None = None,
    run_id: str = "",
    sample_id: str = "",
    operator: str = "",
    intended_output_dir: Path | str | None = None,
    software_version: str = "",
    config_hash: str = "",
    metadata: Mapping[str, object] | None = None,
) -> dict[str, object]:
    timestamp = _utc_timestamp()
    return {
        "schema": VRM_CONTEXT_SCHEMA,
        "launch_id": str(launch_id or uuid.uuid4()),
        "source_app": source_app,
        "launch_action": launch_action,
        "launched_at_utc": timestamp,
        "active_automation": bool(active_automation),
        "nocomm": bool(nocomm),
        "run_id": str(run_id).strip(),
        "sample_id": str(sample_id).strip(),
        "operator": str(operator).strip(),
        "intended_output_dir": (
            str(Path(intended_output_dir)) if intended_output_dir else ""
        ),
        "software_version": str(software_version).strip(),
        "config_hash": str(config_hash).strip(),
        "metadata": dict(metadata or {}),
    }


def validate_vrm_launch_context(context: Mapping[str, object]) -> dict[str, object]:
    """Return a JSON-ready validated handoff without inventing missing identity."""

    payload = dict(context)
    if payload.get("schema") != VRM_CONTEXT_SCHEMA:
        raise VrmLaunchContextError(
            f"unsupported VRM launch-context schema: {payload.get('schema')!r}"
        )
    for field in ("launch_id", "source_app", "launch_action", "launched_at_utc"):
        if not isinstance(payload.get(field), str) or not str(payload[field]).strip():
            raise VrmLaunchContextError(f"VRM launch context requires {field}")
    for field in ("active_automation", "nocomm"):
        if not isinstance(payload.get(field), bool):
            raise VrmLaunchContextError(f"VRM launch context requires boolean {field}")
    if payload["active_automation"]:
        raise VrmLaunchContextError(
            "VRM launch context cannot be issued while main-app automation is active"
        )
    metadata = payload.get("metadata", {})
    if not isinstance(metadata, dict):
        raise VrmLaunchContextError("VRM launch-context metadata must be an object")
    payload["metadata"] = dict(metadata)
    return payload


def vrm_launch_block_reason(
    *,
    active_automation: bool,
    measurement_owner: str | None = None,
    squid_owner: str | None = None,
) -> str:
    """Explain why an external VRM process may not claim the shared SQUID."""

    if active_automation:
        return "A live measurement or queue run is active. Stop it before launching VRM."
    conflicts = []
    if measurement_owner:
        conflicts.append(f"measurement is owned by {measurement_owner!r}")
    if squid_owner:
        conflicts.append(f"SQUID is owned by {squid_owner!r}")
    if conflicts:
        return "VRM cannot launch while " + " and ".join(conflicts) + "."
    return ""


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


def write_vrm_launch_context(
    directory: Path | str | None = None,
    *,
    context: Mapping[str, object] | None = None,
) -> Path:
    target_dir = Path(directory) if directory is not None else default_vrm_context_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    payload = validate_vrm_launch_context(context or build_vrm_launch_context())
    stamp = str(payload.get("launched_at_utc") or _utc_timestamp())
    safe_stamp = (
        stamp.replace(":", "")
        .replace("-", "")
        .replace("+", "Z")
        .replace(".", "_")
    )
    launch_id = str(payload["launch_id"]).replace("/", "_").replace("\\", "_")
    path = target_dir / f"vrm_launch_{safe_stamp}_{launch_id}.json"
    _atomic_write_json(path, payload)
    return path
