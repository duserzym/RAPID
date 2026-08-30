from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Mapping


VRM_CONTEXT_ENV = "RAPID_VRM_CONTEXT_JSON"


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
    metadata: Mapping[str, object] | None = None,
) -> dict[str, object]:
    timestamp = _utc_timestamp()
    return {
        "schema": "rapidpy.vrm.launch_context.v1",
        "source_app": source_app,
        "launch_action": launch_action,
        "launched_at_utc": timestamp,
        "active_automation": bool(active_automation),
        "nocomm": bool(nocomm),
        "metadata": dict(metadata or {}),
    }


def write_vrm_launch_context(
    directory: Path | str | None = None,
    *,
    context: Mapping[str, object] | None = None,
) -> Path:
    target_dir = Path(directory) if directory is not None else default_vrm_context_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    payload = dict(context or build_vrm_launch_context())
    stamp = str(payload.get("launched_at_utc") or _utc_timestamp())
    safe_stamp = (
        stamp.replace(":", "")
        .replace("-", "")
        .replace("+", "Z")
        .replace(".", "_")
    )
    path = target_dir / f"vrm_launch_{safe_stamp}.json"
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return path
