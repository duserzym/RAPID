from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Mapping, Sequence


VRM_CONTEXT_ENV = "RAPID_VRM_CONTEXT_JSON"


def _utc_timestamp_from_epoch(epoch: float) -> str:
    return datetime.fromtimestamp(epoch, tz=timezone.utc).replace(microsecond=0).isoformat()


def vrm_manifest_path(output_csv: Path | str) -> Path:
    path = Path(output_csv)
    if path.suffix:
        return path.with_name(f"{path.stem}.vrm.json")
    return path.with_name(f"{path.name}.vrm.json")


def load_handoff_context(env: Mapping[str, str] | None = None) -> dict[str, object]:
    source = os.environ if env is None else env
    raw_path = source.get(VRM_CONTEXT_ENV, "").strip()
    if not raw_path:
        return {}
    path = Path(raw_path)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return {
            "schema": "rapidpy.vrm.handoff_error.v1",
            "context_path": str(path),
            "error": str(exc),
        }
    if not isinstance(payload, dict):
        return {
            "schema": "rapidpy.vrm.handoff_error.v1",
            "context_path": str(path),
            "error": "handoff context was not a JSON object",
        }
    payload.setdefault("context_path", str(path))
    return payload


def build_vrm_output_manifest(
    *,
    output_csv: Path | str,
    session_start_epoch: float,
    interval_s: float,
    spacing_mode: str,
    display_unit: str,
    baseline_volts: Sequence[float] | None,
    calibration: Mapping[str, float],
    handoff_context: Mapping[str, object] | None = None,
) -> dict[str, object]:
    baseline = list(baseline_volts) if baseline_volts is not None else None
    return {
        "schema": "rapidpy.vrm.session_manifest.v1",
        "output_csv": str(Path(output_csv)),
        "session_start_utc": _utc_timestamp_from_epoch(session_start_epoch),
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
    interval_s: float,
    spacing_mode: str,
    display_unit: str,
    baseline_volts: Sequence[float] | None,
    calibration: Mapping[str, float],
    handoff_context: Mapping[str, object] | None = None,
) -> Path:
    manifest_path = vrm_manifest_path(output_csv)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    payload = build_vrm_output_manifest(
        output_csv=output_csv,
        session_start_epoch=session_start_epoch,
        interval_s=interval_s,
        spacing_mode=spacing_mode,
        display_unit=display_unit,
        baseline_volts=baseline_volts,
        calibration=calibration,
        handoff_context=handoff_context,
    )
    manifest_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    return manifest_path
