from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping


def susceptibility_summary(
    records: Iterable[Mapping[str, object]],
    *,
    sample: str,
    operator: str = "",
    timestamp_iso: str | None = None,
) -> dict[str, Any]:
    """Return a JSON-ready susceptibility run summary."""

    rows = [dict(record) for record in records]
    values = [
        float(row["susceptibility"])
        for row in rows
        if "susceptibility" in row and row["susceptibility"] is not None
    ]
    timestamp = timestamp_iso or datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return {
        "schema": "rapidpy.susceptibility.summary.v1",
        "sample": str(sample),
        "operator": str(operator),
        "timestamp_iso": timestamp,
        "record_count": len(rows),
        "nonzero_count": sum(1 for value in values if value != 0.0),
        "minimum": min(values) if values else None,
        "maximum": max(values) if values else None,
        "records": rows,
        "hardware_validation_required": True,
        "hardware_validation_statement": (
            "Summary records software-captured susceptibility readings; "
            "bridge calibration and live susceptibility hardware acceptance "
            "remain physical validation gates."
        ),
    }


def write_susceptibility_summary_json(
    path: str | Path,
    records: Iterable[Mapping[str, object]],
    *,
    sample: str,
    operator: str = "",
    timestamp_iso: str | None = None,
) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = susceptibility_summary(
        records,
        sample=sample,
        operator=operator,
        timestamp_iso=timestamp_iso,
    )
    target.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target
