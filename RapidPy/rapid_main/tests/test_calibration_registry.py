from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.calibration_registry import (
    CalibrationRegistry,
    CalibrationRegistryError,
)


UTC = timezone.utc


def _write_artifact(path: Path, *, procedure: str, passed: bool = True) -> Path:
    path.write_text(
        json.dumps(
            {
                "procedure_id": procedure,
                "procedure_name": procedure.replace("/", " ").title(),
                "status": "PASS" if passed else "FAIL",
                "passed": passed,
                "samples": [1.0, 1.01],
            }
        ),
        encoding="utf-8",
    )
    return path


class CalibrationRegistryTests(unittest.TestCase):
    def test_approval_snapshots_artifact_and_exposes_provenance(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = _write_artifact(root / "result.json", procedure="calibration/squid")
            registry = CalibrationRegistry(root / "registry")
            now = datetime(2026, 10, 2, 12, 0, tzinfo=UTC)

            record = registry.approve(
                source,
                approved_by="operator-a",
                valid_days=30,
                notes="Reference standard checked.",
                now=now,
            )

            self.assertEqual(record.version, 1)
            self.assertEqual(registry.state_for(record.record_id, now=now).status, "active")
            self.assertEqual(len(registry.events()), 2)
            snapshot = registry.root / record.artifact_path
            self.assertEqual(snapshot.read_bytes(), source.read_bytes())
            self.assertEqual(
                registry.provenance_refs(now=now),
                [
                    {
                        "record_id": record.record_id,
                        "procedure_id": "calibration/squid",
                        "version": 1,
                        "artifact_sha256": record.artifact_sha256,
                        "approved_by": "operator-a",
                        "approved_at_iso": "2026-10-02T12:00:00+00:00",
                        "expires_at_iso": "2026-11-01T12:00:00+00:00",
                    }
                ],
            )

    def test_versions_and_rollback_use_new_events_without_mutating_records(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            registry = CalibrationRegistry(root / "registry")
            first = registry.approve(
                _write_artifact(root / "one.json", procedure="calibration/squid"),
                approved_by="operator-a",
                now=datetime(2026, 1, 1, tzinfo=UTC),
            )
            second = registry.approve(
                _write_artifact(root / "two.json", procedure="calibration/squid"),
                approved_by="operator-b",
                now=datetime(2026, 2, 1, tzinfo=UTC),
            )
            record_text = (registry.records_dir / f"{first.record_id}.json").read_text(
                encoding="utf-8"
            )

            registry.activate(
                first.record_id,
                operator="operator-c",
                reason="Rollback after reference-standard drift.",
                now=datetime(2026, 2, 2, tzinfo=UTC),
            )

            self.assertEqual(second.version, 2)
            self.assertEqual(
                [record.record_id for record in registry.active_records(now=datetime(2026, 2, 2, tzinfo=UTC))],
                [first.record_id],
            )
            self.assertEqual(
                (registry.records_dir / f"{first.record_id}.json").read_text(encoding="utf-8"),
                record_text,
            )
            self.assertEqual(registry.events()[-1].event_type, "activated")

    def test_invalidation_expiry_and_integrity_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            registry = CalibrationRegistry(root / "registry")
            first = registry.approve(
                _write_artifact(root / "one.json", procedure="calibration/squid"),
                approved_by="operator-a",
                valid_days=1,
                now=datetime(2026, 1, 1, tzinfo=UTC),
            )
            self.assertEqual(
                registry.state_for(first.record_id, now=datetime(2026, 1, 3, tzinfo=UTC)).status,
                "expired",
            )
            with self.assertRaisesRegex(CalibrationRegistryError, "expired"):
                registry.activate(
                    first.record_id,
                    operator="operator-a",
                    reason="Unsafe attempt",
                    now=datetime(2026, 1, 3, tzinfo=UTC),
                )

            second = registry.approve(
                _write_artifact(root / "two.json", procedure="calibration/irm"),
                approved_by="operator-b",
                now=datetime(2026, 1, 2, tzinfo=UTC),
            )
            registry.invalidate(
                second.record_id,
                operator="operator-b",
                reason="Reference standard failed post-check.",
                now=datetime(2026, 1, 2, 1, tzinfo=UTC),
            )
            self.assertEqual(registry.state_for(second.record_id).status, "invalidated")
            self.assertNotIn(second.record_id, [item["record_id"] for item in registry.provenance_refs()])

            third = registry.approve(
                _write_artifact(root / "three.json", procedure="calibration/af"),
                approved_by="operator-c",
                now=datetime.now(UTC) - timedelta(minutes=1),
            )
            (registry.root / third.artifact_path).write_text("tampered", encoding="utf-8")
            self.assertEqual(registry.state_for(third.record_id).status, "integrity-error")
            self.assertNotIn(third.record_id, [item["record_id"] for item in registry.provenance_refs()])

    def test_rejects_failed_artifact_and_missing_audit_fields(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            registry = CalibrationRegistry(root / "registry")
            failed = _write_artifact(root / "failed.json", procedure="calibration/squid", passed=False)
            with self.assertRaisesRegex(CalibrationRegistryError, "failed"):
                registry.approve(failed, approved_by="operator-a")
            passing = _write_artifact(root / "pass.json", procedure="calibration/squid")
            with self.assertRaisesRegex(CalibrationRegistryError, "approver"):
                registry.approve(passing, approved_by="")
            record = registry.approve(passing, approved_by="operator-a")
            with self.assertRaisesRegex(CalibrationRegistryError, "reason"):
                registry.invalidate(record.record_id, operator="operator-a", reason="")


if __name__ == "__main__":
    unittest.main(verbosity=2)

