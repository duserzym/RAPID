"""Evidence integrity tests using synthetic operator records."""
from contextlib import redirect_stdout
from datetime import datetime, timedelta, timezone
import hashlib
import io
import json
from pathlib import Path
import re
from tempfile import TemporaryDirectory
import unittest

from rapid_main.__main__ import main
from rapid_main.station_evidence import PROCEDURE_SECTIONS, REQUIRED_CASES, SCHEMA, audit_station_evidence


COMMIT = "1" * 40


class StationEvidenceTests(unittest.TestCase):
    def setUp(self):
        scratch = TemporaryDirectory()
        self.addCleanup(scratch.cleanup)
        self.root = Path(scratch.name)
        self.attachment = self.root / "trace.json"
        self.attachment.write_bytes(b'{"simulated": false}')
        self.manifest = self.root / "campaign.json"

    def record(self, case, serial=1, *, verdict="pass", mode="station"):
        start = datetime(2026, 11, 22, 12, tzinfo=timezone.utc) + timedelta(minutes=serial)
        return {
            "case": case, "record_id": f"test-{serial}", "station": "fixture-station", "software_commit": COMMIT,
            "software_version": "fixture version", "config_hash": "2" * 16, "operator": "Fixture operator",
            "machine": "Fixture host", "device_identity": "Injected device", "evidence_mode": mode,
            "datetime_iso": {"start": start.isoformat(), "end": (start + timedelta(seconds=1)).isoformat()},
            "commands": ["read"], "raw_replies": ["status"], "expected_result": "Fixture expectation",
            "observed_result": "Fixture observation", "verdict": verdict, "reason": "Fixture reason",
            "signature": "Fixture signature", "artifact_paths": [{"original_path": "D:/traces/trace.json",
                "evidence_path": "trace.json", "sha256": hashlib.sha256(self.attachment.read_bytes()).hexdigest()}],
        }

    def audit(self, *records, commit=COMMIT):
        self.manifest.write_text(json.dumps({"schema": SCHEMA, "station": "fixture-station",
                                            "software_commit": COMMIT, "records": records}), encoding="utf-8")
        return audit_station_evidence(self.manifest, commit)

    def test_catalog_covers_the_numbered_procedure_and_reused_ids(self):
        document = Path(__file__).resolve().parents[3] / "docs" / "rapid_hardware_acceptance_procedure_2026-08-29.md"
        sections = re.split(r"^## (?:[2-9]|10|11a?|12)\. ", document.read_text(encoding="utf-8"), flags=re.M)[1:-1]
        self.assertEqual(len(sections), len(PROCEDURE_SECTIONS))
        for content, (section, prefix, count) in zip(sections, PROCEDURE_SECTIONS, strict=True):
            numbered = re.findall(r"^\| ([A-Z]\d+) \|", content, re.M)
            expected = [f"{prefix}{number}" for number in range(1, count + 1)]
            self.assertEqual(numbered, expected, section)
            catalog = {case.rsplit("/", 1)[1] for case in REQUIRED_CASES if case.startswith(section + "/")}
            self.assertEqual(catalog, set(expected))
        self.assertIn("squid-zero/S1", REQUIRED_CASES)
        self.assertIn("susceptibility/S1", REQUIRED_CASES)
        self.assertEqual(sum(case.startswith("identity/") for case in REQUIRED_CASES), 24)

    def test_incomplete_campaign_and_revision_mismatch_are_explicit(self):
        result = self.audit(self.record("squid-zero/S1"))
        self.assertEqual(result["counts"]["pass"], 1)
        self.assertEqual(result["cases"]["susceptibility/S1"]["verdict"], "missing")
        self.assertFalse(result["all_cases_have_passing_evidence"])
        with self.assertRaisesRegex(ValueError, "different software revision"):
            self.audit(commit="3" * 40)
        with redirect_stdout(io.StringIO()) as output:
            self.assertEqual(main(["--check-station-evidence", str(self.manifest), COMMIT]), 1)
        self.assertEqual(json.loads(output.getvalue())["counts"]["missing"], len(REQUIRED_CASES))

    def test_later_failure_is_visible_and_valid_retest_can_pass(self):
        case = "holder-reference/H4"
        first = self.record(case, 1)
        failure = self.record(case, 2, verdict="fail")
        result = self.audit(failure, first)
        self.assertEqual(result["cases"][case], {"verdict": "fail", "record_id": "test-2", "reason": "Fixture reason"})
        later = self.record(case, 3)
        self.assertEqual(self.audit(first, later, failure)["cases"][case]["record_id"], "test-3")
        tied = {**later, "record_id": "test-tie", "verdict": "blocked"}
        self.assertEqual(self.audit(later, tied)["cases"][case]["verdict"], "ambiguous")

    def test_station_case_rejects_simulation_even_when_hashes_match(self):
        self.attachment.write_bytes(b'{"simulated": true}')
        result = self.audit(self.record("holder-reference/H4"), self.record("rockmag/K1", 2, mode="software"))
        self.assertEqual(result["cases"]["holder-reference/H4"]["verdict"], "invalid")
        self.assertEqual(result["cases"]["rockmag/K1"]["verdict"], "pass")
        self.assertIn("Offline evidence", result["errors"][0]["error"])

        self.attachment.write_bytes(b'{"schema":"rapidpy.vrm.session_manifest.v2","provenance":{"simulated":true}}')
        result = self.audit(self.record("vrm/V2"))
        self.assertEqual(result["cases"]["vrm/V2"]["verdict"], "invalid")
        self.assertIn("Offline evidence", result["errors"][0]["error"])

    def test_station_case_rejects_an_attachment_from_another_runtime_revision(self):
        self.attachment.write_text(json.dumps({"simulated": False, "software_version": "4.0.0+" + "3" * 40}),
                                   encoding="utf-8")
        result = self.audit(self.record("holder-reference/H4"))
        self.assertEqual(result["cases"]["holder-reference/H4"]["verdict"], "invalid")
        self.assertIn("different candidate revision", result["errors"][0]["error"])

    def test_hash_mismatch_and_paths_outside_the_evidence_root_are_invalid(self):
        case = "holder-reference/H4"
        original = self.record(case)
        self.attachment.write_bytes(b'{"simulated": false, "changed": true}')
        with self.subTest(issue="modified"):
            self.assertIn("SHA-256 differs", self.audit(original)["errors"][0]["error"])
        for path in ("../trace.json", "E:/trace.json", "other\\trace.json"):
            with self.subTest(path=path):
                record = self.record(case)
                record["artifact_paths"][0]["evidence_path"] = path
                result = self.audit(record)
                self.assertEqual(result["cases"][case]["verdict"], "invalid")
                self.assertFalse(result["all_cases_have_passing_evidence"])

    def test_all_cases_need_passing_evidence_and_unique_records(self):
        records = [self.record(case, number) for number, case in enumerate(REQUIRED_CASES, 1)]
        result = self.audit(*records)
        self.assertTrue(result["all_cases_have_passing_evidence"])
        self.assertFalse(self.audit(*records, records[0])["all_cases_have_passing_evidence"])
