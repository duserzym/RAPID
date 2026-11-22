"""Audit operator-recorded station evidence against the bench procedure."""
from collections import defaultdict
from datetime import datetime, timedelta
import hashlib
import json
from pathlib import Path, PurePosixPath, PureWindowsPath
import re

from rapid_main.specimen_paths import contained_path


SCHEMA = "rapidpy.station_evidence.v1"
_REVISION = re.compile(r"(?:[0-9a-f]{40}|[0-9a-f]{64})")
_DIGEST = re.compile(r"[0-9a-f]{64}")
_CONFIG_HASH = re.compile(r"[0-9a-f]{16}")
_IDENTITY_DEVICES = (
    "squid", "changer-lift-turning", "vacuum", "af", "irm-arm", "dac-mcc", "thermal", "susceptibility",
)

# Section names qualify the reused S1–S4 identifiers in the written procedure.
PROCEDURE_SECTIONS = (
    ("preconditions", "P", 7), ("ports", "C", 3), ("identity", "R", 3), ("motion", "M", 9),
    ("squid-zero", "S", 4), ("discontinuity", "D", 5), ("holder-reference", "H", 11),
    ("restart", "I", 4), ("vrm", "V", 8), ("rockmag", "K", 7), ("susceptibility", "S", 11),
)
REQUIRED_CASES = tuple(
    f"{section}/{device}/{prefix}{number}" if section == "identity" else f"{section}/{prefix}{number}"
    for section, prefix, count in PROCEDURE_SECTIONS
    for device in (_IDENTITY_DEVICES if section == "identity" else (None,))
    for number in range(1, count + 1)
)
_SOFTWARE_CASES = {"discontinuity/D1", "discontinuity/D5", "rockmag/K1", "rockmag/K2"}
_OFFLINE_SCHEMAS = {
    "rapidpy.portable_verification.v2", "rapidpy.release_smoke.v1", "rapidpy.startup_environment.v1",
    "rapidpy.display_check.v1", "rapidpy.performance_check.v1",
}


def _read_json(path: Path) -> object:
    with path.open("r", encoding="utf-8-sig") as stream:
        return json.load(stream)


def _required_text(payload: dict, key: str) -> str:
    value = payload.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{key} must contain text.")
    return value.strip()


def _utc_time(value: str) -> datetime:
    timestamp = datetime.fromisoformat(value)
    if timestamp.utcoffset() != timedelta(0):
        raise ValueError("Record timestamps must use UTC with an explicit offset.")
    return timestamp


def _software_case(case: str) -> bool:
    return case.startswith("preconditions/") or case in _SOFTWARE_CASES


def _check_version(version: object, candidate: str) -> None:
    if isinstance(version, str):
        suffix = version.rpartition("+")[2].lower()
        if _REVISION.fullmatch(suffix) and suffix != candidate:
            raise ValueError("Recorded software version identifies a different candidate revision.")


def _check_attachments(root: Path, record: dict, candidate: str, *, station_test: bool) -> None:
    artifacts = record.get("artifact_paths")
    if not isinstance(artifacts, list) or (record["verdict"] == "pass" and not artifacts):
        raise ValueError("Attach evidence files to a passing case.")
    paths = set()
    for artifact in artifacts:
        if not isinstance(artifact, dict):
            raise ValueError("Each artifact needs an original path, evidence path, and SHA-256.")
        original = _required_text(artifact, "original_path")
        filename = _required_text(artifact, "evidence_path")
        expected = _required_text(artifact, "sha256").lower()
        relative = PurePosixPath(filename)
        if not (PureWindowsPath(original).is_absolute() or PurePosixPath(original).is_absolute()):
            raise ValueError("The original artifact path must be absolute.")
        if (relative.as_posix() != filename or relative.is_absolute() or "\\" in filename
                or ".." in relative.parts or filename.casefold() in paths or not _DIGEST.fullmatch(expected)):
            raise ValueError(f"Invalid or duplicate evidence attachment: {filename}")
        paths.add(filename.casefold())
        path = contained_path(root, filename, output=True)
        if not path.is_file():
            raise ValueError(f"Evidence attachment is missing: {filename}")
        if station_test:
            original_parts = PureWindowsPath(original).parts
            if any(part.casefold() == "simulated" for part in (*relative.parts, *original_parts)):
                raise ValueError(f"Simulation evidence cannot qualify a station case: {filename}")
        with path.open("rb") as stream:
            if path.suffix.lower() == ".json":
                data = stream.read()
                actual = hashlib.sha256(data).hexdigest()
                try:
                    metadata = json.loads(data)
                except ValueError as exc:
                    raise ValueError(f"Evidence attachment is not valid JSON: {filename}") from exc
                if isinstance(metadata, dict):
                    provenance = metadata.get("provenance")
                    if station_test and (metadata.get("simulated") is True
                            or (isinstance(provenance, dict) and provenance.get("simulated") is True)
                            or str(metadata.get("schema")) in _OFFLINE_SCHEMAS):
                        raise ValueError(f"Offline evidence cannot qualify a station case: {filename}")
                    _check_version(metadata.get("software_version"), candidate)
                    if metadata.get("schema") == "rapidpy.portable_verification.v2":
                        inventory = metadata.get("inventory")
                        if not isinstance(inventory, dict) or inventory.get("source_revision") != candidate:
                            raise ValueError(f"Portable verification identifies a different candidate: {filename}")
            else:
                actual = hashlib.file_digest(stream, "sha256").hexdigest()
        if actual != expected:
            raise ValueError(f"Evidence SHA-256 differs: {filename}")


def _check_record(root: Path, record: dict, station: str, candidate: str, case: str) -> datetime:
    for field in ("operator", "machine", "software_version", "device_identity", "expected_result",
                  "observed_result", "signature"):
        _required_text(record, field)
    if _required_text(record, "station") != station or _required_text(record, "software_commit").lower() != candidate:
        raise ValueError("Record station or software revision differs from the evidence campaign.")
    _check_version(record["software_version"], candidate)
    if not _CONFIG_HASH.fullmatch(_required_text(record, "config_hash").lower()):
        raise ValueError("config_hash must be the recorded RAPID configuration fingerprint.")
    timing = record.get("datetime_iso")
    if not isinstance(timing, dict):
        raise ValueError("datetime_iso must contain start and end timestamps.")
    start = _utc_time(_required_text(timing, "start"))
    end = _utc_time(_required_text(timing, "end"))
    if end < start:
        raise ValueError("Record end time precedes its start.")
    for field in ("commands", "raw_replies"):
        values = record.get(field)
        if not isinstance(values, list) or any(not isinstance(item, str) for item in values):
            raise ValueError(f"{field} must be an ordered list of verbatim strings.")
    verdict = record.get("verdict")
    if not isinstance(verdict, str) or verdict not in {"pass", "fail", "blocked"}:
        raise ValueError("verdict must be pass, fail, or blocked.")
    if verdict != "pass":
        _required_text(record, "reason")
    mode = _required_text(record, "evidence_mode")
    if mode not in {"station", "software"} or (verdict == "pass" and not _software_case(case) and mode != "station"):
        raise ValueError("A passing station case requires station evidence.")
    _check_attachments(root, record, candidate, station_test=mode == "station" and not _software_case(case))
    return end


def audit_station_evidence(manifest_path: str | Path, candidate: str) -> dict:
    """Report evidence integrity and recorded outcomes for an explicit release revision."""
    candidate = candidate.lower()
    if not _REVISION.fullmatch(candidate):
        raise ValueError("Provide the full candidate commit hash.")
    manifest_path = Path(manifest_path).resolve(strict=True)
    manifest = _read_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("schema") != SCHEMA:
        raise ValueError(f"The station evidence manifest must use {SCHEMA}.")
    station = _required_text(manifest, "station")
    if _required_text(manifest, "software_commit").lower() != candidate:
        raise ValueError("The evidence campaign is for a different software revision.")
    records = manifest.get("records")
    if not isinstance(records, list):
        raise ValueError("The evidence campaign must contain a records list.")

    attempts = defaultdict(list)
    errors = []
    ids = set()
    invalid_cases = set()
    for position, record in enumerate(records, 1):
        case = None
        try:
            if not isinstance(record, dict):
                raise ValueError("A record must be an object.")
            case = _required_text(record, "case")
            if case not in REQUIRED_CASES:
                raise ValueError(f"Unknown procedure case: {case}")
            record_id = _required_text(record, "record_id")
            if record_id in ids:
                raise ValueError(f"Duplicate record ID: {record_id}")
            ids.add(record_id)
            end = _check_record(manifest_path.parent, record, station, candidate, case)
            attempts[case].append((end, record_id, record["verdict"], str(record.get("reason", ""))))
        except (OSError, ValueError, RuntimeError) as exc:
            errors.append({"record_number": position, "case": case, "error": str(exc)})
            if case in REQUIRED_CASES:
                invalid_cases.add(case)

    cases = {}
    for case in REQUIRED_CASES:
        history = sorted(attempts[case], reverse=True)
        if case in invalid_cases:
            cases[case] = {"verdict": "invalid"}
        elif not history:
            cases[case] = {"verdict": "missing"}
        elif len(history) > 1 and history[0][0] == history[1][0]:
            cases[case] = {"verdict": "ambiguous", "reason": "Latest attempts share an end time."}
        else:
            _, record_id, verdict, reason = history[0]
            cases[case] = {"verdict": verdict, "record_id": record_id}
            if reason:
                cases[case]["reason"] = reason
    counts = {verdict: sum(item["verdict"] == verdict for item in cases.values())
              for verdict in ("pass", "fail", "blocked", "invalid", "ambiguous", "missing")}
    return {
        "schema": "rapidpy.station_evidence_audit.v1", "station": station, "software_commit": candidate,
        "scope": "Evidence integrity and recorded operator verdicts", "errors": errors, "cases": cases,
        "counts": counts, "all_cases_have_passing_evidence": not errors and counts["pass"] == len(cases),
    }
