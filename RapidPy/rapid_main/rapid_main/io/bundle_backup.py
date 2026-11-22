"""Copy an accepted measurement to a separately configured backup folder."""
from dataclasses import astuple
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
from zipfile import ZipFile

from rapid_main.io.bundle_history import (ARCHIVE_MANIFEST, HISTORY_DIRECTORY, _archive_io_path,
                                         open_bundle_archive, publish_bundle_archive)
from rapid_main.io.bundle_publication import _digest
from rapid_main.io.measurement_bundle import SIMULATION_MARKER_FILE, STEP_LEDGER_FILE, validate_measurement_output
from rapid_main.specimen_paths import contained_path


INDEX_FILE = "artifact_index.json"


def backup_measurement(source, destination, *, expected_index_sha256=None):
    """Publish an archive only after checking its accepted artifact index."""
    source = _archive_io_path(source)
    index_path = contained_path(source, INDEX_FILE, output=True)
    if expected_index_sha256 is not None and not re.fullmatch(r"[0-9a-f]{64}", expected_index_sha256):
        raise ValueError("Expected artifact index SHA-256 is invalid.")
    try:
        index_bytes = index_path.read_bytes()
    except FileNotFoundError:
        if expected_index_sha256 is None:
            raise
        index_bytes = b""
    index_digest = hashlib.sha256(index_bytes).hexdigest()
    if expected_index_sha256 is not None and index_digest != expected_index_sha256:
        history = _find_prior_archive(source, expected_index_sha256)
        with open_bundle_archive(history) as archive:
            original_index = archive.read(INDEX_FILE)
            if hashlib.sha256(original_index).hexdigest() != expected_index_sha256:
                raise ValueError("The archived artifact index does not match the queued publication.")
            return _backup_from_index(source, destination, original_index, expected_index_sha256, archive=archive)
    return _backup_from_index(source, destination, index_bytes, index_digest)


def _find_prior_archive(source: Path, expected_index_sha256: str) -> Path:
    history = contained_path(source, HISTORY_DIRECTORY, output=True)
    if history.is_dir():
        for folder in sorted(history.iterdir()):
            if not re.fullmatch(r"[0-9a-f]{32}", folder.name):
                continue
            try:
                manifest = contained_path(history, f"{folder.name}/{ARCHIVE_MANIFEST}", output=True)
                files = json.loads(manifest.read_bytes()).get("files")
            except (OSError, AttributeError, ValueError):
                continue
            if isinstance(files, dict) and files.get(INDEX_FILE) == expected_index_sha256:
                return folder
    raise ValueError("The measurement publication changed before backup started. Its original archive was not found.")


def _backup_from_index(source: Path, destination, index_bytes: bytes, index_digest: str,
                       *, archive: ZipFile | None = None):
    archive_names = set(archive.namelist()) if archive is not None else set()

    def member_digest(name):
        if archive is None:
            return _digest(contained_path(source, name, output=True))
        if name not in archive_names:
            return None
        with archive.open(name) as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()

    index = json.loads(index_bytes)
    if (not isinstance(index, dict) or index.get("schema") != "rapidpy.measurement.artifact_index.v1"
            or index.get("aborted") is not False or index.get("diagnostics_pending")
            or type(index.get("simulated")) is not bool or not isinstance(index.get("sample"), str)
            or not isinstance(index.get("artifacts"), list)):
        raise ValueError("The folder has no finalized measurement publication to back up.")

    specimen = index["sample"]
    scientific = astuple(validate_measurement_output(source, specimen))
    required = {path.relative_to(source).as_posix() for path in scientific} | {"provenance.json"}
    files = {INDEX_FILE: index_digest}
    omitted = []
    seen = set()
    for entry in index["artifacts"]:
        name = entry.get("relative_path") if isinstance(entry, dict) else None
        if not isinstance(name, str) or not name or name == INDEX_FILE or name in seen:
            raise ValueError("The measurement index contains an invalid artifact path.")
        path = PurePosixPath(name)
        if path.is_absolute() or path.as_posix() != name or "\\" in name or ".." in path.parts:
            raise ValueError("The measurement index contains an invalid artifact path.")
        seen.add(name)
        if entry.get("exists") is not True:
            if entry.get("required") is True:
                raise ValueError(f"The measurement index is missing a required artifact: {name}")
            continue
        digest = member_digest(name)
        if digest is None or digest != entry.get("sha256"):
            # Older desktop builds did not refresh this optional digest after plotting.
            if name == "quicklook.json" and entry.get("required") is False:
                omitted.append(name)
                continue
            raise ValueError(f"The published artifact changed or is missing: {name}")
        files[name] = digest
    if not required.issubset(files):
        raise ValueError("The measurement index does not cover all scientific outputs.")
    provenance = json.loads(archive.read("provenance.json") if archive is not None else
                            contained_path(source, "provenance.json", output=True).read_bytes())
    if (not isinstance(provenance, dict) or provenance.get("specimen") != specimen
            or provenance.get("simulated") is not index["simulated"]):
        raise ValueError("The measurement provenance and artifact index disagree.")
    markers = (STEP_LEDGER_FILE, SIMULATION_MARKER_FILE) if index["simulated"] else (STEP_LEDGER_FILE,)
    for name in markers:
        digest = member_digest(name)
        if digest is None:
            raise ValueError(f"The measurement publication is missing {name}.")
        files[name] = digest

    location = Path(destination)
    if not location.is_absolute() or not _archive_io_path(location).is_dir():
        raise ValueError("The backup folder must be an available absolute directory.")
    root = _archive_io_path(location).resolve()
    namespace = f"RapidPy backups/{'SIMULATED' if index['simulated'] else 'Measurements'}"
    target = contained_path(root, namespace, output=True)
    if target.is_relative_to(source) or source.is_relative_to(target):
        raise ValueError("Choose a backup folder outside the specimen publication.")
    saved, archive_id = publish_bundle_archive(archive if archive is not None else source, specimen, files, target)
    return {"archive": str(location / namespace / saved.name), "archive_id": archive_id,
            "index_sha256": index_digest, "omitted": omitted}
