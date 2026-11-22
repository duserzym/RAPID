"""Preserve a specimen's earlier published results before replacing them."""
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
from tempfile import TemporaryDirectory
from typing import Iterator
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from rapidpy_common.hardware_safety import _canonical, _replace_snapshot
from rapidpy_common.paths import absolute_io_path as _archive_io_path
from rapid_main.io.bundle_publication import _digest, _sync_directory
from rapid_main.specimen_paths import contained_path


HISTORY_DIRECTORY = ".rapidpy-history"
ARCHIVE_MANIFEST = "archive_manifest.json"
ARCHIVE_PAYLOAD = "bundle.zip"
_SCHEMA = "rapidpy.measurement.bundle_archive.v1"
_SHA256 = re.compile(r"[0-9a-f]{64}")
_EVIDENCE_DIRECTORIES = {
    ".rapidpy-acquisitions", ".rapidpy-pairs", "af_treatments",
    "pulse_treatments", "susceptibility_acquisitions", ".rapidpy-attempts",
}


def is_bundle_archive(directory: str | Path) -> bool:
    return (_archive_io_path(directory) / ARCHIVE_MANIFEST).exists()


def _archive_identity(payload):
    return hashlib.sha256(_canonical(payload)).hexdigest()


def _indexed_evidence(directory):
    index = contained_path(directory, "artifact_index.json", output=True)
    if not index.is_file():
        return
    try:
        payload = json.loads(index.read_text(encoding="utf-8"))
    except ValueError:
        return
    if not isinstance(payload, dict) or payload.get("schema") != "rapidpy.measurement.artifact_index.v1":
        return
    entries = payload.get("artifacts")
    if not isinstance(entries, list):
        return
    for entry in entries:
        name = entry.get("relative_path") if isinstance(entry, dict) and entry.get("exists") is True else None
        if not isinstance(name, str):
            continue
        parts = name.split("/")
        if (len(parts) > 1 and parts[0] in _EVIDENCE_DIRECTORIES
                and not any(not part or part in (".", "..") or "\\" in part for part in parts)):
            yield contained_path(directory, name, output=True)


def archive_previous_bundle(directory, specimen, scientific_paths, mutable_paths, staging):
    """Archive the source bytes under the writer's publication lease."""
    if not any(path.is_file() for path in scientific_paths):
        return None, {path.relative_to(directory).as_posix(): None for path in scientific_paths}
    candidates = set((*scientific_paths, *mutable_paths, *_indexed_evidence(directory)))
    previous = {path.relative_to(directory).as_posix(): _digest(path) for path in sorted(candidates)}
    files = {name: digest for name, digest in previous.items() if digest is not None}
    root = contained_path(_archive_io_path(directory), HISTORY_DIRECTORY, output=True)
    archive, archive_id = publish_bundle_archive(directory, specimen, files, root, staging=staging)
    manifest = contained_path(archive, ARCHIVE_MANIFEST, output=True)
    digest = _digest(manifest)
    if digest is None:
        raise OSError("The prior measurement archive manifest is missing after publication.")
    return {"relative_path": f"{HISTORY_DIRECTORY}/{archive.name}/{ARCHIVE_MANIFEST}",
            "archive_id": archive_id, "sha256": digest}, previous


def _write_bundle_archive(source, staged, payload, archive_id):
    zip_path = contained_path(staged, ARCHIVE_PAYLOAD, output=True)
    with zip_path.open("xb") as stream:
        with ZipFile(stream, "w", compression=ZIP_DEFLATED, compresslevel=3) as data:
                for name in payload["files"]:
                    if isinstance(source, ZipFile):
                        with source.open(name) as member, data.open(name, "w", force_zip64=True) as target:
                            shutil.copyfileobj(member, target)
                    else:
                        data.write(contained_path(source, name, output=True), name)
        stream.flush()
        os.fsync(stream.fileno())
    manifest = contained_path(staged, ARCHIVE_MANIFEST, output=True)
    with manifest.open("x", encoding="utf-8") as stream:
        json.dump({**payload, "archive_id": archive_id}, stream, indent=2, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())
    with open_bundle_archive(staged, expected_id=archive_id):
        pass
    _sync_directory(staged)


def publish_bundle_archive(source, specimen, files, root, *, staging=None):
    """Write a verified immutable archive. Stage on the destination filesystem."""
    payload = {"schema": _SCHEMA, "specimen": specimen, "files": files}
    archive_id = _archive_identity(payload)
    archive = contained_path(root, archive_id[:32], output=True)
    if archive.exists():
        with open_bundle_archive(archive, expected_id=archive_id):
            pass
        return archive, archive_id

    root.mkdir(parents=True, exist_ok=True)
    _sync_directory(root.parent)
    staging_root = _archive_io_path(staging) if staging is not None else root
    with TemporaryDirectory(prefix=".pending-", dir=staging_root) as temporary:
        staged = contained_path(temporary, "archive", output=True)
        staged.mkdir()
        _write_bundle_archive(source, staged, payload, archive_id)
        try:
            _replace_snapshot(staged, archive)
        except OSError:
            if not archive.is_dir():
                raise
            with open_bundle_archive(archive, expected_id=archive_id):
                pass
        _sync_directory(root)
    return archive, archive_id


@contextmanager
def open_bundle_archive(directory: str | Path, *, expected_id: str | None = None) -> Iterator[ZipFile]:
    """Verify every archived file before yielding the open ZIP for reading."""
    directory = _archive_io_path(directory)
    path = contained_path(directory, ARCHIVE_MANIFEST, output=True)
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ValueError("The measurement archive manifest cannot be read.") from exc
    if not isinstance(manifest, dict) or set(manifest) != {"schema", "specimen", "files", "archive_id"}:
        raise ValueError("Invalid measurement archive manifest.")
    archive_id = manifest.pop("archive_id")
    files = manifest["files"]
    if (manifest["schema"] != _SCHEMA or not isinstance(manifest["specimen"], str)
            or not isinstance(files, dict) or not files or _archive_identity(manifest) != archive_id
            or (expected_id is not None and archive_id != expected_id)):
        raise ValueError("Invalid measurement archive identity.")
    for name, digest in files.items():
        parts = PurePosixPath(name)
        if (parts.is_absolute() or parts.as_posix() != name or "\\" in name or ".." in parts.parts
                or not isinstance(digest, str) or not _SHA256.fullmatch(digest)):
            raise ValueError("Invalid measurement archive file entry.")
    zip_path = contained_path(directory, ARCHIVE_PAYLOAD, output=True)
    with zip_path.open("rb") as stream:
        try:
            with ZipFile(stream) as data:
                if sorted(data.namelist()) != sorted(files):
                    raise ValueError("The measurement archive file list differs from its manifest.")
                for name, digest in files.items():
                    with data.open(name) as member:
                        if hashlib.file_digest(member, "sha256").hexdigest() != digest:
                            raise ValueError(f"Archived measurement file changed: {name}")
                yield data
        except BadZipFile as exc:
            raise ValueError("The measurement archive ZIP is damaged.") from exc
