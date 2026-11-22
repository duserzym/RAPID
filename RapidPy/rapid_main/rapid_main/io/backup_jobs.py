"""Durable pending requests for an accepted measurement backup."""
from dataclasses import asdict, dataclass
import hashlib
import json
import os
from pathlib import Path
import re

from rapid_main.config import _atomic_write_json
from rapid_main.io.bundle_history import _archive_io_path
from rapid_main.io.bundle_publication import _sync_directory
from rapid_main.specimen_paths import contained_path


_SCHEMA = "rapidpy.measurement.backup_job.v1"
_DIGEST = re.compile(r"[0-9a-f]{64}")


@dataclass(frozen=True, slots=True)
class BackupJob:
    sample: str
    source: str
    destination: str
    index_sha256: str

    def __post_init__(self) -> None:
        if not isinstance(self.sample, str) or not self.sample.strip():
            raise ValueError("A pending backup requires a specimen name.")
        if any(not isinstance(value, str) or not Path(value).is_absolute()
               for value in (self.source, self.destination)):
            raise ValueError("The backup source and destination must be absolute folders.")
        if not isinstance(self.index_sha256, str) or not _DIGEST.fullmatch(self.index_sha256):
            raise ValueError("A pending backup requires its original artifact index SHA-256.")

    @property
    def key(self) -> str:
        identity = [os.path.normcase(os.path.normpath(path)) for path in (self.source, self.destination)]
        return hashlib.sha256(json.dumps([*identity, self.index_sha256]).encode("utf-8")).hexdigest()

    def matches_destination(self, destination: str) -> bool:
        return os.path.normcase(os.path.normpath(self.destination)) == os.path.normcase(os.path.normpath(destination))


class BackupJobs:
    def __init__(self, directory: Path) -> None:
        self._profile = _archive_io_path(directory.parent.resolve())
        self._name = directory.name

    @property
    def directory(self) -> Path:
        return contained_path(self._profile, self._name, output=True)

    def save(self, job: BackupJob) -> None:
        path = contained_path(self.directory, f"{job.key}.json", output=True)
        _atomic_write_json(path, {"schema": _SCHEMA, **asdict(job)})
        _sync_directory(path.parent)

    def complete(self, job: BackupJob) -> None:
        path = contained_path(self.directory, f"{job.key}.json", output=True)
        path.unlink(missing_ok=True)
        _sync_directory(path.parent)

    def pending(self) -> tuple[list[BackupJob], list[str]]:
        directory = self.directory
        jobs, unreadable = [], []
        if not directory.exists():
            return jobs, unreadable
        for path in sorted(directory.iterdir()):
            if path.suffix != ".json" or not _DIGEST.fullmatch(path.stem):
                continue
            try:
                payload = json.loads(contained_path(directory, path.name, output=True).read_bytes())
                if not isinstance(payload, dict) or payload.pop("schema", None) != _SCHEMA:
                    raise ValueError("Invalid backup record schema.")
                job = BackupJob(**payload)
                if job.key != path.stem:
                    raise ValueError("Backup record identity differs from its filename.")
                jobs.append(job)
            except (OSError, TypeError, ValueError):
                unreadable.append(path.name)
        return jobs, unreadable
