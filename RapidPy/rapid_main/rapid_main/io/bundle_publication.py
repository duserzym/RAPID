"""Recover a measurement-file replacement interrupted before durable commit."""

import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyStore, _canonical, _replace_snapshot
from rapid_main.specimen_paths import contained_path


JOURNAL = '.rapidpy-publication'
_SEAL = re.compile(r'(?:\.rapidpy-pairs/[0-9a-f]{64}/down/)?\.rapidpy-acquisitions/[0-9a-f]{64}/accepted\.json')
_SHA256 = re.compile(r'[0-9a-f]{64}')
_JOURNAL_FILE = re.compile(r'manifest\.json|[0-9]+\.previous|manifest-[0-9a-f]{32}\.tmp|restore-[0-9a-f]{32}')


def _digest(path):
    if not path.exists():
        return None
    if not path.is_file():
        raise ValueError(f'Publication requires a regular file at {path.name}.')
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def _sync_directory(path):
    if os.name != 'nt':
        descriptor = os.open(path, os.O_RDONLY)
        try:
            os.fsync(descriptor)
        finally:
            os.close(descriptor)


def _copy(source, target):
    with source.open('rb') as original, target.open('xb') as backup:
        shutil.copyfileobj(original, backup)
        backup.flush()
        os.fsync(backup.fileno())


class BundlePublication:
    """Journal old bytes before replacing any of the public filenames."""

    def __init__(self, directory: Path, specimen: str, targets):
        self.directory = directory
        self.specimen = specimen
        self.targets = {path.relative_to(directory).as_posix() for path in targets}

    @property
    def journal(self):
        return contained_path(self.directory, JOURNAL, output=True)

    def lease(self):
        lock = contained_path(self.directory, f'{JOURNAL}.operation.lock', output=True)
        return HardwareSafetyStore(lock.with_name(JOURNAL)).operation_lease()

    def _target(self, name):
        if name not in self.targets and not _SEAL.fullmatch(name):
            raise ValueError('Unexpected destination in the measurement publication journal.')
        return contained_path(self.directory, name, output=True)

    def _write_manifest(self, manifest):
        journal = self.journal
        temp = contained_path(journal, f'manifest-{uuid.uuid4().hex}.tmp', output=True)
        with temp.open('x', encoding='utf-8') as stream:
            json.dump(dict(payload=manifest, sha256=hashlib.sha256(_canonical(manifest)).hexdigest()),
                      stream, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        _replace_snapshot(temp, contained_path(journal, 'manifest.json', output=True))
        _sync_directory(journal)

    def _read_manifest(self):
        path = contained_path(self.journal, 'manifest.json', output=True)
        if not path.exists():
            return None
        envelope = json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(envelope, dict) or set(envelope) != {'payload', 'sha256'}:
            raise ValueError('Invalid measurement publication journal.')
        manifest = envelope['payload']
        if (not isinstance(manifest, dict)
                or envelope['sha256'] != hashlib.sha256(_canonical(manifest)).hexdigest()):
            raise ValueError('Measurement publication journal checksum mismatch.')
        if (manifest.get('schema') != 'rapidpy.bundle_publication.v1'
                or manifest.get('specimen') != self.specimen
                or manifest.get('status') not in {'prepared', 'committed', 'rolled_back'}):
            raise ValueError('Open the original specimen to recover its measurement publication.')
        names = set()
        for name, old, new in manifest['files']:
            if (not isinstance(name, str) or name in names or not isinstance(new, str) or not _SHA256.fullmatch(new)
                    or (old is not None and (not isinstance(old, str) or not _SHA256.fullmatch(old)))):
                raise ValueError('Invalid measurement publication journal.')
            self._target(name)
            names.add(name)
        if not names:
            raise ValueError('Empty measurement publication journal.')
        return manifest

    def recover(self):
        """Restore the accepted files, retaining evidence if any conflicts exist."""
        journal = self.journal
        if not journal.exists():
            return None
        if not journal.is_dir() or any(not _JOURNAL_FILE.fullmatch(path.name)
                or not contained_path(journal, path.name, output=True).is_file() for path in journal.iterdir()):
            raise ValueError('The publication journal contains unexpected files. Preserve it for review.')
        manifest = self._read_manifest()
        status = manifest['status'] if manifest is not None else None
        if status == 'prepared':
            restore = []
            for index, (name, old, new) in enumerate(manifest['files']):
                target = self._target(name)
                current = _digest(target)
                if current == old:
                    continue
                if current != new:
                    raise ValueError(f'Output changed during measurement publication: {name}')
                backup = contained_path(journal, f'{index}.previous', output=True)
                if old is not None and _digest(backup) != old:
                    raise ValueError(f'Measurement recovery backup does not match {name}.')
                restore.append((name, backup, old, new))
            for name, backup, old, new in restore:
                target = self._target(name)
                if _digest(target) != new:
                    raise ValueError(f'Output changed during measurement recovery: {name}')
                if old is None:
                    target.unlink()
                else:
                    temp = contained_path(journal, f'restore-{uuid.uuid4().hex}', output=True)
                    _copy(backup, temp)
                    _replace_snapshot(temp, target)
                _sync_directory(target.parent)
            manifest['status'] = status = 'rolled_back'
            self._write_manifest(manifest)
        shutil.rmtree(journal)
        _sync_directory(self.directory)
        return status

    def publish(self, files, *, archived_previous=None):
        """Publish already prepared files under the caller's exclusive lease."""
        if not files or len({path for _, path in files}) != len(files):
            raise ValueError('Measurement publication needs distinct output files.')
        journal = self.journal
        journal.mkdir()
        _sync_directory(self.directory)
        entries = []
        try:
            for index, (staged, path) in enumerate(files):
                name = path.relative_to(self.directory).as_posix()
                target = self._target(name)
                old = _digest(target)
                if archived_previous is not None and name in archived_previous and old != archived_previous[name]:
                    raise ValueError(f'Measurement output changed after its archive was written: {name}')
                if _SEAL.fullmatch(name) and old is not None:
                    raise ValueError('This acquisition already has an acceptance seal.')
                new = _digest(staged)
                with staged.open('r+b') as stream:
                    os.fsync(stream.fileno())
                if old is not None:
                    backup = contained_path(journal, f'{index}.previous', output=True)
                    _copy(target, backup)
                    if _digest(backup) != old:
                        raise ValueError(f'Output changed while recording its backup: {name}')
                entries.append((name, old, new))
            manifest = dict(schema='rapidpy.bundle_publication.v1', specimen=self.specimen,
                            status='prepared', files=entries)
            self._write_manifest(manifest)
            for (staged, _), (name, old, new) in zip(files, entries):
                target = self._target(name)
                if _digest(target) != old or _digest(staged) != new:
                    raise ValueError(f'Measurement files changed before replacement: {name}')
                target.parent.mkdir(parents=True, exist_ok=True)
                _replace_snapshot(staged, target)
                _sync_directory(target.parent)
            manifest['status'] = 'committed'
            self._write_manifest(manifest)
        except Exception as failure:
            try:
                if self.recover() == 'committed':
                    return
            except Exception as recovery:
                raise OSError(f'Measurement publication failed: {failure}. Recovery is pending: {recovery}') from failure
            raise
        # Once committed, cleanup can be retried by the next owner.
        try:
            self.recover()
        except OSError:
            pass
