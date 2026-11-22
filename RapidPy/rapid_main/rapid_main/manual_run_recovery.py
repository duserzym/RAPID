"""Durable restart guard for an unplanned interruption in Live Measure."""
from dataclasses import astuple, dataclass
from datetime import datetime, timezone
from pathlib import Path
import hashlib
import json
import re
import uuid

from rapid_main.io.bundle_publication import _digest, _sync_directory
from rapid_main.io.measurement_bundle import STAGING_PREFIX, _publication_for, _write_json_atomic, validate_measurement_output
from rapid_main.measurement_checkpoint import clear_accepted_checkpoint, has_finished_checkpoint
from rapid_main.measurement_progress import MeasurementProgress
from rapid_main.specimen_paths import contained_path
from rapidpy_common.hardware_safety import _replace_snapshot
from rapidpy_common.paths import absolute_io_path


DIRECTORY = '.rapidpy-manual-recovery'
_SCHEMA = 'rapidpy.measurement.manual_interruption.v1'


@dataclass(frozen=True)
class InterruptedRun:
    directory: Path
    progress: MeasurementProgress
    staging: str
    digest: str

    def describe(self) -> str:
        identity = self.progress.identity
        steps = []
        for index, label in enumerate(identity['labels']):
            started = self.progress._path(f'{index:05d}-started.json')
            captured = self.progress._path(f'{index:05d}-captured.json')
            if not started.exists() and not captured.exists():
                continue
            try:
                self.progress.read_step(index)
            except (OSError, ValueError, KeyError, TypeError):
                detail = 'Acquisition incomplete or unreadable. Treatment may already have occurred.'
            else:
                detail = 'Readings captured. Review the original run before using them.'
            steps.append(f'{index + 1}. {label} — {detail}')
        if not steps:
            steps.append('The process stopped before a treatment intent was recorded.')
        return (f"Specimen {identity['specimen']}\nRun {identity['run_id']}\n\n"
                + '\n'.join(steps))


def _marker(directory):
    return contained_path(absolute_io_path(directory), f'{DIRECTORY}/pending.json', output=True)


def has_interrupted_run(directory):
    return _marker(directory).exists()


def read_interrupted_run(directory, specimen) -> InterruptedRun | None:
    path = _marker(directory)
    try:
        data = path.read_bytes()
    except FileNotFoundError:
        return None
    record = json.loads(data)
    if (not isinstance(record, dict) or set(record) != {'schema', 'identity_sha256', 'staging', 'attempt_id'}
            or record['schema'] != _SCHEMA or not isinstance(record['staging'], str)
            or not isinstance(record['attempt_id'], str) or not re.fullmatch(r'[0-9a-f]{32}', record['attempt_id'])
            or not re.fullmatch(rf'{re.escape(STAGING_PREFIX)}-[1-9][0-9]*', record['staging'])):
        raise ValueError('The interrupted measurement marker is invalid. Preserve it for review.')
    progress = MeasurementProgress.open_existing(directory, record['identity_sha256'])
    if progress.identity['specimen'] != specimen:
        raise ValueError('Select the original specimen to review its interrupted measurement.')
    return InterruptedRun(Path(directory), progress, record['staging'], hashlib.sha256(data).hexdigest())


def track_manual_run(bundle, progress) -> InterruptedRun:
    """Create the intent while the bundle writer owns the specimen lease."""
    path = _marker(bundle.publish_dir)
    if _digest(path) is not None:
        raise ValueError('Review the interrupted measurement before starting a new run.')
    path.parent.mkdir(parents=True, exist_ok=True)
    _write_json_atomic(path, dict(schema=_SCHEMA, identity_sha256=progress.identity_sha256,
                                 staging=bundle.staging_dir.name, attempt_id=uuid.uuid4().hex))
    _sync_directory(path.parent)
    attempt = read_interrupted_run(bundle.publish_dir, bundle.meta.name)
    if attempt is None:
        raise ValueError('The interrupted measurement marker could not be verified.')
    return attempt


def clear_manual_run(attempt: InterruptedRun) -> None:
    path = _marker(attempt.directory)
    if _digest(path) != attempt.digest:
        raise ValueError('The interrupted measurement changed. Review it again.')
    path.unlink()
    _sync_directory(path.parent)


def _retain_unpublished_output(attempt: InterruptedRun, destination: Path) -> None:
    source = contained_path(absolute_io_path(attempt.directory), attempt.staging, output=True)
    if source.exists():
        if not source.is_dir() or destination.exists():
            raise ValueError('The unpublished staging directory changed. Preserve it for review.')
        _replace_snapshot(source, destination)
        _sync_directory(source.parent)
        _sync_directory(destination.parent)


def _review_directory(attempt: InterruptedRun) -> Path:
    path = contained_path(_marker(attempt.directory).parent, f'reviews/{attempt.digest}', output=True)
    path.mkdir(parents=True, exist_ok=True)
    return path


def acknowledge_interruption(attempt: InterruptedRun, *, operator: str, notes: str) -> Path:
    """Record operator review without issuing commands or accepting partial data."""
    if not operator.strip() or not notes.strip():
        raise ValueError('Identify the reviewing operator and record the specimen and station disposition.')
    specimen = attempt.progress.identity['specimen']
    with _publication_for(attempt.directory, specimen).lease():
        if _digest(_marker(attempt.directory)) != attempt.digest:
            raise ValueError('The interrupted measurement changed. Open it again.')
        _publication_for(attempt.directory, specimen).recover()
        review = _review_directory(attempt)
        _retain_unpublished_output(attempt, contained_path(review, 'unpublished', output=True))
        record = contained_path(review, 'review.json', output=True)
        if record.exists():
            prior = json.loads(record.read_bytes())
            if (not isinstance(prior, dict) or prior.get('marker_sha256') != attempt.digest
                    or prior.get('schema') != 'rapidpy.measurement.manual_review.v1'
                    or prior.get('disposition') != 'operator_reviewed_no_replay'):
                raise ValueError('The retained review differs from the interrupted run.')
        else:
            _write_json_atomic(record, dict(schema='rapidpy.measurement.manual_review.v1',
                marker_sha256=attempt.digest, acquisition=attempt.progress.reference,
                operator=operator.strip(), notes=notes.strip(),
                timestamp_iso=datetime.now(timezone.utc).isoformat(),
                disposition='operator_reviewed_no_replay', summary=attempt.describe()))
        _sync_directory(review)
        clear_manual_run(attempt)
        return record


def _accepted_publication(attempt):
    directory, specimen = attempt.directory, attempt.progress.identity['specimen']
    index_path = contained_path(absolute_io_path(directory), 'artifact_index.json', output=True)
    try:
        index = json.loads(index_path.read_bytes())
    except (FileNotFoundError, ValueError):
        return False
    reference = index.get('acquisition_progress') if isinstance(index, dict) else None
    if (not isinstance(reference, dict) or index.get('aborted') is not False or index.get('sample') != specimen
            or index.get('simulated') is not attempt.progress.identity['simulated']
            or reference.get('identity_sha256') != attempt.progress.identity_sha256):
        return False
    paths = (*astuple(validate_measurement_output(directory, specimen)),
             contained_path(directory, 'provenance.json', output=True))
    entries = index.get('artifacts')
    if not isinstance(entries, list):
        return False
    artifacts = {entry.get('relative_path'): entry.get('sha256') for entry in entries if isinstance(entry, dict)
                 and isinstance(entry.get('relative_path'), str) and entry.get('required') is True
                 and entry.get('exists') is True}
    if any(not artifacts.get(path.relative_to(directory).as_posix())
           or _digest(absolute_io_path(path)) != artifacts[path.relative_to(directory).as_posix()] for path in paths):
        return False
    try:
        attempt.progress.accepted_steps()
    except (OSError, ValueError, KeyError, TypeError):
        return False
    return True


def reconcile_manual_run(directory: Path, specimen: str) -> InterruptedRun | None:
    """Clear leftover intents for verified publications and planned stops."""
    attempt = read_interrupted_run(directory, specimen)
    if attempt is None:
        return None
    publication = _publication_for(directory, specimen)
    with publication.lease():
        attempt = read_interrupted_run(directory, specimen)
        if attempt is None:
            return None
        publication.recover()
        accepted = _accepted_publication(attempt)
        try:
            stopped = not accepted and has_finished_checkpoint(directory, attempt.progress)
        except (OSError, ValueError):
            stopped = False
        if not (accepted or stopped):
            return attempt
        review = _review_directory(attempt)
        _retain_unpublished_output(attempt, contained_path(review, 'unpublished', output=True))
        if accepted:
            clear_accepted_checkpoint(directory, attempt.progress)
        clear_manual_run(attempt)
        return None
