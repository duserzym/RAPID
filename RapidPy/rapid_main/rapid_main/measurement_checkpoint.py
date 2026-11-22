"""Retain and verify a sequence stopped between treatment steps."""
from dataclasses import astuple
import json
import re
import shutil

from rapid_main.acquisition_evidence import canonical
from rapid_main.io.bundle_history import open_bundle_archive, publish_bundle_archive
from rapid_main.io.bundle_publication import _digest, _sync_directory
from rapid_main.io.measurement_bundle import _publication_for, _write_json_atomic
from rapid_main.measurement_progress import MeasurementProgress
from rapid_main.specimen_paths import contained_path
from rapidpy_common.paths import absolute_io_path


DIRECTORY = '.rapidpy-checkpoints'
_SCHEMA = 'rapidpy.measurement.checkpoint.v1'


def _pointer(directory):
    return contained_path(absolute_io_path(directory), f'{DIRECTORY}/pending.json', output=True)


def pending_checkpoint(directory):
    return _digest(_pointer(directory))


def discard_checkpoint(directory, specimen, expected):
    """Remove the pending pointer under the publication lease. Keep its evidence."""
    with _publication_for(directory, specimen).lease():
        clear_checkpoint(directory, expected)


def clear_checkpoint(directory, expected):
    path = _pointer(directory)
    if _digest(path) != expected:
        raise ValueError('The saved run changed. Review it again.')
    path.unlink()
    _sync_directory(path.parent)


def _publication_digests(bundle):
    paths = (*astuple(bundle.paths), bundle.publish_dir / 'artifact_index.json', bundle.publish_dir / 'provenance.json')
    return {path.relative_to(bundle.publish_dir).as_posix(): _digest(absolute_io_path(path)) for path in paths}


def save_checkpoint(bundle, progress, count, *, previous, susceptibility, evidence):
    """Save verified staged bytes after the original worker has returned safely."""
    step_digests = [progress.read_step(index)[1] for index in range(count)]
    files = {path.relative_to(bundle.staging_dir).as_posix(): _digest(path) for path in astuple(bundle._staged)}
    if not all(files.values()):
        raise ValueError('Saved measurement output is incomplete.')
    path = _pointer(bundle.publish_dir)
    if _digest(path) != previous:
        raise ValueError('The saved run changed before the checkpoint could be written.')
    snapshots = contained_path(path.parent, 'snapshots', output=True)
    _, archive_id = publish_bundle_archive(bundle.staging_dir, bundle.meta.name, files, snapshots)
    checkpoint = dict(schema=_SCHEMA, identity_sha256=progress.identity_sha256, count=count,
                      archive_id=archive_id, step_digests=step_digests,
                      publication_files=_publication_digests(bundle),
                      susceptibility=susceptibility, evidence=evidence)
    digest = progress._write(f'checkpoint-{count:05d}.json', checkpoint)
    _write_json_atomic(path, dict(identity_sha256=progress.identity_sha256, count=count, sha256=digest))
    _sync_directory(path.parent)


def restore_checkpoint(bundle, expected, *, labels, source, simulated):
    """Restore an exact prefix with the writer's existing publication lease held."""
    path = _pointer(bundle.publish_dir)
    if _digest(path) != expected:
        raise ValueError('The saved run changed. Review it again.')
    pointer = json.loads(path.read_bytes())
    if not isinstance(pointer, dict) or set(pointer) != {'identity_sha256', 'count', 'sha256'}:
        raise ValueError('Invalid saved run pointer.')
    count = pointer['count']
    if type(count) is not int or not 0 < count < len(labels):
        raise ValueError('Saved run does not have a remaining treatment step.')
    progress = MeasurementProgress.open_existing(bundle.publish_dir, pointer['identity_sha256'])
    checkpoint, digest = progress._read(f'checkpoint-{count:05d}.json')
    identity = progress.identity
    if (checkpoint['schema'] != _SCHEMA or digest != pointer['sha256'] or checkpoint['count'] != count
            or checkpoint['identity_sha256'] != progress.identity_sha256
            or identity['specimen'] != bundle.meta.name or identity['labels'] != list(labels)
            or identity['simulated'] is not simulated or canonical(identity['source']) != canonical(source)):
        raise ValueError('Saved run does not match this specimen, sequence, operator, or calibration.')
    if _publication_digests(bundle) != checkpoint['publication_files']:
        raise ValueError('The accepted specimen output changed after the run was saved.')
    if progress._path('accepted.json').exists() or any(
            progress._path(f'{index:05d}-{record}.json').exists()
            for index in range(count, len(labels)) for record in ('started', 'captured')):
        raise ValueError('A later treatment may already have started. Review the specimen before starting a new run.')
    prior = [progress.read_step(index) for index in range(count)]
    if [digest for _, digest in prior] != checkpoint['step_digests']:
        raise ValueError('The original acquisition prefix changed.')
    for name, expected_digest in checkpoint['evidence'].items():
        if _digest(contained_path(bundle.publish_dir, name, output=True)) != expected_digest:
            raise ValueError(f'The original measurement evidence changed: {name}')

    archive_id = checkpoint['archive_id']
    if not isinstance(archive_id, str) or not re.fullmatch(r'[0-9a-f]{64}', archive_id):
        raise ValueError('Invalid saved run archive identity.')
    archive_path = contained_path(path.parent, f'snapshots/{archive_id[:32]}', output=True)
    targets = {target.relative_to(bundle.staging_dir).as_posix(): target for target in astuple(bundle._staged)}
    with open_bundle_archive(archive_path, expected_id=archive_id) as archive:
        if set(archive.namelist()) != set(targets):
            raise ValueError('Saved run archive does not contain its original measurement files.')
        for name, target in targets.items():
            with archive.open(name) as member, absolute_io_path(target).open('wb') as destination:
                shutil.copyfileobj(member, destination)
    bundle._appended_labels.extend(labels[:count])
    progress._started = True
    return progress, checkpoint, [payload for payload, _ in prior]
