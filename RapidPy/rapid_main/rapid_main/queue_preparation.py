"""Prepare the queue and capture its file inputs without opening hardware."""
from dataclasses import asdict, dataclass
import hashlib
import os
from pathlib import Path
from typing import Callable

from PySide6 import QtCore

from .data_model import SampleIndexRegistrations
from .io.measurement_bundle import validate_measurement_output
from .io.sample_index import read_sample_index_registrations
from .queue_compiler import QueueCommand, QueueOptions, QueueSample, compile_queue, resolve_queue_samples
from .specimen_metadata import capture_specimen_metadata
from .specimen_paths import specimen_run_directory


@dataclass(frozen=True)
class PreparedQueue:
    samples: list[QueueSample]
    options: QueueOptions
    commands: list[QueueCommand]
    source_indexes: dict


def prepare_queue(samples, options, *, labels, routine_labels, output, simulated,
                  from_queue_panel=False, cancelled: Callable[[], bool] = lambda: False) -> PreparedQueue:
    def check_cancelled():
        if cancelled():
            raise InterruptedError('Queue preparation canceled.')

    check_cancelled()
    samples = resolve_queue_samples(samples, labels)
    resolved_paths = {}
    for sample in samples:
        check_cancelled()
        if not sample.source_file:
            continue
        path = Path(sample.source_file)
        if not path.is_absolute() or path.suffix.lower() not in {'.sam', '.csv'}:
            raise ValueError('Source indexes must be absolute SAM or CSV file paths.')
        if path not in resolved_paths:
            resolved_paths[path] = path.resolve()
        resolved = resolved_paths[path]
        identity = os.path.normcase(str(resolved))
        if sample.file_id != identity and not (from_queue_panel and sample.file_id == os.path.normcase(os.path.normpath(path))):
            raise ValueError('Queue source-file identity must match its original index path.')
        sample.file_id = identity
        sample.source_file = str(resolved)

    commands = compile_queue(samples, options, strict=True)
    for routine in routine_labels:
        if any(command.command_type == 'Meas' and command.measurement_labels != routine for command in commands):
            raise ValueError('The reviewed routine identity must match every file measurement sequence.')

    source_indexes = {}
    registrations_by_file: dict[str, SampleIndexRegistrations] = {}
    for sample in samples:
        check_cancelled()
        index = Path(sample.source_file) if sample.source_file else None
        directory = specimen_run_directory(output, sample.sample_name, index)
        validate_measurement_output(directory, sample.sample_name, simulated=simulated)
        if index is None:
            continue
        identity = sample.file_id
        if identity not in source_indexes:
            payload = index.read_bytes()
            check_cancelled()
            registrations = read_sample_index_registrations(index)
            if payload != index.read_bytes():
                raise ValueError('Sample index changed while preparing the queue.')
            registrations_by_file[identity] = registrations
            source_indexes[identity] = dict(source_file=str(index), sha256=hashlib.sha256(payload).hexdigest(),
                                            entries=[asdict(entry) for entry in registrations.entries], specimens={})
        registrations = registrations_by_file[identity]
        matches = [entry for entry in registrations.entries if entry.specimen_name == sample.sample_name]
        if len(matches) != 1:
            raise ValueError('Each queued specimen must match exactly one entry in its original source index.')
        snapshots = source_indexes[identity]['specimens']
        if sample.sample_name not in snapshots:
            snapshots[sample.sample_name] = capture_specimen_metadata(sample.sample_name,
                sample_dir=index.parent, registrations=registrations)
    check_cancelled()
    return PreparedQueue(samples, options, commands, source_indexes)


class QueuePreparationWorker(QtCore.QThread):
    """Run a detached preparation request and expose its result after exit."""

    def __init__(self, samples, options, *, inputs, parent=None):
        super().__init__(parent)
        self._samples = samples
        self._options = options
        self._inputs = inputs
        self.result: PreparedQueue | None = None
        self.error = ''

    def run(self):
        try:
            self.result = prepare_queue(self._samples, self._options, **self._inputs,
                                        cancelled=self.isInterruptionRequested)
        except Exception as exc:
            self.error = str(exc) or type(exc).__name__
