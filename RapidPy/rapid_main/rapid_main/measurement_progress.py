"""Append-only acquisition evidence and explicit per-step execution progress.

Exclusive, fsynced records deliberately fail closed after a partial write. These
records are reproducible evidence, not signatures or authority to replay hardware.
"""
from dataclasses import asdict
import hashlib
import json
import math
import os

from rapid_main.acquisition_evidence import canonical, restore_block
from rapid_main.analysis import reading_cycle_statistics
from rapid_main.block_statistics import block_collection_statistics
from rapid_main.io.legacy_up import LegacyUpBlock, LegacyUpRun, encode_up_file
from rapid_main.specimen_paths import contained_path

DIRECTORY = '.rapidpy-acquisitions'
SCHEMA = 'rapidpy.measurement.progress.v1'


def _digest(value):
    return hashlib.sha256(canonical(value).encode('utf-8')).hexdigest()


class MeasurementProgress:
    def __init__(self, directory, *, run_id, specimen, labels, source=None, simulated=False):
        if not isinstance(run_id, str) or not run_id or not isinstance(specimen, str) or not specimen:
            raise ValueError('Acquisition progress requires explicit run and specimen identities.')
        if not labels or any(not isinstance(label, str) or not label.strip() for label in labels):
            raise ValueError('Acquisition progress requires explicit treatment labels.')
        if type(simulated) is not bool:
            raise ValueError('Acquisition simulation status must be an explicit boolean.')
        self._identity_json = canonical(dict(run_id=run_id, specimen=specimen,
            labels=list(labels), source=source, simulated=simulated))
        self.identity_sha256 = hashlib.sha256(self._identity_json.encode('utf-8')).hexdigest()
        self.root = contained_path(directory, DIRECTORY, output=True)
        self.directory = contained_path(self.root, self.identity_sha256, output=True)
        self._started = False

    @property
    def identity(self):
        return json.loads(self._identity_json)

    @property
    def reference(self):
        return dict(schema=SCHEMA, directory=str(self.directory), identity_sha256=self.identity_sha256)

    def _path(self, name):
        contained_path(self.root.parent, DIRECTORY, output=True)
        contained_path(self.root, self.identity_sha256, output=True)
        return contained_path(self.directory, name, output=True)

    def _write(self, name, payload):
        self._write_bytes(name, canonical(dict(payload=payload, sha256=_digest(payload))).encode('utf-8'))
        return _digest(payload)

    def _write_bytes(self, name, data):
        path = self._path(name)
        path.parent.mkdir(parents=True, exist_ok=True)
        # Never overwrite a previous attempt, including a torn record. The
        # first exclusive creation also arbitrates concurrent owners.
        with path.open('xb') as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        return hashlib.sha256(data).hexdigest()

    def _read(self, name):
        try:
            record = json.loads(self._path(name).read_text(encoding='utf-8'))
            if set(record) != {'payload', 'sha256'} or record['sha256'] != _digest(record['payload']):
                raise ValueError('Acquisition progress digest changed.')
            return record['payload'], record['sha256']
        except (OSError, KeyError, TypeError, json.JSONDecodeError) as exc:
            raise ValueError('Acquisition progress is missing or incomplete: ' + name) from exc

    def begin(self):
        self._write('run.json', dict(schema=SCHEMA, identity=self.identity))
        self._started = True

    def _require_owner(self):
        if not self._started:
            raise ValueError('Only the original acquisition progress writer may append records.')
        self._verify_identity()

    def _verify_identity(self):
        payload, _ = self._read('run.json')
        if payload != dict(schema=SCHEMA, identity=self.identity):
            raise ValueError('Acquisition progress identity changed.')

    def _step(self, index):
        labels = self.identity['labels']
        if type(index) is not int or not 0 <= index < len(labels):
            raise ValueError('Invalid acquisition step index.')
        return dict(identity_sha256=self.identity_sha256, index=index, label=labels[index])

    def start_step(self, index):
        self._require_owner()
        if index:
            self.read_step(index - 1)
        self._write(f'{index:05d}-started.json', self._step(index))

    def capture_step(self, index, *, blocks, readings, susceptibility=None, rmg_treatment=None, sample_height_cm=None):
        self._require_owner()
        started, _ = self._read(f'{index:05d}-started.json')
        if started != self._step(index):
            raise ValueError('Acquisition step intent changed.')
        raw = json.loads(canonical([asdict(block) for block in blocks]))
        payload = dict(**self._step(index), blocks=raw, readings=readings)
        if susceptibility is not None:
            payload['susceptibility'] = susceptibility
        if rmg_treatment is not None:
            payload['rmg_treatment'] = asdict(rmg_treatment)
        if sample_height_cm is not None:
            payload['sample_height_cm'] = sample_height_cm
        payload = json.loads(canonical(payload))  # Detach mutable caller inputs.
        payload.update(self._derive(payload))
        self._write(f'{index:05d}-captured.json', payload)
        if payload['legacy_up'] is not None:
            self._write_bytes(f'{index:05d}.UP', payload['legacy_up'].encode('utf-8'))

    def _derive(self, payload):
        readings = payload['readings']
        cycle = reading_cycle_statistics(readings)
        cycle_payload = asdict(cycle)
        if cycle.signal_to_drift is not None and math.isinf(cycle.signal_to_drift):
            cycle_payload['signal_to_drift'] = '+infinity'
        blocks = tuple(restore_block(raw) for raw in payload['blocks'])
        if blocks:
            if len(blocks) != len(readings):
                raise ValueError('Acquisition collection mixed structured and unstructured readings.')
            ids = [block.audit.block_id for block in blocks if block.audit and block.audit.block_id]
            if len(ids) != len(set(ids)):
                raise ValueError('Acquisition collection repeats a block identity.')
            statistics = block_collection_statistics(blocks)
            from rapid_main.magnetometer import reduce_bracketed_measurement
            derived_readings = [reduce_bracketed_measurement(block).moment_emu for block in blocks]
            if canonical(derived_readings) != canonical(readings):
                raise ValueError('Acquisition readings disagree with retained raw blocks.')
            # The legacy format lacks an H-row acquisition timestamp. Preserve
            # absent times rather than inventing them; the lossless JSON retains
            # all actual observation/audit times and full precision raw vectors.
            up = encode_up_file((LegacyUpRun(self.identity['specimen'], tuple(
                LegacyUpBlock.from_measurement_block(block, timestamps=(None,) * 10)
                for block in blocks)),))
            return dict(kind='bracketed', cycle_statistics=cycle_payload,
                        block_statistics=statistics.payload(), legacy_up=up)
        return dict(kind='unstructured', cycle_statistics=cycle_payload,
                    block_statistics=None, legacy_up=None)

    def read_step(self, index):
        self._verify_identity()
        started, _ = self._read(f'{index:05d}-started.json')
        payload, digest = self._read(f'{index:05d}-captured.json')
        expected = self._step(index)
        required = {*expected, 'blocks', 'readings', 'kind', 'cycle_statistics', 'block_statistics', 'legacy_up'}
        optional = {'susceptibility', 'rmg_treatment', 'sample_height_cm'}
        if not isinstance(payload, dict) or set(payload) - optional != required:
            raise ValueError('Acquisition record has incomplete or unknown fields.')
        if started != expected or any(payload.get(key) != value for key, value in expected.items()):
            raise ValueError('Acquisition step identity changed.')
        if 'rmg_treatment' in payload:
            from rapid_main.io.rmg_writer import RmgTreatment
            if RmgTreatment(**payload['rmg_treatment']).label != expected['label']:
                raise ValueError('RMG treatment evidence belongs to another acquisition step.')
        if 'sample_height_cm' in payload:
            height = payload['sample_height_cm']
            if type(height) not in (int, float) or not math.isfinite(height) or height <= 0:
                raise ValueError('Acquisition specimen height must be finite and positive.')
        derived = self._derive(payload)
        if canonical({key: payload.get(key) for key in derived}) != canonical(derived):
            raise ValueError('Acquisition aggregates disagree with retained readings.')
        if derived['legacy_up'] is not None:
            try:
                exported = self._path(f'{index:05d}.UP').read_bytes()
            except OSError as exc:
                raise ValueError('Legacy .UP acquisition export is missing or incomplete.') from exc
            if exported != derived['legacy_up'].encode('utf-8'):
                raise ValueError('Legacy .UP export disagrees with retained acquisition evidence.')
        return payload, digest

    def accept(self):
        """Seal only after the caller has completed safe return and publication."""
        path, payload = self.prepared_acceptance()
        self._write_bytes(path.name, payload)

    def prepared_acceptance(self):
        """Prepare the same acceptance bytes for the bundle publication journal."""
        self._require_owner()
        refs = [self.read_step(index)[1] for index in range(len(self.identity['labels']))]
        payload = dict(identity_sha256=self.identity_sha256, steps=refs)
        return self._path('accepted.json'), canonical(dict(payload=payload, sha256=_digest(payload))).encode('utf-8')

    def accepted_steps(self):
        self._verify_identity()
        payload, _ = self._read('accepted.json')
        steps = [self.read_step(index) for index in range(len(self.identity['labels']))]
        if payload != dict(identity_sha256=self.identity_sha256, steps=[digest for _, digest in steps]):
            raise ValueError('Acquisition acceptance does not match the original collection.')
        return tuple(payload for payload, _ in steps)
