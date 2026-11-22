"""Bind accepted Up evidence and a Down queue command to one scientific result."""
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path

from rapid_main.acquisition_evidence import canonical, restore_block
from rapid_main.io.legacy_up import parse_up_file, read_up_measurements, select_up_measurements
from rapid_main.measurement_progress import MeasurementProgress
from rapid_main.specimen_paths import contained_path


def _parse_archive(payload):
    try:
        text = payload.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = payload.decode('cp1252')
    return parse_up_file(text)


@dataclass(frozen=True)
class LegacyUpSource:
    path: str
    sha256: str
    file_id: str
    label: str
    range_factor: float
    operator: str
    confirmed_iso: str
    specimen_names: tuple[str, ...]
    susceptibility: tuple[tuple[str, float], ...] = ()

    @classmethod
    def capture(cls, path, *, file_id, label, specimens, range_factor, operator, susceptibility=None):
        if (not isinstance(operator, str) or not operator.strip() or not label or not file_id
                or type(range_factor) not in (int, float) or not math.isfinite(range_factor) or range_factor <= 0):
            raise ValueError('Legacy Up requires a named operator, original step and positive range factor.')
        path = Path(path).resolve(strict=True)
        if path.suffix.lower() != '.up':
            raise ValueError('Select a VB6 .UP file.')
        names = tuple(dict.fromkeys(specimens))
        bridge = {} if susceptibility is None else susceptibility
        if label.strip().upper() == 'SUSC' and (set(bridge) != set(names) or any(
                type(value) not in (int, float) or not math.isfinite(value) for value in bridge.values())):
            raise ValueError('Legacy SUSC pairing requires the original susceptibility for every specimen.')
        payload = path.read_bytes()
        runs = _parse_archive(payload)
        for name in names:
            select_up_measurements(runs, name)
        return cls(str(path), hashlib.sha256(payload).hexdigest(), file_id, label, float(range_factor),
                   operator.strip(), datetime.now(timezone.utc).isoformat(), names, tuple(sorted(bridge.items())))

    def verified_bytes(self):
        payload = Path(self.path).read_bytes()
        if hashlib.sha256(payload).hexdigest() != self.sha256:
            raise ValueError('The selected legacy Up file changed after confirmation.')
        return payload

    def load(self, specimen):
        if specimen not in self.specimen_names:
            raise ValueError('This specimen was not included in the legacy Up confirmation.')
        selected = select_up_measurements(_parse_archive(self.verified_bytes()), specimen)
        # frmSQUID applies the axis gains before Sample.WriteUpMeasurements.
        return tuple(raw.to_measurement_block(range_factor=self.range_factor,
            axis_calibration=(1., 1., 1.)) for raw in selected.blocks)


@dataclass(frozen=True)
class QueuePair:
    queue_run_id: str
    up_index: int | None
    down_index: int
    file_id: str
    row_id: str
    hole: int
    specimen: str
    label: str
    count: int

    @classmethod
    def from_command(cls, queue_run_id, commands, current, *, legacy_up=None):
        matches = [(index, command) for index, command in enumerate(commands)
                   if command.command_type == 'Meas' and command.file_id == current.file_id
                   and command.sample_name == current.sample_name and command.hole == current.hole
                   and command.row_id == current.row_id]
        if len(matches) == 1:
            if not current.legacy_up_required:
                return None
            (down, original), = matches
            if (not isinstance(legacy_up, LegacyUpSource) or not queue_run_id or current is not original
                    or legacy_up.file_id != current.file_id or current.measurement_labels != (legacy_up.label,)
                    or current.sample_name not in legacy_up.specimen_names):
                raise ValueError('Select the verified legacy Up collection for the original Down command.')
            return QueueSide(cls(queue_run_id, None, down, current.file_id, current.row_id,
                current.hole, current.sample_name, legacy_up.label, current.avg_steps), False, legacy_up)
        if (len(matches) != 2 or any(command != current for _, command in matches)
                or len(current.measurement_labels) != 1):
            raise ValueError('Paired queue commands must identify one original specimen and step.')
        (up, first), (down, second) = matches
        if not queue_run_id or not any(command.command_type == 'InitUp' and command.file_id == current.file_id
                                       for command in commands[:up]) or not any(
                command.command_type == 'Flip' and command.file_id == current.file_id for command in commands[up:down]):
            raise ValueError('Paired acquisition requires the original queue run and file flip.')
        if current is not first and current is not second:
            raise ValueError('Paired acquisition requires the original compiled command.')
        return QueueSide(cls(queue_run_id, up, down, current.file_id, current.row_id,
                            current.hole, current.sample_name, current.measurement_labels[0], current.avg_steps),
                         current is first)


@dataclass(frozen=True)
class QueueSide:
    pair: QueuePair
    is_up: bool
    legacy_up: LegacyUpSource | None = None

    @property
    def run_id(self):
        index = self.pair.up_index if self.is_up else self.pair.down_index
        return f'{self.pair.queue_run_id}-command-{index}'

    def _path(self, directory, name):
        key = hashlib.sha256(canonical(asdict(self.pair)).encode()).hexdigest()
        return contained_path(directory, f'.rapidpy-pairs/{key}/{name}', output=True)

    def up_directory(self, directory):
        return self._path(directory, 'up')

    def evidence_directory(self, directory):
        return self.up_directory(directory) if self.is_up else self._path(directory, 'down')

    def _intent(self, directory):
        return self._path(directory, 'up-started.json' if self.is_up else 'down-started.json')

    def check_admission(self, directory, *, simulated):
        if self._intent(directory).exists():
            raise ValueError('This queue pass already has an acquisition attempt. Review original progress before recovery.')
        if not self.is_up:
            return self.load_up(directory, simulated=simulated)
        return None

    @staticmethod
    def _create(path, payload):
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open('xb') as stream:
            stream.write(canonical(payload).encode())
            stream.flush()
            os.fsync(stream.fileno())

    def begin(self, directory):
        self._create(self._intent(directory), self.context)

    @property
    def context(self):
        value = dict(pair=asdict(self.pair), is_up=self.is_up)
        if self.legacy_up is not None:
            value['legacy_up'] = asdict(self.legacy_up)
        return value

    def validate_blocks(self, blocks, *, simulated):
        if len(blocks) != self.pair.count or any(block.is_up is not self.is_up for block in blocks):
            raise ValueError('Paired acquisition direction or block count disagrees with the queue.')
        if not simulated:
            for block in blocks:
                audit = block.audit
                if (audit is None or audit.simulated or audit.is_holder_block or not audit.block_id
                        or not audit.holder_record_id or len(block.observations) != 6
                        or (audit.sample_name, audit.treatment_label, audit.run_id)
                        != (self.pair.specimen, self.pair.label, self.run_id)):
                    raise ValueError('Paired acquisition requires the original live specimen audit.')

    def seal_up(self, directory, progress):
        progress.accepted_steps()
        seal = progress._path('accepted.json').read_bytes()
        self._create(self._path(directory, 'up-ready.json'), dict(
            identity=progress.identity, seal_sha256=hashlib.sha256(seal).hexdigest()))

    def load_up(self, directory, *, simulated):
        try:
            if self.legacy_up is not None:
                source = self.legacy_up
                blocks = source.load(self.pair.specimen)
                archived = self.legacy_archive(directory)
                reference = dict(schema='rapidpy.legacy_up_import.v1', source_file=source.path,
                    sha256=source.sha256, archive=str(archived), sample=self.pair.specimen,
                    label=source.label, blocks=len(blocks), range_factor=source.range_factor,
                    operator=source.operator, confirmed_iso=source.confirmed_iso, audited=False)
                bridge = dict(source.susceptibility)
                susceptibility = (dict(value=bridge[self.pair.specimen], evidence=dict(
                    source='operator_transcribed_legacy_up', operator=source.operator,
                    confirmed_iso=source.confirmed_iso)) if self.pair.label.strip().upper() == 'SUSC' else None)
                return blocks, susceptibility, None, reference
            receipt = json.loads(self._path(directory, 'up-ready.json').read_bytes())
            up_side = QueueSide(self.pair, True)
            identity = receipt['identity']
            if (identity['run_id'] != up_side.run_id or identity['specimen'] != self.pair.specimen
                    or identity['labels'] != [self.pair.label] or identity['simulated'] is not simulated
                    or identity['source']['queue_side'] != up_side.context):
                raise ValueError('Original Up acquisition identity changed.')
            progress = MeasurementProgress(self.up_directory(directory), **identity)
            payload, = progress.accepted_steps()
            seal = progress._path('accepted.json').read_bytes()
            if hashlib.sha256(seal).hexdigest() != receipt['seal_sha256']:
                raise ValueError('Original Up acceptance seal changed.')
            up = read_up_measurements(payload['legacy_up'], self.pair.specimen)
            blocks = tuple(restore_block(raw) for raw in payload['blocks'])
            up_side.validate_blocks(blocks, simulated=simulated)
            if len(up.blocks) != len(blocks):
                raise ValueError('Original .UP export block count changed.')
            susceptibility = payload.get('susceptibility')
            if self.pair.label.strip().upper() == 'SUSC' and (
                    not isinstance(susceptibility, dict) or type(susceptibility.get('value')) not in (int, float)
                    or not math.isfinite(susceptibility['value']) or not isinstance(susceptibility.get('evidence'), dict)):
                raise ValueError('Original Up susceptibility record is missing or invalid.')
            reference = dict(progress.reference, seal_sha256=receipt['seal_sha256'])
            if 'rmg_treatment' in payload:
                reference['rmg_treatment'] = payload['rmg_treatment']
            return blocks, susceptibility, identity['source'], reference
        except (OSError, TypeError, KeyError, AttributeError) as exc:
            raise ValueError('The original accepted Up acquisition is required before the Down pass.') from exc

    def legacy_archive(self, directory):
        return self._path(directory, 'down/legacy-up.UP')

    def archive_legacy_up(self, directory):
        target = self.legacy_archive(directory)
        target.parent.mkdir(parents=True, exist_ok=True)
        payload = self.legacy_up.verified_bytes()
        with target.open('xb') as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())

    def verify_legacy_archive(self, directory):
        payload = self.legacy_archive(directory).read_bytes()
        if hashlib.sha256(payload).hexdigest() != self.legacy_up.sha256:
            raise ValueError('The archived legacy Up input changed during the Down pass.')
