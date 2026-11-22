"""Transactional measurement output bundle.

Every accepted step is appended to a staging copy first. A durable publication
journal protects the last accepted files as the legacy filenames are replaced.

Simulated runs publish into a clearly separated ``SIMULATED`` directory with a
marker file. The legacy fixed-width specimen and ``.rmg`` formats are left
byte-compatible on purpose — a marker line inside them would break the VB6 and
CIT readers — so the marker lives in the directory name, the marker file, and
the JSON provenance record.
"""
from __future__ import annotations

from dataclasses import astuple, dataclass
from copy import deepcopy
import json
import os
import shutil
import uuid
from pathlib import Path
from typing import Any, Mapping

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.bundle_publication import BundlePublication, JOURNAL
from rapidpy_common.hardware_safety import _replace_snapshot
from rapid_main.io.magic_specimen_writer import append_specimen
from rapid_main.io.magic_writer import append_measurement
from rapid_main.io.rmg_writer import RmgTreatment, append_rmg_record
from rapid_main.io.specimen_writer import append_step, write_header
from rapid_main.specimen_paths import contained_path, specimen_output_path

SIMULATED_SUBDIR = "SIMULATED"
SIMULATION_MARKER_FILE = "SIMULATION.txt"
SIMULATION_STATEMENT = (
    "SIMULATED RUN - synthetic values from a no-communication backend.\n"
    "Not hardware evidence and not valid for production measurement.\n"
)
STEP_LEDGER_FILE = ".rapidpy-steps.json"
STAGING_PREFIX = ".rapidpy-pending"
_SIDECARS = (STEP_LEDGER_FILE, 'provenance.json', SIMULATION_MARKER_FILE, 'susceptibility.json',
             'artifact_index.json')


class BundleNotCommittedError(RuntimeError):
    """Raised when a bundle is used after it was committed or aborted."""


@dataclass(slots=True)
class MeasurementBundlePaths:
    specimen_file: Path
    rmg_file: Path
    magic_measurements_file: Path
    magic_specimens_file: Path


class MeasurementBundleWriter:
    """Stage the legacy files and publish them using a recoverable journal."""

    def __init__(
        self,
        output_dir: str | Path,
        meta: SpecimenMeta,
        *,
        simulated: bool = False,
        allow_simulated_production_output: bool = False,
        resume: bool = False,
        provenance: Mapping[str, Any] | None = None,
    ) -> None:
        self.output_dir = Path(output_dir).resolve()
        self.meta = deepcopy(meta)
        self.simulated = bool(simulated)
        self._allow_simulated_production_output = bool(allow_simulated_production_output)
        self.publish_dir = (
            self.output_dir
            if (not self.simulated or self._allow_simulated_production_output)
            else contained_path(self.output_dir, SIMULATED_SUBDIR, output=True)
        )
        self.paths = validate_measurement_output(self.output_dir, self.meta.name, simulated=self.simulated,
            allow_simulated_production_output=self._allow_simulated_production_output)
        self.staging_dir = contained_path(self.publish_dir, f"{STAGING_PREFIX}-{os.getpid()}", output=True)
        self._staged = _paths_for(self.staging_dir, self.meta.name)
        self.publish_dir.mkdir(parents=True, exist_ok=True)
        self._publication = _publication_for(self.publish_dir, self.meta.name)
        self._lease = self._publication.lease()
        self._lease.__enter__()
        try:
            self._publication.recover()
            self._prepare_staging(resume, provenance)
        except BaseException:
            self._release()
            raise

    def _prepare_staging(self, resume, provenance):
        _reset_directory(self.staging_dir)
        self._closed = False
        self._committed = False
        self._appended_labels: list[str] = []
        self._provenance: dict[str, Any] = dict(provenance or {})

        self._resume = bool(resume)
        self._existing_labels = _read_step_ledger(self.publish_dir) if resume else []
        self._skipped_labels: list[str] = []
        if resume:
            self._seed_from_published()
        self._header_written = bool(self._existing_labels)
        self._ensure_headers()

    # -- properties --------------------------------------------------------

    @property
    def appended_labels(self) -> tuple[str, ...]:
        return tuple(self._appended_labels)

    @property
    def skipped_labels(self) -> tuple[str, ...]:
        """Labels already present from an interrupted run, so not re-appended."""
        return tuple(self._skipped_labels)

    @property
    def committed(self) -> bool:
        return self._committed

    # -- writing -----------------------------------------------------------

    def _ensure_headers(self) -> None:
        if not self._header_written:
            write_header(self._staged.specimen_file, self.meta)
            append_specimen(self._staged.magic_specimens_file, self.meta)
            self._header_written = True

    def append_step(self, step: MeasurementStep, susceptibility: float = 0.0, *,
                    rmg_treatment: RmgTreatment | None = None,
                    axis_sd_emu: tuple[float, float, float] | None = None,
                    sample_height_cm: float | None = None) -> bool:
        """Append one accepted step. Returns ``False`` when it was a duplicate."""
        self._require_open()
        _paths_for(self.publish_dir, self.meta.name)
        _paths_for(self.staging_dir, self.meta.name)
        self._ensure_headers()
        label = str(step.demag_label)
        if label in self._existing_labels or label in self._appended_labels:
            self._skipped_labels.append(label)
            return False
        append_step(self._staged.specimen_file, step)
        append_rmg_record(self._staged.rmg_file, step, susceptibility=susceptibility, meta=self.meta,
                          treatment=rmg_treatment, axis_sd_emu=axis_sd_emu, sample_height_cm=sample_height_cm)
        append_measurement(self._staged.magic_measurements_file, self.meta, step)
        self._appended_labels.append(label)
        return True

    def set_provenance(self, **fields: Any) -> None:
        """Record run/holder/calibration provenance published with the bundle."""
        self._require_open()
        self._provenance.update(fields)

    # -- transaction -------------------------------------------------------

    @property
    def staged_susceptibility(self) -> Path:
        return contained_path(self.staging_dir, 'susceptibility.json', output=True)

    def commit(self, *, progress=None, prepare_index=None, finalize_diagnostics=None) -> MeasurementBundlePaths:
        """Publish the current bundle, including its optional acquisition seal."""
        self._require_open()
        try:
            destinations = validate_measurement_output(self.publish_dir, self.meta.name)
            staged = _paths_for(self.staging_dir, self.meta.name)
            _write_json_atomic(contained_path(self.staging_dir, STEP_LEDGER_FILE, output=True), dict(
                schema='rapidpy.measurement.step_ledger.v1', specimen=self.meta.name,
                labels=list(self._existing_labels) + list(self._appended_labels)))
            _write_json_atomic(contained_path(self.staging_dir, 'provenance.json', output=True), self._provenance_payload())
            if self.simulated:
                contained_path(self.staging_dir, SIMULATION_MARKER_FILE, output=True).write_text(
                    SIMULATION_STATEMENT, encoding='utf-8')
            files = [(source, target) for source, target in zip(astuple(staged), astuple(destinations)) if source.exists()]
            for name in _SIDECARS:
                source = contained_path(self.staging_dir, name, output=True)
                if source.exists():
                    files.append((source, contained_path(self.publish_dir, name, output=True)))
            if progress is not None:
                target, payload = progress.prepared_acceptance()
                source = contained_path(self.staging_dir, '.rapidpy-acquisition-seal', output=True)
                with source.open('xb') as stream:
                    stream.write(payload)
                files.append((source, target))
            if prepare_index is not None:
                index = contained_path(self.staging_dir, 'artifact_index.json', output=True)
                prepare_index(index, {target: source for source, target in files})
                files.append((index, contained_path(self.publish_dir, index.name, output=True)))
            self._publication.publish(files)
            self._committed = True
            if finalize_diagnostics is not None:
                finalize_diagnostics()
        finally:
            self.abort()
        return self.paths

    def abort(self) -> None:
        """Discard everything staged; published files are left untouched."""
        if self._closed:
            return
        try:
            self._discard_staging()
        finally:
            self._closed = True
            self._release()

    def _release(self):
        if self._lease is not None:
            lease, self._lease = self._lease, None
            lease.__exit__(None, None, None)

    def _discard_staging(self) -> None:
        contained_path(self.publish_dir, self.staging_dir.name, output=True)
        shutil.rmtree(self.staging_dir, ignore_errors=True)

    def __enter__(self) -> "MeasurementBundleWriter":
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:  # noqa: ANN001 - context protocol
        if self._closed:
            return False
        if exc_type is None:
            self.commit()
        else:
            self.abort()
        return False

    # -- internals ---------------------------------------------------------

    def _provenance_payload(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "schema": "rapidpy.measurement.provenance.v1",
            "specimen": self.meta.name,
            "sample": self.meta.sample,
            "site": self.meta.site,
            "location": self.meta.location,
            "comment": self.meta.comment,
            "volume_cm3": self.meta.volume,
            "core_plate_strike": self.meta.core_plate_strike,
            "core_plate_dip": self.meta.core_plate_dip,
            "bedding_strike": self.meta.bedding_strike,
            "bedding_dip": self.meta.bedding_dip,
            "fold_axis": self.meta.fold_axis,
            "fold_plunge": self.meta.fold_plunge,
            "magnetic_declination_e": self.meta.magnetic_declination_e,
            "core_orientation_available": self.meta.has_core_orientation,
            "simulated": self.simulated,
            "labels": list(self._existing_labels) + list(self._appended_labels),
            "resumed_labels": list(self._existing_labels),
        }
        payload.update(self._provenance)
        if self.simulated:
            payload["simulation_statement"] = SIMULATION_STATEMENT.strip()
        return payload

    def _seed_from_published(self) -> None:
        """Copy published files into staging so a resumed run appends to them."""
        for published, staged in (
            (self.paths.specimen_file, self._staged.specimen_file),
            (self.paths.rmg_file, self._staged.rmg_file),
            (self.paths.magic_measurements_file, self._staged.magic_measurements_file),
            (self.paths.magic_specimens_file, self._staged.magic_specimens_file),
        ):
            if published.exists():
                staged.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(published, staged)

    def _require_open(self) -> None:
        if self._closed:
            raise BundleNotCommittedError(
                "measurement bundle is already committed or aborted"
            )


def _paths_for(directory: Path, name: str) -> MeasurementBundlePaths:
    return MeasurementBundlePaths(
        specimen_file=specimen_output_path(directory, name),
        rmg_file=contained_path(directory, f"{name}.rmg", output=True),
        magic_measurements_file=contained_path(directory, "measurements.txt", output=True),
        magic_specimens_file=contained_path(directory, "specimens.txt", output=True),
    )


def validate_measurement_output(directory: str | Path, name: str, *, simulated: bool = False,
                                allow_simulated_production_output: bool = False) -> MeasurementBundlePaths:
    root = Path(directory).absolute()
    publish = root if not simulated or allow_simulated_production_output else contained_path(root, SIMULATED_SUBDIR, output=True)
    paths = _paths_for(publish, name)
    contained_path(publish, f'{STAGING_PREFIX}-{os.getpid()}', output=True)
    contained_path(publish, JOURNAL, output=True)
    for filename in (STEP_LEDGER_FILE, 'provenance.json', SIMULATION_MARKER_FILE,
                     'artifact_index.json', 'workflow_summary.json', 'quicklook.json',
                     'communication.tsv', 'susceptibility.json', 'rockmag_run.json', 'thermal_run.json'):
        contained_path(publish, filename, output=True)
    return paths


def _publication_for(directory, name):
    targets = (*astuple(_paths_for(directory, name)),
               *(contained_path(directory, filename, output=True) for filename in _SIDECARS))
    return BundlePublication(directory, name, targets)


def recover_measurement_output(directory: Path, name: str) -> None:
    publication = _publication_for(directory, name)
    if publication.journal.exists():
        with publication.lease():
            publication.recover()


def _reset_directory(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)
    path.mkdir(parents=True, exist_ok=True)


def _read_step_ledger(directory: Path) -> list[str]:
    path = contained_path(directory, STEP_LEDGER_FILE, output=True)
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    labels = payload.get("labels")
    if not isinstance(labels, list):
        return []
    return [str(label) for label in labels]


def _write_json_atomic(path: Path, payload: Mapping[str, Any]) -> None:
    temp_path = contained_path(path.parent, f'{path.name}.tmp-{uuid.uuid4().hex}', output=True)
    text = json.dumps(payload, indent=2, sort_keys=True, default=str) + "\n"
    try:
        with temp_path.open('x', encoding='utf-8', newline='\n') as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        _replace_snapshot(temp_path, path)
    except BaseException:
        try:
            temp_path.unlink()
        except OSError:
            pass
        raise
