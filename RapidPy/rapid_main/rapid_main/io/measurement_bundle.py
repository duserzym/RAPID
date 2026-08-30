"""Transactional measurement output bundle.

Every accepted step is appended to a staging copy first. ``commit()`` fsyncs
each staged file and publishes it with an atomic rename, so a rejected block,
an abort, a crash, or an exception can never leave a half-written accepted
measurement in the production output path.

Simulated runs publish into a clearly separated ``SIMULATED`` directory with a
marker file. The legacy fixed-width specimen and ``.rmg`` formats are left
byte-compatible on purpose — a marker line inside them would break the VB6 and
CIT readers — so the marker lives in the directory name, the marker file, and
the JSON provenance record.
"""
from __future__ import annotations

from dataclasses import dataclass
import json
import os
import shutil
from pathlib import Path
from typing import Any, Mapping

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.magic_specimen_writer import append_specimen
from rapid_main.io.magic_writer import append_measurement
from rapid_main.io.rmg_writer import append_rmg_record
from rapid_main.io.specimen_writer import append_step, write_header

SIMULATED_SUBDIR = "SIMULATED"
SIMULATION_MARKER_FILE = "SIMULATION.txt"
SIMULATION_STATEMENT = (
    "SIMULATED RUN - synthetic values from a no-communication backend.\n"
    "Not hardware evidence and not valid for production measurement.\n"
)
STEP_LEDGER_FILE = ".rapidpy-steps.json"
STAGING_PREFIX = ".rapidpy-pending"


class BundleNotCommittedError(RuntimeError):
    """Raised when a bundle is used after it was committed or aborted."""


@dataclass(slots=True)
class MeasurementBundlePaths:
    specimen_file: Path
    rmg_file: Path
    magic_measurements_file: Path
    magic_specimens_file: Path


class MeasurementBundleWriter:
    """Write VB6/CIT + RMG + MagIC files in lockstep, then publish atomically."""

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
        self.output_dir = Path(output_dir)
        self.meta = meta
        self.simulated = bool(simulated)
        self._allow_simulated_production_output = bool(allow_simulated_production_output)
        self.publish_dir = (
            self.output_dir
            if (not self.simulated or self._allow_simulated_production_output)
            else self.output_dir / SIMULATED_SUBDIR
        )
        self.publish_dir.mkdir(parents=True, exist_ok=True)
        self.staging_dir = self.publish_dir / f"{STAGING_PREFIX}-{os.getpid()}"
        _reset_directory(self.staging_dir)

        self.paths = _paths_for(self.publish_dir, meta.name)
        self._staged = _paths_for(self.staging_dir, meta.name)
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

    def append_step(self, step: MeasurementStep, susceptibility: float = 0.0) -> bool:
        """Append one accepted step. Returns ``False`` when it was a duplicate."""
        self._require_open()
        self._ensure_headers()
        label = str(step.demag_label)
        if label in self._existing_labels or label in self._appended_labels:
            self._skipped_labels.append(label)
            return False
        append_step(self._staged.specimen_file, step)
        append_rmg_record(self._staged.rmg_file, step, susceptibility=susceptibility)
        append_measurement(self._staged.magic_measurements_file, self.meta, step)
        self._appended_labels.append(label)
        return True

    def set_provenance(self, **fields: Any) -> None:
        """Record run/holder/calibration provenance published with the bundle."""
        self._require_open()
        self._provenance.update(fields)

    # -- transaction -------------------------------------------------------

    def commit(self) -> MeasurementBundlePaths:
        """Publish every staged file atomically."""
        self._require_open()
        for staged, published in (
            (self._staged.specimen_file, self.paths.specimen_file),
            (self._staged.rmg_file, self.paths.rmg_file),
            (self._staged.magic_measurements_file, self.paths.magic_measurements_file),
            (self._staged.magic_specimens_file, self.paths.magic_specimens_file),
        ):
            if staged.exists():
                _fsync_file(staged)
                os.replace(staged, published)
        _write_json_atomic(
            self.publish_dir / STEP_LEDGER_FILE,
            {
                "schema": "rapidpy.measurement.step_ledger.v1",
                "specimen": self.meta.name,
                "labels": list(self._existing_labels) + list(self._appended_labels),
            },
        )
        _write_json_atomic(self.publish_dir / "provenance.json", self._provenance_payload())
        if self.simulated:
            (self.publish_dir / SIMULATION_MARKER_FILE).write_text(
                SIMULATION_STATEMENT, encoding="utf-8"
            )
        shutil.rmtree(self.staging_dir, ignore_errors=True)
        self._closed = True
        self._committed = True
        return self.paths

    def abort(self) -> None:
        """Discard everything staged; published files are left untouched."""
        shutil.rmtree(self.staging_dir, ignore_errors=True)
        self._closed = True

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
                shutil.copy2(published, staged)

    def _require_open(self) -> None:
        if self._closed:
            raise BundleNotCommittedError(
                "measurement bundle is already committed or aborted"
            )


def _paths_for(directory: Path, name: str) -> MeasurementBundlePaths:
    return MeasurementBundlePaths(
        specimen_file=directory / name,
        rmg_file=directory / f"{name}.rmg",
        magic_measurements_file=directory / "measurements.txt",
        magic_specimens_file=directory / "specimens.txt",
    )


def _reset_directory(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)
    path.mkdir(parents=True, exist_ok=True)


def _read_step_ledger(directory: Path) -> list[str]:
    path = directory / STEP_LEDGER_FILE
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


def _fsync_file(path: Path) -> None:
    """Best-effort durability before the atomic rename.

    ``os.fsync`` needs a writable descriptor on Windows, and some filesystems
    refuse it outright. The rename is the correctness guarantee; the flush is
    an extra durability step, so a refusal must not fail the publish.
    """
    try:
        with open(path, "r+b") as handle:
            handle.flush()
            os.fsync(handle.fileno())
    except OSError:
        return


def _write_json_atomic(path: Path, payload: Mapping[str, Any]) -> None:
    temp_path = path.with_name(f"{path.name}.tmp-{os.getpid()}")
    text = json.dumps(payload, indent=2, sort_keys=True, default=str) + "\n"
    try:
        with open(temp_path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    except BaseException:
        try:
            temp_path.unlink()
        except OSError:
            pass
        raise
