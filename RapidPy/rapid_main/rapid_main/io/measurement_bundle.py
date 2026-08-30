from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.magic_specimen_writer import append_specimen
from rapid_main.io.magic_writer import append_measurement
from rapid_main.io.rmg_writer import append_rmg_record
from rapid_main.io.specimen_writer import append_step, write_header


@dataclass(slots=True)
class MeasurementBundlePaths:
    specimen_file: Path
    rmg_file: Path
    magic_measurements_file: Path
    magic_specimens_file: Path


class MeasurementBundleWriter:
    """Write VB6/CIT + RMG + MagIC files in lockstep per measurement step."""

    def __init__(self, output_dir: str | Path, meta: SpecimenMeta) -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.meta = meta
        self.paths = MeasurementBundlePaths(
            specimen_file=self.output_dir / meta.name,
            rmg_file=self.output_dir / f"{meta.name}.rmg",
            magic_measurements_file=self.output_dir / "measurements.txt",
            magic_specimens_file=self.output_dir / "specimens.txt",
        )
        self._header_written = False
        self._ensure_headers()

    def _ensure_headers(self) -> None:
        if not self._header_written:
            write_header(self.paths.specimen_file, self.meta)
            append_specimen(self.paths.magic_specimens_file, self.meta)
            self._header_written = True

    def append_step(self, step: MeasurementStep, susceptibility: float = 0.0) -> None:
        self._ensure_headers()
        append_step(self.paths.specimen_file, step)
        append_rmg_record(self.paths.rmg_file, step, susceptibility=susceptibility)
        append_measurement(self.paths.magic_measurements_file, self.meta, step)
