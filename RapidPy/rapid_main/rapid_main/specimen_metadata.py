"""Resolve full specimen metadata before a run starts.

RapidPy previously started every measurement with a ``SpecimenMeta`` whose
comment, sample, site, and location were blank strings, and whose orientation
and volume were defaults. VB6 reads those values from the specimen file header
and the ``.sam`` registry, and writes them into every output.

This module recovers the same values:

1. an existing specimen file header wins (it is what VB6 would read back);
2. otherwise the ``.sam`` registry supplies the hierarchy and depth;
3. anything still unknown stays at its documented default, and the caller is
   told which fields were defaulted so it can warn the operator.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, replace
import hashlib
from pathlib import Path
from typing import Iterable

from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations, SpecimenMeta
from rapid_main.io.specimen_reader import read_specimen


def _source_digest(path: Path) -> str | None:
    if not path.is_file():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()


def capture_specimen_metadata(name: str, *, sample_dir: Path,
                              registrations: SampleIndexRegistrations) -> dict:
    """Freeze header selection, including absent candidates, before queue I/O."""
    paths = candidate_specimen_paths(name, sample_dir=sample_dir)
    sources = {str(path.resolve()): _source_digest(path) for path in paths}
    resolution = resolve_specimen_meta(name, sample_dir=sample_dir, registrations=registrations)
    if sources != {str(path.resolve()): _source_digest(path) for path in paths}:
        raise ValueError('Specimen header changed while preparing the queue.')
    return dict(meta=asdict(resolution.meta), source=resolution.source,
                header_path=str(resolution.header_path.resolve()) if resolution.header_path else None,
                registration=asdict(resolution.registration) if resolution.registration else None,
                defaulted_fields=list(resolution.defaulted_fields), sources=sources)


def restore_specimen_metadata(snapshot: dict) -> 'SpecimenMetaResolution':
    """Verify original inputs and return detached frozen scientific metadata."""
    for path, digest in snapshot['sources'].items():
        if _source_digest(Path(path)) != digest:
            raise ValueError('The original specimen header changed after queue preparation.')
    return SpecimenMetaResolution(meta=SpecimenMeta(**snapshot['meta']), source=snapshot['source'],
        header_path=Path(snapshot['header_path']) if snapshot['header_path'] else None,
        registration=SampleIndexRegistration(**snapshot['registration']) if snapshot['registration'] else None,
        defaulted_fields=tuple(snapshot['defaulted_fields']))


@dataclass(frozen=True)
class SpecimenMetaResolution:
    """Resolved metadata plus the provenance of each source."""

    meta: SpecimenMeta
    source: str = "defaults"
    header_path: Path | None = None
    registration: SampleIndexRegistration | None = None
    defaulted_fields: tuple[str, ...] = ()

    @property
    def complete(self) -> bool:
        return not self.defaulted_fields


def candidate_specimen_paths(
    name: str,
    *,
    sample_dir: str | Path | None = None,
    data_dir: str | Path | None = None,
) -> list[Path]:
    """Places VB6 would look for an existing specimen file, in order."""

    candidates: list[Path] = []
    if sample_dir:
        candidates.append(Path(sample_dir) / name)
    if data_dir:
        candidates.append(Path(data_dir) / name / name)
        candidates.append(Path(data_dir) / name)
    return candidates


def find_registration(
    name: str,
    registrations: SampleIndexRegistrations | Iterable[SampleIndexRegistration] | None,
) -> SampleIndexRegistration | None:
    if registrations is None:
        return None
    entries = (
        registrations.entries
        if isinstance(registrations, SampleIndexRegistrations)
        else list(registrations)
    )
    for entry in entries:
        if entry.specimen_name == name:
            return entry
    return None


def resolve_specimen_meta(
    name: str,
    *,
    sample_dir: str | Path | None = None,
    data_dir: str | Path | None = None,
    registrations: SampleIndexRegistrations | Iterable[SampleIndexRegistration] | None = None,
    comment: str = "",
) -> SpecimenMetaResolution:
    """Build the most complete ``SpecimenMeta`` available for ``name``."""

    specimen_name = (name or "").strip() or "UNKNOWN"
    meta = SpecimenMeta(name=specimen_name, comment=comment)
    source = "defaults"
    header_path: Path | None = None

    for candidate in candidate_specimen_paths(
        specimen_name, sample_dir=sample_dir, data_dir=data_dir
    ):
        if not candidate.is_file():
            continue
        try:
            header_meta, _steps = read_specimen(candidate, specimen_name=specimen_name)
        except (OSError, ValueError):
            continue
        meta = header_meta
        source = "specimen-header"
        header_path = candidate
        break

    registration = find_registration(specimen_name, registrations)
    if registration is not None:
        if not meta.site:
            meta = replace(meta, site=registration.formation or registration.sample_set or "")
        if not meta.location:
            meta = replace(meta, location=registration.location or "")
        if not meta.comment and registration.depth_cm:
            meta = replace(meta, comment=f"depth {registration.depth_cm} cm")
        if source == "defaults":
            source = "sample-index"

    if not meta.sample:
        # A specimen belongs to a sample of the same name unless a registry or
        # header says otherwise. Deriving a shorter sample name from the
        # specimen string would be a guess, so it is left explicit.
        meta = replace(meta, sample=specimen_name)

    defaulted: list[str] = []
    if not meta.site:
        defaulted.append("site")
    if not meta.location:
        defaulted.append("location")
    if not meta.comment:
        defaulted.append("comment")
    if meta.volume in (0.0, 1.0):
        defaulted.append("volume")
    if (meta.core_plate_strike, meta.core_plate_dip) == (0.0, 0.0):
        defaulted.append("core_plate_orientation")
    if (meta.bedding_strike, meta.bedding_dip) == (0.0, 0.0):
        defaulted.append("bedding_orientation")

    return SpecimenMetaResolution(
        meta=meta,
        source=source,
        header_path=header_path,
        registration=registration,
        defaulted_fields=tuple(defaulted),
    )
