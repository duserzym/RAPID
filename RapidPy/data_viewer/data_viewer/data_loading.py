from __future__ import annotations

import csv
import hashlib
import json
import math
import re
import sys
from contextlib import ExitStack
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO, StringIO, TextIOWrapper
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile

try:
    from rapid_main.data_model import MeasurementStep, RmgRecord, SpecimenMeta
    from rapid_main.io.bundle_history import ARCHIVE_MANIFEST, ARCHIVE_PAYLOAD, is_bundle_archive, open_bundle_archive
    from rapid_main.io.rmg_reader import parse_rmg
    from rapid_main.io.sam_reader import read_sam, specimen_path
    from rapid_main.io.specimen_reader import parse_specimen, read_specimen
    from rapid_main.specimen_paths import contained_path, relative_specimen_path
except ModuleNotFoundError:
    rapidpy_root = Path(__file__).resolve().parents[2]
    rapid_main_root = rapidpy_root / "rapid_main"
    if str(rapid_main_root) not in sys.path:
        sys.path.insert(0, str(rapid_main_root))
    sys.modules.pop("rapid_main", None)
    from rapid_main.data_model import MeasurementStep, RmgRecord, SpecimenMeta
    from rapid_main.io.bundle_history import ARCHIVE_MANIFEST, ARCHIVE_PAYLOAD, is_bundle_archive, open_bundle_archive
    from rapid_main.io.rmg_reader import parse_rmg
    from rapid_main.io.sam_reader import read_sam, specimen_path
    from rapid_main.io.specimen_reader import parse_specimen, read_specimen
    from rapid_main.specimen_paths import contained_path, relative_specimen_path


@dataclass
class PaleointensityPoint:
    temperature_c: float
    step_kind: str
    nrm_remaining: float
    ptrm_gained: float
    label: str


@dataclass(frozen=True)
class SusceptibilityPoint:
    rmg_step: int
    treatment: str
    level: float | None
    level_kind: str
    susceptibility: float


@dataclass
class ViewerSpecimen:
    name: str
    meta: SpecimenMeta
    steps: list[MeasurementStep]
    experiment_type: str
    source_kind: str
    source_path: Path | None = None
    data_path: Path | None = None
    paleointensity_points: list[PaleointensityPoint] = field(default_factory=list)
    susceptibility_points: list[SusceptibilityPoint] = field(default_factory=list)
    rmg_path: Path | None = None
    coordinate_frames: tuple[str, ...] = ("specimen", "geographic", "tilt-corrected")
    cit_path: Path | None = None

    @property
    def magnitude_label(self) -> str:
        if self.source_kind in {"magic_dir", "measurement_archive"}:
            return "Moment (A·m²)"
        if self.source_kind == "demo":
            return "Magnitude (a.u.)"
        return "Recorded magnitude"


@dataclass
class ViewerDataset:
    source_kind: str
    source_path: Path | None
    specimens: list[ViewerSpecimen]

    @property
    def specimen_names(self) -> list[str]:
        return [specimen.name for specimen in self.specimens]


def load_input(path: str | Path) -> ViewerDataset:
    source = Path(path)
    if source.is_dir():
        return load_magic_directory(source)
    suffix = source.suffix.lower()
    if suffix == ".sam":
        return load_cit_sam(source)
    if suffix == ".csv":
        return load_csv_dataset(source)
    return load_cit_specimen(source)


def _susceptibility_points(records: Iterable[RmgRecord]) -> list[SusceptibilityPoint]:
    points = []
    for step, record in enumerate(records, start=1):
        treatment = record.step_type.upper()
        fields = record.raw_fields
        if len(fields) < 9 or not fields[8].strip():
            continue
        # The historical sidecar puts zero in unmeasured treatment rows.
        # A dedicated SUSC row can contain a valid zero result.
        if record.susceptibility == 0.0 and treatment != "SUSC":
            continue
        level_kind = ""
        if re.fullmatch(r"AF|AFZ|AFMAX|ATAF[XYZ]|UAF[XYZ][12]", treatment):
            level_kind = "af"
        elif treatment in {"TT", "TH", "TEMP"}:
            level_kind = "thermal"
        points.append(SusceptibilityPoint(
            step, record.step_type, record.step_value, level_kind, record.susceptibility,
        ))
    return points


def _load_rmg_sidecar(root: Path, specimen: ViewerSpecimen) -> tuple[bytes, list[RmgRecord]] | None:
    try:
        path = contained_path(root, specimen.name + ".rmg")
        if not path.is_file():
            return None
        data = path.read_bytes()
        records = parse_rmg(StringIO(data.decode("latin-1"), newline=""))
        specimen.susceptibility_points = _susceptibility_points(records)
        specimen.rmg_path = path
        return data, records
    except (OSError, ValueError, csv.Error):
        # A missing, unreadable or unsafe optional sidecar does not hide direction data.
        return None


def _cit_specimen(source: Path, data_path: Path, name: str, *, indexed: bool) -> ViewerSpecimen:
    meta, steps = read_specimen(data_path, name, validate_format=not indexed)
    specimen = ViewerSpecimen(
        name=name,
        meta=meta,
        steps=steps,
        experiment_type=detect_experiment_type(steps),
        source_kind="cit_sam" if indexed else "cit_specimen",
        source_path=source,
        data_path=data_path,
        paleointensity_points=build_paleointensity_points(steps),
    )
    _load_rmg_sidecar(source.parent, specimen)
    return specimen


def load_cit_specimen(path: str | Path) -> ViewerDataset:
    source = Path(path)
    specimen = _cit_specimen(source, source, source.name, indexed=False)
    return ViewerDataset("cit_specimen", source, [specimen])


def load_cit_sam(path: str | Path) -> ViewerDataset:
    sam_path = Path(path)
    specimens = [
        _cit_specimen(sam_path, specimen_path(sam_path, name), name, indexed=True)
        for name in read_sam(sam_path)
    ]
    if not specimens:
        raise ValueError(f"No specimen files were found from {sam_path.name}.")
    return ViewerDataset("cit_sam", sam_path, specimens)


def load_csv_dataset(path: str | Path) -> ViewerDataset:
    csv_path = Path(path)
    steps: list[MeasurementStep] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            label = _first_present(row, "label", "Label", default=str(len(steps)))
            sdx = _to_float(_first_present(row, "n", "N", default="0"))
            sdy = _to_float(_first_present(row, "e", "E", default="0"))
            sdz = _to_float(_first_present(row, "u", "U", default="0"))
            dec, inc, moment = _cartesian_to_dir(sdx, sdy, sdz)
            steps.append(
                MeasurementStep(
                    demag_label=label,
                    gdec=dec,
                    ginc=inc,
                    sdec=dec,
                    sinc=inc,
                    moment=moment,
                    error_angle=0.0,
                    crdec=dec,
                    crinc=inc,
                    sdx=sdx,
                    sdy=sdy,
                    sdz=sdz,
                )
            )
    if not steps:
        raise ValueError("CSV file appears empty or has unrecognised column names.")
    specimen_name = csv_path.stem
    meta = SpecimenMeta(name=specimen_name, comment="Imported CSV")
    return ViewerDataset(
        "csv",
        csv_path,
        [
            ViewerSpecimen(
                name=specimen_name,
                meta=meta,
                steps=steps,
                experiment_type=detect_experiment_type(steps),
                source_kind="csv",
                source_path=csv_path,
                data_path=csv_path,
                paleointensity_points=build_paleointensity_points(steps),
            )
        ],
    )


def load_magic_directory(path: str | Path) -> ViewerDataset:
    root = Path(path)
    archived = is_bundle_archive(root)
    source_kind = "measurement_archive" if archived else "magic_dir"
    source_file = root / (ARCHIVE_PAYLOAD if archived else "measurements.txt")
    specimens_by_name: dict[str, ViewerSpecimen] = {}
    row_evidence: dict[str, list[tuple[str, bool]]] = {}
    with ExitStack() as stack:
        archive = None
        magic_data = None
        if archived:
            archive = stack.enter_context(open_bundle_archive(root))
            if "measurements.txt" not in archive.namelist():
                raise ValueError(f"Could not find measurements.txt in the archive at {root}")
            magic_data = archive.read("measurements.txt")
        else:
            if not source_file.exists():
                raise ValueError(f"Could not find measurements.txt in {root}")
        index = _bundle_metadata(root, "artifact_index.json", archive)
        if not archived and isinstance(index, dict) and index.get("schema") == "rapidpy.measurement.artifact_index.v1":
            magic_data = source_file.read_bytes()
        if magic_data is not None:
            handle = stack.enter_context(TextIOWrapper(BytesIO(magic_data), encoding="utf-8-sig", newline=""))
        else:
            handle = stack.enter_context(source_file.open("r", encoding="utf-8-sig", newline=""))
        header_line = handle.readline().strip().lower()
        if "measurements" not in header_line:
            raise ValueError("measurements.txt does not start with a MagIC measurements header.")
        reader = csv.DictReader(handle, delimiter="\t")
        for row in reader:
            specimen_name = row.get("specimen", "").strip()
            if not specimen_name:
                continue
            specimen = specimens_by_name.get(specimen_name)
            if specimen is None:
                meta = SpecimenMeta(
                    name=specimen_name,
                    sample=row.get("sample", "").strip(),
                    site=row.get("site", "").strip(),
                    location=row.get("location", "").strip(),
                    comment="Imported MagIC measurements",
                )
                specimen = ViewerSpecimen(
                    name=specimen_name,
                    meta=meta,
                    steps=[],
                    experiment_type="Unknown",
                    source_kind=source_kind,
                    source_path=root,
                    data_path=source_file,
                    coordinate_frames=("geographic",),
                )
                specimens_by_name[specimen_name] = specimen
            specimen.steps.append(_measurement_row_to_step(row, specimen_name))
            if magic_data is not None:
                row_evidence.setdefault(specimen_name, []).append(
                    ((row.get("timestamp") or "").strip(), bool(_magic_csd_text(row).strip())),
                )

        members = set(archive.namelist()) if archive is not None else set()
        for specimen in specimens_by_name.values():
            sidecar = None
            if archive is not None:
                try:
                    member = relative_specimen_path(specimen.name + ".rmg").as_posix()
                    if member in members:
                        data = archive.read(member)
                        records = parse_rmg(StringIO(data.decode("latin-1"), newline=""))
                        specimen.susceptibility_points = _susceptibility_points(records)
                        sidecar = data, records
                except (ValueError, csv.Error):
                    pass
            else:
                sidecar = _load_rmg_sidecar(root, specimen)
            evidence = row_evidence.get(specimen.name)
            if magic_data is None or not evidence:
                continue
            legacy_archive = False
            if archive is not None and "artifact_index.json" not in members:
                # Older snapshots can authenticate their members through the archive manifest.
                provenance = _bundle_metadata(root, "provenance.json", archive)
                legacy_archive = (len(specimens_by_name) == 1 and isinstance(provenance, dict)
                                  and provenance.get("schema") == "rapidpy.measurement.provenance.v1"
                                  and provenance.get("specimen") == specimen.name
                                  and provenance.get("labels") == [step.demag_label for step in specimen.steps])
            try:
                name = relative_specimen_path(specimen.name).as_posix()
            except ValueError:
                continue
            if not (legacy_archive or _matches_publication_index(index, specimen.name, {"measurements.txt": magic_data})):
                continue
            _restore_cit_frames(root, archive, index, specimen, name, evidence, legacy_archive)
            rmg_name = name + ".rmg"
            if sidecar and (legacy_archive or _matches_publication_index(index, specimen.name, {rmg_name: sidecar[0]})):
                _restore_rmg_directions(specimen, sidecar[1], evidence)

    specimens = sorted(specimens_by_name.values(), key=lambda specimen: specimen.name)
    if not specimens:
        raise ValueError(f"No specimen rows were found in {source_file}")
    for specimen in specimens:
        specimen.experiment_type = detect_experiment_type(specimen.steps)
        specimen.paleointensity_points = build_paleointensity_points(specimen.steps)
    return ViewerDataset(source_kind, root, specimens)


def _bundle_metadata(root: Path, name: str, archive: ZipFile | None) -> object:
    try:
        data = archive.read(name) if archive is not None else contained_path(root, name).read_bytes()
        return json.loads(data)
    except (KeyError, OSError, ValueError):
        return None


def _restore_cit_frames(root, archive, index, specimen, name, evidence, legacy_archive):
    try:
        path = contained_path(root, name) if archive is None else None
        data = path.read_bytes() if path is not None else archive.read(name)
        specimen.cit_path = path
        if not (legacy_archive or _matches_publication_index(index, specimen.name, {name: data})):
            return
        _, original = parse_specimen(data, specimen.name, validate_format=True)
    except (KeyError, OSError, ValueError):
        return
    if len(original) != len(evidence) or any(
        saved.demag_label != step.demag_label or not stamp
        or stamp.replace("T", " ", 1) != saved.timestamp.strftime("%Y-%m-%d %H:%M:%S")
        for saved, step, (stamp, _) in zip(original, specimen.steps, evidence)
    ):
        return
    for saved, step in zip(original, specimen.steps):
        step.crdec, step.crinc = saved.crdec, saved.crinc
        step.sdec, step.sinc = saved.sdec, saved.sinc
    specimen.coordinate_frames = ("specimen", "geographic", "tilt-corrected")


def _restore_rmg_directions(specimen, records, evidence):
    if len(records) != len(evidence) or any(
        not stamp or stamp.replace("T", " ", 1) != record.timestamp
        for (stamp, _), record in zip(evidence, records)
    ):
        return
    for step, record, (_, has_csd) in zip(specimen.steps, records, evidence):
        if not has_csd and record.csd is not None and 0 < record.csd < 180:
            step.error_angle = record.csd
    # RMG preserves greater precision than the legacy specimen text.
    if any(record.core_dec is None or record.core_inc is None for record in records):
        return
    for step, record in zip(specimen.steps, records):
        step.crdec, step.crinc = record.core_dec, record.core_inc
    if "specimen" not in specimen.coordinate_frames:
        specimen.coordinate_frames = ("specimen", "geographic")


def _matches_publication_index(index: object, specimen: str, sources: dict[str, bytes]) -> bool:
    if (not isinstance(index, dict) or index.get("schema") != "rapidpy.measurement.artifact_index.v1"
            or index.get("aborted") is not False or index.get("sample") != specimen
            or index.get("diagnostics_pending") is True
            or not isinstance(index.get("artifacts"), list)):
        return False
    for name, data in sources.items():
        entries = [entry for entry in index["artifacts"]
                   if isinstance(entry, dict) and entry.get("relative_path") == name]
        if (len(entries) != 1 or entries[0].get("exists") is not True
                or entries[0].get("sha256") != hashlib.sha256(data).hexdigest()):
            return False
    return True


def watch_paths_for_dataset(dataset: ViewerDataset) -> list[Path]:
    paths: list[Path] = []
    if dataset.source_path is not None:
        paths.append(Path(dataset.source_path))
        if dataset.source_kind == "measurement_archive":
            paths.append(Path(dataset.source_path) / ARCHIVE_MANIFEST)
        elif dataset.source_kind == "cit_specimen":
            paths.append(Path(dataset.source_path).parent)
    for specimen in dataset.specimens:
        if specimen.data_path is not None:
            paths.append(Path(specimen.data_path))
        if specimen.rmg_path is not None:
            paths.append(specimen.rmg_path)
        if specimen.cit_path is not None:
            paths.append(specimen.cit_path)

    unique: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        key = str(path.resolve()) if path.exists() else str(path)
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique


def detect_experiment_type(steps: Iterable[MeasurementStep]) -> str:
    labels = [step.demag_label.strip().upper() for step in steps]
    has_izzi = any(_is_izzi_label(label) for label in labels)
    has_af = any(step.treat_ac_field_T() is not None for step in steps)
    has_thermal = any(step.treat_temp_K() is not None for step in steps)
    if has_izzi:
        return "IZZI Thellier"
    if has_af and has_thermal:
        return "AF + thermal"
    if has_af:
        return "AF"
    if has_thermal:
        return "Thermal"
    return "Unknown"


def build_paleointensity_points(steps: Iterable[MeasurementStep]) -> list[PaleointensityPoint]:
    ordered = list(steps)
    if not ordered:
        return []
    nrm0 = ordered[0].moment if ordered[0].moment > 0 else max((step.moment for step in ordered), default=1.0)
    ptrm_anchor = None
    points: list[PaleointensityPoint] = []
    for step in ordered:
        label = step.demag_label.strip().upper()
        if not _is_izzi_label(label):
            continue
        temperature_c = _extract_first_number(label)
        if temperature_c is None:
            temperature_c = (step.treat_temp_K() or 273.15) - 273.15
        kind = _classify_paleointensity_label(label)
        if kind == "ZF":
            points.append(
                PaleointensityPoint(
                    temperature_c=temperature_c,
                    step_kind=kind,
                    nrm_remaining=step.moment / nrm0 if nrm0 else 0.0,
                    ptrm_gained=ptrm_anchor if ptrm_anchor is not None else 0.0,
                    label=step.demag_label,
                )
            )
        elif kind == "IF":
            ptrm_anchor = 1.0 - (step.moment / nrm0 if nrm0 else 0.0)
            points.append(
                PaleointensityPoint(
                    temperature_c=temperature_c,
                    step_kind=kind,
                    nrm_remaining=step.moment / nrm0 if nrm0 else 0.0,
                    ptrm_gained=ptrm_anchor,
                    label=step.demag_label,
                )
            )
        elif kind == "PTRM":
            points.append(
                PaleointensityPoint(
                    temperature_c=temperature_c,
                    step_kind=kind,
                    nrm_remaining=step.moment / nrm0 if nrm0 else 0.0,
                    ptrm_gained=ptrm_anchor if ptrm_anchor is not None else 0.0,
                    label=step.demag_label,
                )
            )
    return points


def _measurement_row_to_step(row: dict[str, str], specimen_name: str) -> MeasurementStep:
    label = _label_from_measurement_row(row, specimen_name)
    mx = _to_float(row.get("magn_x"))
    my = _to_float(row.get("magn_y"))
    # MagIC exports produced by RAPID use north/east/down. The viewer uses up-positive vectors.
    vector_dec, vector_inc, moment = _cartesian_to_dir(mx, my, -_to_float(row.get("magn_z")))
    if moment == 0.0:
        moment = _to_float(_first_present(row, "magn_moment", "measurement_magn_moment", default="0"))
    dec = _to_float(_first_present(row, "dir_dec", "measurement_dec", default=str(vector_dec)))
    inc = _to_float(_first_present(row, "dir_inc", "measurement_inc", default=str(vector_inc)))
    mx, my, mz = _dir_to_cartesian(dec, inc, moment)
    error_angle = _to_float(_magic_csd_text(row))
    timestamp = _parse_timestamp(row.get("timestamp", ""))

    return MeasurementStep(
        demag_label=label,
        gdec=dec,
        ginc=inc,
        sdec=dec,
        sinc=inc,
        moment=moment,
        error_angle=error_angle,
        crdec=dec,
        crinc=inc,
        sdx=mx,
        sdy=my,
        sdz=mz,
        operator=row.get("analysts", "").strip()[:8],
        timestamp=timestamp,
    )


def _magic_csd_text(row: dict[str, str]) -> str:
    return _first_present(row, "dir_csd", "measurement_csd", "meas_csd", "error_angle")

def _label_from_measurement_row(row: dict[str, str], specimen_name: str) -> str:
    measurement_name = row.get("measurement", "").strip()
    if measurement_name:
        prefix = f"{specimen_name}-"
        if measurement_name.upper().startswith(prefix.upper()):
            return re.sub(r"-repeat-(?:[2-9]|[1-9][0-9]+)$", "", measurement_name[len(prefix):])
        return measurement_name.split("-")[-1]

    method_codes = row.get("method_codes", "").upper()
    ac_field_t = _to_float(row.get("treat_ac_field"))
    temp_k = _to_float(row.get("treat_temp"))
    dc_field_t = _to_float(row.get("treat_dc_field"))
    description = " ".join(
        value for value in (row.get("description", ""), row.get("experiment", ""), row.get("sequence", "")) if value
    ).upper()

    if any(token in description for token in ("IZ", "ZI", "ZF", "IF", "PTRM")):
        for token in ("PTRM", "ZF", "IF", "ZI", "IZ"):
            if token in description:
                if temp_k > 0:
                    return f"{token}{int(round(temp_k - 273.15))}"
                return token

    if "LP-PI-TRM" in method_codes and temp_k > 0:
        base = "IF" if dc_field_t > 0 else "ZF"
        return f"{base}{int(round(temp_k - 273.15))}"
    if ac_field_t > 0:
        return f"AF{int(round(ac_field_t * 1e3))}"
    if temp_k > 273.15:
        return f"TT{int(round(temp_k - 273.15))}"
    return "NRM"


def _dir_to_cartesian(dec_deg: float, inc_deg: float, moment: float) -> tuple[float, float, float]:
    dec = math.radians(dec_deg)
    inc = math.radians(inc_deg)
    horiz = moment * math.cos(inc)
    return (
        horiz * math.cos(dec),
        horiz * math.sin(dec),
        -moment * math.sin(inc),
    )


def _cartesian_to_dir(x: float, y: float, z: float) -> tuple[float, float, float]:
    moment = math.sqrt(x * x + y * y + z * z)
    if moment <= 0:
        return 0.0, 0.0, 0.0
    dec = math.degrees(math.atan2(y, x)) % 360.0
    inc = math.degrees(math.asin(-z / moment))
    return dec, inc, moment


def _first_present(row: dict[str, str], *keys: str, default: str = "") -> str:
    for key in keys:
        value = row.get(key)
        if value not in (None, ""):
            return value
    return default


def _to_float(value: str | None) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def _parse_timestamp(value: str) -> datetime:
    if not value:
        return datetime.now()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return datetime.now()


def _is_izzi_label(label: str) -> bool:
    return _classify_paleointensity_label(label) != ""


def _classify_paleointensity_label(label: str) -> str:
    upper = label.upper()
    for token in ("PTRM", "ZF", "IF", "ZI", "IZ"):
        if upper.startswith(token):
            return token
    return ""


def _extract_first_number(label: str) -> float | None:
    digits = []
    started = False
    for char in label:
        if char.isdigit() or (char == "." and started):
            digits.append(char)
            started = True
        elif started:
            break
    if not digits:
        return None
    try:
        return float("".join(digits))
    except ValueError:
        return None
