from __future__ import annotations

import csv
import math
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

try:
    from rapid_main.data_model import MeasurementStep, SpecimenMeta
    from rapid_main.io.sam_reader import read_sam, specimen_path
    from rapid_main.io.specimen_reader import read_specimen
except ModuleNotFoundError:
    rapidpy_root = Path(__file__).resolve().parents[2]
    rapid_main_root = rapidpy_root / "rapid_main"
    if str(rapid_main_root) not in sys.path:
        sys.path.insert(0, str(rapid_main_root))
    sys.modules.pop("rapid_main", None)
    from rapid_main.data_model import MeasurementStep, SpecimenMeta
    from rapid_main.io.sam_reader import read_sam, specimen_path
    from rapid_main.io.specimen_reader import read_specimen


@dataclass
class PaleointensityPoint:
    temperature_c: float
    step_kind: str
    nrm_remaining: float
    ptrm_gained: float
    label: str


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
    raise ValueError(f"Unsupported input: {source}")


def load_cit_sam(path: str | Path) -> ViewerDataset:
    sam_path = Path(path)
    specimens: list[ViewerSpecimen] = []
    for name in read_sam(sam_path):
        specimen_file = specimen_path(sam_path, name)
        meta, steps = read_specimen(specimen_file, name)
        specimens.append(
            ViewerSpecimen(
                name=name,
                meta=meta,
                steps=steps,
                experiment_type=detect_experiment_type(steps),
                source_kind="cit_sam",
                source_path=sam_path,
                data_path=specimen_file,
                paleointensity_points=build_paleointensity_points(steps),
            )
        )
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
    measurements_path = root / "measurements.txt"
    if not measurements_path.exists():
        raise ValueError(f"Could not find measurements.txt in {root}")

    specimens_by_name: dict[str, ViewerSpecimen] = {}
    with measurements_path.open("r", encoding="utf-8-sig", newline="") as handle:
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
                    source_kind="magic_dir",
                    source_path=root,
                    data_path=measurements_path,
                )
                specimens_by_name[specimen_name] = specimen
            specimen.steps.append(_measurement_row_to_step(row, specimen_name))

    specimens = sorted(specimens_by_name.values(), key=lambda specimen: specimen.name)
    if not specimens:
        raise ValueError(f"No specimen rows were found in {measurements_path}")
    for specimen in specimens:
        specimen.experiment_type = detect_experiment_type(specimen.steps)
        specimen.paleointensity_points = build_paleointensity_points(specimen.steps)
    return ViewerDataset("magic_dir", root, specimens)


def watch_paths_for_dataset(dataset: ViewerDataset) -> list[Path]:
    paths: list[Path] = []
    if dataset.source_path is not None:
        paths.append(Path(dataset.source_path))
    for specimen in dataset.specimens:
        if specimen.data_path is not None:
            paths.append(Path(specimen.data_path))

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
    mz = _to_float(row.get("magn_z"))
    if mx == 0.0 and my == 0.0 and mz == 0.0:
        moment = _to_float(_first_present(row, "magn_moment", "measurement_magn_moment", default="0"))
        dec = _to_float(_first_present(row, "dir_dec", "measurement_dec", default="0"))
        inc = _to_float(_first_present(row, "dir_inc", "measurement_inc", default="0"))
        mx, my, mz = _dir_to_cartesian(dec, inc, moment)
    else:
        dec, inc, moment = _cartesian_to_dir(mx, my, mz)

    geographic_dec = _to_float(_first_present(row, "dir_dec", "measurement_dec", default=str(dec)))
    geographic_inc = _to_float(_first_present(row, "dir_inc", "measurement_inc", default=str(inc)))
    error_angle = _to_float(_first_present(row, "dir_csd", "meas_csd", "error_angle", default="0"))
    timestamp = _parse_timestamp(row.get("timestamp", ""))

    return MeasurementStep(
        demag_label=label,
        gdec=geographic_dec,
        ginc=geographic_inc,
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


def _label_from_measurement_row(row: dict[str, str], specimen_name: str) -> str:
    measurement_name = row.get("measurement", "").strip()
    if measurement_name:
        prefix = f"{specimen_name}-"
        if measurement_name.upper().startswith(prefix.upper()):
            return measurement_name[len(prefix):]
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
        moment * math.sin(inc),
    )


def _cartesian_to_dir(x: float, y: float, z: float) -> tuple[float, float, float]:
    moment = math.sqrt(x * x + y * y + z * z)
    if moment <= 0:
        return 0.0, 0.0, 0.0
    dec = math.degrees(math.atan2(y, x)) % 360.0
    inc = math.degrees(math.asin(z / moment))
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