"""
data_model.py — Core dataclasses for RAPID v4 specimen data.

Mirrors the VB6 Sample.cls data structures with added MagIC-derived fields.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


# ---------------------------------------------------------------------------
# Specimen metadata (header lines of the specimen file)
# ---------------------------------------------------------------------------

@dataclass
class SpecimenMeta:
    """Orientation and identification info read from the specimen file header."""
    name: str                          # specimen name (== filename)
    comment: str = ""                  # line 1 of specimen file (free text)
    core_plate_strike: float = 0.0     # °
    core_plate_dip: float = 0.0        # °
    bedding_strike: float = 0.0        # °
    bedding_dip: float = 0.0           # °
    volume: float = 1.0                # cm³
    fold_axis: Optional[float] = None  # ° (optional, col 39-43)
    fold_plunge: Optional[float] = None  # ° (optional, col 45-49)
    # Hierarchical MagIC names (filled in by caller if known)
    sample: str = ""
    site: str = ""
    location: str = ""


# ---------------------------------------------------------------------------
# Sample index registry and queue-order metadata
# ---------------------------------------------------------------------------

@dataclass
class SampleIndexRegistration:
    """One entry in a VB6-style sample-index (.sam) registry.

    The original VB6 registry stores specimen names and optional sample-set
    context lines. This class preserves the practical metadata needed by modern
    queue tooling while remaining tolerant of sparse legacy files.
    """

    specimen_name: str
    sample_set: str = ""
    location: str = ""
    formation: str = ""
    depth_cm: str = ""
    order: int = 0


@dataclass
class SampleIndexRegistrations:
    """Ordered collection wrapper for :class:`SampleIndexRegistration`."""

    entries: list[SampleIndexRegistration]

    @property
    def names(self) -> list[str]:
        """Return registry specimen names in stored order."""
        return [entry.specimen_name for entry in self.entries]


# ---------------------------------------------------------------------------
# Per-step measurement record
# ---------------------------------------------------------------------------

@dataclass
class MeasurementStep:
    """One demagnetisation/measurement step, exactly as written by VB6 WriteData()."""
    demag_label: str          # e.g. "NRM", "AF20", "TT400", "IRM1000", "ARM"
    gdec: float               # geographic declination (°)
    ginc: float               # geographic inclination (°)
    sdec: float               # specimen declination (°)
    sinc: float               # specimen inclination (°)
    moment: float             # moment in emu
    error_angle: float        # CSD / error angle (°)
    crdec: float              # core declination (°)
    crinc: float              # core inclination (°)
    sdx: float                # specimen Cartesian X (emu)
    sdy: float                # specimen Cartesian Y (emu)
    sdz: float                # specimen Cartesian Z (emu)
    operator: str = ""        # operator name (≤ 8 chars)
    timestamp: datetime = field(default_factory=datetime.now)

    # -----------------------------------------------------------------------
    # Derived MagIC fields
    # -----------------------------------------------------------------------

    def magn_moment_Am2(self) -> float:
        """Moment in A·m²  (1 emu = 1e-3 A·m²)."""
        return self.moment * 1e-3

    def treat_ac_field_T(self) -> Optional[float]:
        """AF peak field in Tesla (1 Oe = 1e-4 T).  None if not an AF step."""
        lbl = self.demag_label.upper()
        match = re.match(r'^(?:AF|AFZ|AFMAX)(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) * 1e-4  # Oe → T
        return None

    def treat_dc_field_T(self) -> Optional[float]:
        """DC bias field in Tesla.  Encoded as 'ARM<field>_<bias>' convention."""
        lbl = self.demag_label.upper()
        # ARM steps may encode DC bias after underscore: e.g. "ARM100_1"
        match = re.match(r'^ARM\d+_(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) * 1e-4  # Oe → T
        return None

    def treat_temp_K(self) -> Optional[float]:
        """Treatment temperature in Kelvin.  None if not a thermal step."""
        lbl = self.demag_label.upper()
        match = re.match(r'^(?:TT|TH|TEMP)(\d+\.?\d*)$', lbl)
        if match:
            return float(match.group(1)) + 273.15  # °C → K
        return None

    def treat_ac_field_mT(self) -> Optional[float]:
        """AF field in mT (for display).  1 Oe = 0.1 mT."""
        t = self.treat_ac_field_T()
        return t * 1e3 if t is not None else None

    def method_codes(self) -> str:
        """MagIC LP-* method code(s) for this step."""
        from rapid_main.io.magic_method_codes import label_to_method_codes  # lazy import
        return label_to_method_codes(self.demag_label)

    def measurement_name(self, specimen_name: str) -> str:
        """Unique MagIC measurement identifier."""
        return f"{specimen_name}-{self.demag_label}"


# ---------------------------------------------------------------------------
# Measurement sequence block metadata
# ---------------------------------------------------------------------------

@dataclass
class MeasurementBlock:
    """Named group of measurement/treatment labels.

    VB6 represented rock-magnetic and measurement programs as block-oriented
    structures. RapidPy still executes flattened labels today, but this model
    preserves the block boundary and metadata needed for queue/output parity.
    """

    name: str
    labels: list[str]
    block_type: str = "measurement"
    metadata: dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.name = str(self.name).strip() or "Untitled block"
        self.block_type = str(self.block_type).strip() or "measurement"
        self.labels = [str(label).strip() for label in self.labels if str(label).strip()]
        self.metadata = {str(k): str(v) for k, v in self.metadata.items()}

    @property
    def step_count(self) -> int:
        return len(self.labels)

    def to_queue_labels(self) -> list[str]:
        """Return a copy of the executable labels in this block."""
        return list(self.labels)


@dataclass
class MeasurementBlocks:
    """Ordered collection of :class:`MeasurementBlock` objects."""

    blocks: list[MeasurementBlock]

    @property
    def step_count(self) -> int:
        return sum(block.step_count for block in self.blocks)

    def to_queue_labels(self) -> list[str]:
        """Flatten blocks into the label list consumed by the current runner."""
        labels: list[str] = []
        for block in self.blocks:
            labels.extend(block.to_queue_labels())
        return labels

    def block_names(self) -> list[str]:
        return [block.name for block in self.blocks]


@dataclass
class RockmagStep:
    """One rock-magnetic routine step parsed from a sequence label."""

    label: str
    family: str = ""
    value: float | None = None
    unit: str = ""
    metadata: dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.label = str(self.label).strip().upper()
        if not self.label:
            raise ValueError("RockmagStep label is required")
        self.family = (self.family or self._family_from_label(self.label)).strip().upper()
        if self.value is None:
            self.value, inferred_unit = self._value_and_unit_from_label(self.label)
            if not self.unit:
                self.unit = inferred_unit
        self.unit = str(self.unit).strip()
        self.metadata = {str(k): str(v) for k, v in self.metadata.items()}

    @classmethod
    def from_label(cls, label: str, **metadata: object) -> "RockmagStep":
        return cls(label=label, metadata={str(k): str(v) for k, v in metadata.items()})

    @staticmethod
    def _family_from_label(label: str) -> str:
        if label == "NRM":
            return "NRM"
        if label == "SUSC":
            return "SUSCEPTIBILITY"
        if label == "IRM-BF":
            return "BACKFIELD"
        match = re.match(r"^([A-Z]+)", label)
        return match.group(1) if match else "UNKNOWN"

    @staticmethod
    def _value_and_unit_from_label(label: str) -> tuple[float | None, str]:
        match = re.search(r"(-?\d+(?:\.\d+)?)", label)
        if not match:
            return None, ""
        value = float(match.group(1))
        if label.startswith("AF"):
            return value, "mT"
        if label.startswith("IRM") or label.startswith("ARM"):
            return value, "G"
        if label.startswith("RRM"):
            return value, "rps"
        if label.startswith(("TT", "TH", "TEMP")):
            return value, "C"
        return value, ""


@dataclass
class RockmagSteps:
    """Ordered rock-magnetic routine step collection."""

    steps: list[RockmagStep]
    routine_name: str = "Rockmag routine"

    @classmethod
    def from_labels(cls, labels: list[str], *, routine_name: str = "Rockmag routine") -> "RockmagSteps":
        return cls([RockmagStep.from_label(label) for label in labels], routine_name=routine_name)

    @property
    def labels(self) -> list[str]:
        return [step.label for step in self.steps]

    @property
    def step_count(self) -> int:
        return len(self.steps)

    def families(self) -> list[str]:
        seen: list[str] = []
        for step in self.steps:
            if step.family not in seen:
                seen.append(step.family)
        return seen

    def to_measurement_block(self) -> MeasurementBlock:
        return MeasurementBlock(
            name=self.routine_name,
            labels=self.labels,
            block_type="rockmag",
            metadata={"families": ",".join(self.families())},
        )


# ---------------------------------------------------------------------------
# Transverse probe / angle-vs-field data models
# ---------------------------------------------------------------------------

@dataclass
class AngleVsFieldPoint:
    """One transverse angle/field observation."""

    angle_deg: float
    field_value: float
    weight: float = 1.0

    def __post_init__(self) -> None:
        self.angle_deg = float(self.angle_deg)
        self.field_value = float(self.field_value)
        self.weight = max(0.0, float(self.weight))


@dataclass
class AngleVsFieldCollection:
    """Ordered collection of transverse angle/field points."""

    points: list[AngleVsFieldPoint]

    def __post_init__(self) -> None:
        self.points = sorted(self.points, key=lambda point: point.angle_deg)

    @property
    def angles(self) -> list[float]:
        return [point.angle_deg for point in self.points]

    @property
    def fields(self) -> list[float]:
        return [point.field_value for point in self.points]

    def peak_point(self) -> AngleVsFieldPoint | None:
        if not self.points:
            return None
        return max(self.points, key=lambda point: point.field_value)

    def weighted_center_angle(self) -> float | None:
        weighted = [(point.angle_deg, point.weight) for point in self.points if point.weight > 0]
        if not weighted:
            return None
        total_weight = sum(weight for _angle, weight in weighted)
        return sum(angle * weight for angle, weight in weighted) / total_weight


@dataclass(frozen=True)
class ProbeAngleOptimizationResult:
    angle_deg: float
    field_value: float
    confidence: float
    method: str = "peak-field"


class ProbeAngleOptimizer:
    """Deterministic transverse probe angle optimizer.

    This maps the VB6 optimizer class at the data/model level. Hardware
    auto-positioning remains a separate workflow gate.
    """

    def __init__(self, collection: AngleVsFieldCollection) -> None:
        self.collection = collection

    def optimize(self) -> ProbeAngleOptimizationResult | None:
        peak = self.collection.peak_point()
        if peak is None:
            return None
        fields = self.collection.fields
        if not fields:
            confidence = 0.0
        else:
            low = min(fields)
            high = max(fields)
            confidence = 1.0 if high == 0 else max(0.0, min(1.0, (high - low) / abs(high)))
        return ProbeAngleOptimizationResult(
            angle_deg=peak.angle_deg,
            field_value=peak.field_value,
            confidence=confidence,
        )


# ---------------------------------------------------------------------------
# RMG sidecar record (susceptibility vs. demagnetisation)
# ---------------------------------------------------------------------------

@dataclass
class RmgRecord:
    """One line from a `.rmg` sidecar file (comma-delimited, 9+ columns)."""
    step_type: str          # column 0: AF, TT, IRM, ARM, AFz, AFmax
    step_value: float       # column 1: field (Oe) or temperature (°C)
    susceptibility: float   # column 8: susceptibility (emu/Oe)
    raw_fields: list = field(default_factory=list)  # all raw columns for round-trip fidelity
