"""Measured holder correction state.

VB6 keeps the accepted holder in the module-level ``Holder`` object
(``modMeasure``).  ``Measure_Read`` only replaces it *after* a complete
averaging run over validated blocks, and every specimen block subtracts
``Holder.Sample(j)`` — the holder's **baseline-adjusted** vectors, because
``MeasurementBlocks.AverageBlock`` averages ``BaselineAdjustedSample`` into the
new holder block.

RapidPy previously treated ``Holder`` as changer motion only and displayed
holder/induced statistics as ``N/A``.  This module makes the holder a real,
persisted, auditable measurement:

* a correction is installed **atomically** and only after a block passes every
  check;
* a rejected, aborted, timed-out, or unwritable holder block leaves the
  previous correction in place;
* every sample result can name the exact holder record and version it used;
* absent, stale, or otherwise invalid holder state blocks sample measurement.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping

from rapid_main.magnetometer import (
    BracketedMeasurementResult,
    Position4,
    Vector3,
    holder_average_vector,
    vector_magnitude,
)

HOLDER_SCHEMA = "rapidpy.holder_correction.v1"

#: A holder older than this is treated as stale.  VB6 re-measures the holder
#: every ``samples_between_holder`` specimens; eight hours is a conservative
#: upper bound for one working session.
DEFAULT_MAX_HOLDER_AGE_S = 8 * 60 * 60

#: VB6 applies one holder to both sample directions (``Measure_ReadSample``
#: subtracts ``Holder.Sample(j)`` regardless of ``isUp``).  ``strict`` is an
#: opt-in RapidPy tightening, not VB6 behavior.
DIRECTION_POLICIES = ("shared", "strict")

ZERO_POSITION4: Position4 = (
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
    (0.0, 0.0, 0.0),
)


class HolderStateError(RuntimeError):
    """Raised when holder state is missing, stale, or otherwise unusable."""


@dataclass(frozen=True)
class HolderMetrics:
    """Operator-facing holder quality numbers, in calibrated 2G raw units."""

    magnitude_raw: float = 0.0
    magnitude_emu: float = 0.0
    induced_magnitude_raw: float = 0.0
    asymmetry_ratio: float = 0.0
    drift_magnitude_raw: float = 0.0
    fischer_sd_deg: float = 0.0

    def to_dict(self) -> dict[str, float]:
        return {key: float(value) for key, value in asdict(self).items()}


@dataclass(frozen=True)
class HolderCorrection:
    """One accepted holder measurement and the evidence behind it."""

    holder_id: str
    hole: int = 0
    is_up: bool = True
    positions: Position4 = ZERO_POSITION4
    raw_positions: Position4 = ZERO_POSITION4
    holder_frame_positions: Position4 = ZERO_POSITION4
    zero_before: Vector3 = (0.0, 0.0, 0.0)
    zero_after: Vector3 = (0.0, 0.0, 0.0)
    range_label: str = ""
    range_factor: float = 1.0e-5
    axis_calibration: Vector3 = (1.0, 1.0, 1.0)
    measured_at_iso: str = ""
    software_version: str = ""
    operator: str = ""
    run_id: str = ""
    block_id: str = ""
    config_hash: str = ""
    validation_reason: str = ""
    validation_deltas: Vector3 = (0.0, 0.0, 0.0)
    averaging_cycles: int = 1
    metrics: HolderMetrics = field(default_factory=HolderMetrics)
    susceptibility_raw: float | None = None
    susceptibility_measured_at_iso: str = ""
    susceptibility_evidence_id: str = ""
    simulated: bool = False
    schema: str = HOLDER_SCHEMA
    collection_evidence_json: str = ""
    collection_sha256: str = ""

    @property
    def record_version(self) -> str:
        """Stable identity written next to every corrected sample result."""

        suffix = '#' + self.collection_sha256 if self.collection_sha256 else ''
        return f"{self.holder_id}@{self.measured_at_iso or 'unknown'}{suffix}"

    def age_seconds(self, now: datetime | None = None) -> float | None:
        measured = _parse_iso(self.measured_at_iso)
        if measured is None:
            return None
        reference = now or datetime.now(timezone.utc)
        if reference.tzinfo is None:
            reference = reference.replace(tzinfo=timezone.utc)
        return max(0.0, (reference - measured).total_seconds())

    def is_finite(self) -> bool:
        try:
            self.verify_collection()
        except HolderStateError:
            return False
        if (type(self.averaging_cycles) is not int or not 1 <= self.averaging_cycles <= 32767
                or not math.isfinite(self.range_factor) or self.range_factor <= 0):
            return False
        values: list[float] = []
        for vector in self.positions:
            values.extend(float(axis) for axis in vector)
        for vectors in (self.raw_positions, self.holder_frame_positions):
            for vector in vectors:
                values.extend(float(axis) for axis in vector)
        for vector in (self.zero_before, self.zero_after, self.axis_calibration, self.validation_deltas):
            values.extend(float(axis) for axis in vector)
        values.extend(self.metrics.to_dict().values())
        if not values or not all(math.isfinite(value) for value in values):
            return False
        return self.susceptibility_raw is None or math.isfinite(float(self.susceptibility_raw))

    def verify_collection(self) -> None:
        """Reproduce aggregates and verify their binding to the retained sources."""
        from .holder_collection import canonical, verified_top_fields
        if not self.collection_evidence_json and not self.collection_sha256:
            if self.averaging_cycles != 1:
                raise HolderStateError('Legacy multi-block holder has no collection evidence; remeasure it.')
            return
        try:
            expected = verified_top_fields(self.collection_evidence_json, self.collection_sha256)
            keys = json.loads(expected).keys()
            actual = {key: getattr(self, key) for key in keys}
            actual['metrics'] = self.metrics.to_dict()
            if canonical(actual) != expected:
                raise ValueError('holder correction disagrees with its acquisition collection')
        except (ValueError, TypeError, KeyError, OverflowError) as exc:
            raise HolderStateError(str(exc)) from exc

    def require_susceptibility(self) -> float:
        """Return the accepted scaled bridge value or fail before sample motion."""

        if self.susceptibility_raw is None or not math.isfinite(float(self.susceptibility_raw)):
            raise HolderStateError(
                "The accepted holder has no finite susceptibility correction. "
                "Measure holder susceptibility before running SUSC."
            )
        if not self.susceptibility_measured_at_iso:
            raise HolderStateError("Holder susceptibility has no measurement timestamp.")
        if not self.susceptibility_evidence_id:
            raise HolderStateError("Holder susceptibility has no acquisition evidence identity.")
        return float(self.susceptibility_raw)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["metrics"] = self.metrics.to_dict()
        payload["record_version"] = self.record_version
        for key in ("positions", "raw_positions", "holder_frame_positions"):
            payload[key] = [list(vector) for vector in getattr(self, key)]
        for key in ("zero_before", "zero_after", "axis_calibration", "validation_deltas"):
            payload[key] = list(getattr(self, key))
        return payload

    @classmethod
    def from_dict(cls, payload: Mapping[str, Any]) -> "HolderCorrection":
        metrics = payload.get("metrics") or {}
        known = {f for f in HolderMetrics.__dataclass_fields__}  # type: ignore[attr-defined]
        return cls(
            holder_id=str(payload.get("holder_id", "")),
            hole=int(payload.get("hole", 0) or 0),
            is_up=bool(payload.get("is_up", True)),
            positions=_coerce_position4(payload.get("positions")),
            raw_positions=_coerce_position4(payload.get("raw_positions")),
            holder_frame_positions=_coerce_position4(payload.get("holder_frame_positions")),
            zero_before=_coerce_vector(payload.get("zero_before")),
            zero_after=_coerce_vector(payload.get("zero_after")),
            range_label=str(payload.get("range_label", "")),
            range_factor=float(payload.get("range_factor", 1.0e-5) or 1.0e-5),
            axis_calibration=_coerce_vector(payload.get("axis_calibration"), default=(1.0, 1.0, 1.0)),
            measured_at_iso=str(payload.get("measured_at_iso", "")),
            software_version=str(payload.get("software_version", "")),
            operator=str(payload.get("operator", "")),
            run_id=str(payload.get("run_id", "")),
            block_id=str(payload.get("block_id", "")),
            config_hash=str(payload.get("config_hash", "")),
            validation_reason=str(payload.get("validation_reason", "")),
            validation_deltas=_coerce_vector(payload.get("validation_deltas")),
            averaging_cycles=int(payload.get("averaging_cycles", 1) or 1),
            metrics=HolderMetrics(
                **{key: float(value) for key, value in metrics.items() if key in known}
            ),
            susceptibility_raw=(
                None
                if payload.get("susceptibility_raw") is None
                else float(payload.get("susceptibility_raw"))
            ),
            susceptibility_measured_at_iso=str(
                payload.get("susceptibility_measured_at_iso", "")
            ),
            susceptibility_evidence_id=str(payload.get("susceptibility_evidence_id", "")),
            simulated=bool(payload.get("simulated", False)),
            schema=str(payload.get("schema", HOLDER_SCHEMA)),
            collection_evidence_json=payload.get('collection_evidence_json', ''),
            collection_sha256=payload.get('collection_sha256', ''),
        )

    @classmethod
    def from_collection(cls, blocks, *, holder_id: str, hole: int = 0, measured_at_iso: str):
        from .holder_collection import make_collection
        text, digest, positions, metrics, results = make_collection(
            blocks, holder_id=holder_id, hole=hole, measured_at_iso=measured_at_iso)
        last = results[-1].block
        base = cls.from_result(results[-1], holder_id=holder_id, hole=hole,
            measured_at_iso=measured_at_iso, averaging_cycles=len(results), positions_override=positions)
        correction = replace(base, metrics=HolderMetrics(**metrics),
            axis_calibration=(last.audit.axis_calibration_applied if last.audit else last.axis_calibration),
            collection_evidence_json=text, collection_sha256=digest)
        correction.verify_collection()
        return correction

    @classmethod
    def from_result(
        cls,
        result: BracketedMeasurementResult,
        *,
        holder_id: str,
        hole: int = 0,
        measured_at_iso: str = "",
        averaging_cycles: int = 1,
        positions_override: Position4 | None = None,
    ) -> "HolderCorrection":
        """Build a correction from a validated holder block.

        ``positions`` are the baseline-adjusted holder vectors, matching VB6
        ``MeasurementBlocks.AverageBlock`` which averages
        ``BaselineAdjustedSample`` into the new holder block.
        """

        block = result.block
        if block is None:
            raise HolderStateError("holder result is missing its source block")
        audit = block.audit
        positions = _coerce_position4(positions_override or result.baseline_adjusted_raw)
        magnitude_raw = vector_magnitude(holder_average_vector(positions))
        induced_magnitude = vector_magnitude(result.induced_raw)
        metrics = HolderMetrics(
            magnitude_raw=magnitude_raw,
            magnitude_emu=magnitude_raw * float(block.range_factor),
            induced_magnitude_raw=induced_magnitude,
            asymmetry_ratio=(induced_magnitude / magnitude_raw) if magnitude_raw else 0.0,
            drift_magnitude_raw=vector_magnitude(result.drift_raw),
            fischer_sd_deg=float(result.fischer_sd_deg),
        )
        return cls(
            holder_id=str(holder_id),
            hole=int(hole),
            is_up=bool(block.is_up),
            positions=positions,
            raw_positions=_coerce_position4(block.positions),
            holder_frame_positions=_coerce_position4(result.holder_frame_raw),
            zero_before=_coerce_vector(block.zero_before),
            zero_after=_coerce_vector(block.zero_after),
            range_label=getattr(audit, "range_label", ""),
            range_factor=float(block.range_factor),
            axis_calibration=_coerce_vector(
                getattr(audit, "axis_calibration_applied", (1.0, 1.0, 1.0)),
                default=(1.0, 1.0, 1.0),
            ),
            measured_at_iso=measured_at_iso
            or getattr(audit, "completed_iso", "")
            or datetime.now(timezone.utc).isoformat(),
            software_version=getattr(audit, "software_version", ""),
            operator=getattr(audit, "operator", ""),
            run_id=getattr(audit, "run_id", ""),
            block_id=getattr(audit, "block_id", ""),
            config_hash=getattr(audit, "config_hash", ""),
            validation_reason=result.validation.reason,
            validation_deltas=_coerce_vector(result.validation.deltas),
            averaging_cycles=max(1, int(averaging_cycles)),
            metrics=metrics,
            simulated=bool(getattr(audit, "simulated", False)),
        )


@dataclass(frozen=True)
class HolderStatus:
    """Summary rendered in the operator UI."""

    present: bool
    valid: bool
    reason: str = ""
    holder_id: str = ""
    record_version: str = ""
    age_seconds: float | None = None
    is_up: bool | None = None
    magnitude_emu: float | None = None
    asymmetry_ratio: float | None = None
    simulated: bool = False


class HolderStateStore:
    """Holds the active holder correction and installs replacements atomically."""

    def __init__(
        self,
        path: str | Path | None = None,
        *,
        max_age_s: float = DEFAULT_MAX_HOLDER_AGE_S,
        direction_policy: str = "shared",
        clock: Callable[[], datetime] | None = None,
        allow_simulated: bool = False,
    ) -> None:
        if direction_policy not in DIRECTION_POLICIES:
            raise ValueError(f"direction_policy must be one of {DIRECTION_POLICIES}")
        self._path = Path(path) if path else None
        self._max_age_s = float(max_age_s)
        self._direction_policy = direction_policy
        self._clock = clock or (lambda: datetime.now(timezone.utc))
        self._allow_simulated = bool(allow_simulated)
        self._current: HolderCorrection | None = None
        self._history: list[HolderCorrection] = []
        if self._path is not None and self._path.exists():
            self.load()

    @property
    def path(self) -> Path | None:
        return self._path

    @property
    def current(self) -> HolderCorrection | None:
        return self._current

    @property
    def history(self) -> tuple[HolderCorrection, ...]:
        return tuple(self._history)

    @property
    def direction_policy(self) -> str:
        return self._direction_policy

    def load(self) -> HolderCorrection | None:
        """Load a persisted correction, leaving state untouched on failure."""

        if self._path is None or not self._path.exists():
            return self._current
        try:
            payload = json.loads(self._path.read_text(encoding="utf-8"))
            correction = HolderCorrection.from_dict(payload)
        except (OSError, ValueError, TypeError):
            return self._current
        if not correction.holder_id or not correction.is_finite():
            return self._current
        self._current = correction
        return correction

    def install(self, correction: HolderCorrection) -> HolderCorrection:
        """Replace the active correction only after it is fully persisted.

        The previous correction stays active if validation or the write fails.
        """

        self._validate_installable(correction)
        if self._path is not None:
            _atomic_write_json(self._path, correction.to_dict())
        previous = self._current
        self._current = correction
        if previous is not None:
            self._history.append(previous)
        return correction

    def positions_for(self, is_up: bool = True) -> Position4:
        """Return the holder vectors to subtract, or zeros when unset.

        Callers that require a holder must call :meth:`require_valid` first;
        this accessor exists so a holder block itself (which subtracts nothing)
        and no-holder diagnostics stay simple.
        """

        correction = self._current
        if correction is None:
            return ZERO_POSITION4
        if self._direction_policy == "strict" and bool(correction.is_up) != bool(is_up):
            return ZERO_POSITION4
        return correction.positions

    def status(self, *, is_up: bool | None = None) -> HolderStatus:
        correction = self._current
        if correction is None:
            return HolderStatus(present=False, valid=False, reason="No holder measurement recorded.")
        reason = self._invalid_reason(correction, is_up=is_up)
        return HolderStatus(
            present=True,
            valid=not reason,
            reason=reason,
            holder_id=correction.holder_id,
            record_version=correction.record_version,
            age_seconds=correction.age_seconds(self._clock()),
            is_up=correction.is_up,
            magnitude_emu=correction.metrics.magnitude_emu,
            asymmetry_ratio=correction.metrics.asymmetry_ratio,
            simulated=correction.simulated,
        )

    def require_valid(self, *, is_up: bool | None = None) -> HolderCorrection:
        """Return the active correction or raise before a sample is measured."""

        correction = self._current
        if correction is None:
            raise HolderStateError(
                "No holder correction is available. Measure the holder before measuring samples."
            )
        reason = self._invalid_reason(correction, is_up=is_up)
        if reason:
            raise HolderStateError(reason)
        return correction

    def _invalid_reason(self, correction: HolderCorrection, *, is_up: bool | None) -> str:
        if not correction.holder_id:
            return "Holder correction has no holder identity."
        try:
            correction.verify_collection()
        except HolderStateError as exc:
            return str(exc)
        if not correction.is_finite():
            return "Holder correction contains non-finite values."
        if correction.simulated and not self._allow_simulated:
            return (
                "Holder correction came from a simulated backend and cannot be used "
                "for production measurement."
            )
        age = correction.age_seconds(self._clock())
        if age is None:
            return "Holder correction has no measurement timestamp."
        if self._max_age_s > 0 and age > self._max_age_s:
            return (
                f"Holder correction is stale: measured {age / 3600.0:.2f} h ago "
                f"(limit {self._max_age_s / 3600.0:.2f} h)."
            )
        if (
            is_up is not None
            and self._direction_policy == "strict"
            and bool(correction.is_up) != bool(is_up)
        ):
            measured = "up" if correction.is_up else "down"
            wanted = "up" if is_up else "down"
            return (
                f"Holder correction was measured with the sample {measured}, "
                f"but this step needs {wanted}."
            )
        return ""

    def _validate_installable(self, correction: HolderCorrection) -> None:
        if not isinstance(correction, HolderCorrection):
            raise HolderStateError("holder correction must be a HolderCorrection")
        if not correction.holder_id:
            raise HolderStateError("holder correction requires a holder identity")
        if not correction.measured_at_iso:
            raise HolderStateError("holder correction requires a measurement timestamp")
        correction.verify_collection()
        if not correction.is_finite():
            raise HolderStateError("holder correction contains non-finite values")


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    """Write JSON to a temporary file, fsync, then publish with a rename."""

    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f"{path.name}.tmp-{os.getpid()}")
    text = json.dumps(payload, indent=2, sort_keys=True) + "\n"
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


def _parse_iso(value: str) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _coerce_vector(values: Iterable[float] | None, *, default: Vector3 = (0.0, 0.0, 0.0)) -> Vector3:
    if values is None:
        return default
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise HolderStateError("expected exactly three axis values")
    return (items[0], items[1], items[2])


def _coerce_position4(values: Iterable[Iterable[float]] | None) -> Position4:
    if values is None:
        return ZERO_POSITION4
    items = tuple(_coerce_vector(vector) for vector in values)
    if len(items) != 4:
        raise HolderStateError("expected exactly four holder vectors")
    return (items[0], items[1], items[2], items[3])
