"""Specimen thermal-treatment planning without claiming furnace control.

The legacy ``modThermal`` module protects the AF coil temperature sensors; it
does not implement a specimen furnace protocol.  RapidPy understands specimen
thermal labels such as ``TT400`` in IO and runtime-estimation paths, so this
module provides validated planning and auditable run evidence while keeping the
unknown live furnace boundary fail-closed.
"""
from __future__ import annotations

import json
import math
import os
import re
import tempfile
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from rapid_main.data_model import MeasurementBlock


_THERMAL_LABEL_RE = re.compile(r"^(TT|TH|TEMP)\s*(\d+(?:\.\d+)?)$", re.IGNORECASE)


@dataclass(frozen=True)
class ExternalThermalTreatment:
    """Operator attestation for a completed furnace step and cooled specimen."""

    specimen: str
    label: str
    operator: str
    furnace_log: str
    peak_temperature_c: float
    hold_seconds: int
    handling_temperature_c: float
    completed_at: str
    evidence_id: str = field(default_factory=lambda: "thermal-" + uuid.uuid4().hex)

    def __post_init__(self):
        if any(not isinstance(text, str) or not text.strip() for text in
               (self.specimen, self.label, self.operator, self.furnace_log, self.completed_at)):
            raise ValueError("Specimen, operator, furnace-log reference, label, and completion time are required.")
        if not _THERMAL_LABEL_RE.fullmatch(self.label):
            raise ValueError("Select a thermal step from the reviewed plan.")
        if not re.fullmatch(r"thermal-[0-9a-f]{32}", self.evidence_id):
            raise ValueError("Invalid external treatment evidence ID.")
        for value in (self.peak_temperature_c, self.handling_temperature_c):
            if type(value) not in (float, int) or not math.isfinite(value) or value < 0:
                raise ValueError("Enter the recorded peak and handling temperatures.")
        if self.handling_temperature_c > self.peak_temperature_c:
            raise ValueError("Handling temperature exceeds the recorded peak.")
        if type(self.hold_seconds) is not int or self.hold_seconds < 0:
            raise ValueError("Enter the recorded dwell time in seconds.")
        timestamp = datetime.fromisoformat(self.completed_at)
        if timestamp.tzinfo is None or timestamp > datetime.now(timezone.utc):
            raise ValueError("Enter a completed furnace time with a time zone.")

    def to_artifact(self):
        return {"schema": "rapidpy.thermal.external_treatment.v1",
                "evidence_basis": "operator attestation referencing an external furnace log", **asdict(self)}


def bind_external_thermal_treatment(context, treatment, *, specimen, operator, labels):
    if (context.get("procedure_id") != "thermal/routine-planning" or list(labels) != [treatment.label]
            or treatment.specimen != specimen or treatment.operator != operator):
        raise ValueError("The external treatment does not match the reviewed step, specimen, or operator.")
    planned = list(context.get("queue_labels", ()))
    steps = context.get("steps", ())
    if [step["label"] for step in steps] != planned or len(set(planned)) != len(planned) or treatment.label not in planned:
        raise ValueError("External treatment requires a reviewed plan with distinct thermal steps.")
    step = steps[planned.index(treatment.label)]
    limits = ThermalSafetyLimits(**context["limits"])
    if treatment.peak_temperature_c > limits.max_temperature_c:
        raise ValueError("The recorded peak exceeds the reviewed plan limit.")
    if treatment.handling_temperature_c > limits.cool_to_c:
        raise ValueError("The specimen is above the reviewed cooldown limit.")
    if treatment.hold_seconds < step["hold_seconds"]:
        raise ValueError("The recorded dwell time is shorter than the reviewed step.")
    return {**context, "external_treatment": treatment.to_artifact()}


@dataclass(frozen=True)
class ThermalSafetyLimits:
    """Operator-configurable thermal planning limits."""

    max_temperature_c: float = 700.0
    ambient_temperature_c: float = 25.0
    cool_to_c: float = 50.0
    ramp_rate_c_per_min: float = 20.0
    default_hold_seconds: int = 600

    def __post_init__(self) -> None:
        if not all(math.isfinite(value) for value in (
                self.max_temperature_c, self.ambient_temperature_c, self.cool_to_c, self.ramp_rate_c_per_min)):
            raise ValueError("thermal planning limits must be finite")
        if self.max_temperature_c <= self.ambient_temperature_c:
            raise ValueError("thermal max temperature must exceed ambient temperature")
        if self.cool_to_c < self.ambient_temperature_c:
            raise ValueError("thermal cool-to temperature must be at or above ambient")
        if self.ramp_rate_c_per_min <= 0.0:
            raise ValueError("thermal ramp rate must be positive")
        if self.default_hold_seconds < 0:
            raise ValueError("thermal hold seconds must be non-negative")


@dataclass(frozen=True)
class ThermalStep:
    """One thermal treatment step."""

    temperature_c: float
    prefix: str = "TT"
    hold_seconds: int | None = None

    @classmethod
    def from_label(cls, label: str, *, hold_seconds: int | None = None) -> "ThermalStep":
        match = _THERMAL_LABEL_RE.match(str(label).strip())
        if not match:
            raise ValueError(f"invalid thermal label: {label!r}")
        return cls(
            temperature_c=float(match.group(2)),
            prefix=match.group(1).upper(),
            hold_seconds=hold_seconds,
        )

    def __post_init__(self) -> None:
        if not math.isfinite(self.temperature_c) or self.temperature_c < 0.0:
            raise ValueError("thermal temperature must be non-negative")
        if self.hold_seconds is not None and self.hold_seconds < 0:
            raise ValueError("thermal hold seconds must be non-negative")
        object.__setattr__(self, "prefix", str(self.prefix).strip().upper() or "TT")

    @property
    def label(self) -> str:
        value = int(round(self.temperature_c)) if self.temperature_c == round(self.temperature_c) else self.temperature_c
        text = str(value).rstrip("0").rstrip(".") if isinstance(value, float) else str(value)
        return f"{self.prefix}{text}"

    def treatment_temp_k(self) -> float:
        return self.temperature_c + 273.15


@dataclass(frozen=True)
class ThermalRoutinePlan:
    """Validated thermal plan with queue-compatible labels."""

    name: str
    steps: tuple[ThermalStep, ...]
    limits: ThermalSafetyLimits = field(default_factory=ThermalSafetyLimits)

    def __post_init__(self) -> None:
        if not self.name.strip():
            raise ValueError("thermal routine name is required")
        if not self.steps:
            raise ValueError("thermal routine requires at least one step")
        for step in self.steps:
            if step.temperature_c > self.limits.max_temperature_c:
                raise ValueError(
                    f"thermal step {step.label} exceeds max temperature "
                    f"{self.limits.max_temperature_c:g} C"
                )
        object.__setattr__(self, "name", self.name.strip())

    @property
    def labels(self) -> list[str]:
        return [step.label for step in self.steps]

    @property
    def max_temperature_c(self) -> float:
        return max(step.temperature_c for step in self.steps)

    @property
    def requires_cooldown(self) -> bool:
        return self.max_temperature_c > self.limits.cool_to_c

    def estimated_seconds(self) -> int:
        total = 0.0
        current = self.limits.ambient_temperature_c
        for step in self.steps:
            total += abs(step.temperature_c - current) / self.limits.ramp_rate_c_per_min * 60.0
            total += step.hold_seconds if step.hold_seconds is not None else self.limits.default_hold_seconds
            current = step.temperature_c
        if self.requires_cooldown:
            total += abs(current - self.limits.cool_to_c) / self.limits.ramp_rate_c_per_min * 60.0
        return int(round(total))

    def to_measurement_block(self) -> MeasurementBlock:
        return MeasurementBlock(
            name=self.name,
            labels=self.labels,
            block_type="thermal",
            metadata={
                "max_temperature_c": f"{self.max_temperature_c:g}",
                "requires_cooldown": str(self.requires_cooldown),
                "estimated_seconds": str(self.estimated_seconds()),
            },
        )

    def to_queue_labels(self) -> list[str]:
        """Return the exact labels carried into the measurement runner."""

        return self.to_measurement_block().to_queue_labels()

    def to_artifact(
        self,
        *,
        run_context: str = "",
        operator: str = "",
        notes: str = "",
        timestamp_iso: str | None = None,
    ) -> dict[str, Any]:
        """Return a JSON-ready operator evidence artifact for this plan.

        This records software planning evidence only. Live furnace/oven
        execution, abort handling, alarm behavior, and final temperature
        readback remain hardware acceptance gates.
        """

        timestamp = timestamp_iso or datetime.now(timezone.utc).replace(microsecond=0).isoformat()
        block = self.to_measurement_block()
        return {
            "schema": "rapidpy.thermal.plan.v1",
            "procedure_id": "thermal/routine-planning",
            "procedure_name": "Thermal Routine Planning",
            "timestamp_iso": timestamp,
            "status": "PLANNED",
            "run_context": str(run_context).strip(),
            "operator": str(operator).strip(),
            "notes": str(notes).strip(),
            "hardware_validation_required": True,
            "hardware_validation_statement": (
                "Software planning artifact only; live furnace/oven execution, "
                "safe abort, alarm handling, and temperature readback require "
                "hardware acceptance evidence."
            ),
            "routine_name": self.name,
            "labels": self.labels,
            "queue_labels": self.to_queue_labels(),
            "max_temperature_c": self.max_temperature_c,
            "requires_cooldown": self.requires_cooldown,
            "estimated_seconds": self.estimated_seconds(),
            "limits": {
                "max_temperature_c": self.limits.max_temperature_c,
                "ambient_temperature_c": self.limits.ambient_temperature_c,
                "cool_to_c": self.limits.cool_to_c,
                "ramp_rate_c_per_min": self.limits.ramp_rate_c_per_min,
                "default_hold_seconds": self.limits.default_hold_seconds,
            },
            "steps": [
                {
                    "label": step.label,
                    "prefix": step.prefix,
                    "temperature_c": step.temperature_c,
                    "treatment_temp_k": step.treatment_temp_k(),
                    "hold_seconds": (
                        step.hold_seconds
                        if step.hold_seconds is not None
                        else self.limits.default_hold_seconds
                    ),
                }
                for step in self.steps
            ],
            "queue_block": {
                "name": block.name,
                "block_type": block.block_type,
                "labels": block.to_queue_labels(),
                "metadata": dict(block.metadata),
            },
        }


def compile_thermal_routine(
    temperatures_c: list[float],
    *,
    name: str = "Thermal routine",
    prefix: str = "TT",
    limits: ThermalSafetyLimits | None = None,
    hold_seconds: int | None = None,
) -> ThermalRoutinePlan:
    """Build a validated thermal routine from temperature targets."""

    steps = tuple(ThermalStep(temp, prefix=prefix, hold_seconds=hold_seconds) for temp in temperatures_c)
    return ThermalRoutinePlan(name=name, steps=steps, limits=limits or ThermalSafetyLimits())


def thermal_routine_artifact(
    plan: ThermalRoutinePlan,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
    timestamp_iso: str | None = None,
) -> dict[str, Any]:
    """Return a deterministic JSON-ready artifact for thermal planning."""

    return plan.to_artifact(
        run_context=run_context,
        operator=operator,
        notes=notes,
        timestamp_iso=timestamp_iso,
    )


def write_thermal_routine_artifact(
    path: str | Path,
    plan: ThermalRoutinePlan,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
    timestamp_iso: str | None = None,
) -> Path:
    """Atomically write a thermal routine planning artifact for review."""

    target = Path(path)
    payload = thermal_routine_artifact(
        plan,
        run_context=run_context,
        operator=operator,
        notes=notes,
        timestamp_iso=timestamp_iso,
    )
    _atomic_write_json(target, payload)
    return target


def thermal_integration_decision(*, timestamp_iso: str = "2026-10-02T00:00:00+00:00") -> dict[str, Any]:
    """Return the source-backed decision governing automated thermal treatment."""

    return {
        "schema": "rapidpy.thermal.integration_decision.v1",
        "decision_id": "thermal/manual-external-until-protocol-known",
        "recorded_at_iso": str(timestamp_iso),
        "status": "MANUAL_EXTERNAL_ONLY",
        "automated_live_dispatch": "BLOCKED",
        "summary": (
            "Specimen thermal treatments may be planned and recorded, but RapidPy must not "
            "command a furnace until the production controller and protocol are identified, "
            "implemented with interlocks, and physically accepted."
        ),
        "legacy_source_evidence": [
            {
                "path": "VB6/modThermal.bas",
                "finding": (
                    "Validates two AF coil thermal sensors and pauses/notifies on an invalid "
                    "reading; contains no specimen furnace set-point or readback transport."
                ),
            },
            {
                "path": "VB6/modPaleomag.bas",
                "finding": "Defines thermal-demagnetization identity and TH treatment labels.",
            },
            {
                "path": "VB6/frmMeasure.frm",
                "finding": "Classifies thermal treatment for measurement/output handling.",
            },
            {
                "path": "VB6/frmPlots.frm",
                "finding": "Classifies thermal treatment for plot presentation.",
            },
            {
                "path": "VB6/Paleomag v3.vbp",
                "finding": (
                    "Active project membership includes modThermal, but the audited project "
                    "does not identify a specimen furnace controller or transport module."
                ),
            },
        ],
        "missing_authoritative_inputs": [
            "production furnace/controller make and model",
            "command framing, set-point, readback, alarm, and fault protocol",
            "temperature limits, independent interlocks, and calibration requirements",
            "specimen loading/unloading and safe-transfer procedure",
            "abort, power-loss, cooldown, and safe-return behavior",
        ],
        "reopen_criteria": [
            "operator supplies authoritative controller documentation and wiring identity",
            "typed adapter passes deterministic fake/replay and fault-injection tests",
            "signed bench acceptance proves readback, limits, interlocks, abort, and cooldown",
        ],
        "current_software_behavior": {
            "planning": "SUPPORTED",
            "manual_external_treatment": "OPERATOR_CONTROLLED",
            "external_treatment_measurement": "LIVE_SQUID_WITH_OPERATOR_ATTESTATION",
            "no_communication_execution": "SIMULATED_ONLY",
            "hardware_mode_without_adapter": "BLOCKED_BEFORE_HARDWARE_PREFLIGHT",
        },
    }


def write_thermal_integration_decision(
    path: str | Path,
    *,
    timestamp_iso: str = "2026-10-02T00:00:00+00:00",
) -> Path:
    """Atomically publish the thermal integration/retirement decision."""

    target = Path(path)
    _atomic_write_json(target, thermal_integration_decision(timestamp_iso=timestamp_iso))
    return target


def thermal_run_artifact(
    routine_context: Mapping[str, Any],
    *,
    run_id: str,
    sample: str,
    operator: str,
    labels_requested: Sequence[str],
    completed_labels: Sequence[str],
    skipped_duplicate_labels: Sequence[str] = (),
    errors: Sequence[str] = (),
    aborted: bool,
    simulated: bool,
    final_phase: str,
    software_version: str = "",
    config_hash: str = "",
    timestamp_iso: str | None = None,
) -> dict[str, Any]:
    """Build run evidence tied to one compiled thermal planning artifact."""

    context = dict(routine_context)
    if context.get("procedure_id") != "thermal/routine-planning":
        raise ValueError("thermal run context must be a routine-planning artifact")
    planned = [str(label) for label in context.get("queue_labels", context.get("labels", ()))]
    requested = [str(label) for label in labels_requested]
    external = context.get("external_treatment")
    if external is not None:
        treatment = ExternalThermalTreatment(**{
            key: external[key] for key in ExternalThermalTreatment.__dataclass_fields__
        })
        if simulated:
            raise ValueError("External treatment evidence cannot be attached to a simulated run.")
        bind_external_thermal_treatment(context, treatment, specimen=sample, operator=operator, labels=requested)
    elif planned != requested:
        raise ValueError("thermal run labels do not match the compiled routine context")
    completed = [str(label) for label in completed_labels]
    if completed != requested[: len(completed)]:
        raise ValueError("completed thermal labels are not a prefix of the requested plan")
    timestamp = timestamp_iso or datetime.now(timezone.utc).isoformat()
    return {
        "schema": "rapidpy.thermal.run.v1",
        "status": "ABORTED" if aborted else "COMPLETED",
        "timestamp_iso": timestamp,
        "run_id": str(run_id),
        "sample": str(sample),
        "operator": str(operator),
        "software_version": str(software_version),
        "config_hash": str(config_hash),
        "simulated": bool(simulated),
        "simulation_statement": (
            "Simulated thermal-label execution; no specimen furnace was controlled."
            if simulated
            else ""
        ),
        "automation_status": "MANUAL_EXTERNAL_ONLY",
        "integration_decision_id": "thermal/manual-external-until-protocol-known",
        "routine": context,
        "labels_requested": requested,
        "labels_completed": completed,
        "skipped_duplicate_labels": [str(label) for label in skipped_duplicate_labels],
        "errors": [str(error) for error in errors],
        "final_phase": str(final_phase),
        "hardware_validation_required": True,
        "hardware_validation_statement": (
            "The furnace treatment and handling temperature are operator attestations. "
            "The measurement is independently recorded by the SQUID."
            if external is not None else
            "This record proves software planning and runner behavior only. It is not proof "
            "of specimen furnace treatment, temperature readback, interlocks, or cooldown."
        ),
    }


def write_thermal_run_artifact(
    path: str | Path,
    routine_context: Mapping[str, Any],
    **run_evidence: Any,
) -> Path:
    """Atomically publish one thermal run artifact."""

    target = Path(path)
    _atomic_write_json(target, thermal_run_artifact(routine_context, **run_evidence))
    return target


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            json.dump(dict(payload), handle, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except Exception:
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
        raise
