"""Specimen thermal-treatment planning without claiming furnace control.

The legacy ``modThermal`` module protects the AF coil temperature sensors; it
does not implement a specimen furnace protocol.  RapidPy understands specimen
thermal labels such as ``TT400`` in IO and runtime-estimation paths, so this
module provides validated planning and auditable run evidence while keeping the
unknown live furnace boundary fail-closed.
"""
from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from rapid_main.data_model import MeasurementBlock


_THERMAL_LABEL_RE = re.compile(r"^(TT|TH|TEMP)\s*(\d+(?:\.\d+)?)$", re.IGNORECASE)


@dataclass(frozen=True)
class ThermalSafetyLimits:
    """Operator-configurable thermal planning limits."""

    max_temperature_c: float = 700.0
    ambient_temperature_c: float = 25.0
    cool_to_c: float = 50.0
    ramp_rate_c_per_min: float = 20.0
    default_hold_seconds: int = 600

    def __post_init__(self) -> None:
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
        if self.temperature_c < 0.0:
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
    if planned != requested:
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
