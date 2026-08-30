"""Rock-magnetic routine compiler mapped from VB6 ``frmRockmagRoutine``.

The current runner consumes a flat list of treatment labels. The legacy form,
however, was routine-oriented: an operator selected a rock-magnetic program and
the application expanded it into ordered treatment/measurement steps. This
module preserves that form-level behavior while keeping the output compatible
with the existing RapidPy queue/measurement path.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence

from rapid_main.data_model import MeasurementBlock, MeasurementBlocks, RockmagStep, RockmagSteps


@dataclass(frozen=True)
class RockmagRoutineSpec:
    """Operator-selected rockmag routine definition."""

    name: str
    labels: tuple[str, ...]
    repeats: int = 1
    metadata: dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        clean_labels = tuple(str(label).strip().upper() for label in self.labels if str(label).strip())
        if not self.name.strip():
            raise ValueError("rockmag routine name is required")
        if not clean_labels:
            raise ValueError("rockmag routine requires at least one label")
        if int(self.repeats) < 1:
            raise ValueError("rockmag routine repeats must be positive")
        object.__setattr__(self, "name", self.name.strip())
        object.__setattr__(self, "labels", clean_labels)
        object.__setattr__(self, "repeats", int(self.repeats))
        object.__setattr__(self, "metadata", {str(k): str(v) for k, v in self.metadata.items()})


@dataclass(frozen=True)
class RockmagRoutinePlan:
    """Compiled routine with block metadata and runner-compatible labels."""

    spec: RockmagRoutineSpec
    steps: RockmagSteps
    blocks: MeasurementBlocks

    @property
    def labels(self) -> list[str]:
        return self.steps.labels

    @property
    def step_count(self) -> int:
        return self.steps.step_count

    def to_queue_labels(self) -> list[str]:
        return self.blocks.to_queue_labels()

    def to_artifact(
        self,
        *,
        run_context: str = "",
        operator: str = "",
        notes: str = "",
        timestamp_iso: str | None = None,
    ) -> dict[str, Any]:
        return rockmag_routine_artifact(
            self,
            run_context=run_context,
            operator=operator,
            notes=notes,
            timestamp_iso=timestamp_iso,
        )


def compile_rockmag_routine(spec: RockmagRoutineSpec) -> RockmagRoutinePlan:
    """Compile a routine spec into steps and family-preserving blocks."""

    expanded_labels: list[str] = []
    for repeat_index in range(spec.repeats):
        expanded_labels.extend(spec.labels)
        if repeat_index < spec.repeats - 1:
            expanded_labels.append(f"REPEAT{repeat_index + 2}")

    steps = RockmagSteps.from_labels(expanded_labels, routine_name=spec.name)
    blocks = MeasurementBlocks(_blocks_by_family(steps.steps, spec))
    return RockmagRoutinePlan(spec=spec, steps=steps, blocks=blocks)


def rockmag_routine_artifact(
    plan: RockmagRoutinePlan,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
    timestamp_iso: str | None = None,
) -> dict[str, Any]:
    """Return a deterministic JSON-ready artifact for Rockmag plan review."""

    timestamp = timestamp_iso or datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return {
        "procedure_id": "rockmag/routine-planning",
        "procedure_name": "Rockmag Routine Planning",
        "status": "PLANNED",
        "timestamp_iso": timestamp,
        "run_context": str(run_context).strip(),
        "operator": str(operator).strip(),
        "notes": str(notes).strip(),
        "routine_name": plan.spec.name,
        "repeats": plan.spec.repeats,
        "metadata": dict(plan.spec.metadata),
        "labels": list(plan.labels),
        "queue_labels": list(plan.to_queue_labels()),
        "step_count": plan.step_count,
        "block_count": len(plan.blocks.blocks),
        "blocks": [
            {
                "name": block.name,
                "block_type": block.block_type,
                "labels": list(block.labels),
                "metadata": dict(block.metadata),
            }
            for block in plan.blocks.blocks
        ],
        "hardware_validation_required": True,
        "hardware_validation_statement": (
            "Software planning artifact only; live Rockmag execution, "
            "hardware treatment/readback, and run-bundle capture require "
            "physical acceptance evidence."
        ),
    }


def write_rockmag_routine_artifact(
    path: str | Path,
    plan: RockmagRoutinePlan,
    *,
    run_context: str = "",
    operator: str = "",
    notes: str = "",
    timestamp_iso: str | None = None,
) -> Path:
    """Write a Rockmag routine planning artifact for operator review."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = rockmag_routine_artifact(
        plan,
        run_context=run_context,
        operator=operator,
        notes=notes,
        timestamp_iso=timestamp_iso,
    )
    target.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target


def rockmag_the_works(
    *,
    af_fields_mT: Sequence[float] = (20.0, 40.0, 60.0, 80.0, 100.0),
    irm_fields_g: Sequence[float] = (100.0, 300.0, 1000.0),
    arm_fields_g: Sequence[float] = (50.0, 100.0),
    include_backfield: bool = True,
    include_susceptibility: bool = True,
) -> RockmagRoutineSpec:
    """Return a broad rockmag routine approximating the legacy 'works' flow."""

    labels = ["NRM"]
    labels.extend(_field_labels("AF", af_fields_mT))
    labels.extend(_field_labels("IRM", irm_fields_g))
    labels.extend(_field_labels("ARM", arm_fields_g))
    if include_backfield:
        labels.append("IRM-BF")
    if include_susceptibility:
        labels.append("SUSC")
    return RockmagRoutineSpec(
        name="Rockmag the Works",
        labels=tuple(labels),
        metadata={
            "template": "the-works",
            "af_count": str(len(af_fields_mT)),
            "irm_count": str(len(irm_fields_g)),
            "arm_count": str(len(arm_fields_g)),
        },
    )


def rockmag_af_demag(
    fields_mT: Iterable[float],
    *,
    include_nrm: bool = True,
    name: str = "Rockmag AF demag",
) -> RockmagRoutineSpec:
    """Return an AF-focused routine spec."""

    labels = ["NRM"] if include_nrm else []
    labels.extend(_field_labels("AF", fields_mT))
    return RockmagRoutineSpec(name=name, labels=tuple(labels), metadata={"template": "af-demag"})


def _blocks_by_family(
    steps: Sequence[RockmagStep],
    spec: RockmagRoutineSpec,
) -> list[MeasurementBlock]:
    blocks: list[MeasurementBlock] = []
    current_family = ""
    current_labels: list[str] = []

    def flush() -> None:
        if not current_labels:
            return
        blocks.append(
            MeasurementBlock(
                name=f"{spec.name}: {current_family}",
                labels=list(current_labels),
                block_type="rockmag",
                metadata={
                    "routine": spec.name,
                    "family": current_family,
                    **spec.metadata,
                },
            )
        )

    for step in steps:
        if step.family != current_family:
            flush()
            current_family = step.family
            current_labels = []
        current_labels.append(step.label)
    flush()
    return blocks


def _field_labels(prefix: str, values: Iterable[float]) -> list[str]:
    labels: list[str] = []
    for value in values:
        numeric = float(value)
        if numeric < 0.0:
            raise ValueError(f"{prefix} field values must be non-negative")
        labels.append(f"{prefix}{_format_field(numeric)}")
    return labels


def _format_field(value: float) -> str:
    if value == round(value):
        return str(int(round(value)))
    return f"{value:.3f}".rstrip("0").rstrip(".")
