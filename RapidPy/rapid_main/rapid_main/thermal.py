"""Thermal treatment planning mapped from VB6 ``modThermal``.

RapidPy already understands thermal labels such as ``TT400`` in IO and runtime
estimation paths. This module adds the missing workflow foundation: validated
thermal treatment steps, routine expansion, safety limits, and queue-compatible
labels without claiming live furnace control.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

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
    """Write a thermal routine planning artifact for operator review."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = thermal_routine_artifact(
        plan,
        run_context=run_context,
        operator=operator,
        notes=notes,
        timestamp_iso=timestamp_iso,
    )
    target.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return target
