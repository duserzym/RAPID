"""Operator-facing status taxonomy for rapid_main workflows.

This is the RapidPy replacement for the VB6 ``modStatusCode`` role: a small,
stable set of status codes that can be logged, tested, and attached to workflow
events without changing the human-readable UI messages.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from rapid_main.workflow import WorkflowPhase


class StatusSeverity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    SAFETY = "safety"


class StatusCode(str, Enum):
    RUN_IDLE = "RUN_IDLE"
    RUN_PREFLIGHT = "RUN_PREFLIGHT"
    RUN_LOADING = "RUN_LOADING"
    RUN_TREATING = "RUN_TREATING"
    RUN_POSITIONING = "RUN_POSITIONING"
    RUN_MEASURING = "RUN_MEASURING"
    RUN_VALIDATING = "RUN_VALIDATING"
    RUN_SAVING = "RUN_SAVING"
    RUN_RETURNING = "RUN_RETURNING"
    RUN_COMPLETE = "RUN_COMPLETE"
    RUN_ACTIVE = "RUN_ACTIVE"
    RUN_PAUSED = "RUN_PAUSED"
    RUN_HALTED = "RUN_HALTED"
    RUN_ERROR = "RUN_ERROR"
    PREFLIGHT_WARNING = "PREFLIGHT_WARNING"
    HARDWARE_ERROR = "HARDWARE_ERROR"
    SQUID_READ_ERROR = "SQUID_READ_ERROR"
    FILE_WRITE_ERROR = "FILE_WRITE_ERROR"
    OUTPUT_INIT_ERROR = "OUTPUT_INIT_ERROR"
    POSITION_ERROR = "POSITION_ERROR"
    SAFE_RETURN_ERROR = "SAFE_RETURN_ERROR"


@dataclass(frozen=True, slots=True)
class OperatorStatus:
    code: StatusCode
    severity: StatusSeverity
    message: str
    phase: WorkflowPhase | None = None

    def format_for_log(self) -> str:
        return f"[{self.severity.value.upper()}:{self.code.value}] {self.message}"


_PHASE_STATUS: dict[WorkflowPhase, tuple[StatusCode, StatusSeverity, str]] = {
    WorkflowPhase.IDLE: (StatusCode.RUN_IDLE, StatusSeverity.INFO, "Workflow idle."),
    WorkflowPhase.PREFLIGHT: (StatusCode.RUN_PREFLIGHT, StatusSeverity.INFO, "Running preflight checks."),
    WorkflowPhase.LOADING: (StatusCode.RUN_LOADING, StatusSeverity.INFO, "Loading sample workflow."),
    WorkflowPhase.TREATING: (StatusCode.RUN_TREATING, StatusSeverity.INFO, "Applying treatment step."),
    WorkflowPhase.POSITIONING: (StatusCode.RUN_POSITIONING, StatusSeverity.INFO, "Validating sample position."),
    WorkflowPhase.MEASURING: (StatusCode.RUN_MEASURING, StatusSeverity.INFO, "Reading measurement channels."),
    WorkflowPhase.VALIDATING: (StatusCode.RUN_VALIDATING, StatusSeverity.INFO, "Validating measurement step."),
    WorkflowPhase.SAVING: (StatusCode.RUN_SAVING, StatusSeverity.INFO, "Saving measurement output."),
    WorkflowPhase.RETURNING: (StatusCode.RUN_RETURNING, StatusSeverity.SAFETY, "Returning hardware to safe state."),
    WorkflowPhase.COMPLETE: (StatusCode.RUN_COMPLETE, StatusSeverity.INFO, "Workflow complete."),
    WorkflowPhase.RUNNING: (StatusCode.RUN_ACTIVE, StatusSeverity.INFO, "Workflow running."),
    WorkflowPhase.PAUSED: (StatusCode.RUN_PAUSED, StatusSeverity.WARNING, "Workflow paused."),
    WorkflowPhase.HALTED: (StatusCode.RUN_HALTED, StatusSeverity.SAFETY, "Workflow halted."),
    WorkflowPhase.ERROR: (StatusCode.RUN_ERROR, StatusSeverity.ERROR, "Workflow error."),
}


def status_for_phase(phase: WorkflowPhase, message: str | None = None) -> OperatorStatus:
    code, severity, default_message = _PHASE_STATUS[phase]
    return OperatorStatus(code=code, severity=severity, message=message or default_message, phase=phase)


def warning_status(message: str) -> OperatorStatus:
    return OperatorStatus(
        code=StatusCode.PREFLIGHT_WARNING,
        severity=StatusSeverity.WARNING,
        message=message,
        phase=WorkflowPhase.PREFLIGHT,
    )


def error_status(message: str, *, phase: WorkflowPhase | None = None) -> OperatorStatus:
    text = message.lower()
    if "squid read" in text:
        code = StatusCode.SQUID_READ_ERROR
    elif "file write" in text:
        code = StatusCode.FILE_WRITE_ERROR
    elif "initialize output bundle" in text:
        code = StatusCode.OUTPUT_INIT_ERROR
    elif "position" in text:
        code = StatusCode.POSITION_ERROR
    elif "safe state" in text:
        code = StatusCode.SAFE_RETURN_ERROR
    else:
        code = StatusCode.HARDWARE_ERROR
    return OperatorStatus(
        code=code,
        severity=StatusSeverity.ERROR,
        message=message,
        phase=phase,
    )
