"""Shared workflow phase primitives for rapid_main run-state signaling."""
from __future__ import annotations

from enum import Enum, auto


class WorkflowPhase(str, Enum):
    """Canonical flow phases used by measurement and future treatment workflows."""

    IDLE = "idle"
    PREFLIGHT = "preflight"
    LOADING = "loading"
    TREATING = "treating"
    POSITIONING = "positioning"
    MEASURING = "measuring"
    VALIDATING = "validating"
    SAVING = "saving"
    RETURNING = "returning"
    COMPLETE = "complete"
    RUNNING = "running"
    PAUSED = "paused"
    HALTED = "halted"
    ERROR = "error"


class WorkflowTransitionError(RuntimeError):
    """Raised when a requested workflow transition violates the state model."""


_ALLOWED_TRANSITIONS: dict[WorkflowPhase, set[WorkflowPhase]] = {
    WorkflowPhase.IDLE: {
        WorkflowPhase.PREFLIGHT,
        WorkflowPhase.RUNNING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.PREFLIGHT: {
        WorkflowPhase.LOADING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.LOADING: {
        WorkflowPhase.TREATING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.TREATING: {
        WorkflowPhase.MEASURING,
        WorkflowPhase.POSITIONING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.POSITIONING: {
        WorkflowPhase.MEASURING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.MEASURING: {
        WorkflowPhase.SAVING,
        WorkflowPhase.VALIDATING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.VALIDATING: {
        WorkflowPhase.SAVING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.SAVING: {
        WorkflowPhase.RETURNING,
        WorkflowPhase.COMPLETE,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.RETURNING: {
        WorkflowPhase.COMPLETE,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.PAUSED: {
        WorkflowPhase.RUNNING,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.RUNNING: {
        WorkflowPhase.PAUSED,
        WorkflowPhase.HALTED,
        WorkflowPhase.ERROR,
    },
    WorkflowPhase.COMPLETE: {
        WorkflowPhase.IDLE,
        WorkflowPhase.HALTED,
    },
    WorkflowPhase.HALTED: {
        WorkflowPhase.IDLE,
    },
    WorkflowPhase.ERROR: {
        WorkflowPhase.IDLE,
    },
}


class WorkflowStateMachine:
    """Minimal state machine used by `rapid_main` workflow controllers."""

    def __init__(self, initial: WorkflowPhase = WorkflowPhase.IDLE) -> None:
        self._confirmed = initial
        self._requested = initial

    @property
    def requested(self) -> WorkflowPhase:
        return self._requested

    @property
    def confirmed(self) -> WorkflowPhase:
        return self._confirmed

    def request(self, phase: WorkflowPhase) -> None:
        """Request a new state and validate the transition path."""
        if phase not in _ALLOWED_TRANSITIONS.get(self._confirmed, set()):
            raise WorkflowTransitionError(
                f"Cannot transition {self._confirmed.value} -> {phase.value}"
            )
        self._requested = phase

    def confirm(self, phase: WorkflowPhase) -> None:
        """Confirm current state as physically observed on hardware."""
        if phase not in _ALLOWED_TRANSITIONS.get(self._requested, set()):
            if phase != self._requested:
                raise WorkflowTransitionError(
                    f"Cannot confirm {phase.value} when requested is {self._requested.value}"
                )
        self._confirmed = phase

    def advance(self, phase: WorkflowPhase) -> None:
        """Shortcut: request then confirm in one step."""
        self.request(phase)
        self.confirm(phase)

    def is_terminal(self) -> bool:
        return self._confirmed in {WorkflowPhase.COMPLETE, WorkflowPhase.HALTED, WorkflowPhase.ERROR}

