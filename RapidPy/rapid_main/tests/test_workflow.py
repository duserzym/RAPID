from __future__ import annotations

import unittest

from rapid_main.workflow import WorkflowPhase, WorkflowStateMachine, WorkflowTransitionError


class TestWorkflowStateMachine(unittest.TestCase):
    def test_allowed_transition(self) -> None:
        fsm = WorkflowStateMachine()
        self.assertEqual(fsm.confirmed, WorkflowPhase.IDLE)
        fsm.advance(WorkflowPhase.PREFLIGHT)
        fsm.advance(WorkflowPhase.LOADING)
        fsm.advance(WorkflowPhase.TREATING)
        fsm.advance(WorkflowPhase.MEASURING)
        self.assertEqual(fsm.requested, WorkflowPhase.MEASURING)
        self.assertEqual(fsm.confirmed, WorkflowPhase.MEASURING)

    def test_rejects_invalid_transition(self) -> None:
        fsm = WorkflowStateMachine()
        with self.assertRaises(WorkflowTransitionError):
            fsm.request(WorkflowPhase.SAVING)


class TestWorkflowPhaseMachineInWorkerContract(unittest.TestCase):
    """Smoke check: phase values are stable and easy to consume from UI signals."""

    def test_phase_values_are_non_empty(self) -> None:
        for phase in WorkflowPhase:
            self.assertTrue(phase.value)


if __name__ == "__main__":
    unittest.main(verbosity=2)

