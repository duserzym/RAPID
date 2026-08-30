from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import PreflightResult
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.status_codes import (
    OperatorStatus,
    StatusCode,
    StatusSeverity,
    error_status,
    status_for_phase,
    warning_status,
)
from rapid_main.workflow import WorkflowPhase


def _meta(name: str = "STATUS") -> SpecimenMeta:
    return SpecimenMeta(name=name, comment="status test")


class WarningBackend:
    def preflight(self) -> PreflightResult:
        return PreflightResult(ok=True, blockers=(), warnings=("coil warm-up required",))

    def is_available(self) -> bool:
        return True

    def set_demag_step(self, label: str) -> None:
        del label

    def read_squid(self) -> tuple[float, float, float]:
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        return 0.0


class SquidFailBackend(WarningBackend):
    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def read_squid(self) -> tuple[float, float, float]:
        raise RuntimeError("serial timeout")


class ReturnBackend(WarningBackend):
    def __init__(self) -> None:
        self.return_calls = 0

    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def return_to_safe_state(self) -> None:
        self.return_calls += 1


class StatusCodeTest(unittest.TestCase):
    def test_phase_status_maps_to_stable_code(self) -> None:
        status = status_for_phase(WorkflowPhase.RETURNING)

        self.assertEqual(status.code, StatusCode.RUN_RETURNING)
        self.assertEqual(status.severity, StatusSeverity.SAFETY)
        self.assertEqual(status.phase, WorkflowPhase.RETURNING)
        self.assertIn("RUN_RETURNING", status.format_for_log())

    def test_warning_and_error_helpers_classify_operator_events(self) -> None:
        warning = warning_status("coil warm-up required")
        squid = error_status("SQUID read error at step AF20: serial timeout")
        output = error_status("Failed to initialize output bundle: denied")

        self.assertEqual(warning.code, StatusCode.PREFLIGHT_WARNING)
        self.assertEqual(warning.severity, StatusSeverity.WARNING)
        self.assertEqual(squid.code, StatusCode.SQUID_READ_ERROR)
        self.assertEqual(output.code, StatusCode.OUTPUT_INIT_ERROR)

    def test_measurement_worker_emits_structured_warning_status(self) -> None:
        statuses: list[OperatorStatus] = []

        with tempfile.TemporaryDirectory() as td:
            worker = MeasurementWorker(
                meta=_meta("WARN_STATUS"),
                labels=["NRM"],
                output_dir=Path(td) / "WARN_STATUS",
                backend=WarningBackend(),
            )
            worker.status_event.connect(statuses.append)
            worker.run()

        self.assertIn(StatusCode.PREFLIGHT_WARNING, [status.code for status in statuses])
        self.assertIn(StatusCode.RUN_COMPLETE, [status.code for status in statuses])

    def test_measurement_worker_emits_classified_squid_error_status(self) -> None:
        statuses: list[OperatorStatus] = []
        errors: list[str] = []

        with tempfile.TemporaryDirectory() as td:
            worker = MeasurementWorker(
                meta=_meta("SQUID_STATUS"),
                labels=["AF20"],
                output_dir=Path(td) / "SQUID_STATUS",
                backend=SquidFailBackend(),
            )
            worker.status_event.connect(statuses.append)
            worker.error_occurred.connect(errors.append)
            worker.run()

        self.assertTrue(errors)
        self.assertIn(StatusCode.SQUID_READ_ERROR, [status.code for status in statuses])
        self.assertIn(StatusCode.RUN_ERROR, [status.code for status in statuses])

    def test_measurement_worker_marks_returning_as_safety_status(self) -> None:
        statuses: list[OperatorStatus] = []
        backend = ReturnBackend()

        with tempfile.TemporaryDirectory() as td:
            worker = MeasurementWorker(
                meta=_meta("RETURN_STATUS"),
                labels=["NRM"],
                output_dir=Path(td) / "RETURN_STATUS",
                backend=backend,
            )
            worker.status_event.connect(statuses.append)
            worker.run()

        returning = [status for status in statuses if status.code == StatusCode.RUN_RETURNING]
        self.assertEqual(len(returning), 1)
        self.assertEqual(returning[0].severity, StatusSeverity.SAFETY)
        self.assertEqual(backend.return_calls, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
