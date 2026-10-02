from __future__ import annotations

import json
import tempfile
import time
import threading
import unittest
from datetime import datetime, timezone
from pathlib import Path

from PySide6 import QtCore

from rapid_main.data_model import SpecimenMeta
from rapid_main.communication_log import CommunicationDirection, CommunicationEvent
from rapid_main.acquisition import RecoveryRecord
from rapid_main.hardware_contracts import PreflightResult
from rapid_main.magnetometer import BracketedMeasurementBlock, CommandEvent, ZeroPairValidation
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.workflow import WorkflowPhase


class TimeoutBackend:
    """Backend used to exercise timeout and cancellation behavior in the worker."""

    preflight_timeout = 0.05
    step_timeout = 0.05
    read_timeout = 0.05
    susceptibility_timeout = 0.05

    def __init__(
        self,
        *,
        preflight_delay: float = 0.0,
        step_delay: float = 0.0,
        read_delay: float = 0.0,
        available: bool = True,
    ) -> None:
        self.preflight_delay = preflight_delay
        self.step_delay = step_delay
        self.read_delay = read_delay
        self.available = available
        self.preflight_calls = 0
        self.step_calls = 0
        self.read_calls = 0
        self.susc_calls = 0

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        if self.preflight_delay:
            time.sleep(self.preflight_delay)
        return PreflightResult.pass_ok()

    def set_demag_step(self, label: str) -> None:
        del label
        self.step_calls += 1
        if self.step_delay:
            time.sleep(self.step_delay)

    def read_squid(self) -> tuple[float, float, float]:
        self.read_calls += 1
        if self.read_delay:
            time.sleep(self.read_delay)
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        self.susc_calls += 1
        if self.read_delay:
            time.sleep(self.read_delay)
        return 0.01

    def is_available(self) -> bool:
        return self.available


class WarningBackend:
    def __init__(self) -> None:
        self.preflight_calls = 0

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        return PreflightResult(ok=True, blockers=(), warnings=("coil warm-up required",))

    def set_demag_step(self, label: str) -> None:
        del label

    def read_squid(self) -> tuple[float, float, float]:
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        return 0.02

    def is_available(self) -> bool:
        return True


class DeterministicBackend:
    """Minimal backend that exercises the full successful run path."""

    def __init__(self, labels: list[tuple[str, float, float, float]] | None = None) -> None:
        self.readings = labels or [
            ("AF10", 1.0, 2.0, 3.0),
            ("AF20", 1.5, 1.8, 2.4),
            ("AF40", 2.0, 2.5, 1.9),
        ]
        self.step_calls: list[str] = []
        self.squid_calls = 0
        self.susc_calls = 0
        self.preflight_calls = 0

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        return PreflightResult.pass_ok()

    def set_demag_step(self, label: str) -> None:
        self.step_calls.append(label)

    def read_squid(self) -> tuple[float, float, float]:
        self.squid_calls += 1
        idx = min(self.squid_calls - 1, len(self.readings) - 1)
        _label, x, y, z = self.readings[idx]
        return (x, y, z)

    def read_susceptibility(self) -> float:
        self.susc_calls += 1
        return 0.1 / (1 + self.susc_calls)

    def is_available(self) -> bool:
        return True


class ReturnAwareBackend(DeterministicBackend):
    """Backend with an explicit safe-state return hook."""

    def __init__(
        self,
        labels: list[tuple[str, float, float, float]] | None = None,
        *,
        return_fail: bool = False,
    ) -> None:
        super().__init__(labels=labels)
        self.return_to_safe_state_calls = 0
        self.return_fail = return_fail

    def return_to_safe_state(self) -> None:
        self.return_to_safe_state_calls += 1
        if self.return_fail:
            raise RuntimeError("return failed for test")


class ValidatingPositionBackend:
    """Backend that validates positioning step before treatment reads."""

    def __init__(self, *, allow_position: bool, labels: list[tuple[str, float, float, float]] | None = None) -> None:
        self._allow_position = allow_position
        self.position_checked = 0
        self.preflight_calls = 0
        self.read_calls = 0
        self.susc_calls = 0
        self.step_calls = 0
        self.readings = labels or [("NRM", 1.0, 2.0, 3.0)]

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        return PreflightResult.pass_ok()

    def validate_position(self) -> bool:
        self.position_checked += 1
        return self._allow_position

    def set_demag_step(self, label: str) -> None:
        del label
        self.step_calls += 1

    def read_squid(self) -> tuple[float, float, float]:
        self.read_calls += 1
        _label, x, y, z = self.readings[min(self.read_calls - 1, len(self.readings) - 1)]
        return x, y, z

    def read_susceptibility(self) -> float:
        self.susc_calls += 1
        return 0.1 / (1 + self.susc_calls)

    def is_available(self) -> bool:
        return True


class ReturnAwareFailingPositionBackend(ReturnAwareBackend):
    """Backend that fails during treatment but still supports a return hook."""

    def __init__(self, fail_label: str = "AF20", **kwargs) -> None:
        super().__init__(**kwargs)
        self.fail_label = fail_label

    def set_demag_step(self, label: str) -> None:
        if label == self.fail_label:
            self.step_calls.append(label)
            raise RuntimeError("AF step failure for return-path test")
        super().set_demag_step(label)


class PreflightEnableBackend:
    """Backend that becomes available only after preflight."""

    def __init__(self) -> None:
        self.preflight_calls = 0
        self.is_available_calls = 0
        self.step_calls = 0
        self.read_calls = 0
        self.susc_calls = 0
        self._available = False

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        self._available = True
        return PreflightResult.pass_ok()

    def is_available(self) -> bool:
        self.is_available_calls += 1
        return self._available

    def set_demag_step(self, label: str) -> None:
        del label
        self.step_calls += 1

    def read_squid(self) -> tuple[float, float, float]:
        self.read_calls += 1
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        self.susc_calls += 1
        return 0.01


class HaltingBackend:
    """Backend that blocks inside step execution so halt timing can be verified."""

    def __init__(self, *, step_delay: float = 1.0) -> None:
        self.step_delay = step_delay
        self.preflight_calls = 0
        self.step_calls = 0
        self.read_calls = 0
        self.susc_calls = 0
        self.step_started = threading.Event()

    def preflight(self) -> PreflightResult:
        self.preflight_calls += 1
        return PreflightResult.pass_ok()

    def set_demag_step(self, label: str) -> None:
        del label
        self.step_calls += 1
        self.step_started.set()
        time.sleep(self.step_delay)

    def read_squid(self) -> tuple[float, float, float]:
        self.read_calls += 1
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        self.susc_calls += 1
        return 0.01

    def is_available(self) -> bool:
        return True


class FluxRecoveryBackend(DeterministicBackend):
    """Returns one invalid bracketed block, then a valid legacy moment tuple."""

    flux_discontinuity_retries = 2

    def __init__(self) -> None:
        super().__init__([("NRM", 1.0, 2.0, 3.0)])
        self.recovery_evidence: list[ZeroPairValidation] = []

    def read_squid(self) -> BracketedMeasurementBlock | tuple[float, float, float]:
        self.squid_calls += 1
        if self.squid_calls == 1:
            return BracketedMeasurementBlock(
                zero_before=(-1.0436457, 0.0, 0.0),
                positions=((0.0, 0.0, 0.0),) * 4,
                zero_after=(-0.95364007, 0.0, 0.0),
            )
        return (1.0, 2.0, 3.0)

    def recover_flux_count_discontinuity(self, evidence: ZeroPairValidation) -> None:
        self.recovery_evidence.append(evidence)


def _meta(name: str = "MEASURE_WORKER") -> SpecimenMeta:
    return SpecimenMeta(
        name=name,
        comment="test",
        sample="S1",
        site="SITE1",
        location="LOC",
    )


class TestMeasurementWorkerPreflightTimeout(unittest.TestCase):
    def test_preflight_timeout_causes_error(self) -> None:
        output_dir = Path("unused")
        backend = TimeoutBackend(preflight_delay=0.2)
        errors: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            output_dir = out
            worker = MeasurementWorker(meta=_meta("TIMEOUT1"), labels=["AF20"], output_dir=out, backend=backend)
            worker.error_occurred.connect(errors.append)
            worker.run_finished.connect(finished.append)
            worker.run()
            workflow_payload = json.loads((output_dir / "workflow_summary.json").read_text(encoding="utf-8"))
            artifact_payload = json.loads((output_dir / "artifact_index.json").read_text(encoding="utf-8"))
            self.assertTrue(workflow_payload["aborted"])
            self.assertEqual(workflow_payload["final_phase"], "error")
            self.assertTrue(artifact_payload["aborted"])
            self.assertEqual(artifact_payload["schema"], "rapidpy.measurement.artifact_index.v1")

        self.assertEqual(len(errors), 1)
        self.assertIn("preflight", errors[0].lower())
        self.assertIn("exceeded timeout", errors[0].lower())
        self.assertEqual(finished, [True])
        self.assertEqual(backend.preflight_calls, 1)
        # preflight failure occurs before bundle creation, but still leaves run evidence.

    def test_halt_before_start_skips_hardware_calls(self) -> None:
        backend = TimeoutBackend(preflight_delay=0.2)
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(meta=_meta("HALT1"), labels=["AF20"], output_dir=out, backend=backend)
            worker.run_finished.connect(finished.append)
            worker.halt()
            worker.run()
            workflow_payload = json.loads((out / "workflow_summary.json").read_text(encoding="utf-8"))
            artifact_payload = json.loads((out / "artifact_index.json").read_text(encoding="utf-8"))
            self.assertEqual(backend.preflight_calls, 0)
            self.assertEqual(finished, [True])
            self.assertEqual(workflow_payload["final_phase"], "halted")
            self.assertTrue(artifact_payload["aborted"])

    def test_step_command_timeout_becomes_error(self) -> None:
        backend = TimeoutBackend(step_delay=0.2)
        errors: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(meta=_meta("STEPTO"), labels=["AF20"], output_dir=out, backend=backend)
            worker.error_occurred.connect(errors.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(len(errors), 1)
        self.assertIn("Hardware error at step AF20", errors[0])
        self.assertIn("set_demag_step", errors[0])
        self.assertEqual(finished, [True])
        self.assertEqual(backend.step_calls, 1)

    def test_backend_unavailable_stops_run_and_skips_calls(self) -> None:
        backend = TimeoutBackend(preflight_delay=0.0, available=False)
        errors: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(meta=_meta("UNAV1"), labels=["AF20"], output_dir=out, backend=backend)
            worker.error_occurred.connect(errors.append)
            worker.run_finished.connect(finished.append)
            worker.run()
            self.assertTrue((out / "workflow_summary.json").exists())
            self.assertTrue((out / "artifact_index.json").exists())

        self.assertEqual(len(errors), 1)
        self.assertIn("backend is not available", errors[0].lower())
        self.assertEqual(finished, [True])
        self.assertEqual(backend.preflight_calls, 1)

    def test_preflight_can_enable_hardware_and_allow_run(self) -> None:
        backend = PreflightEnableBackend()
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("PREFLT"),
                labels=["AF20"],
                output_dir=out / "PREFLT",
                backend=backend,
            )
            worker.error_occurred.connect(errors.append)
            worker.phase_changed.connect(phases.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(backend.preflight_calls, 1)
        self.assertGreaterEqual(backend.is_available_calls, 1)
        self.assertEqual(backend.step_calls, 1)
        self.assertEqual(errors, [])
        self.assertEqual(finished, [False])
        self.assertEqual(phases[0], "preflight")

    def test_preflight_warning_is_emitted(self) -> None:
        backend = WarningBackend()
        warnings: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(meta=_meta("WARN1"), labels=["AF20"], output_dir=out, backend=backend)
            worker.preflight_warning.connect(warnings.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(backend.preflight_calls, 1)
        self.assertEqual(warnings, ["coil warm-up required"])
        self.assertEqual(finished, [False])

    def test_position_validation_failure_prevents_run(self) -> None:
        backend = ValidatingPositionBackend(allow_position=False)
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("POSITION_FAIL"),
                labels=["AF20"],
                output_dir=out / "POSITION_FAIL",
                backend=backend,
            )
            worker.error_occurred.connect(errors.append)
            worker.phase_changed.connect(phases.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(errors, ["Hardware error at step AF20: backend position validation failed"])
        self.assertEqual(finished, [True])
        self.assertEqual(backend.position_checked, 1)
        self.assertEqual(backend.step_calls, 1)
        self.assertEqual(phases[0], "preflight")
        self.assertIn("error", phases)

    def test_successful_run_invokes_return_to_safe_state(self) -> None:
        backend = ReturnAwareBackend(
            [
                ("NRM", 1.0, 2.0, 3.0),
                ("AF20", 2.0, 2.2, 2.6),
            ]
        )
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("SAFE_RETURN"),
                labels=["NRM", "AF20"],
                output_dir=out / "SAFE_RETURN",
                backend=backend,
            )
            worker.error_occurred.connect(errors.append)
            worker.phase_changed.connect(phases.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(errors, [])
        self.assertEqual(finished, [False])
        self.assertEqual(backend.return_to_safe_state_calls, 1)
        self.assertEqual(phases[-2], "returning")
        self.assertEqual(phases[-1], "complete")

    def test_return_to_safe_state_runs_after_run_failure(self) -> None:
        backend = ReturnAwareFailingPositionBackend(
            fail_label="AF20",
            return_fail=False,
            labels=[
                ("NRM", 1.0, 2.0, 3.0),
                ("AF20", 2.0, 2.2, 2.6),
            ],
        )
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("SAFE_RETURN_FAIL"),
                labels=["NRM", "AF20"],
                output_dir=out / "SAFE_RETURN_FAIL",
                backend=backend,
            )
            worker.error_occurred.connect(errors.append)
            worker.phase_changed.connect(phases.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertTrue(any("Hardware error at step AF20" in msg for msg in errors))
        self.assertEqual(backend.return_to_safe_state_calls, 1)
        self.assertEqual(finished, [True])
        self.assertEqual(phases[0], "preflight")
        self.assertIn("returning", phases)

    def test_successful_run_generates_bundle_and_steps(self) -> None:
        backend = DeterministicBackend([
            ("NRM", 1.0, 2.0, 3.0),
            ("AF20", 2.0, 2.2, 2.6),
        ])
        labels = ["NRM", "AF20"]
        step_results: list[object] = []
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("GOOD1"),
                labels=labels,
                output_dir=out / "GOOD1",
                backend=backend,
            )
            worker.step_complete.connect(step_results.append)
            worker.error_occurred.connect(errors.append)
            worker.run_finished.connect(finished.append)
            worker.phase_changed.connect(phases.append)
            worker.run()
            self.assertEqual(errors, [])
            self.assertEqual(finished, [False])
            self.assertEqual(phases[0], "preflight")
            self.assertEqual(phases[-1], "complete")
            self.assertEqual(len(step_results), 2)
            self.assertEqual(backend.preflight_calls, 1)
            self.assertEqual(backend.step_calls, labels)
            self.assertEqual(backend.squid_calls, 2)
            self.assertEqual(backend.susc_calls, 2)
            self.assertTrue((out / "GOOD1").exists())
            self.assertTrue((out / "GOOD1" / "GOOD1").exists())
            self.assertTrue((out / "GOOD1" / "GOOD1.rmg").exists())
            self.assertTrue((out / "GOOD1" / "measurements.txt").exists())
            susceptibility_summary = out / "GOOD1" / "susceptibility.json"
            self.assertTrue(susceptibility_summary.exists())
            susceptibility_payload = json.loads(susceptibility_summary.read_text(encoding="utf-8"))
            self.assertEqual(susceptibility_payload["schema"], "rapidpy.susceptibility.summary.v1")
            self.assertEqual(susceptibility_payload["sample"], "GOOD1")
            self.assertEqual(susceptibility_payload["record_count"], 2)
            self.assertEqual(susceptibility_payload["records"][0]["label"], "NRM")
            self.assertAlmostEqual(susceptibility_payload["records"][0]["susceptibility"], 0.05)
            self.assertTrue(susceptibility_payload["hardware_validation_required"])
            workflow_summary = out / "GOOD1" / "workflow_summary.json"
            self.assertTrue(workflow_summary.exists())
            workflow_payload = json.loads(workflow_summary.read_text(encoding="utf-8"))
            self.assertEqual(workflow_payload["schema"], "rapidpy.measurement.workflow_summary.v1")
            self.assertEqual(workflow_payload["sample"], "GOOD1")
            self.assertFalse(workflow_payload["aborted"])
            self.assertEqual(workflow_payload["final_phase"], "complete")
            self.assertIn("preflight", [row["phase"] for row in workflow_payload["phases"]])
            self.assertIn("returning", [row["phase"] for row in workflow_payload["phases"]])
            self.assertTrue(workflow_payload["hardware_validation_required"])
            artifact_index = out / "GOOD1" / "artifact_index.json"
            self.assertTrue(artifact_index.exists())
            artifact_payload = json.loads(artifact_index.read_text(encoding="utf-8"))
            self.assertEqual(artifact_payload["schema"], "rapidpy.measurement.artifact_index.v1")
            self.assertFalse(artifact_payload["aborted"])
            artifacts_by_name = {row["name"]: row for row in artifact_payload["artifacts"]}
            self.assertTrue(artifacts_by_name["vb6_specimen_file"]["exists"])
            self.assertTrue(artifacts_by_name["rmg_file"]["exists"])
            self.assertTrue(artifacts_by_name["magic_measurements"]["exists"])
            self.assertTrue(artifacts_by_name["magic_specimens"]["exists"])
            self.assertTrue(artifacts_by_name["susceptibility_summary"]["exists"])
            self.assertTrue(artifacts_by_name["workflow_summary"]["exists"])
            self.assertTrue(artifacts_by_name["communication_transcript"]["exists"])
            self.assertFalse(artifacts_by_name["quicklook_summary"]["required"])
            self.assertTrue(artifact_payload["hardware_validation_required"])
            transcript = out / "GOOD1" / "communication.tsv"
            self.assertTrue(transcript.exists())
            transcript_text = transcript.read_text(encoding="utf-8")
            self.assertIn("timestamp\tchannel\tdirection\tport\tpayload\tdetail", transcript_text)
            self.assertIn("\tTX\tGOOD1\tNRM\tset_demag_step", transcript_text)
            self.assertIn("\tRX\tGOOD1\t(1.0, 2.0, 3.0)\tread_squid", transcript_text)
            self.assertIn("\tRX\tGOOD1\t0.05\tread_susceptibility", transcript_text)
            self.assertEqual(
                (out / "GOOD1" / "GOOD1.rmg").read_text(encoding="latin-1").count("\n"), 2
            )
        self.assertIn("measuring", phases)
        allowed = {phase.value for phase in WorkflowPhase}
        self.assertTrue(set(phases).issubset(allowed))

    def test_backend_transport_events_are_published_once(self) -> None:
        class TranscriptBackend(DeterministicBackend):
            def __init__(self) -> None:
                super().__init__([("NRM", 1.0, 2.0, 3.0)])
                self._events: list[CommunicationEvent] = []

            def read_squid(self):
                self._events.extend((
                    CommunicationEvent(
                        timestamp=datetime(2026, 10, 2, tzinfo=timezone.utc),
                        channel="SQUID-2G",
                        direction=CommunicationDirection.TX,
                        port="COM7",
                        payload="XSC",
                        detail="counter axis=X latch_id=A-000001",
                    ),
                    CommunicationEvent(
                        timestamp=datetime(2026, 10, 2, tzinfo=timezone.utc),
                        channel="SQUID-2G",
                        direction=CommunicationDirection.RX,
                        port="COM7",
                        payload="12",
                        detail="counter axis=X latch_id=A-000001",
                    ),
                ))
                return super().read_squid()

            def communication_events(self):
                return tuple(self._events)

        class VacuumSource:
            simulated = False

            def __init__(self) -> None:
                self.events: list[CommunicationEvent] = []

            def communication_events(self):
                return tuple(self.events)

        backend = TranscriptBackend()
        vacuum = VacuumSource()

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("TRANSCRIPT"),
                labels=["NRM"],
                output_dir=out / "TRANSCRIPT",
                backend=backend,
                communication_sources=(vacuum,),
            )
            vacuum.events.extend((
                CommunicationEvent(
                    timestamp=datetime(2026, 10, 2, tzinfo=timezone.utc),
                    channel="VACUUM",
                    direction=CommunicationDirection.TX,
                    port="COM8",
                    payload="10MFF",
                    detail="vacuum command",
                ),
                CommunicationEvent(
                    timestamp=datetime(2026, 10, 2, tzinfo=timezone.utc),
                    channel="VACUUM",
                    direction=CommunicationDirection.RX,
                    port="COM8",
                    payload="MOTOR-ON",
                    detail="vacuum response",
                ),
            ))
            worker.run()
            worker._write_communication_transcript()

            transcript = (out / "TRANSCRIPT" / "communication.tsv").read_text(
                encoding="utf-8"
            )

        self.assertEqual(transcript.count("\tSQUID-2G\tTX\tCOM7\tXSC\t"), 1)
        self.assertEqual(transcript.count("\tSQUID-2G\tRX\tCOM7\t12\t"), 1)
        self.assertEqual(transcript.count("\tVACUUM\tTX\tCOM8\t10MFF\t"), 1)
        self.assertEqual(transcript.count("\tVACUUM\tRX\tCOM8\tMOTOR-ON\t"), 1)
        self.assertIn("\tmeasurement-worker\tINFO\tTRANSCRIPT\t\tbundle initialized", transcript)

    def test_transport_recovery_records_are_retained_in_workflow_summary(self) -> None:
        class RecoveryEvidenceBackend(DeterministicBackend):
            def __init__(self) -> None:
                super().__init__([("NRM", 1.0, 2.0, 3.0)])
                prior = self._record()
                prior = RecoveryRecord(
                    attempt=prior.attempt,
                    started_iso=prior.started_iso,
                    completed_iso=prior.completed_iso,
                    validation=None,
                    commands=prior.commands,
                    detail="previous run timeout",
                )
                self._transport_recoveries = (prior,)

            def read_squid(self):
                self._transport_recoveries += (self._record(),)
                return super().read_squid()

            @property
            def transport_recovery_records(self):
                return self._transport_recoveries

            @staticmethod
            def _record():
                return RecoveryRecord(
                    attempt=1,
                    started_iso="2026-10-02T00:00:00+00:00",
                    completed_iso="2026-10-02T00:00:03+00:00",
                    validation=None,
                    commands=(
                        CommandEvent(
                            index=0,
                            kind="squid.clear_reset",
                            detail="transport recovery",
                            started_iso="2026-10-02T00:00:01+00:00",
                            completed_iso="2026-10-02T00:00:02+00:00",
                            ok=True,
                            reply="",
                        ),
                    ),
                    detail="position-2 axis Y timed out",
                )

        backend = RecoveryEvidenceBackend()
        warnings: list[str] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("RECOVERY_EVIDENCE"),
                labels=["NRM"],
                output_dir=out / "RECOVERY_EVIDENCE",
                backend=backend,
            )
            worker.preflight_warning.connect(warnings.append)
            worker.run()
            payload = json.loads(
                (out / "RECOVERY_EVIDENCE" / "workflow_summary.json").read_text(
                    encoding="utf-8"
                )
            )

        self.assertEqual(payload["transport_recoveries"], 1)
        self.assertEqual(
            payload["transport_recovery_records"][0]["detail"],
            "position-2 axis Y timed out",
        )
        self.assertEqual(
            payload["transport_recovery_records"][0]["commands"][0]["kind"],
            "squid.clear_reset",
        )
        self.assertTrue(any("discarded and reacquired from zero-before" in row for row in warnings))

    def test_magnetometer_reading_quality_flags_are_operator_warnings(self) -> None:
        from rapid_main.magnetometer import MagnetometerReading

        class QualityBackend(DeterministicBackend):
            def read_squid(self) -> MagnetometerReading:
                self.squid_calls += 1
                return MagnetometerReading(
                    raw_volts=(9.8, 0.0, 0.0),
                    corrected_volts=(9.8, 0.0, 0.0),
                    moment_emu=(1.0e-5, 0.0, 0.0),
                    moment_magnitude_emu=1.0e-5,
                    flags=("near-saturation", "below-minimum-signal"),
                )

        backend = QualityBackend([("NRM", 0.0, 0.0, 0.0)])
        warnings: list[str] = []
        step_results: list[object] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("MAGQUAL"),
                labels=["NRM"],
                output_dir=out / "MAGQUAL",
                backend=backend,
            )
            worker.preflight_warning.connect(warnings.append)
            worker.step_complete.connect(step_results.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(finished, [False])
        self.assertEqual(len(step_results), 1)
        self.assertEqual(
            warnings,
            ["SQUID read quality at step NRM: near-saturation, below-minimum-signal"],
        )

    def test_worker_averages_configured_squid_cycle_and_emits_quality_stats(self) -> None:
        backend = DeterministicBackend(
            [
                ("NRM", 1.0, 2.0, 3.0),
                ("NRM", 3.0, 2.0, 1.0),
                ("NRM", 2.0, 4.0, 2.0),
            ]
        )
        results: list[object] = []

        with tempfile.TemporaryDirectory() as td:
            worker = MeasurementWorker(
                meta=_meta("CYCLE"),
                labels=["NRM"],
                output_dir=Path(td) / "CYCLE",
                backend=backend,
                samples_per_position=3,
            )
            worker.step_complete.connect(results.append)
            worker.run()

        self.assertEqual(backend.squid_calls, 3)
        self.assertEqual(len(results), 1)
        result = results[0]
        self.assertAlmostEqual(result.step.sdx, 2.0)
        self.assertAlmostEqual(result.step.sdy, 8.0 / 3.0)
        self.assertAlmostEqual(result.step.sdz, 2.0)
        self.assertEqual(result.cycle_stats.count, 3)
        self.assertEqual(result.cycle_stats.axis_ranges, (2.0, 2.0, 2.0))

    def test_worker_rejects_block_then_uses_backend_rezero_hook_before_retry(self) -> None:
        backend = FluxRecoveryBackend()
        with tempfile.TemporaryDirectory() as td:
            worker = MeasurementWorker(
                meta=_meta("FLUX_RECOVERY"),
                labels=["NRM"],
                output_dir=Path(td) / "FLUX_RECOVERY",
                backend=backend,
            )
            vector = worker._read_validated_squid_sample("NRM")

        self.assertEqual(vector, (1.0, 2.0, 3.0))
        self.assertEqual(backend.squid_calls, 2)
        self.assertEqual(len(backend.recovery_evidence), 1)
        self.assertEqual(backend.recovery_evidence[0].flux_step_axes, ("X",))

    def test_halt_during_step_marks_run_halted_without_error(self) -> None:
        backend = HaltingBackend(step_delay=0.5)
        errors: list[str] = []
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("HALT_STEP"),
                labels=["AF20"],
                output_dir=out / "HALT_STEP",
                backend=backend,
            )
            worker.error_occurred.connect(
                errors.append, QtCore.Qt.ConnectionType.DirectConnection
            )
            worker.phase_changed.connect(phases.append, QtCore.Qt.ConnectionType.DirectConnection)
            worker.run_finished.connect(
                finished.append, QtCore.Qt.ConnectionType.DirectConnection
            )

            thread = threading.Thread(target=worker.run)
            thread.daemon = True
            thread.start()

            self.assertTrue(backend.step_started.wait(1.0), "set_demag_step did not begin")
            worker.halt()
            thread.join(2.0)

        self.assertFalse(thread.is_alive(), "worker thread did not stop after halt")
        self.assertEqual(errors, [])
        self.assertEqual(finished, [True])
        self.assertIn("halted", phases)
        self.assertEqual(phases[-1], WorkflowPhase.RETURNING.value)

    def test_phase_sequence_covers_treatment_position_and_validation(self) -> None:
        backend = DeterministicBackend([
            ("NRM", 1.0, 2.0, 3.0),
        ])
        phases: list[str] = []
        finished: list[bool] = []

        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "out"
            worker = MeasurementWorker(
                meta=_meta("STATE1"),
                labels=["NRM"],
                output_dir=out / "STATE1",
                backend=backend,
            )
            worker.phase_changed.connect(phases.append)
            worker.run_finished.connect(finished.append)
            worker.run()

        self.assertEqual(finished, [False])
        idx_treating = phases.index("treating")
        idx_positioning = phases.index("positioning")
        idx_measuring = phases.index("measuring")
        idx_validating = phases.index("validating")
        idx_saving = max(i for i, p in enumerate(phases) if p == "saving")
        idx_returning = phases.index("returning")
        idx_complete = phases.index("complete")
        self.assertTrue(idx_treating < idx_positioning < idx_measuring < idx_validating < idx_saving)
        self.assertLess(idx_saving, idx_returning)
        self.assertLess(idx_returning, idx_complete)


if __name__ == "__main__":
    unittest.main(verbosity=2)
