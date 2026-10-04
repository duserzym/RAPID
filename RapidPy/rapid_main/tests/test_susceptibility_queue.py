"""Queue- and worker-level automated susceptibility acquisition.

These tests drive ``QueueHardwareBackend`` and ``MeasurementWorker`` with
injected fakes only. No serial port is opened and no axis is moved.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from rapid_main.communication_log import CommunicationDirection, CommunicationEvent
from rapid_main.config import AppConfig
from tests.af_fakes import configured_af
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import (
    HardwareError,
    PreflightResult,
    QueueAutomationError,
    QueueHardwareBackend,
    SusceptibilitySafeStateError,
)
from rapid_main.holder_state import HolderStateStore
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.susceptibility_acquisition import (
    SUSCEPTIBILITY_ACQUISITION_SCHEMA,
    SusceptibilityAcquisitionRecord,
)

from tests.acquisition_fakes import (
    FakeClock,
    FakeMotorSerialClient,
    FakeRawSquidClient,
    counts_for,
    counts_with_step,
)

CALIBRATION = (0.09000563, 0.106, 0.066)
ZERO_POS = -25886
MEAS_POS = -30607
COIL_POS = -20000
SAMPLE_TOP = 0
SAMPLE_BOTTOM = -2000  # SampleHeight = 2000 steps
TARGET = int(COIL_POS + (SAMPLE_TOP - SAMPLE_BOTTOM) / 2)
NOW = datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)


def _observations(*, zero_after_step_axis: int | None = None):
    zero_after = (
        counts_with_step((0.0, 0.0, 0.0), CALIBRATION, axis=zero_after_step_axis)
        if zero_after_step_axis is not None
        else counts_for((0.0, 0.0, 0.0), CALIBRATION)
    )
    return [
        counts_for((0.0, 0.0, 0.0), CALIBRATION),
        counts_for((0.20, 0.05, 0.30), CALIBRATION),
        counts_for((0.06, 0.19, 0.31), CALIBRATION),
        counts_for((-0.21, 0.04, 0.29), CALIBRATION),
        counts_for((-0.05, -0.20, 0.30), CALIBRATION),
        zero_after,
    ]


class _StubSquidAdapter:
    simulated = False

    def __init__(self, raw_client) -> None:
        self.raw_client = raw_client

    def is_connected(self) -> bool:
        return True

    def test_connection(self) -> bool:
        return True


class _Bridge:
    """Bartington bridge fake that records calls into a shared journal."""

    simulated = False

    def __init__(self, journal: list, values=(0.25,), *, connected: bool = True) -> None:
        self.journal = journal
        self.values = list(values)
        self.connected = connected
        self.fail_measure: Exception | None = None
        self.events: list[CommunicationEvent] = []

    def is_connected(self) -> bool:
        return self.connected

    def test_connection(self) -> bool:
        self.journal.append("bridge_connect")
        self.connected = True
        return True

    def zero(self) -> str:
        self.journal.append("bridge_zero")
        self._event(CommunicationDirection.TX, "Z\\r\\n")
        self._event(CommunicationDirection.RX, "OK\\r")
        return "OK\r"

    def measure(self) -> float:
        self.journal.append("bridge_measure")
        self._event(CommunicationDirection.TX, "M\\r\\n")
        if self.fail_measure is not None:
            self._event(CommunicationDirection.ERROR, str(self.fail_measure))
            raise self.fail_measure
        value = self.values.pop(0) if len(self.values) > 1 else self.values[0]
        self._event(CommunicationDirection.RX, f"{value}\\r")
        return value

    def communication_events(self):
        return tuple(self.events)

    def _event(self, direction, payload) -> None:
        self.events.append(
            CommunicationEvent(NOW, "SUSCEPTIBILITY", direction, payload, "COM7")
        )


class _JournalMotor(FakeMotorSerialClient):
    """Motor fake that also records lift calls into the shared journal."""

    journal: list = []

    def updown_move(self, axis, target: int, speed_index: int, wait_for_stop: bool = True):
        self.journal.append(("lift", int(target), int(speed_index)))
        return super().updown_move(axis, target, speed_index, wait_for_stop)

    def home_to_top(self, axis):
        self.journal.append("home")
        result = super().home_to_top(axis)
        if result.success:
            self.positions[axis.motor_id] = 0
        return result


def _config(tmp: Path, *, enabled: bool = True) -> AppConfig:
    cfg = AppConfig()
    cfg.general.nocomm = False
    cfg.general.data_dir = str(tmp)
    cfg.general.operator = "opr"
    cfg.changer.port = "COM3"
    cfg.squid.samples_per_pos = 1
    cfg.calibration.cal_x, cfg.calibration.cal_y, cfg.calibration.cal_z = CALIBRATION
    cfg.motion.zero_pos = ZERO_POS
    cfg.motion.meas_pos = MEAS_POS
    cfg.motion.sample_top = SAMPLE_TOP
    cfg.motion.sample_bottom = SAMPLE_BOTTOM
    cfg.susceptibility.enabled = enabled
    cfg.susceptibility.port = "COM7"
    cfg.susceptibility.coil_position = COIL_POS
    cfg.susceptibility.moment_factor_cgs = 2.0e-5
    return cfg


class _BackendCase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.holder_path = self.tmp / "holder_correction.json"
        self.store = HolderStateStore(self.holder_path, clock=lambda: NOW)
        self.journal: list = []

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _backend(self, *, bridge=None, cfg=None, observations=None) -> QueueHardwareBackend:
        cfg = cfg or _config(self.tmp)
        self.bridge = bridge if bridge is not None else _Bridge(self.journal, (0.25, 1.75))
        client = FakeRawSquidClient(observations or _observations() * 4)
        motor_cls = type("_Motor", (_JournalMotor,), {"journal": self.journal})
        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient", motor_cls):
            with mock.patch(
                "rapid_main.diagnostic_services.build_squid_backend",
                return_value=_StubSquidAdapter(client),
            ):
                with mock.patch(
                    "rapid_main.diagnostic_services.build_irm_arm_backend", return_value=object()
                ):
                    with mock.patch(
                        "rapid_main.diagnostic_services.build_af_demag_backend",
                        return_value=object(),
                    ):
                        backend = QueueHardwareBackend(
                            cfg,
                            holder_store=self.store,
                            susceptibility_backend=self.bridge,
                            clock=FakeClock(start=NOW),
                        )
        backend._client.is_connected = True
        backend._connected = True
        backend.preflight()
        return backend


class HolderSusceptibilityTests(_BackendCase):
    def test_holder_measures_bridge_before_squid_and_installs_both_atomically(self) -> None:
        backend = self._backend()

        backend.holder(0)

        correction = self.store.current
        self.assertIsNotNone(correction)
        self.assertEqual(correction.susceptibility_raw, 0.25)
        record = backend.susceptibility_acquisition_records[-1]
        self.assertTrue(record.is_holder)
        self.assertEqual(correction.susceptibility_evidence_id, record.acquisition_id)
        self.assertEqual(correction.susceptibility_measured_at_iso, record.completed_iso)
        # VB6 order: home (RapidPy safety), zero, slow move to coil, measure,
        # verified home, and only then the bracketed SQUID holder block.
        self.assertEqual(
            self.journal[:5],
            ["home", "bridge_zero", ("lift", TARGET, 0), "bridge_measure", "home"],
        )
        self.assertTrue(
            any(isinstance(item, tuple) and item[0] == "lift" for item in self.journal[5:]),
            "the bracketed SQUID holder block must follow the bridge acquisition",
        )
        evidence = self.tmp / "holder_susceptibility" / f"{record.acquisition_id}.json"
        payload = json.loads(evidence.read_text(encoding="utf-8"))
        self.assertEqual(payload["schema"], SUSCEPTIBILITY_ACQUISITION_SCHEMA)
        self.assertEqual(payload["bridge_scaled_value"], 0.25)
        self.assertTrue(payload["is_holder"])
        persisted = json.loads(self.holder_path.read_text(encoding="utf-8"))
        self.assertEqual(persisted["susceptibility_evidence_id"], record.acquisition_id)

    def test_holder_bridge_failure_keeps_previous_holder_and_records_evidence(self) -> None:
        backend = self._backend()
        backend.holder(0)
        before = self.holder_path.read_text(encoding="utf-8")
        first = self.store.current
        self.bridge.fail_measure = TimeoutError("bridge reply timeout")

        with self.assertRaisesRegex(HardwareError, "bridge reply timeout"):
            backend.holder(0)

        self.assertIs(self.store.current, first)
        self.assertEqual(self.holder_path.read_text(encoding="utf-8"), before)
        failed = backend.susceptibility_acquisition_records[-1]
        self.assertEqual(failed.outcome, "failed")
        self.assertTrue(failed.safe_state_confirmed)
        self.assertTrue(
            (self.tmp / "holder_susceptibility" / f"{failed.acquisition_id}.json").exists()
        )

    def test_rejected_magnetic_holder_discards_staged_bridge_value(self) -> None:
        backend = self._backend()
        backend.holder(0)
        first = self.store.current
        backend._bracketed = None
        backend._measurement = _StubSquidAdapter(
            FakeRawSquidClient(_observations(zero_after_step_axis=1) * 8)
        )
        backend._ensure_bracketed()

        with self.assertRaisesRegex(QueueAutomationError, "rejected"):
            backend.holder(0)

        self.assertIs(self.store.current, first)
        self.assertEqual(self.store.current.susceptibility_raw, 0.25)

    def test_holder_blocks_before_motion_when_bridge_configuration_is_invalid(self) -> None:
        cfg = _config(self.tmp)
        cfg.susceptibility.coil_position = 0
        backend = self._backend(cfg=cfg)
        self.journal.clear()
        calls_before = list(backend._client.calls)

        with self.assertRaisesRegex(QueueAutomationError, "coil position"):
            backend.holder(3)

        self.assertEqual(self.journal, [])
        self.assertEqual(backend._client.calls, calls_before)
        self.assertIsNone(self.store.current)

    def test_disabled_bridge_measures_magnetic_holder_only(self) -> None:
        backend = self._backend(cfg=_config(self.tmp, enabled=False))

        backend.holder(0)

        self.assertIsNone(self.store.current.susceptibility_raw)
        self.assertNotIn("bridge_zero", self.journal)


class HolderStagingGuardTests(unittest.TestCase):
    def _holder_record(self, **overrides):
        from dataclasses import replace

        values = dict(
            is_holder=True,
            bridge_scaled_value=0.25,
            holder_scaled_value=None,
            holder_evidence_id="",
        )
        values.update(overrides)
        return replace(_record("susc-holder-1"), **values)

    def test_only_a_completed_safe_finite_holder_record_is_staged(self) -> None:
        from rapid_main.holder_measurement import _with_susceptibility
        from rapid_main.holder_state import HolderCorrection, HolderStateError

        correction = HolderCorrection(holder_id="holder", measured_at_iso=NOW.isoformat())
        staged = _with_susceptibility(correction, self._holder_record())
        self.assertEqual(staged.susceptibility_raw, 0.25)
        self.assertEqual(staged.susceptibility_evidence_id, "susc-holder-1")
        self.assertIsNone(correction.susceptibility_raw)

        rejected = {
            "not a holder": dict(is_holder=False),
            "did not complete": dict(outcome="failed"),
            "safe state": dict(safe_state_confirmed=False),
            "non-finite": dict(bridge_scaled_value=float("inf")),
            "evidence identity": dict(acquisition_id=""),
            "simulation state": dict(simulated=True),
        }
        for reason, overrides in rejected.items():
            with self.subTest(reason=reason):
                with self.assertRaisesRegex(HolderStateError, reason):
                    _with_susceptibility(correction, self._holder_record(**overrides))


class SampleSusceptibilityTests(_BackendCase):
    def _with_holder(self) -> QueueHardwareBackend:
        backend = self._backend()
        backend.holder(0)
        self.journal.clear()
        return backend

    def test_sample_value_subtracts_holder_bridge_value_and_applies_factor(self) -> None:
        backend = self._with_holder()
        backend.set_measurement_context(sample_name="RS01A")

        value = backend.read_susceptibility()

        self.assertAlmostEqual(value, (1.75 - 0.25) * 2.0e-5)
        self.assertEqual(
            self.journal,
            ["home", "bridge_zero", ("lift", TARGET, 0), "bridge_measure", "home"],
        )
        record = backend.susceptibility_acquisition_records[-1]
        self.assertFalse(record.is_holder)
        self.assertEqual(record.sample_id, "RS01A")
        self.assertEqual(record.holder_scaled_value, 0.25)
        self.assertEqual(record.holder_evidence_id, self.store.current.susceptibility_evidence_id)
        self.assertEqual(record.measured_position, TARGET)
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(
            [event["direction"] for event in record.communication_events],
            ["TX", "RX", "TX", "RX"],
        )

    def test_plan_validation_accepts_susc_only_with_every_input_present(self) -> None:
        backend = self._with_holder()

        backend._af_demag = mock.Mock(simulated=False)
        backend._config.af_demag = configured_af()
        backend._af_demag.is_connected.return_value = True
        self.assertTrue(backend.validate_treatment_plan(("NRM", "SUSC", "AF20")).ok)
        backend.set_demag_step("SUSC")  # measurement-only, no actuator call
        self.assertEqual(self.journal, [])

    def test_plan_validation_blocks_without_holder_susceptibility(self) -> None:
        backend = self._backend()

        result = backend.validate_treatment_plan(("SUSC", "SUSC"))

        self.assertFalse(result.ok)
        self.assertEqual(len(result.blockers), 1)
        self.assertIn("Holder susceptibility unavailable", result.blockers[0])
        with self.assertRaisesRegex(HardwareError, "No holder correction"):
            backend.read_susceptibility()
        self.assertEqual(self.journal, [])

    def test_plan_validation_blocks_without_shared_or_live_bridge(self) -> None:
        backend = self._with_holder()
        backend._susceptibility = None
        self.assertIn(
            "No susceptibility bridge backend",
            backend.validate_treatment_plan(("SUSC",)).blockers[0],
        )

        simulated = _Bridge(self.journal)
        simulated.simulated = True
        backend._susceptibility = simulated
        self.assertIn("simulated", backend.validate_treatment_plan(("SUSC",)).blockers[0])

        class _Unavailable:
            simulated = False
            reason = "port COM7 missing"

            def is_connected(self) -> bool:
                return False

        backend._susceptibility = _Unavailable()
        self.assertIn("COM7 missing", backend.validate_treatment_plan(("SUSC",)).blockers[0])
        self.assertEqual(self.journal, [])

    def test_plan_validation_blocks_invalid_geometry_scale_and_factor(self) -> None:
        cases = {
            "coil position": lambda c: setattr(c.susceptibility, "coil_position", 0),
            "cross the configured": lambda c: setattr(c.susceptibility, "coil_position", -500),
            "moment factor": lambda c: setattr(c.susceptibility, "moment_factor_cgs", float("nan")),
            "scale factor": lambda c: setattr(c.susceptibility, "scale_factor", 0.0),
            "Sample height": lambda c: setattr(c.motion, "sample_top", SAMPLE_BOTTOM),
            "disabled": lambda c: setattr(c.susceptibility, "enabled", False),
        }
        for expected, mutate in cases.items():
            with self.subTest(expected=expected):
                backend = self._with_holder()
                mutate(backend._config)
                result = backend.validate_treatment_plan(("SUSC",))
                self.assertFalse(result.ok)
                self.assertIn(expected, result.blockers[0])
                with self.assertRaises(HardwareError):
                    backend.set_demag_step("SUSC")

    def test_stale_or_simulated_holder_blocks_sample_acquisition(self) -> None:
        backend = self._with_holder()
        self.store._clock = lambda: NOW + timedelta(days=2)

        with self.assertRaisesRegex(HardwareError, "stale"):
            backend.read_susceptibility()
        self.assertEqual(self.journal, [])

    def test_disconnected_shared_bridge_connects_only_during_acquisition(self) -> None:
        backend = self._with_holder()
        self.bridge.connected = False

        backend.read_susceptibility()

        self.assertEqual(self.journal[0], "bridge_connect")
        self.assertEqual(self.journal[1], "home")

    def test_measure_failure_returns_home_and_never_yields_a_value(self) -> None:
        backend = self._with_holder()
        self.bridge.fail_measure = ValueError("non-numeric reply 'ERR'")

        with self.assertRaisesRegex(HardwareError, "non-numeric reply"):
            backend.read_susceptibility()

        self.assertEqual(self.journal[-1], "home")
        record = backend.susceptibility_acquisition_records[-1]
        self.assertEqual(record.outcome, "failed")
        self.assertIsNone(record.susceptibility)
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.communication_events[-1]["direction"], "ERROR")

    def test_motion_mismatch_stops_before_measure(self) -> None:
        backend = self._with_holder()
        backend._client.lift_failure_target = TARGET

        with self.assertRaisesRegex(HardwareError, "Susceptibility acquisition failed"):
            backend.read_susceptibility()

        self.assertNotIn("bridge_measure", self.journal)
        self.assertEqual(self.journal[-1], "home")

    def test_safe_return_failure_is_a_distinct_high_severity_error(self) -> None:
        backend = self._with_holder()
        self.bridge.fail_measure = TimeoutError("measure timeout")
        original_home = backend._client.home_to_top
        calls = {"n": 0}

        def failing_second_home(axis):
            calls["n"] += 1
            if calls["n"] == 2:
                backend._client.home_failure = True
            return original_home(axis)

        backend._client.home_to_top = failing_second_home

        with self.assertRaises(SusceptibilitySafeStateError) as ctx:
            backend.read_susceptibility()

        message = str(ctx.exception)
        self.assertIn("SAFE-STATE NOT CONFIRMED", message)
        self.assertIn("measure timeout", message)
        record = backend.susceptibility_acquisition_records[-1]
        self.assertFalse(record.safe_state_confirmed)
        self.assertTrue(record.safe_return_error)

    def test_operator_halt_between_phases_cancels_and_returns_home(self) -> None:
        backend = self._with_holder()
        halted = {"value": False}
        backend.set_halt_check(lambda: halted["value"])
        original_zero = self.bridge.zero

        def zero_then_halt():
            halted["value"] = True
            return original_zero()

        self.bridge.zero = zero_then_halt

        with self.assertRaisesRegex(HardwareError, "cancelled"):
            backend.read_susceptibility()

        self.assertNotIn(("lift", TARGET, 0), self.journal)
        self.assertEqual(self.journal[-1], "home")

    def test_bridge_traffic_joins_backend_communication_events(self) -> None:
        backend = self._with_holder()
        backend.read_susceptibility()

        channels = [event.channel for event in backend.communication_events()]
        self.assertIn("SUSCEPTIBILITY", channels)


def _meta(name: str) -> SpecimenMeta:
    return SpecimenMeta(name=name)


def _record(acquisition_id: str, *, outcome="completed", simulated=False, value=0.5):
    return SusceptibilityAcquisitionRecord(
        acquisition_id=acquisition_id,
        sample_id="S1",
        is_holder=False,
        started_iso=NOW.isoformat(),
        completed_iso=NOW.isoformat(),
        outcome=outcome,
        coil_position=COIL_POS,
        sample_height=2000,
        target_position=TARGET,
        start_position=0,
        measured_position=TARGET,
        final_position=0,
        speed_index=0,
        bridge_scaled_value=1.0,
        holder_scaled_value=0.5,
        holder_evidence_id="susc-holder",
        moment_factor_cgs=1.0,
        susceptibility=value if outcome == "completed" else None,
        safe_state_confirmed=True,
        error="" if outcome == "completed" else "bridge timeout",
        simulated=simulated,
    )


class _WorkerBackend:
    """Live-declared backend whose bridge evidence the worker must publish."""

    simulated = False

    def __init__(self, *, fail_susc: bool = False, simulated_record: bool = False) -> None:
        self.calls: list[str] = []
        self.records: list[SusceptibilityAcquisitionRecord] = [_record("susc-before-run")]
        self.fail_susc = fail_susc
        self.simulated_record = simulated_record
        self.halt_check = None
        self.halt_probes_during_susc = []

    @property
    def susceptibility_acquisition_records(self):
        return tuple(self.records)

    def set_halt_check(self, check) -> None:
        self.halt_check = check

    def validate_treatment_plan(self, labels):
        return PreflightResult.pass_ok()

    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def is_available(self) -> bool:
        return True

    def set_demag_step(self, label: str) -> None:
        self.calls.append(f"treat:{label}")

    def read_squid(self):
        self.calls.append("squid")
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        self.calls.append("susc")
        self.halt_probes_during_susc.append((callable(self.halt_check), self.halt_check() if callable(self.halt_check) else None))
        index = len(self.records)
        if self.fail_susc:
            self.records.append(_record(f"susc-{index}", outcome="failed"))
            raise HardwareError("Susceptibility acquisition failed: bridge timeout")
        self.records.append(_record(f"susc-{index}", simulated=self.simulated_record))
        return 0.5

    def return_to_safe_state(self) -> None:
        self.calls.append("safe")


class WorkerSusceptibilityTests(unittest.TestCase):
    def _run(self, backend, labels, out: Path):
        errors: list[str] = []
        finished: list[bool] = []
        worker = MeasurementWorker(
            meta=_meta("S1"), labels=labels, output_dir=out, backend=backend, run_id="run-9"
        )
        worker.error_occurred.connect(errors.append)
        worker.run_finished.connect(finished.append)
        worker.run()
        return errors, finished

    def test_susceptibility_runs_before_treatment_and_squid_like_vb6(self) -> None:
        backend = _WorkerBackend()
        with tempfile.TemporaryDirectory() as td:
            errors, finished = self._run(backend, ["NRM", "SUSC"], Path(td) / "S1")

        self.assertEqual(errors, [])
        self.assertEqual(finished, [False])
        self.assertEqual(
            backend.calls, ["treat:NRM", "squid", "susc", "treat:SUSC", "squid", "safe"]
        )
        self.assertEqual(backend.halt_probes_during_susc, [(True, False)])
        self.assertIsNone(backend.halt_check)

    def test_failed_cancellation_hook_cleanup_emits_one_aborted_result(self) -> None:
        class FailedClear(_WorkerBackend):
            def set_halt_check(self, check):
                if check is None:
                    raise RuntimeError('cancellation hook could not be cleared')
                super().set_halt_check(check)
        with tempfile.TemporaryDirectory() as td:
            errors, finished = self._run(FailedClear(), ['NRM'], Path(td) / 'S1')
        self.assertEqual(finished, [True])
        self.assertTrue(any('cancellation hook could not be cleared' in error for error in errors))

    def test_current_run_acquisitions_are_published_and_indexed_with_digest(self) -> None:
        backend = _WorkerBackend()
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "S1"
            errors, _finished = self._run(backend, ["SUSC"], out)
            self.assertEqual(errors, [])
            artifact = out / "susceptibility_acquisitions" / "susc-1.json"
            payload = json.loads(artifact.read_text(encoding="utf-8"))
            self.assertEqual(payload["schema"], SUSCEPTIBILITY_ACQUISITION_SCHEMA)
            self.assertFalse(
                (out / "susceptibility_acquisitions" / "susc-before-run.json").exists()
            )
            index = json.loads((out / "artifact_index.json").read_text(encoding="utf-8"))
            rows = {row["name"]: row for row in index["artifacts"]}
            row = rows["susceptibility_acquisition:susc-1"]
            self.assertTrue(row["required"])
            self.assertEqual(row["relative_path"], "susceptibility_acquisitions/susc-1.json")
            self.assertEqual(row["size_bytes"], artifact.stat().st_size)
            self.assertEqual(row["sha256"], hashlib.sha256(artifact.read_bytes()).hexdigest())
            summary = json.loads((out / "susceptibility.json").read_text(encoding="utf-8"))
            self.assertEqual(summary["records"][0]["acquisition_id"], "susc-1")
            self.assertEqual(summary["records"][0]["holder_evidence_id"], "susc-holder")
            workflow = json.loads((out / "workflow_summary.json").read_text(encoding="utf-8"))
            self.assertEqual(workflow["susceptibility_acquisition_ids"], ["susc-1"])

    def test_failed_acquisition_aborts_without_output_but_keeps_evidence(self) -> None:
        backend = _WorkerBackend(fail_susc=True)
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "S1"
            errors, finished = self._run(backend, ["SUSC"], out)
            self.assertFalse((out / "S1").exists())
            self.assertFalse((out / "S1.rmg").exists())
            self.assertFalse((out / "susceptibility.json").exists())
            evidence = json.loads(
                (out / "susceptibility_acquisitions" / "susc-1.json").read_text(encoding="utf-8")
            )
            self.assertEqual(evidence["outcome"], "failed")
            self.assertIsNone(evidence["susceptibility"])

        self.assertEqual(finished, [True])
        self.assertIn("Susceptibility read error at step SUSC", errors[0])
        # Neither treatment nor SQUID runs after a rejected bridge read.
        self.assertEqual(backend.calls, ["susc", "safe"])

    def test_shared_bridge_traffic_is_transcribed_once(self) -> None:
        bridge = _Bridge([])

        class _Backend(_WorkerBackend):
            def communication_events(self):
                return bridge.communication_events()

            def read_susceptibility(self) -> float:
                bridge.zero()
                bridge.measure()
                return super().read_susceptibility()

        backend = _Backend()
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "S1"
            worker = MeasurementWorker(
                meta=_meta("S1"),
                labels=["SUSC"],
                output_dir=out,
                backend=backend,
                communication_sources=(bridge,),
            )
            worker.run()
            transcript = (out / "communication.tsv").read_text(encoding="utf-8")

        self.assertEqual(transcript.count("	SUSCEPTIBILITY	TX	"), 2)
        self.assertEqual(transcript.count("	SUSCEPTIBILITY	RX	"), 2)

    def test_live_backend_returning_simulated_acquisition_is_rejected(self) -> None:
        backend = _WorkerBackend(simulated_record=True)
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "S1"
            errors, finished = self._run(backend, ["SUSC"], out)
            self.assertFalse((out / "S1").exists())

        self.assertEqual(finished, [True])
        self.assertIn("simulated susceptibility", errors[0])


class SusceptibilityDecisionRecordTests(unittest.TestCase):
    def test_decision_record_agrees_with_the_implementation(self) -> None:
        from rapid_main.susceptibility_acquisition import SusceptibilityAcquisitionConfig

        path = Path(__file__).resolve().parents[3] / "docs" / "susceptibility_integration_decision.json"
        decision = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(decision["schema"], "rapidpy.susceptibility.integration_decision.v1")
        self.assertEqual(decision["evidence_schema"], SUSCEPTIBILITY_ACQUISITION_SCHEMA)
        self.assertEqual(decision["status"], "PENDING_PHYSICAL_ACCEPTANCE")
        self.assertEqual(
            SusceptibilityAcquisitionConfig(-1, 2, 1.0).speed_index, 0, "VB6 slow rod speed"
        )
        self.assertTrue(decision["open_physical_acceptance_questions"])
        self.assertEqual(
            {row["path"] for row in decision["legacy_source_evidence"]},
            {
                "VB6/modSusceptibility.bas",
                "VB6/modMeasure.bas",
                "VB6/frmSusceptibilityMeter.frm",
                "VB6/frmCalRod.frm",
            },
        )
        for row in decision["legacy_source_evidence"]:
            self.assertTrue((path.parents[1] / row["path"]).exists(), row["path"])


if __name__ == "__main__":
    unittest.main()
