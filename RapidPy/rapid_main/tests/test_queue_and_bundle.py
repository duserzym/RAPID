from __future__ import annotations

import sys
import tempfile
import unittest
from unittest import mock
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from rapid_main.data_model import (
    AngleVsFieldCollection,
    AngleVsFieldPoint,
    MeasurementBlock,
    MeasurementBlocks,
    MeasurementStep,
    ProbeAngleOptimizer,
    RockmagStep,
    RockmagSteps,
    SpecimenMeta,
)
from rapid_main.io.cit_sam import CitSamEntry, CitSamHeader, read_cit_sam, write_cit_sam
from rapid_main.io.measurement_bundle import MeasurementBundleWriter
from rapid_main.queue_compiler import (
    QueueOptions,
    QueueSample,
    QueueValidationResult,
    compile_queue,
    validate_queue_samples,
)
from rapid_main.hardware_contracts import (
    HardwareError,
    HardwareBackend,
    MeasurementBackend,
    NoCommBackend,
    _parse_demag_label,
    _parse_thermal_label,
    PreflightResult,
    QueueHardwareBackend,
)
from rapid_main.config import AppConfig


def _step(label: str = "AF20") -> MeasurementStep:
    return MeasurementStep(
        demag_label=label,
        gdec=10.0,
        ginc=30.0,
        sdec=12.0,
        sinc=28.0,
        moment=1.11e-6,
        error_angle=4.2,
        crdec=11.0,
        crinc=29.0,
        sdx=9.0e-7,
        sdy=7.0e-7,
        sdz=5.0e-7,
        operator="tester",
        timestamp=datetime(2026, 5, 20, 10, 0, 0),
    )


def _meta(name: str = "TEST01") -> SpecimenMeta:
    return SpecimenMeta(
        name=name,
        comment="bundle test",
        core_plate_strike=100.0,
        core_plate_dip=35.0,
        bedding_strike=320.0,
        bedding_dip=15.0,
        volume=8.5,
        sample="SAMP-A",
        site="SITE-A",
        location="LAB-A",
    )


class TestCitSamIO(unittest.TestCase):
    def test_roundtrip(self) -> None:
        header = CitSamHeader(
            format_id="CIT",
            comment="Example locality",
            latitude_n=36.2,
            longitude_e=245.3,
            magnetic_declination_e=14.0,
            fold_axis_azimuth=42.2,
            fold_axis_plunge=45.8,
            bedding_strike=330.0,
            bedding_dip=20.0,
        )
        entries = [
            CitSamEntry("erb1.0a", 12.3, "aa"),
            CitSamEntry("erb2.0a", 23.4, "ab"),
        ]

        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "test.sam"
            write_cit_sam(path, header, entries)
            got_header, got_entries = read_cit_sam(path)

        self.assertEqual(got_header.format_id, "CIT")
        self.assertEqual(got_header.comment, "Example locality")
        self.assertAlmostEqual(got_header.latitude_n, 36.2, places=1)
        self.assertEqual(len(got_entries), 2)
        self.assertEqual(got_entries[0].specimen_name, "erb1.0a")


class TestQueueCompiler(unittest.TestCase):
    def test_compile_core_path(self) -> None:
        samples = [
            QueueSample("S1", "F1", 1, do_up=True, do_both=True, measurement_step_count=1),
            QueueSample("S2", "F1", 2, do_up=True, do_both=True, measurement_step_count=1),
            QueueSample("S3", "F2", 3, do_up=True, do_both=False, measurement_step_count=3),
        ]
        options = QueueOptions(samples_between_holder=2)
        queue = compile_queue(samples, options)

        command_types = [cmd.command_type for cmd in queue]
        self.assertIn("InitUp", command_types)
        self.assertEqual(command_types[0], "InitUp")
        self.assertEqual(command_types.count("Meas"), 5)
        self.assertGreaterEqual(command_types.count("Holder"), 2)
        self.assertEqual(command_types[-1], "Goto")

    def test_validate_queue_samples_detects_duplicate_hole(self) -> None:
        samples = [
            QueueSample("S1", "F1", 4, measurement_step_count=1),
            QueueSample("S2", "F2", 4, measurement_step_count=1),
        ]
        result = validate_queue_samples(samples)
        self.assertIsInstance(result, QueueValidationResult)
        self.assertTrue(result.is_valid())
        self.assertEqual(len(result.warnings), 1)
        self.assertIn("duplicate hole 4", result.warnings[0])

    def test_compile_queue_strict_mode_rejects_invalid_samples(self) -> None:
        samples = [
            QueueSample("", "F1", 0, measurement_step_count=0),
            QueueSample("S2", "", 1, measurement_step_count=1),
        ]
        options = QueueOptions()
        with self.assertRaises(ValueError) as context:
            compile_queue(samples, options, strict=True)
        self.assertIn("sample_name is required", str(context.exception))
        self.assertIn("hole must be >= 1", str(context.exception))


class TestMeasurementBlocks(unittest.TestCase):
    def test_measurement_block_normalizes_and_flattens_labels(self) -> None:
        block = MeasurementBlock(
            name="  Rockmag warmup  ",
            block_type=" rockmag ",
            labels=[" NRM ", "", "AF20", " TT200 "],
            metadata={"template": "Rockmag the Works", "passes": 2},
        )

        self.assertEqual(block.name, "Rockmag warmup")
        self.assertEqual(block.block_type, "rockmag")
        self.assertEqual(block.step_count, 3)
        self.assertEqual(block.to_queue_labels(), ["NRM", "AF20", "TT200"])
        self.assertEqual(block.metadata["passes"], "2")

    def test_measurement_blocks_preserve_order_for_runner_labels(self) -> None:
        blocks = MeasurementBlocks(
            [
                MeasurementBlock("NRM", ["NRM"]),
                MeasurementBlock("AF", ["AF10", "AF20"]),
                MeasurementBlock("IRM", ["IRM100"]),
            ]
        )

        self.assertEqual(blocks.block_names(), ["NRM", "AF", "IRM"])
        self.assertEqual(blocks.step_count, 4)
        self.assertEqual(blocks.to_queue_labels(), ["NRM", "AF10", "AF20", "IRM100"])


class TestRockmagStepModels(unittest.TestCase):
    def test_rockmag_step_parses_common_sequence_labels(self) -> None:
        af = RockmagStep.from_label("AF20")
        irm = RockmagStep.from_label("IRM100")
        arm = RockmagStep.from_label("ARM5")
        rrm = RockmagStep.from_label("RRM-0.5")
        susc = RockmagStep.from_label("SUSC")

        self.assertEqual((af.family, af.value, af.unit), ("AF", 20.0, "mT"))
        self.assertEqual((irm.family, irm.value, irm.unit), ("IRM", 100.0, "G"))
        self.assertEqual((arm.family, arm.value, arm.unit), ("ARM", 5.0, "G"))
        self.assertEqual((rrm.family, rrm.value, rrm.unit), ("RRM", -0.5, "rps"))
        self.assertEqual((susc.family, susc.value, susc.unit), ("SUSCEPTIBILITY", None, ""))

    def test_rockmag_steps_preserve_order_and_block_conversion(self) -> None:
        routine = RockmagSteps.from_labels(
            ["NRM", "AF20", "IRM100", "IRM-BF", "SUSC"],
            routine_name="Rockmag the Works",
        )
        block = routine.to_measurement_block()

        self.assertEqual(routine.step_count, 5)
        self.assertEqual(routine.labels, ["NRM", "AF20", "IRM100", "IRM-BF", "SUSC"])
        self.assertEqual(routine.families(), ["NRM", "AF", "IRM", "BACKFIELD", "SUSCEPTIBILITY"])
        self.assertEqual(block.name, "Rockmag the Works")
        self.assertEqual(block.block_type, "rockmag")
        self.assertEqual(block.to_queue_labels(), routine.labels)
        self.assertEqual(block.metadata["families"], "NRM,AF,IRM,BACKFIELD,SUSCEPTIBILITY")

    def test_rockmag_step_rejects_empty_label(self) -> None:
        with self.assertRaises(ValueError):
            RockmagStep.from_label("  ")


class TestAngleVsFieldModels(unittest.TestCase):
    def test_angle_vs_field_collection_sorts_and_finds_peak(self) -> None:
        collection = AngleVsFieldCollection(
            [
                AngleVsFieldPoint(30.0, 4.2),
                AngleVsFieldPoint(10.0, 1.5),
                AngleVsFieldPoint(20.0, 6.1),
            ]
        )

        self.assertEqual(collection.angles, [10.0, 20.0, 30.0])
        self.assertEqual(collection.fields, [1.5, 6.1, 4.2])
        peak = collection.peak_point()
        self.assertIsNotNone(peak)
        self.assertEqual(peak.angle_deg, 20.0)

    def test_angle_vs_field_collection_weighted_center(self) -> None:
        collection = AngleVsFieldCollection(
            [
                AngleVsFieldPoint(0.0, 1.0, weight=1.0),
                AngleVsFieldPoint(10.0, 2.0, weight=3.0),
                AngleVsFieldPoint(100.0, 9.0, weight=0.0),
            ]
        )

        self.assertAlmostEqual(collection.weighted_center_angle(), 7.5)

    def test_angle_vs_field_collection_handles_empty_or_zero_weight(self) -> None:
        self.assertIsNone(AngleVsFieldCollection([]).peak_point())
        self.assertIsNone(AngleVsFieldCollection([]).weighted_center_angle())
        self.assertIsNone(
            AngleVsFieldCollection(
                [AngleVsFieldPoint(15.0, 3.0, weight=-4.0)]
            ).weighted_center_angle()
        )

    def test_probe_angle_optimizer_selects_peak_field_angle(self) -> None:
        optimizer = ProbeAngleOptimizer(
            AngleVsFieldCollection(
                [
                    AngleVsFieldPoint(0.0, 2.0),
                    AngleVsFieldPoint(10.0, 8.0),
                    AngleVsFieldPoint(20.0, 5.0),
                ]
            )
        )

        result = optimizer.optimize()

        self.assertIsNotNone(result)
        assert result is not None
        self.assertEqual(result.angle_deg, 10.0)
        self.assertEqual(result.field_value, 8.0)
        self.assertAlmostEqual(result.confidence, 0.75)
        self.assertEqual(result.method, "peak-field")

    def test_probe_angle_optimizer_handles_empty_collection(self) -> None:
        self.assertIsNone(ProbeAngleOptimizer(AngleVsFieldCollection([])).optimize())


class TestMeasurementBundleWriter(unittest.TestCase):
    def test_bundle_writes_all_targets(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            writer = MeasurementBundleWriter(td, _meta("BUNDLE01"))
            writer.append_step(_step("NRM"), susceptibility=0.01)
            writer.append_step(_step("AF20"), susceptibility=0.005)

            # Nothing is published until the transaction commits.
            self.assertFalse(writer.paths.specimen_file.exists())
            writer.commit()

            specimen = writer.paths.specimen_file
            rmg = writer.paths.rmg_file
            meas = writer.paths.magic_measurements_file
            specs = writer.paths.magic_specimens_file

            self.assertTrue(specimen.exists())
            self.assertTrue(rmg.exists())
            self.assertTrue(meas.exists())
            self.assertTrue(specs.exists())

            rmg_lines = rmg.read_text(encoding="latin-1").splitlines()
            self.assertEqual(len(rmg_lines), 2)
            self.assertIn("0.01", rmg_lines[0])

            meas_lines = meas.read_text(encoding="utf-8").splitlines()
            self.assertEqual(meas_lines[0], "tab delimited\tmeasurements")
            self.assertGreaterEqual(len(meas_lines), 4)

            spec_lines = specs.read_text(encoding="utf-8").splitlines()
            self.assertEqual(spec_lines[0], "tab delimited\tspecimens")
            self.assertEqual(len(spec_lines), 3)


class TestHardwareContracts(unittest.TestCase):
    def test_hardware_backend_alias(self) -> None:
        self.assertIs(HardwareBackend, MeasurementBackend)

    def test_preflight_result_helpers(self) -> None:
        ok = PreflightResult.pass_ok()
        self.assertTrue(ok.ok)
        self.assertEqual(ok.blockers, ())
        self.assertEqual(ok.warnings, ())

        blocked = PreflightResult.blocked("not connected", warnings=("review needed",))
        self.assertFalse(blocked.ok)
        self.assertEqual(blocked.blockers, ("not connected",))
        self.assertEqual(blocked.warnings, ("review needed",))

    def test_build_measurement_backend_nocomm_is_noop_default(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        from rapid_main.hardware_contracts import build_measurement_backend

        backend = build_measurement_backend(cfg)
        self.assertIsInstance(backend, NoCommBackend)

    def test_hardware_backend_reuses_injected_susceptibility_owner(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = False
        shared = object()
        from rapid_main.hardware_contracts import build_measurement_backend

        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient"):
            with mock.patch("rapid_main.diagnostic_services.build_squid_backend", return_value=object()):
                with mock.patch("rapid_main.diagnostic_services.build_irm_arm_backend", return_value=object()):
                    with mock.patch("rapid_main.diagnostic_services.build_af_demag_backend", return_value=object()):
                        backend = build_measurement_backend(
                            cfg,
                            susceptibility_backend=shared,
                        )

        self.assertIs(backend.susceptibility_backend, shared)

    def test_queue_backend_preflight_blocks_without_configured_lift_positions(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = False
        cfg.changer.port = "COM3"

        backend = QueueHardwareBackend(cfg)
        backend._connected = True

        class _StubMeasurement:
            def is_connected(self) -> bool:
                return True

            def test_connection(self) -> bool:
                return True

        backend._measurement = _StubMeasurement()
        result = backend.preflight()

        self.assertFalse(result.ok)
        self.assertTrue(
            any("lift positions are not configured" in blocker for blocker in result.blockers),
            result.blockers,
        )

    def test_queue_backend_preflight_treats_warnings_as_non_blocking(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = False
        cfg.changer.port = "COM3"
        cfg.motion.zero_pos = -25886
        cfg.motion.meas_pos = -30607

        backend = QueueHardwareBackend(cfg)
        backend._connected = True

        class _StubMeasurement:
            def is_connected(self) -> bool:
                return True

            def test_connection(self) -> bool:
                return True

        backend._measurement = _StubMeasurement()
        # Simulate a station where every component constructed successfully.
        backend._backend_errors.clear()
        backend._collect_preflight_warnings = lambda: ["measurement service degraded"]  # type: ignore[method-assign]
        result = backend.preflight()
        self.assertTrue(result.ok, f"warnings should not block preflight: {result.warnings}")
        self.assertIn("measurement service degraded", result.warnings)
        self.assertEqual(result.blockers, ())
        # A stub SQUID cannot supply a raw 2G client, so bracketed acquisition
        # is reported as unavailable rather than silently skipped.
        self.assertTrue(
            any("Bracketed SQUID acquisition unavailable" in warning for warning in result.warnings),
            result.warnings,
        )

    def test_parse_demag_label_recognizes_af_irm_arm_patterns(self) -> None:
        self.assertEqual(_parse_demag_label("AF50"), ("AF", 50.0, None))
        self.assertEqual(_parse_demag_label("AFZ"), ("AFZ", None, None))
        self.assertEqual(_parse_demag_label("AFMAX"), ("AFMAX", None, None))
        self.assertEqual(_parse_demag_label("IRM1000"), ("IRM", 1000.0, None))
        self.assertEqual(_parse_demag_label("ARM120_2.5"), ("ARM", 120.0, 2.5))
        self.assertEqual(_parse_demag_label("NRM"), ("NRM", None, None))

    def test_parse_thermal_label_recognizes_temperature_steps(self) -> None:
        self.assertEqual(_parse_thermal_label("TT400"), ("TT", 400.0))
        self.assertEqual(_parse_thermal_label("TH 350"), ("TH", 350.0))
        self.assertEqual(_parse_thermal_label("TEMP25.5"), ("TEMP", 25.5))
        self.assertEqual(_parse_thermal_label("AF20"), ("AF20", None))

    def test_queue_backend_routes_af_irm_and_arm_to_treatment_backend(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)

        class _RecordMeasurement:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def set_demag_step(self, label: str) -> None:
                self.calls.append(label)

            def is_connected(self) -> bool:
                return True

            def preflight(self) -> None:
                return None

            def test_connection(self) -> bool:
                return True

            def read_squid(self):
                return (0.0, 0.0, 0.0)

            def read_susceptibility(self) -> float:
                return 0.0

        class _RecordIrmArm:
            def __init__(self) -> None:
                self.calls: list[tuple[str, float | int | str | None, str | None, str | None, int | None]] = []

            def apply_irm(self, *, max_field_mT: float, axis: str, ramp_label: str, steps: int) -> str:
                self.calls.append(("irm", max_field_mT, axis, ramp_label, steps))
                return "ok"

            def apply_arm(self, *, peak_af_mT: float, bias_mT: float, steps: int | None = None) -> str:
                self.calls.append(("arm", peak_af_mT, None, None, int(steps or 0)))
                return "ok"

            def is_connected(self) -> bool:
                return True

        class _RecordAfDemag:
            def __init__(self) -> None:
                self.calls: list[object] = []
                self.reset_count = 0

            def apply_af(self, command: object) -> str:
                self.calls.append(command)
                return "af ok"

            def reset_field(self) -> str:
                self.reset_count += 1
                return "af reset"

            def is_connected(self) -> bool:
                return True

        meas = _RecordMeasurement()
        arm = _RecordIrmArm()
        af = _RecordAfDemag()
        backend._measurement = meas
        backend._irm_arm = arm
        backend._af_demag = af

        backend.set_demag_step("AF25")
        backend.set_demag_step("IRM40")
        backend.set_demag_step("ARM100_1.2")
        backend.set_demag_step("NRM")

        self.assertEqual(len(af.calls), 1)
        self.assertEqual(getattr(af.calls[0], "field_mT"), 25.0)
        self.assertEqual(
            arm.calls,
            [
                ("irm", 40.0, cfg.irm_arm.irm_axis, cfg.irm_arm.irm_ramp, cfg.irm_arm.irm_steps),
                ("arm", 100.0, None, None, cfg.irm_arm.irm_steps),
            ],
        )
        self.assertEqual(meas.calls, ["NRM"])

    def test_queue_backend_rejects_irm_without_field(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)
        backend._measurement = type("M", (), {"is_connected": lambda self: True, "read_squid": lambda self: (0,0,0), "read_susceptibility": lambda self: 0.0, "preflight": lambda self: type("x",(object,),{"ok":True})(), "test_connection": lambda self: True, "set_demag_step": lambda self,label: (_ for _ in ()).throw(Exception("should not be called"))})()

        class _NoopIrmArm:
            def is_connected(self) -> bool:
                return True

        backend._irm_arm = _NoopIrmArm()

        with self.assertRaises(ValueError):
            backend.set_demag_step("IRM")

    def test_queue_backend_routes_thermal_steps_when_backend_supports_it(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)
        cfg.general.nocomm = False

        class _ThermalMeasurement:
            def __init__(self) -> None:
                self.thermal_calls: list[tuple[float, str]] = []
                self.measurement_calls: list[str] = []

            def apply_thermal(self, *, temperature_c: float, label: str) -> None:
                self.thermal_calls.append((temperature_c, label))

            def set_demag_step(self, label: str) -> None:
                self.measurement_calls.append(label)

            def is_connected(self) -> bool:
                return True

            def test_connection(self) -> bool:
                return True

            def read_squid(self) -> tuple[float, float, float]:
                return (0.0, 0.0, 0.0)

            def read_susceptibility(self) -> float:
                return 0.0

        meas = _ThermalMeasurement()
        backend._measurement = meas

        backend.set_demag_step("TT400")
        backend.set_demag_step("NRM")

        self.assertTrue(backend.validate_treatment_plan(("TT400", "NRM")).ok)
        self.assertEqual(meas.thermal_calls, [(400.0, "TT400")])
        self.assertEqual(meas.measurement_calls, ["NRM"])

    def test_queue_backend_blocks_live_thermal_without_furnace_adapter(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)
        cfg.general.nocomm = False

        class _PlanningOnlyMeasurement:
            def __init__(self) -> None:
                self.labels: list[str] = []

            def set_demag_step(self, label: str) -> None:
                self.labels.append(label)

        measurement = _PlanningOnlyMeasurement()
        backend._measurement = measurement

        preflight = backend.validate_treatment_plan(("NRM", "TT400", "TH500"))

        self.assertFalse(preflight.ok)
        self.assertEqual(len(preflight.blockers), 2)
        self.assertTrue(all("furnace/oven adapter" in item for item in preflight.blockers))
        with self.assertRaisesRegex(HardwareError, "planning-only"):
            backend.set_demag_step("TT400")
        self.assertEqual(measurement.labels, [])

    def test_queue_backend_validates_live_measurement_and_treatment_routes(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)
        cfg.general.nocomm = False

        class _RecordingMeasurement:
            def __init__(self) -> None:
                self.labels: list[str] = []

            def set_demag_step(self, label: str) -> None:
                self.labels.append(label)

        measurement = _RecordingMeasurement()
        backend._measurement = measurement

        measurement_only = ("NRM", "NRM-X", "NRM-Y", "NRM-Z", "REPEAT2")
        self.assertTrue(backend.validate_treatment_plan(measurement_only).ok)
        for label in measurement_only:
            backend.set_demag_step(label)
        self.assertEqual(measurement.labels, list(measurement_only))

        unsupported = (
            "IRM-BF",
            "IRM",
            "RRM-0.5",
            "PTRM400",
            "ZF400",
            "IF400",
            "CUSTOM",
            "AFBOGUS",
            "SUSC",
        )
        preflight = backend.validate_treatment_plan(unsupported)

        self.assertFalse(preflight.ok)
        self.assertEqual(len(preflight.blockers), len(unsupported))
        for label in unsupported:
            if label == "SUSC":
                self.assertTrue(
                    any("Susceptibility step SUSC" in item for item in preflight.blockers)
                )
            else:
                self.assertTrue(any(repr(label) in item for item in preflight.blockers))
        for label in unsupported:
            if label == "IRM":
                with self.assertRaisesRegex(ValueError, "requires numeric field"):
                    backend.set_demag_step(label)
            elif label == "SUSC":
                with self.assertRaisesRegex(HardwareError, "SUSC cannot run in hardware mode"):
                    backend.set_demag_step(label)
            else:
                with self.assertRaisesRegex(HardwareError, "no production actuator route"):
                    backend.set_demag_step(label)
        self.assertEqual(measurement.labels, list(measurement_only))
        self.assertEqual(backend._treatment_label, "REPEAT2")

    def test_queue_backend_allows_explicit_simulator_only_label(self) -> None:
        cfg = AppConfig()
        cfg.general.nocomm = True
        backend = QueueHardwareBackend(cfg)

        class _RecordingMeasurement:
            def __init__(self) -> None:
                self.labels: list[str] = []

            def set_demag_step(self, label: str) -> None:
                self.labels.append(label)

        measurement = _RecordingMeasurement()
        backend._measurement = measurement

        self.assertTrue(backend.validate_treatment_plan(("CUSTOM-SIM",)).ok)
        backend.set_demag_step("CUSTOM-SIM")

        self.assertEqual(measurement.labels, ["CUSTOM-SIM"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
