"""Transactional output, simulation isolation, and metadata provenance."""
from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.data_model import MeasurementStep, SampleIndexRegistration, SampleIndexRegistrations, SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend, PreflightResult
from rapid_main.io.measurement_bundle import (
    SIMULATED_SUBDIR,
    SIMULATION_MARKER_FILE,
    BundleNotCommittedError,
    MeasurementBundleWriter,
)
from rapid_main.magnetometer import BlockAudit, BracketedMeasurementBlock
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.specimen_metadata import resolve_specimen_meta


def _meta(name: str = "OUT01") -> SpecimenMeta:
    return SpecimenMeta(name=name, comment="test specimen", volume=10.5, sample="OUT01", site="SITE", location="LOC")


def _step(label: str, moment: float = 1.0e-6) -> MeasurementStep:
    return MeasurementStep(
        demag_label=label,
        gdec=10.0,
        ginc=20.0,
        sdec=10.0,
        sinc=20.0,
        moment=moment,
        error_angle=1.0,
        crdec=10.0,
        crinc=20.0,
        sdx=1.0e-6,
        sdy=2.0e-6,
        sdz=3.0e-6,
        operator="opr",
        timestamp=datetime(2026, 8, 29, 12, 0, 0),
    )


class TransactionalBundleTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_abort_leaves_no_accepted_output(self) -> None:
        writer = MeasurementBundleWriter(self.tmp, _meta())
        writer.append_step(_step("NRM"))

        writer.abort()

        self.assertFalse(writer.paths.specimen_file.exists())
        self.assertFalse(writer.paths.rmg_file.exists())
        self.assertFalse(writer.paths.magic_measurements_file.exists())
        self.assertEqual(list(self.tmp.glob(".rapidpy-pending*")), [])

    def test_exception_inside_the_context_manager_aborts(self) -> None:
        with self.assertRaises(RuntimeError):
            with MeasurementBundleWriter(self.tmp, _meta()) as writer:
                writer.append_step(_step("NRM"))
                raise RuntimeError("simulated crash mid-run")

        self.assertFalse((self.tmp / "OUT01").exists())

    def test_commit_publishes_every_target_and_clears_staging(self) -> None:
        with MeasurementBundleWriter(self.tmp, _meta()) as writer:
            writer.append_step(_step("NRM"))
            writer.append_step(_step("AF20"))

        self.assertTrue((self.tmp / "OUT01").exists())
        self.assertTrue((self.tmp / "OUT01.rmg").exists())
        self.assertTrue((self.tmp / "measurements.txt").exists())
        self.assertTrue((self.tmp / "specimens.txt").exists())
        self.assertEqual(list(self.tmp.glob(".rapidpy-pending*")), [])

    def test_use_after_commit_is_refused(self) -> None:
        writer = MeasurementBundleWriter(self.tmp, _meta())
        writer.append_step(_step("NRM"))
        writer.commit()

        with self.assertRaises(BundleNotCommittedError):
            writer.append_step(_step("AF20"))

    def test_resume_skips_labels_already_published(self) -> None:
        with MeasurementBundleWriter(self.tmp, _meta()) as first:
            first.append_step(_step("NRM"))
            first.append_step(_step("AF20"))
        first_specimen = (self.tmp / "OUT01").read_text(encoding="latin-1")

        with MeasurementBundleWriter(self.tmp, _meta(), resume=True) as second:
            self.assertFalse(second.append_step(_step("NRM")))
            self.assertFalse(second.append_step(_step("AF20")))
            self.assertTrue(second.append_step(_step("AF40")))
            self.assertEqual(second.skipped_labels, ("NRM", "AF20"))

        resumed = (self.tmp / "OUT01").read_text(encoding="latin-1")
        self.assertTrue(resumed.startswith(first_specimen))
        self.assertEqual(resumed.count("NRM"), 1)
        self.assertEqual(resumed.count("AF20"), 1)
        self.assertEqual(resumed.count("AF40"), 1)

    def test_provenance_records_metadata_and_holder_association(self) -> None:
        with MeasurementBundleWriter(self.tmp, _meta()) as writer:
            writer.append_step(_step("NRM"))
            writer.set_provenance(
                holder_record_id="holder@2026-08-29T12:00:00+00:00",
                run_id="run-42",
                config_hash="cfg-abc",
            )

        payload = json.loads((self.tmp / "provenance.json").read_text(encoding="utf-8"))
        self.assertEqual(payload["specimen"], "OUT01")
        self.assertEqual(payload["sample"], "OUT01")
        self.assertEqual(payload["site"], "SITE")
        self.assertEqual(payload["location"], "LOC")
        self.assertEqual(payload["comment"], "test specimen")
        self.assertEqual(payload["volume_cm3"], 10.5)
        self.assertEqual(payload["holder_record_id"], "holder@2026-08-29T12:00:00+00:00")
        self.assertEqual(payload["run_id"], "run-42")
        self.assertEqual(payload["labels"], ["NRM"])
        self.assertFalse(payload["simulated"])


class SimulationIsolationTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_simulated_output_never_lands_in_the_production_path(self) -> None:
        with MeasurementBundleWriter(self.tmp, _meta(), simulated=True) as writer:
            writer.append_step(_step("NRM"))

        self.assertFalse((self.tmp / "OUT01").exists())
        published = self.tmp / SIMULATED_SUBDIR
        self.assertTrue((published / "OUT01").exists())
        marker = (published / SIMULATION_MARKER_FILE).read_text(encoding="utf-8")
        self.assertIn("SIMULATED RUN", marker)
        payload = json.loads((published / "provenance.json").read_text(encoding="utf-8"))
        self.assertTrue(payload["simulated"])

    def test_simulated_output_can_be_enabled_for_a_deliberate_test(self) -> None:
        with MeasurementBundleWriter(
            self.tmp, _meta(), simulated=True, allow_simulated_production_output=True
        ) as writer:
            writer.append_step(_step("NRM"))

        self.assertTrue((self.tmp / "OUT01").exists())
        self.assertIn("SIMULATED RUN", (self.tmp / SIMULATION_MARKER_FILE).read_text(encoding="utf-8"))

    def test_worker_labels_a_simulated_run_everywhere(self) -> None:
        worker = MeasurementWorker(
            meta=_meta("SIM01"),
            labels=["NRM"],
            output_dir=self.tmp,
            backend=NoCommBackend(),
            operator="opr",
        )
        warnings: list[str] = []
        worker.preflight_warning.connect(warnings.append)
        worker.run()

        self.assertTrue(any("SIMULATED RUN" in message for message in warnings))
        summary = json.loads((self.tmp / "workflow_summary.json").read_text(encoding="utf-8"))
        self.assertTrue(summary["simulated"])
        self.assertIn("SIMULATED RUN", summary["simulation_statement"])
        index = json.loads((self.tmp / "artifact_index.json").read_text(encoding="utf-8"))
        self.assertTrue(index["simulated"])
        self.assertTrue((self.tmp / SIMULATED_SUBDIR / "SIM01").exists())
        self.assertFalse((self.tmp / "SIM01").exists())


class _RejectingBackend:
    """Backend whose every block carries a one-flux-count step on X."""

    simulated = False
    flux_discontinuity_retries = 2

    def __init__(self, *, recover: bool) -> None:
        self._recover_enabled = recover
        self.recover_calls = 0
        self.reads = 0
        self.safe_state_calls = 0
        if recover:
            self.recover_flux_count_discontinuity = self._recover

    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def is_available(self) -> bool:
        return True

    def set_demag_step(self, label: str) -> None:
        del label

    def read_susceptibility(self) -> float:
        return 0.0

    def return_to_safe_state(self) -> None:
        self.safe_state_calls += 1

    def read_squid(self) -> BracketedMeasurementBlock:
        self.reads += 1
        return BracketedMeasurementBlock(
            zero_before=(0.0, 0.0, 0.0),
            positions=(
                (1.0, 0.0, 0.0),
                (1.0, 0.0, 0.0),
                (1.0, 0.0, 0.0),
                (1.0, 0.0, 0.0),
            ),
            zero_after=(0.09000563, 0.0, 0.0),
            audit=BlockAudit(block_id=f"block-{self.reads}", holder_record_id="holder@t0"),
        )

    def _recover(self, validation) -> None:
        self.recover_calls += 1


class RejectedBlockLeavesNoOutputTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run(self, backend) -> tuple[list[str], list[bool]]:
        worker = MeasurementWorker(
            meta=_meta("REJ01"),
            labels=["NRM", "AF20"],
            output_dir=self.tmp,
            backend=backend,
            operator="opr",
        )
        errors: list[str] = []
        finished: list[bool] = []
        worker.error_occurred.connect(errors.append)
        worker.run_finished.connect(finished.append)
        worker.run()
        return errors, finished

    def test_retry_exhaustion_writes_no_accepted_measurement(self) -> None:
        backend = _RejectingBackend(recover=True)

        errors, finished = self._run(backend)

        self.assertEqual(finished, [True])
        self.assertTrue(any("SQUID read error" in message for message in errors), errors)
        self.assertEqual(backend.recover_calls, 2)
        self.assertEqual(backend.reads, 3)
        self.assertFalse((self.tmp / "REJ01").exists())
        self.assertFalse((self.tmp / "REJ01.rmg").exists())
        self.assertFalse((self.tmp / "measurements.txt").exists())
        self.assertEqual(list(self.tmp.glob(".rapidpy-pending*")), [])
        self.assertEqual(backend.safe_state_calls, 1)

    def test_no_recovery_hook_fails_on_the_first_rejected_block(self) -> None:
        backend = _RejectingBackend(recover=False)

        errors, finished = self._run(backend)

        self.assertEqual(finished, [True])
        self.assertEqual(backend.reads, 1)
        self.assertEqual(backend.recover_calls, 0)
        self.assertTrue(any("bracketing zeros" in message for message in errors), errors)
        self.assertFalse((self.tmp / "REJ01").exists())

    def test_rejected_run_still_writes_evidence_artifacts(self) -> None:
        backend = _RejectingBackend(recover=False)

        self._run(backend)

        summary = json.loads((self.tmp / "workflow_summary.json").read_text(encoding="utf-8"))
        self.assertTrue(summary["aborted"])
        index = json.loads((self.tmp / "artifact_index.json").read_text(encoding="utf-8"))
        self.assertTrue(index["aborted"])
        self.assertEqual(index["published_paths"], {})


class SpecimenMetadataTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_reads_the_existing_specimen_header(self) -> None:
        with MeasurementBundleWriter(self.tmp, _meta("HDR01")) as writer:
            writer.append_step(_step("NRM"))

        resolution = resolve_specimen_meta("HDR01", sample_dir=self.tmp)

        self.assertEqual(resolution.source, "specimen-header")
        self.assertEqual(resolution.meta.name, "HDR01")
        self.assertEqual(resolution.meta.comment, "test specimen")
        self.assertAlmostEqual(resolution.meta.volume, 10.5)

    def test_falls_back_to_the_sample_index_registry(self) -> None:
        registrations = SampleIndexRegistrations(
            [
                SampleIndexRegistration(
                    specimen_name="REG01",
                    sample_set="SET-A",
                    location="Iceland",
                    formation="Basalt",
                    depth_cm="42",
                    order=1,
                )
            ]
        )

        resolution = resolve_specimen_meta("REG01", registrations=registrations)

        self.assertEqual(resolution.source, "sample-index")
        self.assertEqual(resolution.meta.site, "Basalt")
        self.assertEqual(resolution.meta.location, "Iceland")
        self.assertEqual(resolution.meta.sample, "REG01")
        self.assertIn("depth 42 cm", resolution.meta.comment)

    def test_reports_defaulted_fields_when_nothing_is_known(self) -> None:
        resolution = resolve_specimen_meta("BLANK01")

        self.assertEqual(resolution.source, "defaults")
        self.assertEqual(resolution.meta.sample, "BLANK01")
        self.assertIn("site", resolution.defaulted_fields)
        self.assertIn("location", resolution.defaulted_fields)
        self.assertIn("volume", resolution.defaulted_fields)
        self.assertFalse(resolution.complete)


if __name__ == "__main__":
    unittest.main(verbosity=2)
