"""Replay recorded 2G blocks end to end, with no hardware attached.

These are the software half of acceptance step D1: replay the archived
discontinuous block and prove the holder correction and the output directory
are unchanged afterwards.
"""
from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.holder_measurement import HolderMeasurementService
from rapid_main.holder_state import HolderStateStore
from rapid_main.magnetometer import (
    FluxCountDiscontinuityError,
    reduce_bracketed_measurement,
    validate_block_observations,
)
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.replay import (
    REPLAY_SCHEMA,
    ReplayFixtureError,
    ReplaySquidBackend,
    load_replay_fixture,
    replay_fixture_from_dict,
)

from tests.test_output_integrity import _meta

FIXTURE_DIR = Path(__file__).resolve().parent / "fixtures"
STABLE = FIXTURE_DIR / "holder_stable.json"
ARCHIVED = FIXTURE_DIR / "holder_archived_x_flux_step.json"


class ReplayFixtureLoadingTests(unittest.TestCase):
    def test_both_fixtures_are_present_and_declare_the_schema(self) -> None:
        for path in (STABLE, ARCHIVED):
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["schema"], REPLAY_SCHEMA)
            self.assertEqual(len(payload["observations"]), 6)
            self.assertTrue(payload["description"])
            self.assertTrue(payload["source"])

    def test_recorded_counter_and_dvm_reproduce_the_archived_values(self) -> None:
        fixture = load_replay_fixture(ARCHIVED)

        self.assertAlmostEqual(fixture.block.zero_before[0], -1.0436457, places=7)
        self.assertAlmostEqual(fixture.block.zero_after[0], -0.95364007, places=7)
        # The step is one counter increment, not a DVM excursion.
        before = fixture.block.observations[0].axes[0]
        after = fixture.block.observations[5].axes[0]
        self.assertEqual(before.counts, 0.0)
        self.assertEqual(after.counts, -1.0)
        self.assertAlmostEqual(before.dvm, after.dvm, places=12)

    def test_fixture_observations_are_coherent(self) -> None:
        for path in (STABLE, ARCHIVED):
            validate_block_observations(load_replay_fixture(path).block)

    def test_unknown_schema_is_refused(self) -> None:
        with self.assertRaisesRegex(ReplayFixtureError, "unsupported replay schema"):
            replay_fixture_from_dict({"schema": "something.else.v1"})

    def test_truncated_fixture_is_refused(self) -> None:
        payload = json.loads(STABLE.read_text(encoding="utf-8"))
        payload["observations"] = payload["observations"][:3]

        with self.assertRaisesRegex(ReplayFixtureError, "exactly 6 observations"):
            replay_fixture_from_dict(payload)

    def test_missing_axis_is_refused(self) -> None:
        payload = json.loads(STABLE.read_text(encoding="utf-8"))
        del payload["observations"][2]["axes"]["Y"]

        with self.assertRaisesRegex(ReplayFixtureError, "missing axis Y"):
            replay_fixture_from_dict(payload)


class ReplayReductionTests(unittest.TestCase):
    def test_stable_holder_fixture_reduces(self) -> None:
        fixture = load_replay_fixture(STABLE)

        result = reduce_bracketed_measurement(fixture.block)

        self.assertTrue(fixture.expectation.accepted)
        self.assertTrue(result.validation.valid)
        self.assertGreater(result.average_magnitude_raw, 0.0)

    def test_archived_fixture_is_rejected_as_expected(self) -> None:
        fixture = load_replay_fixture(ARCHIVED)

        with self.assertRaises(FluxCountDiscontinuityError) as ctx:
            reduce_bracketed_measurement(fixture.block)

        self.assertFalse(fixture.expectation.accepted)
        self.assertEqual(ctx.exception.validation.discontinuous_axes, fixture.expectation.discontinuous_axes)
        self.assertEqual(ctx.exception.validation.flux_step_axes, fixture.expectation.flux_step_axes)
        self.assertIn(fixture.expectation.reason_contains, str(ctx.exception))


class ReplayHolderCommandTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.store = HolderStateStore(self.tmp / "holder.json", allow_simulated=True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_stable_replay_installs_a_holder_correction(self) -> None:
        backend = ReplaySquidBackend([load_replay_fixture(STABLE)])
        service = HolderMeasurementService(backend.read_squid, self.store)

        outcome = service.measure(holder_id="REPLAY-STABLE")

        self.assertTrue(outcome.installed)
        self.assertEqual(self.store.current.holder_id, "REPLAY-STABLE")

    def test_archived_replay_never_touches_the_installed_holder(self) -> None:
        HolderMeasurementService(
            ReplaySquidBackend([load_replay_fixture(STABLE)]).read_squid, self.store
        ).measure(holder_id="REPLAY-STABLE")
        before = (self.tmp / "holder.json").read_text(encoding="utf-8")

        outcome = HolderMeasurementService(
            ReplaySquidBackend([load_replay_fixture(ARCHIVED)]).read_squid,
            self.store,
            recover=None,
        ).measure(holder_id="REPLAY-ARCHIVED")

        self.assertFalse(outcome.installed)
        self.assertIn("rejected", outcome.rejection_reason)
        self.assertEqual(self.store.current.holder_id, "REPLAY-STABLE")
        self.assertEqual((self.tmp / "holder.json").read_text(encoding="utf-8"), before)


class ReplayWorkerTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run(self, backend) -> tuple[list[str], list[bool]]:
        worker = MeasurementWorker(
            meta=_meta("REPLAY01"),
            labels=["NRM"],
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

    def test_archived_replay_writes_no_accepted_measurement(self) -> None:
        backend = ReplaySquidBackend([load_replay_fixture(ARCHIVED)])

        errors, finished = self._run(backend)

        self.assertEqual(finished, [True])
        self.assertTrue(any("bracketing zeros" in message for message in errors), errors)
        self.assertFalse((self.tmp / "REPLAY01").exists())
        self.assertFalse((self.tmp / "SIMULATED" / "REPLAY01").exists())
        self.assertEqual(list(self.tmp.glob("**/.rapidpy-pending*")), [])

    def test_stable_replay_publishes_into_the_simulated_directory(self) -> None:
        backend = ReplaySquidBackend([load_replay_fixture(STABLE)])

        _errors, finished = self._run(backend)

        self.assertEqual(finished, [False])
        # A replay is not hardware evidence, so it must stay out of production.
        self.assertTrue((self.tmp / "SIMULATED" / "REPLAY01").exists())
        self.assertFalse((self.tmp / "REPLAY01").exists())
        summary = json.loads(
            (self.tmp / "SIMULATED" / "workflow_summary.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertTrue(summary["simulated"])
        for name in (
            "workflow_summary.json",
            "communication.tsv",
            "susceptibility.json",
            "artifact_index.json",
        ):
            self.assertFalse((self.tmp / name).exists(), name)

    def test_replay_backend_has_no_recovery_hook(self) -> None:
        backend = ReplaySquidBackend([load_replay_fixture(ARCHIVED)])

        self.assertIsNone(getattr(backend, "recover_flux_count_discontinuity", None))
        self.assertEqual(backend.flux_discontinuity_retries, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
