"""Tests for measured holder correction state and the holder command."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.holder_measurement import HolderMeasurementService
from rapid_main.holder_state import (
    DEFAULT_MAX_HOLDER_AGE_S,
    HolderCorrection,
    HolderStateError,
    HolderStateStore,
    ZERO_POSITION4,
)
from rapid_main.magnetometer import (
    BlockAudit,
    BracketedMeasurementBlock,
    FluxCountDiscontinuityError,
    reduce_bracketed_measurement,
)

NOW = datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)


def _clock(offset_s: float = 0.0):
    return lambda: NOW + timedelta(seconds=offset_s)


def _holder_block(
    *,
    zero_after=(0.0, 0.0, 0.0),
    positions=None,
    is_up: bool = True,
    simulated: bool = False,
    block_id: str = "block-holder-1",
) -> BracketedMeasurementBlock:
    return BracketedMeasurementBlock(
        zero_before=(0.0, 0.0, 0.0),
        positions=positions
        or (
            (0.20, 0.05, 0.30),
            (0.06, 0.19, 0.31),
            (-0.21, 0.04, 0.29),
            (-0.05, -0.20, 0.30),
        ),
        zero_after=zero_after,
        is_up=is_up,
        range_factor=1.0e-5,
        audit=BlockAudit(
            block_id=block_id,
            sample_name="Holder",
            run_id="run-1",
            operator="opr",
            software_version="rapidpy-test",
            config_hash="cfg-1",
            range_label="1",
            range_factor=1.0e-5,
            axis_calibration_applied=(0.09000563, 0.106, 0.066),
            completed_iso=NOW.isoformat(),
            is_holder_block=True,
            simulated=simulated,
        ),
    )


def _correction(**overrides) -> HolderCorrection:
    result = reduce_bracketed_measurement(_holder_block())
    correction = HolderCorrection.from_result(
        result,
        holder_id="HOLDER-A",
        hole=1,
        measured_at_iso=NOW.isoformat(),
    )
    if overrides:
        from dataclasses import replace

        correction = replace(correction, **overrides)
    return correction


class HolderCorrectionTests(unittest.TestCase):
    def test_builds_from_a_validated_block_with_full_evidence(self) -> None:
        correction = _correction()

        self.assertEqual(correction.holder_id, "HOLDER-A")
        self.assertEqual(correction.hole, 1)
        self.assertTrue(correction.is_up)
        self.assertEqual(correction.block_id, "block-holder-1")
        self.assertEqual(correction.run_id, "run-1")
        self.assertEqual(correction.operator, "opr")
        self.assertEqual(correction.software_version, "rapidpy-test")
        self.assertEqual(correction.range_label, "1")
        self.assertEqual(correction.axis_calibration, (0.09000563, 0.106, 0.066))
        self.assertEqual(len(correction.positions), 4)
        self.assertEqual(len(correction.raw_positions), 4)
        self.assertEqual(correction.zero_before, (0.0, 0.0, 0.0))
        self.assertGreater(correction.metrics.magnitude_raw, 0.0)
        self.assertAlmostEqual(
            correction.metrics.magnitude_emu,
            correction.metrics.magnitude_raw * 1.0e-5,
        )
        self.assertGreaterEqual(correction.metrics.asymmetry_ratio, 0.0)
        self.assertIn("HOLDER-A@", correction.record_version)

    def test_round_trips_through_json(self) -> None:
        correction = _correction(
            susceptibility_raw=0.0125,
            susceptibility_measured_at_iso=NOW.isoformat(),
            susceptibility_evidence_id="susc-holder-1",
        )

        restored = HolderCorrection.from_dict(json.loads(json.dumps(correction.to_dict())))

        self.assertEqual(restored.holder_id, correction.holder_id)
        self.assertEqual(restored.positions, correction.positions)
        self.assertEqual(restored.metrics.to_dict(), correction.metrics.to_dict())
        self.assertEqual(restored.record_version, correction.record_version)
        self.assertEqual(restored.require_susceptibility(), 0.0125)
        self.assertEqual(restored.susceptibility_evidence_id, "susc-holder-1")

    def test_legacy_holder_loads_but_cannot_authorize_susceptibility(self) -> None:
        correction = _correction()
        payload = correction.to_dict()
        payload.pop("susceptibility_raw")
        payload.pop("susceptibility_measured_at_iso")
        payload.pop("susceptibility_evidence_id")

        restored = HolderCorrection.from_dict(payload)

        with self.assertRaisesRegex(HolderStateError, "no finite susceptibility"):
            restored.require_susceptibility()


class HolderStateStoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "holder.json"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_install_publishes_atomically_and_reloads(self) -> None:
        store = HolderStateStore(self.path, clock=_clock())
        correction = _correction()

        store.install(correction)

        self.assertTrue(self.path.exists())
        self.assertEqual(list(self.path.parent.glob("*.tmp-*")), [])
        reloaded = HolderStateStore(self.path, clock=_clock())
        self.assertIsNotNone(reloaded.current)
        self.assertEqual(reloaded.current.holder_id, "HOLDER-A")
        self.assertEqual(reloaded.positions_for(is_up=True), correction.positions)

    def test_missing_holder_blocks_sample_measurement(self) -> None:
        store = HolderStateStore(self.path, clock=_clock())

        self.assertEqual(store.positions_for(), ZERO_POSITION4)
        with self.assertRaisesRegex(HolderStateError, "No holder correction"):
            store.require_valid()
        status = store.status()
        self.assertFalse(status.present)
        self.assertFalse(status.valid)

    def test_stale_holder_blocks_sample_measurement(self) -> None:
        store = HolderStateStore(
            self.path,
            clock=_clock(DEFAULT_MAX_HOLDER_AGE_S + 60),
            max_age_s=DEFAULT_MAX_HOLDER_AGE_S,
        )
        store.install(_correction())

        with self.assertRaisesRegex(HolderStateError, "stale"):
            store.require_valid()
        self.assertFalse(store.status().valid)

    def test_simulated_holder_is_refused_for_production(self) -> None:
        store = HolderStateStore(self.path, clock=_clock())
        store.install(_correction(simulated=True))

        with self.assertRaisesRegex(HolderStateError, "simulated"):
            store.require_valid()

        permissive = HolderStateStore(self.path, clock=_clock(), allow_simulated=True)
        self.assertTrue(permissive.require_valid().simulated)

    def test_direction_policy_shared_matches_vb6_and_strict_blocks(self) -> None:
        shared = HolderStateStore(self.path, clock=_clock())
        shared.install(_correction(is_up=True))
        # VB6 subtracts Holder.Sample(j) regardless of isUp.
        self.assertIsNotNone(shared.require_valid(is_up=False))

        strict_path = self.path.with_name("holder-strict.json")
        strict = HolderStateStore(strict_path, clock=_clock(), direction_policy="strict")
        strict.install(_correction(is_up=True))
        with self.assertRaisesRegex(HolderStateError, "measured with the sample up"):
            strict.require_valid(is_up=False)
        self.assertEqual(strict.positions_for(is_up=False), ZERO_POSITION4)

    def test_non_finite_correction_is_never_installed(self) -> None:
        store = HolderStateStore(self.path, clock=_clock())
        bad = _correction(
            positions=(
                (float("nan"), 0.0, 0.0),
                (0.0, 0.0, 0.0),
                (0.0, 0.0, 0.0),
                (0.0, 0.0, 0.0),
            )
        )

        with self.assertRaisesRegex(HolderStateError, "non-finite"):
            store.install(bad)
        self.assertIsNone(store.current)
        self.assertFalse(self.path.exists())

    def test_write_failure_keeps_the_previous_correction(self) -> None:
        store = HolderStateStore(self.path, clock=_clock())
        first = _correction()
        store.install(first)

        replacement = _correction(holder_id="HOLDER-B")
        store._path = self.path / "not-a-directory" / "holder.json"  # force a write error
        with self.assertRaises(OSError):
            store.install(replacement)

        self.assertEqual(store.current.holder_id, "HOLDER-A")
        self.assertEqual(
            HolderCorrection.from_dict(json.loads(self.path.read_text(encoding="utf-8"))).holder_id,
            "HOLDER-A",
        )


class _RecordingRecovery:
    def __init__(self, *, fail: bool = False) -> None:
        self.calls: list[object] = []
        self._fail = fail

    def __call__(self, validation: object) -> object:
        self.calls.append(validation)
        if self._fail:
            raise RuntimeError("2G counter reset refused")
        return validation


class HolderMeasurementServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "holder.json"
        self.store = HolderStateStore(self.path, clock=_clock())

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_installs_only_after_a_validated_block(self) -> None:
        service = HolderMeasurementService(
            lambda: _holder_block(), self.store, clock=_clock()
        )

        outcome = service.measure(holder_id="HOLDER-A", hole=3)

        self.assertTrue(outcome.installed)
        self.assertIsNone(outcome.previous)
        self.assertEqual(outcome.correction.hole, 3)
        self.assertEqual(self.store.current.holder_id, "HOLDER-A")
        self.assertTrue(self.path.exists())

    def test_rejected_block_keeps_the_previous_holder_and_file(self) -> None:
        seed = HolderMeasurementService(lambda: _holder_block(), self.store, clock=_clock())
        seed.measure(holder_id="HOLDER-A")
        before = self.path.read_text(encoding="utf-8")

        # One flux count on X between the bracketing zeros.
        rejected = HolderMeasurementService(
            lambda: _holder_block(zero_after=(0.09000563, 0.0, 0.0), block_id="block-bad"),
            self.store,
            clock=_clock(),
        )
        outcome = rejected.measure(holder_id="HOLDER-B")

        self.assertFalse(outcome.installed)
        self.assertTrue(outcome.retained_previous)
        self.assertIn("rejected", outcome.rejection_reason)
        self.assertEqual(self.store.current.holder_id, "HOLDER-A")
        self.assertEqual(self.path.read_text(encoding="utf-8"), before)

    def test_recovery_then_success_installs_the_replacement(self) -> None:
        blocks = [
            _holder_block(zero_after=(0.09000563, 0.0, 0.0), block_id="block-bad"),
            _holder_block(block_id="block-good"),
        ]
        recovery = _RecordingRecovery()
        service = HolderMeasurementService(
            lambda: blocks.pop(0),
            self.store,
            recover=recovery,
            flux_discontinuity_retries=2,
            clock=_clock(),
        )

        outcome = service.measure(holder_id="HOLDER-C")

        self.assertTrue(outcome.installed)
        self.assertEqual(outcome.recovery_attempts, 1)
        self.assertEqual(len(recovery.calls), 1)
        self.assertEqual(recovery.calls[0].discontinuous_axes, ("X",))
        self.assertEqual(self.store.current.holder_id, "HOLDER-C")

    def test_retry_exhaustion_never_installs(self) -> None:
        recovery = _RecordingRecovery()
        service = HolderMeasurementService(
            lambda: _holder_block(zero_after=(0.09000563, 0.0, 0.0)),
            self.store,
            recover=recovery,
            flux_discontinuity_retries=2,
            clock=_clock(),
        )

        outcome = service.measure(holder_id="HOLDER-D")

        self.assertFalse(outcome.installed)
        self.assertEqual(outcome.recovery_attempts, 2)
        self.assertIsNone(self.store.current)
        self.assertFalse(self.path.exists())

    def test_no_recovery_hook_fails_immediately(self) -> None:
        service = HolderMeasurementService(
            lambda: _holder_block(zero_after=(0.0, 0.106, 0.0)),
            self.store,
            recover=None,
            flux_discontinuity_retries=5,
            clock=_clock(),
        )

        outcome = service.measure(holder_id="HOLDER-E")

        self.assertFalse(outcome.installed)
        self.assertEqual(outcome.recovery_attempts, 0)
        self.assertIn("Y=", outcome.rejection_reason)

    def test_failed_recovery_is_reported_and_nothing_is_installed(self) -> None:
        recovery = _RecordingRecovery(fail=True)
        service = HolderMeasurementService(
            lambda: _holder_block(zero_after=(0.0, 0.0, 0.066)),
            self.store,
            recover=recovery,
            clock=_clock(),
        )

        outcome = service.measure(holder_id="HOLDER-F")

        self.assertFalse(outcome.installed)
        self.assertIn("recovery failed", outcome.rejection_reason)
        self.assertIsNone(self.store.current)

    def test_timeout_during_acquisition_keeps_previous_holder(self) -> None:
        seed = HolderMeasurementService(lambda: _holder_block(), self.store, clock=_clock())
        seed.measure(holder_id="HOLDER-A")

        def _timeout() -> BracketedMeasurementBlock:
            raise TimeoutError("read_squid exceeded timeout of 8.00s")

        service = HolderMeasurementService(_timeout, self.store, clock=_clock())
        outcome = service.measure(holder_id="HOLDER-G")

        self.assertFalse(outcome.installed)
        self.assertIn("timeout", outcome.rejection_reason.lower())
        self.assertEqual(self.store.current.holder_id, "HOLDER-A")

    def test_averages_baseline_adjusted_vectors_across_cycles(self) -> None:
        blocks = [
            _holder_block(
                positions=(
                    (0.10, 0.0, 0.0),
                    (0.10, 0.0, 0.0),
                    (0.10, 0.0, 0.0),
                    (0.10, 0.0, 0.0),
                )
            ),
            _holder_block(
                positions=(
                    (0.30, 0.0, 0.0),
                    (0.30, 0.0, 0.0),
                    (0.30, 0.0, 0.0),
                    (0.30, 0.0, 0.0),
                )
            ),
        ]
        service = HolderMeasurementService(
            lambda: blocks.pop(0),
            self.store,
            averaging_cycles=2,
            clock=_clock(),
        )

        outcome = service.measure(holder_id="HOLDER-H")

        self.assertTrue(outcome.installed)
        self.assertEqual(outcome.correction.averaging_cycles, 2)
        for vector in outcome.correction.positions:
            self.assertAlmostEqual(vector[0], 0.20)

    def test_holder_block_subtracts_no_holder_of_its_own(self) -> None:
        block = _holder_block()

        self.assertEqual(block.holder_positions, ZERO_POSITION4)
        result = reduce_bracketed_measurement(block)
        self.assertEqual(result.holder_average_raw, (0.0, 0.0, 0.0))


class HolderRejectionEvidenceTests(unittest.TestCase):
    def test_rejection_raises_before_any_reduced_value_exists(self) -> None:
        block = _holder_block(zero_after=(0.09000563, 0.0, 0.0))

        with self.assertRaises(FluxCountDiscontinuityError) as ctx:
            reduce_bracketed_measurement(block)

        self.assertEqual(ctx.exception.validation.discontinuous_axes, ("X",))
        self.assertEqual(ctx.exception.validation.flux_step_axes, ("X",))


if __name__ == "__main__":
    unittest.main(verbosity=2)
