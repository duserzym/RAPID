"""Holder measurement command: acquire, validate, then install atomically.

VB6 ``Measure_Read`` measures the holder with a *blank* holder correction
(``blankHolder``), averages ``avgSteps`` validated blocks, and only then
replaces the module-level ``Holder``.  A rejected block halts the run and
leaves the previous holder active.

This module reproduces that command as a testable service so queue-level
behavior — not just the reduction math — can be asserted.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime, timezone
import math
from typing import Callable

from rapid_main.holder_state import (
    HolderCorrection,
    HolderStateError,
    HolderStateStore,
)
from rapid_main.magnetometer import (
    BracketedMeasurementBlock,
    BracketedMeasurementResult,
    FluxCountDiscontinuityError,
    ObservationIntegrityError,
    reduce_bracketed_measurement,
)

AcquireBlock = Callable[[], BracketedMeasurementBlock]
RecoverHook = Callable[[object], object]


@dataclass(frozen=True)
class HolderMeasurementOutcome:
    """Result of one holder-measurement command."""

    installed: bool
    correction: HolderCorrection | None
    previous: HolderCorrection | None
    blocks: tuple[BracketedMeasurementBlock, ...] = ()
    results: tuple[BracketedMeasurementResult, ...] = ()
    recovery_attempts: int = 0
    rejection_reason: str = ""

    @property
    def retained_previous(self) -> bool:
        """True when the command finished without replacing the holder."""

        return not self.installed


class HolderMeasurementService:
    """Acquire, validate, and install a replacement holder correction."""

    def __init__(
        self,
        acquire: AcquireBlock,
        store: HolderStateStore,
        *,
        recover: RecoverHook | None = None,
        averaging_cycles: int = 1,
        flux_discontinuity_retries: int = 2,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self._acquire = acquire
        self._store = store
        self._recover = recover
        self._averaging_cycles = max(1, int(averaging_cycles))
        self._retries = max(0, int(flux_discontinuity_retries))
        self._clock = clock or (lambda: datetime.now(timezone.utc))

    @property
    def store(self) -> HolderStateStore:
        return self._store

    def measure(
        self,
        *,
        holder_id: str,
        hole: int = 0,
        susceptibility: object | None = None,
    ) -> HolderMeasurementOutcome:
        """Run the full holder command; never mutate state on any failure.

        ``susceptibility`` is an optional completed holder bridge acquisition
        (``SusceptibilityAcquisitionRecord``). VB6 measures it before the
        magnetic holder block; it is staged here and installed in the same
        atomic replacement, so a later magnetic failure discards it.
        """

        previous = self._store.current
        blocks: list[BracketedMeasurementBlock] = []
        results: list[BracketedMeasurementResult] = []
        recovery_attempts = 0

        for _cycle in range(self._averaging_cycles):
            attempt = 0
            while True:
                try:
                    block = self._acquire()
                    result = reduce_bracketed_measurement(block)
                except FluxCountDiscontinuityError as exc:
                    if attempt >= self._retries or self._recover is None:
                        return HolderMeasurementOutcome(
                            installed=False,
                            correction=None,
                            previous=previous,
                            blocks=tuple(blocks),
                            results=tuple(results),
                            recovery_attempts=recovery_attempts,
                            rejection_reason=(
                                f"Holder block rejected and not replaced: {exc}"
                            ),
                        )
                    attempt += 1
                    recovery_attempts += 1
                    try:
                        self._recover(exc.validation)
                    except Exception as recovery_exc:  # recovery is hardware work
                        return HolderMeasurementOutcome(
                            installed=False,
                            correction=None,
                            previous=previous,
                            blocks=tuple(blocks),
                            results=tuple(results),
                            recovery_attempts=recovery_attempts,
                            rejection_reason=(
                                f"Holder recovery failed after a rejected block: {recovery_exc}"
                            ),
                        )
                    continue
                except (ObservationIntegrityError, ValueError, TimeoutError, OSError) as exc:
                    return HolderMeasurementOutcome(
                        installed=False,
                        correction=None,
                        previous=previous,
                        blocks=tuple(blocks),
                        results=tuple(results),
                        recovery_attempts=recovery_attempts,
                        rejection_reason=f"Holder acquisition failed: {exc}",
                    )
                blocks.append(block)
                results.append(result)
                break

        try:
            correction = HolderCorrection.from_collection(blocks, holder_id=holder_id,
                hole=hole, measured_at_iso=self._clock().isoformat())
        except (HolderStateError, ValueError, TypeError, KeyError, OverflowError) as exc:
            return HolderMeasurementOutcome(installed=False, correction=None, previous=previous,
                blocks=tuple(blocks), results=tuple(results), recovery_attempts=recovery_attempts,
                rejection_reason=f'Holder collection was not accepted: {exc}')
        if susceptibility is not None:
            try:
                correction = _with_susceptibility(correction, susceptibility)
            except HolderStateError as exc:
                return HolderMeasurementOutcome(
                    installed=False,
                    correction=None,
                    previous=previous,
                    blocks=tuple(blocks),
                    results=tuple(results),
                    recovery_attempts=recovery_attempts,
                    rejection_reason=f"Holder susceptibility was not accepted: {exc}",
                )
        try:
            self._store.install(correction)
        except (HolderStateError, OSError) as exc:
            return HolderMeasurementOutcome(
                installed=False,
                correction=None,
                previous=previous,
                blocks=tuple(blocks),
                results=tuple(results),
                recovery_attempts=recovery_attempts,
                rejection_reason=f"Holder correction was not installed: {exc}",
            )
        return HolderMeasurementOutcome(
            installed=True,
            correction=correction,
            previous=previous,
            blocks=tuple(blocks),
            results=tuple(results),
            recovery_attempts=recovery_attempts,
        )


def _with_susceptibility(correction: HolderCorrection, record: object) -> HolderCorrection:
    """Attach a completed holder bridge acquisition to the staged correction."""

    if not bool(getattr(record, "is_holder", False)):
        raise HolderStateError("the susceptibility acquisition was not a holder measurement")
    if str(getattr(record, "outcome", "")) != "completed":
        raise HolderStateError("the holder susceptibility acquisition did not complete")
    if not bool(getattr(record, "safe_state_confirmed", False)):
        raise HolderStateError("the lift safe state was not confirmed after the bridge read")
    value = getattr(record, "bridge_scaled_value", None)
    if value is None or not math.isfinite(float(value)):
        raise HolderStateError("the holder bridge value is missing or non-finite")
    evidence_id = str(getattr(record, "acquisition_id", "")).strip()
    measured_at = str(getattr(record, "completed_iso", "")).strip()
    if not evidence_id or not measured_at:
        raise HolderStateError("the holder bridge value has no evidence identity or timestamp")
    if bool(getattr(record, "simulated", False)) != bool(correction.simulated):
        raise HolderStateError(
            "the holder bridge and magnetic holder blocks disagree about simulation state"
        )
    return replace(
        correction,
        susceptibility_raw=float(value),
        susceptibility_measured_at_iso=measured_at,
        susceptibility_evidence_id=evidence_id,
    )
