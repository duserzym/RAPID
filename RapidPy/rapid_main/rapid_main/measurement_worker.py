"""
measurement_worker.py — QThread-based measurement flow engine for RAPID v4.

Implements the measurement loop that was previously ``frmMeasure`` +
``modFlow.bas`` in VB6. Runs on a worker thread so the UI stays responsive.

Responsibilities
----------------
* Iterate through a sequence of step labels
* For each step: command hardware (via injected backend interfaces), read SQUID,
  build a ``MeasurementStep``, write VB6 specimen file + ``.rmg`` sidecar +
  MagIC measurements.txt simultaneously (dual-write)
* Emit Qt signals for UI updates: step progress, live readings, completion

The worker delegates hardware I/O to a backend contract shared across the app
(``rapid_main.hardware_contracts``). In ``NO_COMM``/offline mode this defaults to
the simulator backend.

Usage::

    worker = MeasurementWorker(meta, labels, output_dir, backend=NoCommBackend())
    worker.step_started.connect(main_window.set_step)
    worker.step_complete.connect(on_step_done)
    worker.run_finished.connect(on_run_done)
    worker.start()

    # To pause/stop:
    worker.pause()
    worker.resume()
    worker.halt()
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeout
import hashlib
import json
import threading
from concurrent.futures import Future
from datetime import datetime, timezone
from pathlib import Path
import time
import math
from typing import Any, Mapping, Optional, Callable, Sequence, TypeVar

from PySide6 import QtCore

from rapid_main.analysis import ReadingCycleStatistics, reading_cycle_statistics
from rapid_main.communication_log import CommunicationEvent, CommunicationLogger
from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.hardware_contracts import MeasurementBackend, NoCommBackend
from rapid_main.geometry import Cartesian3D, cartesian3d_to_angular3d
from rapid_main.magnetometer import (
    BlockAudit,
    BracketedMeasurementBlock,
    BracketedMeasurementResult,
    FluxCountDiscontinuityError,
    MagnetometerReading,
    ObservationIntegrityError,
    reduce_bracketed_measurement,
)
from rapid_main.status_codes import OperatorStatus, error_status, status_for_phase, warning_status
from rapid_main.susceptibility import write_susceptibility_summary_json
from rapid_main.rockmag import write_rockmag_run_artifact
from rapid_main.workflow import WorkflowPhase, WorkflowStateMachine
from rapid_main import software_version
from rapid_main.io.measurement_bundle import (
    SIMULATION_STATEMENT,
    SIMULATED_SUBDIR,
    MeasurementBundleWriter,
)


class StepResult:
    """Holds the result of a single measurement step (passed via signals)."""

    def __init__(
        self,
        step: MeasurementStep,
        susceptibility: float,
        step_idx: int,
        total_steps: int,
        cycle_stats: ReadingCycleStatistics | None = None,
        block_result: "BracketedMeasurementResult | None" = None,
        holder_status: object | None = None,
    ) -> None:
        self.step = step
        self.susceptibility = susceptibility
        self.step_idx = step_idx
        self.total_steps = total_steps
        self.cycle_stats = cycle_stats
        # Present only when the backend returns a full bracketed block, which
        # is the only case where VB6-equivalent holder/induced ratios exist.
        self.block_result = block_result
        self.holder_status = holder_status


T = TypeVar("T")


class MeasurementTimeoutError(TimeoutError):
    """Raised when a hardware call exceeds the configured timeout."""


class MeasurementHaltRequested(RuntimeError):
    """Raised when operator requests an in-flight halt."""


class MeasurementWorker(QtCore.QThread):
    """
    Worker thread that executes the full measurement sequence.

    Signals
    -------
    step_started(step_idx: int, label: str)
        Emitted just before each step begins.
    step_complete(result: StepResult)
        Emitted after each step measurement is recorded.
    run_finished(aborted: bool)
        Emitted when the full sequence ends (or is halted).
    error_occurred(msg: str)
        Emitted if a hardware or I/O error occurs.
    """

    # Qt signals
    step_started = QtCore.Signal(int, str)  # idx, label
    step_complete = QtCore.Signal(object)  # StepResult
    run_finished = QtCore.Signal(bool)  # aborted?
    error_occurred = QtCore.Signal(str)  # error message
    preflight_warning = QtCore.Signal(str)  # preflight warning
    phase_changed = QtCore.Signal(str)  # workflow-state label
    status_event = QtCore.Signal(object)  # OperatorStatus

    def __init__(
        self,
        meta: SpecimenMeta,
        labels: list[str],
        output_dir: Path,
        backend: MeasurementBackend | None = None,
        operator: str = "",
        samples_per_position: int = 1,
        run_id: str = "",
        resume: bool = False,
        allow_simulated_production_output: bool = False,
        calibration_records: Sequence[Mapping[str, Any]] | None = None,
        communication_sources: Sequence[object] | None = None,
        routine_context: Mapping[str, Any] | None = None,
        parent: Optional[QtCore.QObject] = None,
    ) -> None:
        super().__init__(parent)
        self._meta = meta
        self._labels = list(labels)
        self._output_dir = Path(output_dir)
        self._backend = backend or NoCommBackend()
        sources: list[object] = [self._backend]
        for source in tuple(communication_sources or ()):
            if source is not None and all(source is not item for item in sources):
                sources.append(source)
        self._communication_sources = tuple(sources)
        self._operator = operator
        self._samples_per_position = max(1, int(samples_per_position))
        self._run_id = str(run_id)
        self._resume = bool(resume)
        self._allow_simulated_production_output = bool(allow_simulated_production_output)
        self._calibration_records = [dict(record) for record in (calibration_records or ())]
        self._routine_context = dict(routine_context) if routine_context is not None else None
        # A backend that declares itself simulated taints every artifact it
        # produces: the run is labelled and kept out of the production path.
        self._simulated = bool(getattr(self._backend, "simulated", False))
        self._last_block_audit: BlockAudit | None = None
        self._last_block_result: BracketedMeasurementResult | None = None
        self._holder_record_id = ""
        self._holder_recorded_iso = ""
        self._skipped_labels: list[str] = []
        self._completed_labels: list[str] = []
        self._error_messages: list[str] = []
        self._recovery_count = 0
        self._published_paths: dict[str, str] = {}
        self._publish_dir = (
            self._output_dir
            if not self._simulated or self._allow_simulated_production_output
            else self._output_dir / SIMULATED_SUBDIR
        )

        # Control flags (thread-safe via threading.Event)
        self._pause_event = threading.Event()
        self._pause_event.set()  # not paused initially
        self._halt_flag = False
        self._timeout_epsilon = 0.05
        self._state = WorkflowStateMachine()
        self._comm_logger: CommunicationLogger | None = None
        self._communication_event_counts = {
            id(source): len(self._communication_events(source))
            for source in self._communication_sources
        }
        self._transport_recovery_start_count = len(
            self._backend_transport_recovery_records()
        )
        self._phase_history: list[dict[str, object]] = []

    # Control API
    def pause(self) -> None:
        """Suspend execution after the current step completes."""
        self._pause_event.clear()

    def resume(self) -> None:
        """Resume a paused run."""
        self._pause_event.set()

    def halt(self) -> None:
        """Stop execution after the current step completes."""
        self._halt_flag = True
        self._pause_event.set()  # unblock if paused

    def run(self) -> None:
        """QThread entry point — executes the full sequence."""
        if self._halt_flag:
            self._emit_phase(WorkflowPhase.HALTED)
            self._finish_run(aborted=True)
            return

        plan_validator = getattr(self._backend, "validate_treatment_plan", None)
        if callable(plan_validator):
            try:
                self._emit_phase(WorkflowPhase.PREFLIGHT)
                plan_preflight = self._call_with_timeout(
                    lambda: plan_validator(tuple(self._labels)),
                    timeout=self._get_backend_timeout("preflight_timeout"),
                    phase="treatment plan preflight",
                )
            except MeasurementHaltRequested:
                self._emit_phase(WorkflowPhase.HALTED)
                self._finish_run(aborted=True)
                return
            except Exception as exc:
                self._emit_error(
                    f"Treatment plan preflight failed: {exc}",
                    phase=WorkflowPhase.PREFLIGHT,
                )
                self._emit_phase(WorkflowPhase.ERROR)
                self._finish_run(aborted=True)
                return
            if not plan_preflight.ok:
                reasons = (
                    "; ".join(plan_preflight.blockers)
                    if plan_preflight.blockers
                    else "treatment plan preflight failed"
                )
                self._emit_error(
                    f"Treatment plan preflight failed: {reasons}",
                    phase=WorkflowPhase.PREFLIGHT,
                )
                self._emit_phase(WorkflowPhase.ERROR)
                self._finish_run(aborted=True)
                return
            if plan_preflight.warnings:
                self._emit_warning("; ".join(plan_preflight.warnings))

        try:
            self._emit_phase(WorkflowPhase.PREFLIGHT)
            preflight = self._call_with_timeout(
                self._backend.preflight,
                timeout=self._get_backend_timeout("preflight_timeout"),
                phase="preflight",
            )
        except MeasurementHaltRequested:
            self._emit_phase(WorkflowPhase.HALTED)
            self._finish_run(aborted=True)
            return
        except Exception as exc:
            self._emit_error(f"Preflight failed: {exc}", phase=WorkflowPhase.PREFLIGHT)
            self._emit_phase(WorkflowPhase.ERROR)
            self._finish_run(aborted=True)
            return

        if not preflight.ok:
            reasons = "; ".join(preflight.blockers) if preflight.blockers else "preflight failed"
            self._emit_error(f"Preflight failed: {reasons}", phase=WorkflowPhase.PREFLIGHT)
            self._emit_phase(WorkflowPhase.ERROR)
            self._finish_run(aborted=True)
            return

        if preflight.warnings:
            self._emit_warning("; ".join(preflight.warnings))

        if not self._backend.is_available():
            self._emit_error("Hardware backend is not available.", phase=WorkflowPhase.PREFLIGHT)
            self._emit_phase(WorkflowPhase.ERROR)
            self._finish_run(aborted=True)
            return

        if self._halt_flag:
            self._emit_phase(WorkflowPhase.HALTED)
            self._finish_run(aborted=True)
            return

        if self._simulated:
            self._emit_warning(
                "SIMULATED RUN: this backend produces synthetic values. Output is "
                "written to the SIMULATED directory and is not hardware evidence."
            )

        self._apply_measurement_context()
        self._emit_phase(WorkflowPhase.LOADING)

        # Prepare bundle writer (VB6 specimen + RMG + MagIC measurement/specimen)
        try:
            self._emit_phase(WorkflowPhase.SAVING)
            bundle = MeasurementBundleWriter(
                self._output_dir,
                self._meta,
                simulated=self._simulated,
                allow_simulated_production_output=self._allow_simulated_production_output,
                resume=self._resume,
                provenance=self._base_provenance(),
            )
            self._comm_logger = CommunicationLogger("measurement-worker", port=self._meta.name)
            self._comm_logger.info(
                "bundle initialized (simulated)" if self._simulated else "bundle initialized"
            )
        except OSError as exc:
            self._emit_error(f"Failed to initialize output bundle: {exc}", phase=WorkflowPhase.SAVING)
            self._emit_phase(WorkflowPhase.ERROR)
            self._finish_run(aborted=True)
            return

        total = len(self._labels)
        aborted = False
        susceptibility_records: list[dict[str, object]] = []

        for idx, label in enumerate(self._labels):
            # Halt has top priority
            if self._halt_flag:
                aborted = True
                self._emit_phase(WorkflowPhase.HALTED)
                break

            self.step_started.emit(idx, label)

            # Apply treatment step
            try:
                self._emit_phase(WorkflowPhase.TREATING)
                self._comm_sent(label, detail="set_demag_step")
                self._call_with_timeout(
                    lambda: self._backend.set_demag_step(label),
                    timeout=self._get_backend_timeout("step_timeout"),
                    phase="set_demag_step",
                )
                self._emit_phase(WorkflowPhase.POSITIONING)
                self._validate_position()
            except MeasurementHaltRequested:
                aborted = True
                self._emit_phase(WorkflowPhase.HALTED)
                break
            except Exception as exc:
                self._emit_error(f"Hardware error at step {label}: {exc}", phase=WorkflowPhase.TREATING)
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True
                break

            # Pause handling
            self._pause_event.wait()
            if self._halt_flag:
                aborted = True
                self._emit_phase(WorkflowPhase.HALTED)
                break

            # Read the configured SQUID cycle, retaining its quality evidence.
            try:
                self._emit_phase(WorkflowPhase.MEASURING)
                cycle_stats = self._read_squid_cycle(label)
                sdx, sdy, sdz = cycle_stats.mean_vector
            except MeasurementHaltRequested:
                aborted = True
                self._emit_phase(WorkflowPhase.HALTED)
                break
            except Exception as exc:
                self._emit_error(f"SQUID read error at step {label}: {exc}", phase=WorkflowPhase.MEASURING)
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True
                break

            susc = 0.0
            if label.strip().upper() == "SUSC":
                try:
                    self._emit_phase(WorkflowPhase.MEASURING)
                    raw_susc = self._call_with_timeout(
                        self._backend.read_susceptibility,
                        timeout=self._get_backend_timeout("susceptibility_timeout"),
                        phase="read_susceptibility",
                    )
                    susc = float(raw_susc)
                    if not math.isfinite(susc):
                        raise ValueError(f"non-finite susceptibility reading: {raw_susc!r}")
                    self._comm_received(susc, detail="read_susceptibility")
                except MeasurementHaltRequested:
                    aborted = True
                    self._emit_phase(WorkflowPhase.HALTED)
                    break
                except Exception as exc:
                    self._emit_error(
                        f"Susceptibility read error at step {label}: {exc}",
                        phase=WorkflowPhase.MEASURING,
                    )
                    self._emit_phase(WorkflowPhase.ERROR)
                    aborted = True
                    break

            step = _build_step(
                label=label,
                sdx=sdx,
                sdy=sdy,
                sdz=sdz,
                operator=self._operator,
                timestamp=datetime.now(),
            )

            try:
                self._emit_phase(WorkflowPhase.VALIDATING)
                self._validate_step(step)
                self._emit_phase(WorkflowPhase.SAVING)
                if not bundle.append_step(step, susceptibility=susc):
                    self._skipped_labels.append(label)
                    self._emit_warning(
                        f"Step {label} is already present from an interrupted run; "
                        "it was not written twice."
                    )
                if label.strip().upper() == "SUSC":
                    susceptibility_records.append(
                        {
                            "step_index": idx,
                            "label": label,
                            "susceptibility": float(susc),
                        }
                    )
            except Exception as exc:
                self._emit_error(f"File write error at step {label}: {exc}", phase=WorkflowPhase.SAVING)
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True
                break

            result = StepResult(
                step=step,
                susceptibility=susc,
                step_idx=idx,
                total_steps=total,
                cycle_stats=cycle_stats,
                block_result=self._last_block_result,
                holder_status=self._backend_holder_status(),
            )
            self.step_complete.emit(result)
            self._completed_labels.append(label)

        # Publish only a complete, accepted run. An abort discards the staged
        # copy so the production output path is never partially updated.
        if aborted:
            bundle.abort()
        else:
            try:
                self._emit_phase(WorkflowPhase.SAVING)
                bundle.set_provenance(**self._run_provenance())
                published = bundle.commit()
                self._published_paths = {
                    "specimen_file": str(published.specimen_file),
                    "rmg_file": str(published.rmg_file),
                    "magic_measurements": str(published.magic_measurements_file),
                    "magic_specimens": str(published.magic_specimens_file),
                }
                self._publish_dir = bundle.publish_dir
            except Exception as exc:
                bundle.abort()
                self._emit_error(
                    f"Failed to publish measurement bundle: {exc}", phase=WorkflowPhase.SAVING
                )
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True

        # Return to safe state on both normal completion and interrupted execution
        self._emit_phase(WorkflowPhase.RETURNING)
        try:
            return_to_safe_state = getattr(self._backend, "return_to_safe_state")
            if callable(return_to_safe_state):
                self._comm_info("return_to_safe_state")
                self._call_with_timeout(
                    lambda: return_to_safe_state(),
                    timeout=self._get_backend_timeout("return_timeout"),
                    phase="return_to_safe_state",
                )
        except AttributeError:
            pass
        except Exception as exc:
            self._emit_error(f"Failed to return to safe state: {exc}", phase=WorkflowPhase.RETURNING)
            self._emit_phase(WorkflowPhase.ERROR)
            aborted = True

        if not aborted:
            try:
                write_susceptibility_summary_json(
                    self._publish_dir / "susceptibility.json",
                    susceptibility_records,
                    sample=self._meta.name,
                    operator=self._operator,
                )
            except Exception as exc:
                self._emit_error(
                    f"Failed to write susceptibility summary: {exc}",
                    phase=WorkflowPhase.SAVING,
                )
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True

        if not aborted:
            self._emit_phase(WorkflowPhase.COMPLETE)
        self._finish_run(aborted=aborted)

    @property
    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def _get_backend_timeout(self, key: str) -> float | None:
        """Read timeout values from backend without hard coupling.

        Backends can provide:
        - ``preflight_timeout``
        - ``step_timeout``
        - ``read_timeout``
        - ``susceptibility_timeout``
        - ``return_timeout``
        """
        value = getattr(self._backend, key, None)
        if value is None:
            return None
        try:
            parsed = float(value)
        except (TypeError, ValueError):
            return None
        return None if parsed <= 0 else parsed

    def _call_with_timeout(
        self,
        call: Callable[[], T],
        *,
        timeout: float | None,
        phase: str,
    ) -> T:
        if timeout is None:
            # No timeout configured: run directly to avoid overhead.
            return call()

        deadline = time.monotonic() + timeout
        with ThreadPoolExecutor(max_workers=1) as executor:
            future: Future[T] = executor.submit(call)
            while True:
                if self._halt_flag:
                    return self._raise_halted(phase)
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                try:
                    return future.result(timeout=min(remaining, self._timeout_epsilon))
                except FutureTimeout:
                    continue

        future.cancel()
        raise MeasurementTimeoutError(f"{phase} exceeded timeout of {timeout:.2f}s")

    def _raise_halted(self, phase: str) -> T:
        raise MeasurementHaltRequested(f"{phase} canceled by user halt")

    def _emit_phase(self, phase: WorkflowPhase) -> None:
        """Emit a phase transition that obeys the shared state machine."""
        try:
            self._state.advance(phase)
        except Exception:
            # State corruption at this level should not abort the entire run;
            # recover to the requested state so we keep a best-effort phase stream.
            self._state = WorkflowStateMachine(phase)
        status = status_for_phase(phase)
        self._phase_history.append(
            {
                "timestamp_iso": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
                "phase": phase.value,
                "status_code": getattr(status.code, "value", str(status.code)),
                "severity": getattr(status.severity, "value", str(status.severity)),
            }
        )
        self.phase_changed.emit(phase.value)
        self.status_event.emit(status)

    def _emit_warning(self, message: str) -> None:
        self._comm_warning(message)
        self.status_event.emit(warning_status(message))
        self.preflight_warning.emit(message)

    def _emit_error(self, message: str, *, phase: WorkflowPhase | None = None) -> None:
        self._error_messages.append(str(message))
        if self._comm_logger is not None:
            self._comm_logger.error(message)
        self.status_event.emit(error_status(message, phase=phase))
        self.error_occurred.emit(message)

    def _validate_position(self) -> None:
        """Optional real positioning validation through backend hook."""
        validator = getattr(self._backend, "validate_position", None)
        if validator is None or not callable(validator):
            return

        result = self._call_with_timeout(
            lambda: validator(),
            timeout=self._get_backend_timeout("position_timeout"),
            phase="validate_position",
        )

        if isinstance(result, bool):
            if not result:
                raise ValueError("backend position validation failed")
            return

        if isinstance(result, tuple) and len(result) == 2:
            ok, message = result
            if bool(ok):
                return
            detail = str(message).strip() if message else "backend reported position invalid"
            raise ValueError(f"position validation failed: {detail}")

        if result is not None and not isinstance(result, bool):
            detail = str(result).strip()
            raise ValueError(f"position validation failed: {detail}")

    def _validate_step(self, step: MeasurementStep) -> None:
        if not math.isfinite(step.sdx + step.sdy + step.sdz):
            raise ValueError("measurement step contains non-finite values")
        if step.moment <= 0:
            raise ValueError("measurement step has non-positive moment")

    def _coerce_squid_reading(self, reading: object, label: str) -> tuple[float, float, float]:
        """Accept legacy tuple reads or calibrated magnetometer evidence."""
        quality_flags: tuple[str, ...] = ()
        if isinstance(reading, BracketedMeasurementBlock):
            result = reduce_bracketed_measurement(reading)
            self._record_block_evidence(reading, result)
            sdx, sdy, sdz = result.moment_emu
        elif isinstance(reading, MagnetometerReading):
            sdx, sdy, sdz = reading.moment_emu
            quality_flags = reading.flags
        else:
            try:
                sdx, sdy, sdz = reading  # type: ignore[misc]
            except Exception as exc:
                raise ValueError(f"SQUID read returned unsupported payload: {reading!r}") from exc

        if quality_flags:
            self._emit_warning(
                f"SQUID read quality at step {label}: {', '.join(quality_flags)}"
            )
        return (float(sdx), float(sdy), float(sdz))

    def _read_squid_cycle(self, label: str) -> ReadingCycleStatistics:
        """Read and summarize the configured number of SQUID samples."""

        readings: list[tuple[float, float, float]] = []
        for _ in range(self._samples_per_position):
            vector = self._read_validated_squid_sample(label)
            readings.append(vector)
            self._comm_received(vector, detail="read_squid")
        return reading_cycle_statistics(readings)

    def _read_validated_squid_sample(self, label: str) -> tuple[float, float, float]:
        """Read one sample, recovering only when a backend can safely re-zero."""

        retry_value = getattr(self._backend, "flux_discontinuity_retries", 2)
        try:
            retry_limit = max(0, int(retry_value))
        except (TypeError, ValueError):
            retry_limit = 2
        for attempt in range(retry_limit + 1):
            transport_recoveries_before = len(self._transport_recovery_evidence())
            squid_reading = self._call_with_timeout(
                self._backend.read_squid,
                timeout=self._get_backend_timeout("read_timeout"),
                phase="read_squid",
            )
            transport_recoveries = self._transport_recovery_evidence()
            if len(transport_recoveries) > transport_recoveries_before:
                latest = transport_recoveries[-1]
                self._emit_warning(
                    f"SQUID transport recovered at step {label}; the incomplete block "
                    f"was discarded and reacquired from zero-before "
                    f"(attempt {latest['attempt']}): {latest['detail']}"
                )
            try:
                return self._coerce_squid_reading(squid_reading, label)
            except ObservationIntegrityError as exc:
                # Incoherent transport evidence is never recoverable by
                # re-zeroing: the block and the reply stream disagree.
                self._comm_warning(f"incoherent SQUID observation at {label}: {exc}")
                raise
            except FluxCountDiscontinuityError as exc:
                self._comm_warning(f"rejected SQUID block at {label}: {exc}")
                recover = getattr(self._backend, "recover_flux_count_discontinuity", None)
                if attempt >= retry_limit or not callable(recover):
                    raise
                self._emit_warning(
                    f"Rejected discontinuous SQUID block at step {label}; "
                    f"re-zeroing and retrying ({attempt + 1}/{retry_limit}): {exc}"
                )
                self._call_with_timeout(
                    lambda: recover(exc.validation),
                    timeout=self._get_backend_timeout("read_timeout"),
                    phase="recover_flux_count_discontinuity",
                )
                self._recovery_count += 1
        raise RuntimeError("unreachable SQUID retry state")

    def _apply_measurement_context(self) -> None:
        """Give the backend the identity that belongs in each block audit."""
        setter = getattr(self._backend, "set_measurement_context", None)
        if setter is None or not callable(setter):
            return
        try:
            setter(
                sample_name=self._meta.name,
                run_id=self._run_id,
                operator=self._operator,
            )
        except Exception as exc:
            self._comm_warning(f"set_measurement_context failed: {exc}")

    def _record_block_evidence(
        self,
        block: BracketedMeasurementBlock,
        result: BracketedMeasurementResult,
    ) -> None:
        """Keep the audit identifiers of the last accepted block."""
        audit = block.audit
        if audit is None:
            self._last_block_result = result
            return
        if audit.simulated and not self._simulated:
            raise ObservationIntegrityError(
                "A backend declared as live returned a simulated measurement block. "
                "The block was rejected before accepted state or output could change."
            )
        self._last_block_result = result
        self._last_block_audit = audit
        if audit.holder_record_id:
            self._holder_record_id = audit.holder_record_id
            self._holder_recorded_iso = audit.holder_recorded_iso

    def _backend_holder_status(self) -> object | None:
        """Holder identity/age/validity for the operator UI, when available."""
        reader = getattr(self._backend, "holder_status", None)
        if reader is None or not callable(reader):
            return None
        try:
            return reader()
        except Exception:
            return None

    def _base_provenance(self) -> dict[str, object]:
        return {
            "run_id": self._run_id,
            "operator": self._operator,
            "labels_requested": list(self._labels),
            "samples_per_position": self._samples_per_position,
            "resumed": self._resume,
            "simulated": self._simulated,
            "software_version": software_version(),
            "calibration_record_ids": [
                str(record.get("record_id", ""))
                for record in self._calibration_records
                if record.get("record_id")
            ],
            "calibration_records": list(self._calibration_records),
            "rockmag_routine": dict(self._routine_context) if self._routine_context else None,
        }

    def _run_provenance(self) -> dict[str, object]:
        audit = self._last_block_audit
        transport_recoveries = self._transport_recovery_evidence()
        payload: dict[str, object] = {
            "holder_record_id": self._holder_record_id,
            "holder_recorded_iso": self._holder_recorded_iso,
            "flux_recoveries": self._recovery_count,
            "transport_recoveries": len(transport_recoveries),
            "transport_recovery_records": transport_recoveries,
            "skipped_duplicate_labels": list(self._skipped_labels),
            "simulated": self._simulated,
        }
        if audit is not None:
            payload.update(
                {
                    "last_block_id": audit.block_id,
                    "config_hash": audit.config_hash,
                    "range_label": audit.range_label,
                    "range_factor": audit.range_factor,
                    "axis_calibration_applied": list(audit.axis_calibration_applied),
                    "zero_position": audit.zero_position,
                    "measurement_position": audit.measurement_position,
                    "software_version": audit.software_version or software_version(),
                }
            )
        return payload

    def _transport_recovery_evidence(self) -> list[dict[str, object]]:
        """Serialize backend whole-block recovery records for run evidence."""

        records = self._backend_transport_recovery_records()[
            self._transport_recovery_start_count :
        ]
        evidence: list[dict[str, object]] = []
        for record in records:
            commands = []
            for command in tuple(getattr(record, "commands", ()) or ()):
                commands.append(
                    {
                        "index": int(getattr(command, "index", len(commands))),
                        "kind": str(getattr(command, "kind", "")),
                        "detail": str(getattr(command, "detail", "")),
                        "started_iso": str(getattr(command, "started_iso", "")),
                        "completed_iso": str(getattr(command, "completed_iso", "")),
                        "ok": bool(getattr(command, "ok", False)),
                        "reply": str(getattr(command, "reply", "")),
                    }
                )
            evidence.append(
                {
                    "attempt": int(getattr(record, "attempt", len(evidence) + 1)),
                    "started_iso": str(getattr(record, "started_iso", "")),
                    "completed_iso": str(getattr(record, "completed_iso", "")),
                    "detail": str(getattr(record, "detail", "")),
                    "commands": commands,
                }
            )
        return evidence

    def _backend_transport_recovery_records(self) -> tuple[object, ...]:
        """Snapshot backend recovery records without making evidence optionality fatal."""

        try:
            records = getattr(self._backend, "transport_recovery_records", ())
            if callable(records):
                records = records()
            return tuple(records or ())
        except Exception:
            return ()

    def _comm_info(self, detail: str) -> None:
        if self._comm_logger is not None:
            self._comm_logger.info(detail)

    def _comm_warning(self, detail: str) -> None:
        if self._comm_logger is not None:
            self._comm_logger.info(f"warning: {detail}")

    def _comm_sent(self, payload: object, *, detail: str) -> None:
        if self._comm_logger is not None:
            self._comm_logger.sent(payload, detail=detail)

    def _comm_received(self, payload: object, *, detail: str) -> None:
        if self._comm_logger is not None:
            self._comm_logger.received(payload, detail=detail)

    def _write_workflow_summary(self, *, aborted: bool) -> None:
        transport_recoveries = self._transport_recovery_evidence()
        payload = {
            "schema": "rapidpy.measurement.workflow_summary.v1",
            "sample": self._meta.name,
            "operator": self._operator,
            "labels": list(self._labels),
            "completed_labels": list(self._completed_labels),
            "errors": list(self._error_messages),
            "rockmag_routine": (
                {
                    "procedure_id": self._routine_context.get("procedure_id", ""),
                    "routine_name": self._routine_context.get("routine_name", ""),
                    "queue_labels": list(self._routine_context.get("queue_labels", ())),
                }
                if self._routine_context is not None
                else None
            ),
            "aborted": bool(aborted),
            "simulated": self._simulated,
            "simulation_statement": (
                SIMULATION_STATEMENT.strip() if self._simulated else ""
            ),
            "run_id": self._run_id,
            "flux_recoveries": self._recovery_count,
            "transport_recoveries": len(transport_recoveries),
            "transport_recovery_records": transport_recoveries,
            "skipped_duplicate_labels": list(self._skipped_labels),
            "holder_record_id": self._holder_record_id,
            "calibration_record_ids": [
                str(record.get("record_id", ""))
                for record in self._calibration_records
                if record.get("record_id")
            ],
            "phase_count": len(self._phase_history),
            "final_phase": self._phase_history[-1]["phase"] if self._phase_history else "",
            "phases": list(self._phase_history),
            "hardware_validation_required": True,
            "hardware_validation_statement": (
                "Software workflow phase evidence only; live hardware transition, "
                "safe-state, and recovery behavior require physical acceptance testing."
            ),
        }
        try:
            self._publish_dir.mkdir(parents=True, exist_ok=True)
            (self._publish_dir / "workflow_summary.json").write_text(
                json.dumps(payload, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
        except Exception as exc:
            if self._comm_logger is not None:
                self._comm_logger.error(f"workflow summary write failed: {exc}")

    def _write_communication_transcript(self) -> None:
        if self._comm_logger is None:
            return
        try:
            for source in self._communication_sources:
                if bool(getattr(source, "simulated", False)):
                    continue
                events = self._communication_events(source)
                previous = self._communication_event_counts.get(id(source), 0)
                if len(events) >= previous:
                    self._comm_logger.transcript.extend(events[previous:])
                    self._communication_event_counts[id(source)] = len(events)
            self._comm_logger.write_text(self._publish_dir / "communication.tsv")
        except Exception as exc:
            self.error_occurred.emit(f"Failed to write communication transcript: {exc}")

    @staticmethod
    def _communication_events(source: object) -> tuple[CommunicationEvent, ...]:
        provider = getattr(source, "communication_events", None)
        if not callable(provider):
            return ()
        try:
            return tuple(
                event for event in provider() if isinstance(event, CommunicationEvent)
            )
        except Exception:
            return ()

    def _finish_run(self, *, aborted: bool) -> None:
        self._write_workflow_summary(aborted=aborted)
        self._write_communication_transcript()
        self._write_rockmag_run_artifact(aborted=aborted)
        self._write_artifact_index(aborted=aborted)
        self.run_finished.emit(aborted)

    def _write_rockmag_run_artifact(self, *, aborted: bool) -> None:
        if self._routine_context is None:
            return
        try:
            write_rockmag_run_artifact(
                self._publish_dir / "rockmag_run.json",
                self._routine_context,
                run_id=self._run_id,
                sample=self._meta.name,
                operator=self._operator,
                labels_requested=self._labels,
                completed_labels=self._completed_labels,
                skipped_duplicate_labels=self._skipped_labels,
                errors=self._error_messages,
                aborted=aborted,
                simulated=self._simulated,
                final_phase=(
                    str(self._phase_history[-1]["phase"]) if self._phase_history else ""
                ),
                software_version=software_version(),
                config_hash=(
                    str(self._last_block_audit.config_hash)
                    if self._last_block_audit is not None
                    else ""
                ),
            )
        except Exception as exc:
            self._error_messages.append(f"rockmag run artifact write failed: {exc}")
            if self._comm_logger is not None:
                self._comm_logger.error(self._error_messages[-1])

    def _write_artifact_index(self, *, aborted: bool) -> None:
        artifact_path = self._publish_dir / "artifact_index.json"

        def entry(
            name: str,
            path: Path,
            *,
            required: bool,
            producer: str,
            description: str,
        ) -> dict[str, object]:
            exists = path.exists()
            size = path.stat().st_size if exists else 0
            digest = ""
            if exists:
                sha256 = hashlib.sha256()
                with path.open("rb") as handle:
                    for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                        sha256.update(chunk)
                digest = sha256.hexdigest()
            return {
                "name": name,
                "relative_path": self._relative_artifact_path(path),
                "required": required,
                "exists": exists,
                "size_bytes": size,
                "sha256": digest,
                "producer": producer,
                "description": description,
            }

        payload = {
            "schema": "rapidpy.measurement.artifact_index.v1",
            "run_id": self._run_id,
            "sample": self._meta.name,
            "operator": self._operator,
            "aborted": bool(aborted),
            "simulated": self._simulated,
            "simulation_statement": (
                SIMULATION_STATEMENT.strip() if self._simulated else ""
            ),
            "published_paths": dict(self._published_paths),
            "rockmag_routine": (
                str(self._routine_context.get("routine_name", ""))
                if self._routine_context is not None
                else ""
            ),
            "calibration_record_ids": [
                str(record.get("record_id", ""))
                for record in self._calibration_records
                if record.get("record_id")
            ],
            "generated_at_iso": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
            "artifacts": [
                entry(
                    "vb6_specimen_file",
                    self._publish_dir / self._meta.name,
                    required=not aborted,
                    producer="MeasurementBundleWriter",
                    description="Legacy specimen output compatible with VB6/CIT review paths.",
                ),
                entry(
                    "rmg_file",
                    self._publish_dir / f"{self._meta.name}.rmg",
                    required=not aborted,
                    producer="MeasurementBundleWriter",
                    description="RMG sidecar with per-step treatment and susceptibility values.",
                ),
                entry(
                    "magic_measurements",
                    self._publish_dir / "measurements.txt",
                    required=not aborted,
                    producer="MeasurementBundleWriter",
                    description="MagIC measurements table emitted in lockstep with specimen output.",
                ),
                entry(
                    "magic_specimens",
                    self._publish_dir / "specimens.txt",
                    required=not aborted,
                    producer="MeasurementBundleWriter",
                    description="MagIC specimen metadata table for the run bundle.",
                ),
                entry(
                    "susceptibility_summary",
                    self._publish_dir / "susceptibility.json",
                    required=not aborted,
                    producer="MeasurementWorker",
                    description="Per-step susceptibility readings and summary statistics.",
                ),
                entry(
                    "workflow_summary",
                    self._publish_dir / "workflow_summary.json",
                    required=True,
                    producer="MeasurementWorker",
                    description="Phase/status trace for completion, abort, or preflight failure evidence.",
                ),
                entry(
                    "communication_transcript",
                    self._publish_dir / "communication.tsv",
                    required=False,
                    producer="CommunicationLogger",
                    description="Transport-neutral transcript emitted when bundle execution begins.",
                ),
                entry(
                    "quicklook_summary",
                    self._publish_dir / "quicklook.json",
                    required=False,
                    producer="MeasurementPanel",
                    description="Panel-level quicklook plot contract, written after completed UI runs.",
                ),
            ],
            "hardware_validation_required": True,
            "hardware_validation_statement": (
                "Artifact presence proves software bundle generation only; live hardware run "
                "association and physical safe-state behavior require bench acceptance testing."
            ),
        }
        if self._routine_context is not None:
            payload["artifacts"].append(
                entry(
                    "rockmag_run",
                    self._publish_dir / "rockmag_run.json",
                    required=True,
                    producer="MeasurementWorker",
                    description=(
                        "Compiled rockmag routine identity, requested/completed labels, "
                        "outcome, failures, and simulation/hardware acceptance boundary."
                    ),
                )
            )
        try:
            self._publish_dir.mkdir(parents=True, exist_ok=True)
            artifact_path.write_text(
                json.dumps(payload, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
        except Exception as exc:
            if self._comm_logger is not None:
                self._comm_logger.error(f"artifact index write failed: {exc}")

    def _relative_artifact_path(self, path: Path) -> str:
        try:
            return path.relative_to(self._publish_dir).as_posix()
        except ValueError:
            return path.name


def _build_step(
    label: str,
    sdx: float,
    sdy: float,
    sdz: float,
    operator: str,
    timestamp: datetime,
) -> MeasurementStep:
    """Convert raw SQUID Cartesian readings to a ``MeasurementStep``."""
    moment = (sdx**2 + sdy**2 + sdz**2) ** 0.5

    direction = cartesian3d_to_angular3d(
        vector=Cartesian3D(sdx, sdy, sdz)
    )
    sdec = direction.dec
    sinc = direction.inc
    gdec = sdec
    ginc = sinc

    return MeasurementStep(
        demag_label=label,
        gdec=gdec,
        ginc=ginc,
        sdec=sdec,
        sinc=sinc,
        moment=moment,
        error_angle=0.0,
        crdec=gdec,
        crinc=ginc,
        sdx=sdx,
        sdy=sdy,
        sdz=sdz,
        operator=operator[:8] if operator else "",
        timestamp=timestamp,
    )
