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
from copy import deepcopy
from dataclasses import asdict
import json
import re
import threading
from concurrent.futures import Future
from datetime import datetime, timezone
from pathlib import Path
import time
import math
import uuid
from typing import Any, Mapping, Optional, Callable, Sequence, TypeVar

from PySide6 import QtCore

from rapid_main.analysis import ReadingCycleStatistics, reading_cycle_statistics
from rapid_main.acquisition_evidence import canonical
from rapid_main.block_statistics import BlockCollectionStatistics, block_collection_statistics, paired_block_statistics
from rapid_main.communication_log import CommunicationEvent, CommunicationLogger
from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.specimen_metadata import validate_specimen_provenance
from rapid_main.specimen_paths import contained_path
from rapid_main.measurement_progress import MeasurementProgress
from rapid_main.queue_pair import QueueSide
from rapid_main.io.rmg_writer import RmgTreatment, describe_rmg_treatment
from rapid_main.hardware_contracts import MeasurementBackend, NoCommBackend
from rapid_main.specimen_orientation import specimen_directions
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
from rapid_main.susceptibility_acquisition import write_susceptibility_acquisition
from rapid_main.af_treatment import write_af_treatment_record
from rapid_main.rockmag import write_rockmag_run_artifact
from rapid_main.thermal import write_thermal_run_artifact
from rapid_main.workflow import WorkflowPhase, WorkflowStateMachine
from rapid_main import software_version
from rapid_main.worker_ownership import worker_claim
from rapid_main.io.measurement_bundle import (
    SIMULATION_STATEMENT,
    SIMULATED_SUBDIR,
    MeasurementBundleWriter,
    validate_measurement_output,
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
        collection_stats: BlockCollectionStatistics | None = None,
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
        self.collection_stats = collection_stats


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
        thermal_context: Mapping[str, Any] | None = None,
        parent: Optional[QtCore.QObject] = None,
        *,
        specimen_provenance: Mapping[str, Any] | None = None,
        queue_side: QueueSide | None = None,
    ) -> None:
        super().__init__(parent)
        self._specimen_provenance = deepcopy(dict(specimen_provenance)) if specimen_provenance is not None else None
        if self._specimen_provenance is not None:
            validate_specimen_provenance(self._specimen_provenance, meta)
        self._meta = deepcopy(meta)
        self._labels = list(labels)
        self._output_dir = Path(output_dir).resolve()
        self._backend = backend or NoCommBackend()
        self._pending_run_result = None
        sources: list[object] = [self._backend]
        for source in tuple(communication_sources or ()):
            if source is not None and all(source is not item for item in sources):
                sources.append(source)
        self._communication_sources = tuple(sources)
        self._operator = operator
        self._samples_per_position = max(1, int(samples_per_position))
        self._queue_side = queue_side
        if queue_side is not None and (meta.name != queue_side.pair.specimen
                or labels != [queue_side.pair.label] or self._samples_per_position != queue_side.pair.count):
            raise ValueError('Measurement inputs differ from the original paired queue command.')
        self._run_id = queue_side.run_id if queue_side is not None else str(run_id)
        self._resume = bool(resume)
        self._allow_simulated_production_output = bool(allow_simulated_production_output)
        self._calibration_records = [dict(record) for record in (calibration_records or ())]
        self._routine_context = dict(routine_context) if routine_context is not None else None
        self._thermal_context = dict(thermal_context) if thermal_context is not None else None
        if self._routine_context is not None and self._thermal_context is not None:
            raise ValueError("a measurement run cannot have both rockmag and thermal plan identity")
        # A backend that declares itself simulated taints every artifact it
        # produces: the run is labelled and kept out of the production path.
        self._simulated = bool(getattr(self._backend, "simulated", False))
        validate_measurement_output(self._output_dir, self._meta.name, simulated=self._simulated,
            allow_simulated_production_output=self._allow_simulated_production_output)
        self._last_block_audit: BlockAudit | None = None
        self._last_block_result: BracketedMeasurementResult | None = None
        self._cycle_blocks: list[BracketedMeasurementBlock] = []
        self._cycle_readings: list[tuple[float, float, float]] = []
        self._progress: MeasurementProgress | None = None
        self._progress_run_id = self._run_id or uuid.uuid4().hex
        self._last_collection_stats: BlockCollectionStatistics | None = None
        self._collection_statistics: list[dict[str, object]] = []
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
            else contained_path(self._output_dir, SIMULATED_SUBDIR, output=True)
        )
        self._pair_directory = self._publish_dir
        self._up_reference = None
        if queue_side is not None:
            self._publish_dir = queue_side.evidence_directory(self._pair_directory)

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
        # Bridge acquisitions made before this run (e.g. the holder command)
        # belong to their own evidence; only current-run records are published.
        self._susceptibility_record_start_count = len(
            self._backend_susceptibility_records()
        )
        self._susceptibility_artifacts: list[tuple[str, Path]] = []
        self._af_record_start_count = len(self._backend_af_records())
        self._af_artifacts: list[tuple[str, Path]] = []
        self._pulse_record_start_count = len(self._backend_pulse_records())
        self._pulse_artifacts: list[tuple[str,Path]] = []
        self._phase_history: list[dict[str, object]] = []
        self._published_event_ids: set[int] = set()

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
        try:
            with worker_claim(self._backend):
                try:
                    try:
                        self._run_owned()
                    finally:
                        clear = getattr(self._backend, 'set_halt_check', None)
                        if callable(clear):
                            clear(None)
                except Exception as exc:
                    self._record_worker_failure(exc)
        except Exception as exc:
            self._record_worker_failure(exc)
        finally:
            self.run_finished.emit(self._pending_run_result is not False or bool(self._error_messages))

    def _record_worker_failure(self, exc):
        phase = WorkflowPhase.PREFLIGHT if self._pending_run_result is None else WorkflowPhase.RETURNING
        self._emit_error(f'Measurement worker failed: {exc}', phase=phase)
        self._emit_phase(WorkflowPhase.ERROR)
        self._pending_run_result = True
        try:
            self._finish_run(aborted=True)
        except Exception as artifact_error:
            self._emit_error(f'Failed to record worker failure: {artifact_error}', phase=WorkflowPhase.SAVING)

    def _run_owned(self) -> None:
        if self._halt_flag:
            self._emit_phase(WorkflowPhase.HALTED)
            self._finish_run(aborted=True)
            return

        # The old resume flag only suppressed duplicate file writes after
        # reapplying treatment. Until a trusted original-stage recovery route
        # exists it must never authorize replay of scientific hardware actions.
        if self._resume:
            raise ValueError('Automatic measurement replay is unavailable; restore original step progress before recovery.')
        source = dict(specimen_source=self._specimen_provenance, specimen_metadata=asdict(self._meta),
                      samples_per_position=self._samples_per_position, operator=self._operator,
                      calibration_records=self._calibration_records, routine=self._routine_context,
                      thermal=self._thermal_context)
        side = self._queue_side
        apply_treatment = side is None or side.is_up
        up_blocks = ()
        up_susceptibility = None
        if side is not None:
            original_up = side.check_admission(self._pair_directory, simulated=self._simulated)
            if original_up is not None:
                up_blocks, up_susceptibility, up_source, self._up_reference = original_up
                if up_source is not None:
                    for field in ('specimen_source', 'specimen_metadata', 'calibration_records'):
                        if canonical(source[field]) != canonical(up_source[field]):
                            raise ValueError('Specimen source or calibration changed after the Up pass.')
            source['queue_side'] = side.context
            source['up_acquisition'] = self._up_reference
            side.begin(self._pair_directory)
            if side.legacy_up is not None:
                side.archive_legacy_up(self._pair_directory)
        self._progress = MeasurementProgress(self._publish_dir, run_id=self._progress_run_id,
            specimen=self._meta.name, labels=self._labels, source=source, simulated=self._simulated)
        self._progress.begin()

        plan_validator = getattr(self._backend, "validate_treatment_plan", None)
        if apply_treatment and callable(plan_validator):
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

        self._apply_measurement_context(carried_label=None if apply_treatment else self._labels[0])
        self._emit_phase(WorkflowPhase.LOADING)

        # Prepare bundle writer (VB6 specimen + RMG + MagIC measurement/specimen)
        bundle = None
        try:
            self._emit_phase(WorkflowPhase.SAVING)
            if side is None or not side.is_up:
                bundle = MeasurementBundleWriter(
                    self._output_dir, self._meta, simulated=self._simulated,
                    allow_simulated_production_output=self._allow_simulated_production_output,
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
            try:
                self._progress.start_step(idx)
            except Exception as exc:
                self._emit_error(f'Cannot persist original step intent: {exc}', phase=WorkflowPhase.SAVING)
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True
                break

            # VB6 modMeasure.Measure runs the requested bridge acquisition
            # before PerformStep (treatment) and Measure_Read (SQUID).
            susc = 0.0
            susceptibility_evidence: dict[str, object] = {}
            if label.strip().upper() == "SUSC":
                try:
                    self._emit_phase(WorkflowPhase.MEASURING)
                    if up_susceptibility is not None:
                        susc = up_susceptibility['value']
                        susceptibility_evidence = dict(up_susceptibility['evidence'], original_up=self._up_reference)
                    else:
                        susc, susceptibility_evidence = self._read_susceptibility(label)
                except MeasurementHaltRequested:
                    aborted = True
                    self._emit_phase(WorkflowPhase.HALTED)
                    break
                except Exception as exc:
                    if self._halt_flag:
                        aborted = True
                        self._emit_phase(WorkflowPhase.HALTED)
                        break
                    self._emit_error(
                        f"Susceptibility read error at step {label}: {exc}",
                        phase=WorkflowPhase.MEASURING,
                    )
                    self._emit_phase(WorkflowPhase.ERROR)
                    aborted = True
                    break

            # Apply treatment step
            try:
                self._emit_phase(WorkflowPhase.TREATING)
                if apply_treatment:
                    af_start, pulse_start = len(self._backend_af_records()), len(self._backend_pulse_records())
                    self._comm_sent(label, detail="set_demag_step")
                    self._call_with_timeout(
                        lambda: self._backend.set_demag_step(label),
                        timeout=self._get_backend_timeout("step_timeout"),
                        phase="set_demag_step",
                    )
                    rmg_treatment = self._rmg_treatment(label, af_start=af_start, pulse_start=pulse_start)
                else:
                    self._comm_info(f'Paired Down retains the original Up treatment {label}')
                    carried = self._up_reference.get('rmg_treatment')
                    rmg_treatment = RmgTreatment(**carried) if carried is not None else describe_rmg_treatment(label)
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

            try:
                self._emit_phase(WorkflowPhase.VALIDATING)
                if side is not None:
                    side.validate_blocks(self._cycle_blocks, simulated=self._simulated)
                height_reader = getattr(self._backend, 'measured_specimen_height_cm', None)
                sample_height_cm = height_reader() if callable(height_reader) else None
                self._progress.capture_step(idx, blocks=self._cycle_blocks, readings=self._cycle_readings,
                    rmg_treatment=rmg_treatment,
                    sample_height_cm=sample_height_cm,
                    susceptibility=(dict(value=susc, evidence=susceptibility_evidence)
                                    if label.strip().upper() == 'SUSC' else None))
                if up_blocks:
                    self._last_collection_stats = paired_block_statistics(up_blocks, self._cycle_blocks,
                        legacy_up=side.legacy_up is not None)
                    sdx, sdy, sdz = self._last_collection_stats.moment_emu
                    cycle_stats = reading_cycle_statistics([
                        *(reduce_bracketed_measurement(block).moment_emu for block in up_blocks),
                        *self._cycle_readings])
                step = _build_step(label, sdx, sdy, sdz, self._operator, datetime.now(),
                    error_angle=(self._last_collection_stats.fischer_sd_deg
                                 if self._last_collection_stats is not None else 0.0),
                    meta=self._meta, collection=self._last_collection_stats,
                    range_factor=self._cycle_blocks[0].range_factor if self._cycle_blocks else None)
                self._validate_step(step)
                self._emit_phase(WorkflowPhase.SAVING)
                if bundle is not None and not bundle.append_step(step, susceptibility=susc,
                        rmg_treatment=rmg_treatment,
                        axis_sd_emu=(self._last_collection_stats.axis_sd_emu
                                     if self._last_collection_stats is not None else None),
                        sample_height_cm=sample_height_cm):
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
                            **susceptibility_evidence,
                        }
                    )
                if self._last_collection_stats is not None:
                    self._collection_statistics.append({"step_index": idx, "label": label,
                                                       **self._last_collection_stats.payload()})
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
                collection_stats=self._last_collection_stats,
            )
            self.step_complete.emit(result)
            self._completed_labels.append(label)

        # Return to safe state on both normal completion and interrupted execution
        self._emit_phase(WorkflowPhase.RETURNING)
        try:
            return_to_safe_state = getattr(self._backend, "return_to_safe_state", None)
            if callable(return_to_safe_state):
                self._comm_info("return_to_safe_state")
                self._call_with_timeout(
                    lambda: return_to_safe_state(),
                    timeout=self._get_backend_timeout("return_timeout"),
                    phase="return_to_safe_state",
                )
            elif side is not None:
                raise ValueError('Paired acquisition requires a verified safe return.')
        except Exception as exc:
            self._emit_error(f"Failed to return to safe state: {exc}", phase=WorkflowPhase.RETURNING)
            self._emit_phase(WorkflowPhase.ERROR)
            aborted = True

        if not aborted:
            try:
                if self._up_reference is not None:
                    _, _, _, current_up = side.load_up(self._pair_directory, simulated=self._simulated)
                    if current_up != self._up_reference:
                        raise ValueError('The original Up acquisition changed during the Down pass.')
                    if side.legacy_up is not None:
                        side.verify_legacy_archive(self._pair_directory)
                write_susceptibility_summary_json(
                    self._artifact_path("susceptibility.json"),
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

        # Retain raw evidence on failures, but publish the scientific bundle and
        # seal completed acquisition evidence only after successful safe return.
        if aborted:
            if bundle is not None:
                bundle.abort()
        else:
            try:
                if bundle is not None:
                    bundle.set_provenance(**self._run_provenance())
                    published = bundle.commit()
                    self._publish_dir = bundle.publish_dir
                    if self._up_reference is not None:
                        write_susceptibility_summary_json(
                            self._artifact_path('susceptibility.json'), susceptibility_records,
                            sample=self._meta.name, operator=self._operator)
                    self._published_paths = {
                        "specimen_file": str(published.specimen_file),
                        "rmg_file": str(published.rmg_file),
                        "magic_measurements": str(published.magic_measurements_file),
                        "magic_specimens": str(published.magic_specimens_file),
                    }
                self._progress.accept()
            except Exception as exc:
                if bundle is not None:
                    bundle.abort()
                self._emit_error(f"Failed to publish measurement bundle: {exc}", phase=WorkflowPhase.SAVING)
                self._emit_phase(WorkflowPhase.ERROR)
                aborted = True

        if not aborted:
            self._emit_phase(WorkflowPhase.COMPLETE)
        self._finish_run(aborted=aborted)
        if side is not None and side.is_up and self._pending_run_result is False:
            side.seal_up(self._pair_directory, self._progress)

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
        if not all(math.isfinite(value) for value in (step.sdx, step.sdy, step.sdz, step.moment,
                step.gdec, step.ginc, step.crdec, step.crinc, step.sdec, step.sinc,
                step.intensity_emu_per_cm3)):
            raise ValueError("measurement step contains non-finite values")
        if step.moment <= 0:
            raise ValueError("measurement step has non-positive moment")

    def _coerce_squid_reading(self, reading: object, label: str) -> tuple[float, float, float]:
        """Accept legacy tuple reads or calibrated magnetometer evidence."""
        quality_flags: tuple[str, ...] = ()
        if isinstance(reading, BracketedMeasurementBlock):
            result = reduce_bracketed_measurement(reading)
            self._record_block_evidence(reading, result)
            self._cycle_blocks.append(reading)
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
        self._cycle_readings = readings
        self._cycle_blocks = []
        self._last_collection_stats = None
        self._last_block_result = None
        for _ in range(self._samples_per_position):
            vector = self._read_validated_squid_sample(label)
            readings.append(vector)
            self._comm_received(vector, detail="read_squid")
        if self._cycle_blocks:
            if len(self._cycle_blocks) != len(readings):
                raise ObservationIntegrityError("SQUID cycle mixed raw blocks and unstructured readings")
            self._last_collection_stats = block_collection_statistics(self._cycle_blocks)
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
            read_squid = self._backend.read_squid
            if self._queue_side is not None and isinstance(self._backend, NoCommBackend):
                legacy_up = self._queue_side.legacy_up
                if legacy_up is None:
                    read_squid = lambda: self._backend.read_paired_squid(is_up=self._queue_side.is_up)
                else:
                    read_squid = lambda: self._backend.read_paired_squid(is_up=False, range_factor=legacy_up.range_factor)
            squid_reading = self._call_with_timeout(
                read_squid,
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

    def _read_susceptibility(self, label: str) -> tuple[float, dict[str, object]]:
        """Read one bridge value and return it with its acquisition identity."""

        before = len(self._backend_susceptibility_records())
        raw_susc = self._call_with_timeout(
            self._backend.read_susceptibility,
            timeout=self._get_backend_timeout("susceptibility_timeout"),
            phase="read_susceptibility",
        )
        susc = float(raw_susc)
        if not math.isfinite(susc):
            raise ValueError(f"non-finite susceptibility reading: {raw_susc!r}")
        evidence: dict[str, object] = {}
        records = self._backend_susceptibility_records()[before:]
        if records:
            record = records[-1]
            if bool(getattr(record, "simulated", False)) and not self._simulated:
                raise ObservationIntegrityError(
                    "A backend declared as live returned a simulated susceptibility "
                    "acquisition. The value was rejected before output could change."
                )
            evidence = {
                "acquisition_id": str(getattr(record, "acquisition_id", "")),
                "bridge_scaled_value": getattr(record, "bridge_scaled_value", None),
                "holder_scaled_value": getattr(record, "holder_scaled_value", None),
                "holder_evidence_id": str(getattr(record, "holder_evidence_id", "")),
                "moment_factor_cgs": getattr(record, "moment_factor_cgs", None),
                "safe_state_confirmed": bool(getattr(record, "safe_state_confirmed", False)),
            }
        self._comm_received(susc, detail="read_susceptibility")
        return susc, evidence

    def _backend_susceptibility_records(self) -> tuple[object, ...]:
        try:
            records = getattr(self._backend, "susceptibility_acquisition_records", ())
            if callable(records):
                records = records()
            return tuple(records or ())
        except Exception:
            return ()

    def _backend_af_records(self) -> tuple[object, ...]:
        provider = getattr(self._backend, "af_treatment_records", ())
        return tuple(provider() if callable(provider) else provider or ())

    def _backend_pulse_records(self):
        provider = getattr(self._backend,"pulse_treatment_records",())
        return tuple(provider() if callable(provider) else provider or ())

    def _rmg_treatment(self, label, *, af_start, pulse_start):
        def accepted(record):
            return (record.run_id == self._run_id and record.sample_id == self._meta.name
                    and record.simulated is self._simulated and record.safe_state_confirmed is True
                    and not record.error)

        af = next((record for record in reversed(self._backend_af_records()[af_start:])
                   if accepted(record) and record.plan.label == label), None)
        pulse = next((record for record in reversed(self._backend_pulse_records()[pulse_start:])
                      if accepted(record) and record.circuit is not None
                      and record.circuit.safe_state_confirmed is True), None)
        return describe_rmg_treatment(label, af=af, pulse=pulse)

    def _write_pulse_treatment_artifacts(self):
        written = {identity for identity,_path in self._pulse_artifacts}
        for record in self._backend_pulse_records()[self._pulse_record_start_count:]:
            try:
                identity = record.treatment_id
                if identity in written:
                    continue
                if not re.fullmatch(r"irm-[0-9a-f]{32}",identity):
                    raise ValueError("Invalid pulse IRM artifact identity.")
                path = self._artifact_path(f"pulse_treatments/{identity}.json")
                write_af_treatment_record(path,record)
            except Exception as exc:
                message = f"Pulse IRM artifact write failed: {exc}"
                self._error_messages.append(message)
                self.error_occurred.emit(message)
                continue
            self._pulse_artifacts.append((identity,path))
            written.add(identity)

    def _write_af_treatment_artifacts(self) -> None:
        written = {name for name, _path in self._af_artifacts}
        for record in self._backend_af_records()[self._af_record_start_count:]:
            identity = record.treatment_id
            if identity in written:
                continue
            try:
                if not re.fullmatch(r"af-[0-9a-f]{32}", identity):
                    raise ValueError("Invalid AF treatment artifact identity.")
                path = self._artifact_path(f"af_treatments/{identity}.json")
                write_af_treatment_record(path, record)
            except Exception as exc:
                self._error_messages.append(f"AF treatment artifact write failed: {exc}")
                self.error_occurred.emit(self._error_messages[-1])
                if self._comm_logger is not None:
                    self._comm_logger.error(self._error_messages[-1])
                continue
            self._af_artifacts.append((identity, path))
            written.add(identity)

    def _write_susceptibility_acquisition_artifacts(self) -> None:
        """Publish every current-run bridge acquisition, accepted or failed."""

        records = self._backend_susceptibility_records()[
            self._susceptibility_record_start_count :
        ]
        written = {name for name, _path in self._susceptibility_artifacts}
        for record in records:
            acquisition_id = str(getattr(record, "acquisition_id", "")).strip()
            if not acquisition_id or acquisition_id in written:
                continue
            target = self._artifact_path(f"susceptibility_acquisitions/{acquisition_id}.json")
            try:
                write_susceptibility_acquisition(target, record)
            except Exception as exc:
                self._error_messages.append(
                    f"susceptibility acquisition artifact write failed: {exc}"
                )
                if self._comm_logger is not None:
                    self._comm_logger.error(self._error_messages[-1])
                continue
            self._susceptibility_artifacts.append((acquisition_id, target))
            written.add(acquisition_id)

    def _apply_measurement_context(self, *, carried_label: str | None = None) -> None:
        """Give the backend the identity that belongs in each block audit."""
        halt_setter = getattr(self._backend, "set_halt_check", None)
        if callable(halt_setter):
            try:
                halt_setter(lambda: bool(self._halt_flag))
            except Exception as exc:
                self._comm_warning(f"set_halt_check failed: {exc}")
        setter = getattr(self._backend, "set_measurement_context", None)
        if setter is None or not callable(setter):
            if carried_label is not None and not self._simulated:
                raise ValueError('A native Down pass requires a read-only original treatment context.')
            return
        context = dict(sample_name=self._meta.name, run_id=self._run_id, operator=self._operator)
        if carried_label is not None:
            context['treatment_label'] = carried_label
        try:
            setter(**context)
        except Exception as exc:
            if carried_label is not None and not self._simulated:
                raise ValueError('Cannot bind the original treatment label for the Down acquisition.') from exc
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
            "thermal_routine": dict(self._thermal_context) if self._thermal_context else None,
            "specimen_source": deepcopy(self._specimen_provenance),
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
            "block_collection_statistics": deepcopy(self._collection_statistics),
            "coordinate_correction": {
                "method": "vb6_measure_unfold" if self._meta.has_core_orientation else "core_only",
                "magnetic_declination_e": self._meta.magnetic_declination_e,
                "volume_cm3_used": self._meta.volume if self._meta.volume > 0 else 1.,
                "cit_intensity_unit": "emu/cm^3",
                "magic_cartesian_frame": "geographic" if self._meta.has_core_orientation else "core",
            },
            "acquisition_progress": self._progress.reference if self._progress is not None else None,
            "up_acquisition": self._up_reference,
            "queue_side": self._queue_side.context if self._queue_side is not None else None,
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
            "thermal_routine": (
                {
                    "procedure_id": self._thermal_context.get("procedure_id", ""),
                    "routine_name": self._thermal_context.get("routine_name", ""),
                    "queue_labels": list(
                        self._thermal_context.get(
                            "queue_labels", self._thermal_context.get("labels", ())
                        )
                    ),
                    "automation_status": "MANUAL_EXTERNAL_ONLY",
                }
                if self._thermal_context is not None
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
            "susceptibility_acquisition_ids": [
                name for name, _path in self._susceptibility_artifacts
            ],
            "af_treatment_ids": [name for name, _path in self._af_artifacts],
            "pulse_treatment_ids": [name for name,_path in self._pulse_artifacts],
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
            (self._artifact_path("workflow_summary.json")).write_text(
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
                    # A device can be both a direct source and merged into the
                    # backend snapshot (e.g. the shared susceptibility bridge).
                    # The transcript keeps each event object exactly once.
                    fresh = [
                        event
                        for event in events[previous:]
                        if id(event) not in self._published_event_ids
                    ]
                    self._comm_logger.transcript.extend(fresh)
                    self._published_event_ids.update(id(event) for event in fresh)
                    self._communication_event_counts[id(source)] = len(events)
            self._comm_logger.write_text(self._artifact_path("communication.tsv"))
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
        self._write_susceptibility_acquisition_artifacts()
        self._write_af_treatment_artifacts()
        self._write_pulse_treatment_artifacts()
        aborted = aborted or bool(self._error_messages)
        self._write_workflow_summary(aborted=aborted)
        self._write_communication_transcript()
        self._write_rockmag_run_artifact(aborted=aborted)
        self._write_thermal_run_artifact(aborted=aborted)
        self._write_artifact_index(aborted=aborted)
        self._pending_run_result = aborted

    def _write_rockmag_run_artifact(self, *, aborted: bool) -> None:
        if self._routine_context is None:
            return
        try:
            write_rockmag_run_artifact(
                self._artifact_path("rockmag_run.json"),
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

    def _write_thermal_run_artifact(self, *, aborted: bool) -> None:
        if self._thermal_context is None:
            return
        try:
            write_thermal_run_artifact(
                self._artifact_path("thermal_run.json"),
                self._thermal_context,
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
            self._error_messages.append(f"thermal run artifact write failed: {exc}")
            if self._comm_logger is not None:
                self._comm_logger.error(self._error_messages[-1])

    def _write_artifact_index(self, *, aborted: bool) -> None:
        artifact_path = self._artifact_path("artifact_index.json")
        bundle_required = not aborted and (self._queue_side is None or not self._queue_side.is_up)

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
            "acquisition_progress": self._progress.reference if self._progress is not None else None,
            "up_acquisition": self._up_reference,
            "queue_side": self._queue_side.context if self._queue_side is not None else None,
            "specimen_source": deepcopy(self._specimen_provenance),
            "rockmag_routine": (
                str(self._routine_context.get("routine_name", ""))
                if self._routine_context is not None
                else ""
            ),
            "thermal_routine": (
                str(self._thermal_context.get("routine_name", ""))
                if self._thermal_context is not None
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
                    "measurement_provenance",
                    self._artifact_path('provenance.json'),
                    required=bundle_required,
                    producer='MeasurementBundleWriter',
                    description='Scientific metadata and original queue source provenance.',
                ),
                entry(
                    "vb6_specimen_file",
                    self._artifact_path(self._meta.name),
                    required=bundle_required,
                    producer="MeasurementBundleWriter",
                    description="Legacy specimen output compatible with VB6/CIT review paths.",
                ),
                entry(
                    "rmg_file",
                    self._artifact_path(f"{self._meta.name}.rmg"),
                    required=bundle_required,
                    producer="MeasurementBundleWriter",
                    description="RMG sidecar with per-step treatment and susceptibility values.",
                ),
                entry(
                    "magic_measurements",
                    self._artifact_path("measurements.txt"),
                    required=bundle_required,
                    producer="MeasurementBundleWriter",
                    description="MagIC measurements table emitted in lockstep with specimen output.",
                ),
                entry(
                    "magic_specimens",
                    self._artifact_path("specimens.txt"),
                    required=bundle_required,
                    producer="MeasurementBundleWriter",
                    description="MagIC specimen metadata table for the run bundle.",
                ),
                entry(
                    "susceptibility_summary",
                    self._artifact_path("susceptibility.json"),
                    required=not aborted,
                    producer="MeasurementWorker",
                    description="Per-step susceptibility readings and summary statistics.",
                ),
                entry(
                    "workflow_summary",
                    self._artifact_path("workflow_summary.json"),
                    required=True,
                    producer="MeasurementWorker",
                    description="Phase/status trace for completion, abort, or preflight failure evidence.",
                ),
                entry(
                    "communication_transcript",
                    self._artifact_path("communication.tsv"),
                    required=False,
                    producer="CommunicationLogger",
                    description="Transport-neutral transcript emitted when bundle execution begins.",
                ),
                entry(
                    "quicklook_summary",
                    self._artifact_path("quicklook.json"),
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
                    self._artifact_path("rockmag_run.json"),
                    required=True,
                    producer="MeasurementWorker",
                    description=(
                        "Compiled rockmag routine identity, requested/completed labels, "
                        "outcome, failures, and simulation/hardware acceptance boundary."
                    ),
                )
            )
        for treatment_id,path in self._pulse_artifacts:
            payload["artifacts"].append(entry(f"pulse_treatment:{treatment_id}",path,required=True,
                producer="PulseTreatmentService",description="Capacitor/field plan, relay readback, charge/fire/discharge evidence, specimen motion and safe return."))
        if self._progress is not None and self._progress.directory.is_dir():
            for path in sorted(self._progress.directory.iterdir()):
                path = self._progress._path(path.name)
                if path.is_file():
                    payload['artifacts'].append(entry('acquisition_progress:' + path.name, path,
                        required=True, producer='MeasurementProgress',
                        description='Original step intent, lossless raw acquisition or acceptance seal.'))
        if self._queue_side is not None and self._queue_side.legacy_up is not None and self._up_reference:
            payload['artifacts'].append(entry('legacy_up_source', self._queue_side.legacy_archive(self._pair_directory),
                required=not aborted, producer='VB6 Sample.WriteUpMeasurements',
                description='Byte-for-byte capture of the original Up import.'))
        for treatment_id, path in self._af_artifacts:
            payload["artifacts"].append(entry(
                f"af_treatment:{treatment_id}", path, required=True,
                producer="AfTreatmentService",
                description="Calibrated AF multi-pass plan, positions, per-phase results, faults, and safe return.",
            ))
        for acquisition_id, path in self._susceptibility_artifacts:
            payload["artifacts"].append(
                entry(
                    f"susceptibility_acquisition:{acquisition_id}",
                    path,
                    required=True,
                    producer="SusceptibilityAcquisitionService",
                    description=(
                        "Immutable bridge acquisition: raw/scaled value, holder identity, "
                        "factor, positions, phases, exact bridge traffic, and safe-state outcome."
                    ),
                )
            )
        if self._thermal_context is not None:
            payload["artifacts"].append(
                entry(
                    "thermal_run",
                    self._artifact_path("thermal_run.json"),
                    required=True,
                    producer="MeasurementWorker",
                    description=(
                        "Compiled thermal plan identity, requested/completed labels, blocker or "
                        "outcome, and the manual-external/hardware acceptance boundary."
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

    def _artifact_path(self, name: str) -> Path:
        return contained_path(self._publish_dir, name, output=True)

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
    error_angle: float = 0.0,
    *,
    meta: SpecimenMeta | None = None,
    collection: BlockCollectionStatistics | None = None,
    range_factor: float | None = None,
) -> MeasurementStep:
    """Convert raw SQUID Cartesian readings to a ``MeasurementStep``."""
    vector = (sdx, sdy, sdz)
    moment = math.hypot(*vector)
    # VB6's core-inclination epsilon is in calibrated SQUID units, before
    # RangeFact converts the reading to emu.
    direction_vector = tuple(component / range_factor for component in vector) if range_factor else vector
    directions = specimen_directions(direction_vector, meta or SpecimenMeta(''),
                                     label=label, raw_squid=range_factor is not None)
    volume = meta.volume if meta is not None and meta.volume > 0 else 1.

    return MeasurementStep(
        demag_label=label,
        gdec=directions.geographic.dec,
        ginc=directions.geographic.inc,
        sdec=directions.structural.dec,
        sinc=directions.structural.inc,
        moment=moment,
        error_angle=error_angle,
        crdec=directions.core.dec,
        crinc=directions.core.inc,
        sdx=sdx,
        sdy=sdy,
        sdz=sdz,
        operator=operator[:8] if operator else "",
        timestamp=timestamp,
        intensity_emu_per_cm3=moment / volume,
        axis_sd_raw=collection.axis_sd_raw if collection is not None else None,
        core_orientation_available=meta.has_core_orientation if meta is not None else None,
    )
