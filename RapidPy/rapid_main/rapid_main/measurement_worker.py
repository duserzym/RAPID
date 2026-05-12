"""
measurement_worker.py — QThread-based measurement flow engine for RAPID v4.

Implements the measurement loop that was previously ``frmMeasure`` +
``modFlow.bas`` in VB6.  Runs on a worker thread so the UI stays responsive.

Responsibilities
----------------
* Iterate through a sequence of step labels
* For each step: command hardware (via injected backend interfaces), read SQUID,
  build a ``MeasurementStep``, write VB6 specimen file + ``.rmg`` sidecar +
  MagIC measurements.txt simultaneously (dual-write)
* Emit Qt signals for UI updates: step progress, live readings, completion

The worker is designed to be hardware-agnostic: actual I/O is delegated to
``HardwareBackend`` (abstract) so the engine runs without hardware in
``NO_COMM`` / simulation mode.

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

import threading
from datetime import datetime
from pathlib import Path
from typing import Optional, Protocol

from PySide6 import QtCore

from rapid_main.data_model import MeasurementStep, RmgRecord, SpecimenMeta
from rapid_main.io.magic_writer import append_measurement
from rapid_main.io.rmg_writer import append_rmg_record
from rapid_main.io.specimen_writer import append_step, write_header


# ─────────────────────────────────────────────────────────────────────────────
# Hardware backend protocol
# ─────────────────────────────────────────────────────────────────────────────

class HardwareBackend(Protocol):
    """
    Interface that the measurement engine uses to talk to hardware.

    Implementations:
      • ``NoCommBackend`` (this file) — returns synthetic data, no I/O
      • Future: ``SquidBackend``, ``AdwinBackend``, etc.
    """

    def read_squid(self) -> tuple[float, float, float]:
        """
        Return (x_emu, y_emu, z_emu) raw SQUID readings in emu.
        May block until the measurement settles.
        """
        ...

    def set_demag_step(self, label: str) -> None:
        """Apply the demagnetisation step corresponding to *label*."""
        ...

    def read_susceptibility(self) -> float:
        """Return current susceptibility reading in emu/Oe (0.0 if unavailable)."""
        ...

    def is_available(self) -> bool:
        """True if hardware communication is functioning."""
        ...


# ─────────────────────────────────────────────────────────────────────────────
# No-comm (simulation) backend
# ─────────────────────────────────────────────────────────────────────────────

class NoCommBackend:
    """Simulated backend — returns synthetic sinusoidal data, no hardware I/O."""

    def __init__(self) -> None:
        self._step_count = 0

    def read_squid(self) -> tuple[float, float, float]:
        import math
        t = self._step_count
        decay = math.exp(-t * 0.15)
        x = 1.23e-6 * decay * math.cos(math.radians(t * 12.5))
        y = 1.23e-6 * decay * math.sin(math.radians(t * 12.5))
        z = 0.5e-6  * decay
        return (x, y, z)

    def set_demag_step(self, label: str) -> None:
        self._step_count += 1

    def read_susceptibility(self) -> float:
        import math
        return 1e-3 * math.exp(-self._step_count * 0.1)

    def is_available(self) -> bool:
        return True


# ─────────────────────────────────────────────────────────────────────────────
# Measurement result
# ─────────────────────────────────────────────────────────────────────────────

class StepResult:
    """Holds the result of a single measurement step (passed via signals)."""

    def __init__(
        self,
        step: MeasurementStep,
        susceptibility: float,
        step_idx: int,
        total_steps: int,
    ) -> None:
        self.step = step
        self.susceptibility = susceptibility
        self.step_idx = step_idx
        self.total_steps = total_steps


# ─────────────────────────────────────────────────────────────────────────────
# Measurement worker (QThread)
# ─────────────────────────────────────────────────────────────────────────────

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
    step_started  = QtCore.Signal(int, str)          # idx, label
    step_complete = QtCore.Signal(object)             # StepResult
    run_finished  = QtCore.Signal(bool)               # aborted?
    error_occurred = QtCore.Signal(str)               # error message

    def __init__(
        self,
        meta: SpecimenMeta,
        labels: list[str],
        output_dir: Path,
        backend: HardwareBackend | None = None,
        operator: str = "",
        parent: Optional[QtCore.QObject] = None,
    ) -> None:
        super().__init__(parent)
        self._meta       = meta
        self._labels     = list(labels)
        self._output_dir = Path(output_dir)
        self._backend    = backend or NoCommBackend()
        self._operator   = operator

        # Control flags (thread-safe via threading.Event)
        self._pause_event  = threading.Event()
        self._pause_event.set()   # not paused initially
        self._halt_flag    = False

    # ── Control API (call from main thread) ───────────────────────────────────

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

    # ── Main loop ─────────────────────────────────────────────────────────────

    def run(self) -> None:
        """QThread entry point — executes the full sequence."""
        # Prepare output paths
        self._output_dir.mkdir(parents=True, exist_ok=True)
        specimen_file = self._output_dir / self._meta.name
        rmg_file      = self._output_dir / f"{self._meta.name}.rmg"
        magic_file    = self._output_dir / "measurements.txt"

        # Write specimen file header (creates / overwrites)
        try:
            write_header(specimen_file, self._meta)
        except OSError as exc:
            self.error_occurred.emit(f"Failed to write specimen header: {exc}")
            self.run_finished.emit(True)
            return

        total = len(self._labels)
        aborted = False

        for idx, label in enumerate(self._labels):
            # ── Check for halt ──
            if self._halt_flag:
                aborted = True
                break

            # ── Emit step start ──
            self.step_started.emit(idx, label)

            # ── Apply demagnetisation step ──
            try:
                self._backend.set_demag_step(label)
            except Exception as exc:
                self.error_occurred.emit(f"Hardware error at step {label}: {exc}")
                aborted = True
                break

            # ── Wait if paused ──
            self._pause_event.wait()
            if self._halt_flag:
                aborted = True
                break

            # ── Read SQUID ──
            try:
                sdx, sdy, sdz = self._backend.read_squid()
            except Exception as exc:
                self.error_occurred.emit(f"SQUID read error at step {label}: {exc}")
                aborted = True
                break

            # ── Read susceptibility ──
            susc = 0.0
            try:
                susc = self._backend.read_susceptibility()
            except Exception:
                pass  # susceptibility is optional

            # ── Build MeasurementStep ──
            step = _build_step(
                label=label,
                sdx=sdx, sdy=sdy, sdz=sdz,
                meta=self._meta,
                operator=self._operator,
                timestamp=datetime.now(),
            )

            # ── Dual write: specimen + .rmg + MagIC ──
            try:
                append_step(specimen_file, step)
                append_rmg_record(rmg_file, step, susceptibility=susc)
                append_measurement(magic_file, self._meta, step)
            except OSError as exc:
                self.error_occurred.emit(f"File write error at step {label}: {exc}")
                aborted = True
                break

            # ── Emit step complete ──
            result = StepResult(step, susc, idx, total)
            self.step_complete.emit(result)

        self.run_finished.emit(aborted)

    # ── Properties ────────────────────────────────────────────────────────────

    @property
    def is_paused(self) -> bool:
        return not self._pause_event.is_set()


# ─────────────────────────────────────────────────────────────────────────────
# Helper: build MeasurementStep from raw SQUID readings
# ─────────────────────────────────────────────────────────────────────────────

def _build_step(
    label: str,
    sdx: float,
    sdy: float,
    sdz: float,
    meta: SpecimenMeta,
    operator: str,
    timestamp: datetime,
) -> MeasurementStep:
    """
    Convert raw SQUID Cartesian readings to a ``MeasurementStep``.

    Orientation correction (specimen → geographic) is a placeholder — full
    tilt/strike correction will be added in Phase 3D hardware integration.
    """
    import math

    moment = math.sqrt(sdx**2 + sdy**2 + sdz**2)

    # Specimen coordinates → declination/inclination
    # (simplified: no orientation correction applied yet)
    horiz = math.sqrt(sdx**2 + sdy**2)
    sdec  = math.degrees(math.atan2(sdy, sdx)) % 360.0
    sinc  = math.degrees(math.atan2(sdz, horiz))

    # Geographic = specimen for now (Phase 3D will apply orientation matrix)
    gdec = sdec
    ginc = sinc

    # Core dec/inc = geographic (placeholder)
    crdec = gdec
    crinc = ginc

    error_angle = 0.0   # placeholder — SQUID RMS noise not yet computed

    return MeasurementStep(
        demag_label=label,
        gdec=gdec, ginc=ginc,
        sdec=sdec, sinc=sinc,
        moment=moment,
        error_angle=error_angle,
        crdec=crdec, crinc=crinc,
        sdx=sdx, sdy=sdy, sdz=sdz,
        operator=operator[:8] if operator else "",
        timestamp=timestamp,
    )
