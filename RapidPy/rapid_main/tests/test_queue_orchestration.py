from __future__ import annotations

import unittest
import json

from PySide6 import QtCore, QtWidgets
from unittest.mock import patch

from rapid_main.app import (
    MainWindow,
    _QSETTINGS_QUEUE_ACTIVE,
    _QSETTINGS_QUEUE_PROGRESS,
    _QSETTINGS_QUEUE_ROWS,
    _QSETTINGS_QUEUE_RESUME_POS,
)
from rapid_main.hardware_contracts import PreflightResult
from rapid_main.queue_compiler import QueueCommand, QueueOptions, QueueSample


class _QueueOnlyBackend:
    def __init__(self, *, with_queue: bool = True) -> None:
        self.calls: list[tuple[str, object]] = []
        self._with_queue = with_queue

    def read_squid(self) -> tuple[float, float, float]:
        return (0.0, 0.0, 0.0)

    def set_demag_step(self, label: str) -> None:
        self.calls.append(("set_demag_step", label))

    def read_susceptibility(self) -> float:
        return 0.0

    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def is_available(self) -> bool:
        return True

    def return_to_safe_state(self) -> None:
        self.calls.append(("return_to_safe_state", ""))

    def init_up(self, file_id: str) -> None:
        if self._with_queue:
            self.calls.append(("init_up", file_id))

    def holder(self, hole: int) -> None:
        if self._with_queue:
            self.calls.append(("holder", hole))

    def goto_hole(self, hole: int) -> None:
        if self._with_queue:
            self.calls.append(("goto_hole", hole))

    def flip(self) -> None:
        if self._with_queue:
            self.calls.append(("flip", ""))


class _QueueBackendWithoutCommands:
    def __init__(self) -> None:
        self.calls: list[tuple[str, object]] = []

    def read_squid(self) -> tuple[float, float, float]:
        return (0.0, 0.0, 0.0)

    def set_demag_step(self, label: str) -> None:
        self.calls.append(("set_demag_step", label))

    def read_susceptibility(self) -> float:
        return 0.0

    def preflight(self) -> PreflightResult:
        return PreflightResult.pass_ok()

    def is_available(self) -> bool:
        return True

    def return_to_safe_state(self) -> None:
        self.calls.append(("return_to_safe_state", ""))


class _MeasurementStub:
    def __init__(self) -> None:
        self.started_samples: list[tuple[str, bool]] = []
        self._run_finished: list[tuple[bool, str]] = []
        self._last_error = False

    def start_measurement_for_sample(self, sample: str, *, queue_run: bool = False) -> bool:
        self.started_samples.append((sample, queue_run))
        return True

    def is_active(self) -> bool:
        return False

    def halt_run(self) -> None:
        return

    def take_last_run_error(self) -> bool:
        had_error = self._last_error
        self._last_error = False
        return had_error


class _SampleQueueStub:
    def __init__(self) -> None:
        self.started: list[str] = []
        self.done: list[str] = []
        self.failed: list[str] = []
        self._interrupts: list[str] = []

    def row_snapshot(self) -> list[dict[str, str]]:
        return []

    def start_queue_sample(self, sample: str) -> None:
        self.started.append(sample)

    def mark_queue_sample_done(self, sample: str) -> None:
        self.done.append(sample)

    def set_queue_sample_failed(self, sample: str) -> None:
        self.failed.append(sample)

    def recover_interrupted_samples(self) -> int:
        count = len(self._interrupts)
        self._interrupts.clear()
        return count


class _HighPressureVacuumBackend:
    def is_connected(self) -> bool:
        return True

    def read_pressure(self) -> float:
        return 250.0

    def set_pump(self, on: bool) -> None:
        del on

    def is_pump_on(self) -> bool:
        return True

    def status(self) -> str:
        return "vacuum pressure high"


class _DisconnectedSquidBackend:
    simulated = False

    def is_connected(self) -> bool:
        return False

    def test_connection(self) -> bool:
        return False

    def status(self) -> str:
        return "SQUID transport not connected"

    def read_squid(self) -> tuple[float, float, float]:
        return (0.0, 0.0, 0.0)

    def read_susceptibility(self) -> float:
        return 0.0


class _ConnectedSquidBackend:
    simulated = False

    def is_connected(self) -> bool:
        return True

    def test_connection(self) -> bool:
        return True

    def status(self) -> str:
        return "SQUID connected"

    def read_squid(self) -> tuple[float, float, float]:
        return (0.0, 0.0, 0.0)

    def read_susceptibility(self) -> float:
        return 0.0


class QueueOrchestrationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if QtWidgets.QApplication.instance() is None:
            cls._qt_app = QtWidgets.QApplication([])
        else:
            cls._qt_app = None

    def setUp(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        settings.clear()

    def _run_single_sample_queue(self, backend: object, *, nocomm: bool) -> tuple[MainWindow, _MeasurementStub, _QueueOnlyBackend]:
        mw = MainWindow()
        mw.config.general.nocomm = nocomm
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM", "AF50", "SUSC"]
        samples = [
            QueueSample(
                sample_name="HBK-01",
                file_id="FILE-A",
                hole=7,
                do_up=True,
                do_both=False,
                measurement_step_count=1,
            )
        ]
        options = QueueOptions(ascending=True, load_return=True, do_return=True, repeat_holder=True)
        self.assertTrue(mw.start_queue_run(samples, options))
        return mw, measurement, backend  # type: ignore[return-value]

    def test_queue_pause_stops_advancing_until_resume(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM", "AF50", "SUSC"]
        samples = [
            QueueSample(
                sample_name="HBK-01",
                file_id="FILE-A",
                hole=7,
                do_up=True,
                do_both=False,
                measurement_step_count=1,
            )
        ]
        options = QueueOptions(ascending=True, load_return=True, do_return=False, repeat_holder=True)
        self.assertTrue(mw.start_queue_run(samples, options))

        self.assertIn(("holder", 0), backend.calls)
        self.assertEqual(measurement.started_samples, [("HBK-01", True)])
        self.assertTrue(mw._queue_paused is False)

        mw.toggle_queue_pause()
        self.assertTrue(mw._queue_paused)
        self.assertEqual(mw._queue_pos, 2)

        self.assertEqual(len(backend.calls), 1)
        mw._on_queue_sample_finished(False, "HBK-01")
        # Sample finish while paused should not advance queue commands.
        self.assertEqual(len(backend.calls), 1)
        self.assertTrue(mw._queue_active)
        self.assertEqual(backend.calls[-1], ("holder", 0))
        self.assertIn("HBK-01", sample_queue.done)

        mw.toggle_queue_pause()
        self.assertFalse(mw._queue_paused)
        # Resume should continue queue execution and mark completion.
        self.assertIn(("goto_hole", -1), backend.calls)
        self.assertFalse(mw._queue_active)
        self.assertEqual(mw._ownership.is_owned("changer"), False)

    def test_queue_lease_is_released_at_completion(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw, measurement, _ = self._run_single_sample_queue(backend, nocomm=False)
        self.assertIsNotNone(mw._queue_lease)
        self.assertTrue(mw._ownership.is_owned("changer"))
        mw._on_queue_sample_finished(False, "HBK-01")
        self.assertFalse(mw._queue_active)
        self.assertFalse(mw._ownership.is_owned("changer"))
        self.assertIsNone(mw._queue_lease)

    def test_queue_pause_state_cleared_after_cancel(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw, measurement, _ = self._run_single_sample_queue(backend, nocomm=False)
        self.assertFalse(mw._queue_paused)
        mw.toggle_queue_pause()
        self.assertTrue(mw._queue_paused)
        mw.cancel_queue_run("test")
        self.assertFalse(mw._queue_paused)
        self.assertFalse(mw._queue_active)
        self.assertIsNone(mw._queue_lease)
        self.assertEqual(measurement.started_samples, [("HBK-01", True)])

    def test_queue_runs_queue_commands_when_queue_backend_available(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw, measurement, _ = self._run_single_sample_queue(backend, nocomm=False)
        # Queue should have executed queue commands then started the measurement sample.
        self.assertEqual(measurement.started_samples, [("HBK-01", True)])
        self.assertIn(("holder", 0), backend.calls)
        self.assertEqual(backend.calls[0], ("holder", 0))
        # The measurement callback completion should advance queue.
        mw._on_queue_sample_finished(False, "HBK-01")
        self.assertIn(("goto_hole", -1), backend.calls)
        self.assertIn(("return_to_safe_state", ""), backend.calls)
        self.assertEqual(mw._queue_active, False)
        self.assertEqual(mw._ownership.is_owned("changer"), False)
        self.assertEqual(sample_queue := mw._sample_queue, mw._sample_queue)
        self.assertIn("HBK-01", sample_queue.done)
        self.assertEqual(len(sample_queue.failed), 0)
        self.assertEqual(backend.calls[0][0], "holder")
        self.assertIn(backend.calls[-1][0], {"goto_hole", "return_to_safe_state"})

    def test_queue_rejects_missing_queue_capability_in_hardware_mode(self) -> None:
        backend = _QueueBackendWithoutCommands()
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM"]
        with patch.object(QtWidgets.QMessageBox, "critical"):
            started = mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )
        # With hardware mode (no-comm disabled), missing required queue
        # movement methods should fail immediately.
        self.assertFalse(started)
        self.assertFalse(mw._queue_active)
        self.assertEqual(measurement.started_samples, [])
        self.assertEqual(sample_queue.failed, [])
        self.assertEqual(sample_queue.started, [])
        self.assertEqual(backend.calls, [])

    def test_queue_start_blocks_on_vacuum_pressure_fault(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw.config.vacuum.warn_threshold = 20.0
        mw._measurement_backend = backend
        mw._vacuum_backend = _HighPressureVacuumBackend()  # type: ignore[assignment]
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM"]

        with patch.object(QtWidgets.QMessageBox, "critical"):
            started = mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )

        self.assertFalse(started)
        self.assertFalse(mw._queue_active)
        self.assertFalse(mw._ownership.is_owned("changer"))
        self.assertEqual(measurement.started_samples, [])
        self.assertEqual(sample_queue.started, [])
        self.assertEqual(backend.calls, [])
        self.assertIn("Vacuum pressure", mw._sb_status.text())

    def test_queue_start_blocks_on_squid_comm_fault_in_hardware_mode(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _DisconnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM"]

        with patch.object(QtWidgets.QMessageBox, "critical"):
            started = mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )

        self.assertFalse(started)
        self.assertFalse(mw._queue_active)
        self.assertFalse(mw._ownership.is_owned("changer"))
        self.assertEqual(measurement.started_samples, [])
        self.assertEqual(sample_queue.started, [])
        self.assertEqual(backend.calls, [])
        self.assertIn("SQUID", mw._sb_status.text())

    def test_no_comm_mode_allows_missing_queue_methods(self) -> None:
        backend = _QueueOnlyBackend(with_queue=False)
        mw, measurement, _ = self._run_single_sample_queue(backend, nocomm=True)
        self.assertEqual(measurement.started_samples, [("HBK-01", True)])
        self.assertEqual(backend.calls, [])
        mw._on_queue_sample_finished(False, "HBK-01")
        self.assertEqual(mw._queue_active, False)

    def test_queue_sample_abort_marks_halted_and_releases_lease(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM", "AF50", "SUSC"]
        self.assertTrue(
            mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )
        )

        flow_states: list[str] = []
        mw.set_flow_state = lambda state: flow_states.append(state)
        self.assertIsNotNone(mw._queue_lease)
        self.assertTrue(mw._queue_active)

        mw._on_queue_sample_finished(True, "HBK-01")
        self.assertFalse(mw._queue_active)
        self.assertIsNone(mw._queue_lease)
        self.assertFalse(mw._ownership.is_owned("changer"))
        self.assertEqual(sample_queue.failed, ["HBK-01"])
        self.assertIn("halted", flow_states)

    def test_queue_sample_error_marks_error_and_releases_lease(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM", "AF50", "SUSC"]
        self.assertTrue(
            mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )
        )

        flow_states: list[str] = []
        mw.set_flow_state = lambda state: flow_states.append(state)
        measurement._last_error = True

        self.assertIsNotNone(mw._queue_lease)
        self.assertTrue(mw._queue_active)

        mw._on_queue_sample_finished(False, "HBK-01")
        self.assertFalse(mw._queue_active)
        self.assertIsNone(mw._queue_lease)
        self.assertFalse(mw._ownership.is_owned("changer"))
        self.assertEqual(sample_queue.failed, ["HBK-01"])
        self.assertIn("error", flow_states)

    def test_queue_resume_progress_is_persisted_and_visible_after_restore(self) -> None:
        backend = _QueueOnlyBackend(with_queue=True)
        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._measurement_backend = backend
        mw._squid_backend = _ConnectedSquidBackend()  # type: ignore[assignment]
        measurement = _MeasurementStub()
        sample_queue = _SampleQueueStub()
        mw._measurement = measurement  # type: ignore[assignment]
        mw._sample_queue = sample_queue  # type: ignore[assignment]
        mw._sequence_labels = ["NRM", "AF50", "SUSC"]
        self.assertTrue(
            mw.start_queue_run(
                [
                    QueueSample(
                        sample_name="HBK-01",
                        file_id="FILE-A",
                        hole=7,
                        do_up=True,
                        do_both=False,
                        measurement_step_count=1,
                    )
                ],
                QueueOptions(),
            )
        )

        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        self.assertTrue(bool(settings.value(_QSETTINGS_QUEUE_ACTIVE, False, type=bool)))
        progress_before = int(settings.value(_QSETTINGS_QUEUE_PROGRESS, 0, type=int))
        self.assertGreaterEqual(progress_before, 1)
        self.assertIsNotNone(mw._queue_lease)

        mw.cancel_queue_run("test")
        self.assertFalse(mw._ownership.is_owned("changer"))

        restored = MainWindow()
        self.assertEqual(
            restored._queue_pos,
            int(settings.value(_QSETTINGS_QUEUE_PROGRESS, 0, type=int)),
        )
        self.assertEqual(
            restored._queue_resume_pos,
            int(settings.value(_QSETTINGS_QUEUE_RESUME_POS, 0, type=int)),
        )
        self.assertEqual(restored._queue_pos, 0)
        self.assertEqual(restored._queue_resume_pos, 0)

    def test_restore_queue_state_recovers_running_rows_and_flags_resume(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        interrupted_rows = [
            {
                "position": "1",
                "sample_name": "HBK-INT-1",
                "sample_set": "set-A",
                "treatment": "NRM",
                "status": "Running",
            },
            {
                "position": "2",
                "sample_name": "HBK-INT-2",
                "sample_set": "set-B",
                "treatment": "AF50",
                "status": "Done",
            },
            {
                "position": "3",
                "sample_name": "HBK-INT-3",
                "sample_set": "set-C",
                "treatment": "SUSC",
                "status": "Running",
            },
        ]
        settings.setValue(_QSETTINGS_QUEUE_ROWS, json.dumps(interrupted_rows))
        settings.setValue(_QSETTINGS_QUEUE_PROGRESS, 3)
        settings.setValue(_QSETTINGS_QUEUE_RESUME_POS, 3)
        settings.setValue(_QSETTINGS_QUEUE_ACTIVE, False)

        captured: list[str] = []
        with patch.object(MainWindow, "set_status", lambda self, text: captured.append(text)):
            restored = MainWindow()

        rows = restored._sample_queue.row_snapshot()
        statuses = {row["sample_name"]: row["status"] for row in rows}

        self.assertEqual(statuses.get("HBK-INT-1"), "Interrupted")
        self.assertEqual(statuses.get("HBK-INT-2"), "Done")
        self.assertEqual(statuses.get("HBK-INT-3"), "Interrupted")
        self.assertEqual(restored._queue_pos, 3)
        self.assertEqual(restored._queue_resume_pos, 3)
        self.assertEqual(
            restored._sample_queue.interrupted_sample_names(),
            ["HBK-INT-1", "HBK-INT-3"],
        )
        self.assertTrue(any("choose Resume, Re-run, Skip, or Abort" in msg for msg in captured))


if __name__ == "__main__":
    unittest.main()
