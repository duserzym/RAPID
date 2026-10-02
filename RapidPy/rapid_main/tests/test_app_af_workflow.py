from __future__ import annotations

import unittest
from unittest.mock import Mock, patch

from PySide6 import QtWidgets


class _AfWorkflowMixin:
    @classmethod
    def setUpClass(cls) -> None:
        if QtWidgets.QApplication.instance() is None:
            cls._qt_app = QtWidgets.QApplication([])
        else:
            cls._qt_app = None

    @classmethod
    def tearDownClass(cls) -> None:
        if cls._qt_app is not None:
            cls._qt_app.quit()
            cls._qt_app = None



class TestAfWorkflowPlan(unittest.TestCase):
    def test_af_demo_labels_are_stable(self) -> None:
        try:
            from rapid_main.app import MainWindow
        except ModuleNotFoundError as exc:  # pragma: no cover - optional-dependency CI environments
            self.skipTest(f"Skipping AF workflow app import check: {exc}")

        self.assertEqual(
            MainWindow._af_demo_labels(),
            ["NRM", "AF25", "AF50", "AF100", "AF200", "AF400", "AF800", "SUSC"],
        )


class TestAfWorkflowQueue(_AfWorkflowMixin, unittest.TestCase):
    def test_af_demo_queue_builds_and_starts_queue_sample(self) -> None:
        try:
            from rapid_main.app import MainWindow
            from rapid_main.queue_compiler import QueueSample
        except ModuleNotFoundError as exc:  # pragma: no cover - optional-dependency CI environments
            self.skipTest(f"Skipping AF workflow app test: {exc}")

        started = {}
        mw = MainWindow()
        mw.start_queue_run = Mock(side_effect=lambda samples, options: (
            started.setdefault("samples", samples),
            started.setdefault("options", options),
            True
        )[2])
        started_measurements: list[tuple] = []

        class _PanelStub:
            def set_specimen_context(self, *args, **kwargs) -> None:
                setattr(self, "_last_context", (args, kwargs))

            def start_measurement_for_sample(self, sample: str, *, queue_run: bool = False) -> bool:
                started_measurements.append((sample, queue_run))
                return True

        mw._measurement = _PanelStub()  # type: ignore[assignment]
        mw.config.general.nocomm = True

        self.assertTrue(mw._prepare_af_workflow(auto_start=True, queue_mode=True))
        self.assertEqual(mw._sequence_labels, mw._af_demo_labels())
        self.assertEqual(  # type: ignore[index]
            started["samples"][0].sample_name,
            "SIMULATED_AF_EXAMPLE",
        )
        self.assertIsInstance(started["samples"][0], QueueSample)  # type: ignore[index]
        self.assertEqual(started["options"].samples_between_holder, 8)  # type: ignore[index]
        self.assertEqual(started_measurements, [])

    def test_af_demo_queue_start_failure_reports_failure(self) -> None:
        try:
            from rapid_main.app import MainWindow
        except ModuleNotFoundError as exc:  # pragma: no cover - optional-dependency CI environments
            self.skipTest(f"Skipping AF workflow app test: {exc}")

        mw = MainWindow()
        mw.config.general.nocomm = True
        mw.start_queue_run = Mock(return_value=False)
        class _PanelStub:
            def set_specimen_context(self, *args, **kwargs) -> None:
                return None
        mw._measurement = _PanelStub()  # type: ignore[assignment]

        self.assertFalse(mw._prepare_af_workflow(auto_start=True, queue_mode=True))

    def test_simulated_af_example_is_refused_in_hardware_mode(self) -> None:
        try:
            from rapid_main.app import MainWindow
        except ModuleNotFoundError as exc:  # pragma: no cover - optional-dependency CI environments
            self.skipTest(f"Skipping AF workflow app test: {exc}")

        mw = MainWindow()
        mw.config.general.nocomm = False
        mw.start_queue_run = Mock(return_value=True)
        original_labels = list(mw._sequence_labels)
        original_sample = mw._current_sample

        with patch.object(QtWidgets.QMessageBox, "warning") as warning:
            started = mw._prepare_af_workflow(auto_start=True, queue_mode=True)

        self.assertFalse(started)
        mw.start_queue_run.assert_not_called()
        self.assertEqual(mw._sequence_labels, original_labels)
        self.assertEqual(mw._current_sample, original_sample)
        warning.assert_called_once()

    def test_hardware_af_setup_uses_current_real_sample_without_autostart(self) -> None:
        try:
            from rapid_main.app import MainWindow
        except ModuleNotFoundError as exc:  # pragma: no cover - optional-dependency CI environments
            self.skipTest(f"Skipping AF workflow app test: {exc}")

        mw = MainWindow()
        mw.config.general.nocomm = False
        mw._current_sample = "SPEC-REAL"
        panel = Mock()
        mw._measurement = panel  # type: ignore[assignment]

        self.assertTrue(mw._prepare_af_workflow(auto_start=False))

        panel.set_specimen_context.assert_called_once_with(
            sample="SPEC-REAL",
            depth="—",
            treatment="AF workflow",
        )
        panel.start_measurement_for_sample.assert_not_called()
