from __future__ import annotations

import unittest
from unittest.mock import Mock

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
        self.assertEqual(started["samples"][0].sample_name, "AF_DEMO")  # type: ignore[index]
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
