from __future__ import annotations

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

VIEWER_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = VIEWER_ROOT.parent
sys.path[:0] = [str(VIEWER_ROOT), str(RAPIDPY_ROOT / "rapid_main"), str(RAPIDPY_ROOT)]
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6 import QtCore, QtTest, QtWidgets

from data_viewer.app import ZijderveldWindow
from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.magic_writer import write_measurements
from rapid_main.io.specimen_writer import append_step, write_header


class TestRecentTreatmentSteps(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_plain_magic_only_offers_its_recorded_frame(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            meta = SpecimenMeta("TEST")
            step = MeasurementStep("NRM", 80, 35, 100, -23, 7e-9, 2, 30, 12, 0, 0, 0)
            write_measurements(root / "measurements.txt", meta, [step])
            write_header(root / meta.name, meta)
            append_step(root / meta.name, step)
            window = ZijderveldWindow()
            try:
                with patch.object(QtWidgets.QFileDialog, "getExistingDirectory", return_value=directory):
                    window._open_directory()
                self.assertEqual(window._coord_combo.count(), 1)
                self.assertEqual(window._coord_combo.currentText(), "Recorded")
                with patch.object(QtWidgets.QFileDialog, "getOpenFileName", return_value=(str(root / meta.name), "")):
                    window._open_file()
                self.assertEqual(window._coord_combo.count(), 3)
                self.assertEqual(window._coord_combo.currentText(), "Specimen")
            finally:
                window.close()
                window.deleteLater()
                self.app.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)

    def test_recent_window_follows_auto_watch_and_preserves_repeated_fit_bound(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "TEST"
            write_header(path, SpecimenMeta(path.name))
            labels = ["NRM", "AF5", "AF10", "AF10", "AF20", "AF40"]
            for index, label in enumerate(labels):
                self._append_step(path, label, index)

            window = ZijderveldWindow()
            try:
                with patch.object(QtWidgets.QFileDialog, "getOpenFileName", return_value=(str(path), "")):
                    window._open_file()
                self.assertEqual(window._table.rowCount(), 6)
                self._set_recent_steps(window, 5)
                self.assertEqual(window._table.item(0, 0).text(), "AF5")
                self.assertEqual(window._table.rowCount(), 5)
                window._fit_start.setCurrentIndex(2)
                self.assertEqual(window._fit_start.currentText(), "4. AF10")

                self._append_step(path, "AF60", 6)
                deadline = time.monotonic() + 5
                while time.monotonic() < deadline and not window._watch_lbl.text().startswith("Reloaded"):
                    QtTest.QTest.qWait(50)

                self.assertTrue(window._watch_lbl.text().startswith("Reloaded"), window._watch_lbl.text())
                self.assertEqual(window._fit_start.currentText(), "4. AF10")
                self.assertEqual(window._fit_end.currentText(), "7. AF60")
                self.assertEqual(window._step_count_lbl.text(), "Steps: 5 of 7 | Fit window: 4 steps")
                self.assertEqual([window._table.item(row, 0).text() for row in range(5)], labels[2:] + ["AF60"])
                self.assertEqual(len(window._zij._horiz_line.getData()[0]), 5)
                self.assertEqual(len(window._stereo._specimen.steps), 5)
                self.assertEqual(len(window._vector3d._specimen.steps), 5)
                self.assertEqual(list(window._int_widget._curve.getData()[0]), [2, 3, 4, 5, 6])

                self._set_recent_steps(window, 2, accepted=False)
                self.assertEqual(window._table.rowCount(), 5)
                self._set_recent_steps(window, 1)
                self.assertEqual(window._fit_start.currentText(), "7. AF60")
                self.assertEqual(window._fit_end.currentText(), "7. AF60")
                self.assertIsNone(window._zij._fit_line.getData()[0])
                self._set_recent_steps(window, 0)
                self.assertEqual(window._table.rowCount(), 7)
                self.assertEqual(window._fit_start.currentText(), "1. NRM")
            finally:
                window.close()
                window.deleteLater()
                self.app.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)

    @staticmethod
    def _set_recent_steps(window: ZijderveldWindow, count: int, accepted: bool = True) -> None:
        menu_actions = window.menuBar().actions()
        view = next(action.menu() for action in menu_actions if action.text() == "&View")
        recent = next(action for action in view.actions() if action.text() == "&Recent steps…")
        with patch.object(QtWidgets.QInputDialog, "getInt", return_value=(count, accepted)):
            recent.trigger()

    @staticmethod
    def _append_step(path: Path, label: str, index: int) -> None:
        append_step(path, MeasurementStep(label, 80, 35, 80, 35, (7 - index) * 1e-9, 2, 80, 35, 0, 0, 0))
