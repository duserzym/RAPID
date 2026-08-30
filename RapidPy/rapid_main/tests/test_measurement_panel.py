from __future__ import annotations

import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from PySide6 import QtWidgets

from rapid_main.data_model import MeasurementStep
from rapid_main.dialogs.plots import PlotsDialog
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.measurement_worker import StepResult
from rapid_main.panels.measurement import _resolve_measurement_backend
from rapid_main.panels.measurement import MeasurementPanel


class _WindowWithBackend:
    def __init__(self, backend: object) -> None:
        self._backend = backend

    def measurement_backend(self) -> object:
        return self._backend


class _WindowWithoutBackend:
    pass


class _WindowWithBadBackend:
    def measurement_backend(self) -> object:
        return object()


class _WindowWithNonCallableBackendAttr:
    measurement_backend = NoCommBackend()


class TestMeasurementPanelHelpers(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_measurement_backend_uses_window_provider(self) -> None:
        backend = NoCommBackend()
        window = _WindowWithBackend(backend)

        resolved = _resolve_measurement_backend(window, NoCommBackend())

        self.assertIs(resolved, backend)

    def test_measurement_backend_falls_back_when_provider_absent(self) -> None:
        fallback = NoCommBackend()
        resolved = _resolve_measurement_backend(_WindowWithoutBackend(), fallback)
        self.assertIs(resolved, fallback)

    def test_measurement_backend_falls_back_when_provider_returns_invalid_backend(self) -> None:
        fallback = NoCommBackend()
        resolved = _resolve_measurement_backend(_WindowWithBadBackend(), fallback)
        self.assertIs(resolved, fallback)

    def test_measurement_backend_falls_back_when_provider_is_not_callable(self) -> None:
        fallback = NoCommBackend()
        resolved = _resolve_measurement_backend(_WindowWithNonCallableBackendAttr(), fallback)
        self.assertIs(resolved, fallback)

    def test_completed_steps_are_applied_to_plots_dialog(self) -> None:
        panel = MeasurementPanel()
        panel._completed_steps = [
            MeasurementStep(
                demag_label="NRM",
                gdec=0.0,
                ginc=0.0,
                sdec=0.0,
                sinc=0.0,
                moment=1.0,
                error_angle=0.0,
                crdec=0.0,
                crinc=0.0,
                sdx=1.0,
                sdy=0.0,
                sdz=0.0,
                timestamp=datetime(2026, 7, 15, 12, 0, 0),
            ),
            MeasurementStep(
                demag_label="AF20",
                gdec=90.0,
                ginc=45.0,
                sdec=90.0,
                sinc=45.0,
                moment=2.0,
                error_angle=0.0,
                crdec=90.0,
                crinc=45.0,
                sdx=0.0,
                sdy=1.0,
                sdz=-1.0,
                timestamp=datetime(2026, 7, 15, 12, 1, 0),
            ),
        ]
        dialog = PlotsDialog(panel)

        applied = panel._apply_completed_steps_to_plots(dialog)

        self.assertTrue(applied)
        self.assertEqual(dialog._last_plot_data["labels"], ["NRM", "AF20"])
        self.assertEqual(dialog._last_plot_data["north"], [1.0, 0.0])
        self.assertEqual(dialog._last_plot_data["east"], [0.0, 1.0])
        self.assertEqual(dialog._last_plot_data["up"], [0.0, -1.0])
        self.assertEqual(dialog._demo_lbl.text(), "Current measurement run")
        dialog.deleteLater()

    def test_successful_run_writes_quicklook_sidecar_to_output_bundle(self) -> None:
        panel = MeasurementPanel()
        with tempfile.TemporaryDirectory() as td:
            bundle_dir = Path(td) / "SAMPLE_A"
            panel._current_sample = "SAMPLE_A"
            panel._current_output_dir = bundle_dir
            panel._completed_steps = [
                MeasurementStep(
                    demag_label="NRM",
                    gdec=0.0,
                    ginc=0.0,
                    sdec=0.0,
                    sinc=0.0,
                    crdec=0.0,
                    crinc=0.0,
                    moment=1.0,
                    error_angle=0.0,
                    sdx=1.0,
                    sdy=0.0,
                    sdz=0.0,
                    timestamp=datetime(2026, 7, 15, 12, 0, 0),
                ),
            ]

            panel._on_run_finished(False)
            payload = json.loads((bundle_dir / "quicklook.json").read_text(encoding="utf-8"))

        self.assertEqual(payload["step_count"], 1)
        self.assertEqual(payload["labels"], ["NRM"])
        self.assertEqual(payload["vectors"]["north"], [1.0])
        self.assertEqual(payload["intensity"], [1.0])
        panel.deleteLater()
        panel.deleteLater()

    def test_completed_cycle_populates_live_stats_grid(self) -> None:
        panel = MeasurementPanel()
        step = MeasurementStep(
            demag_label="NRM",
            gdec=45.0,
            ginc=35.0,
            sdec=45.0,
            sinc=35.0,
            moment=4.0,
            error_angle=0.0,
            crdec=45.0,
            crinc=35.0,
            sdx=2.0,
            sdy=2.0,
            sdz=2.0,
        )

        panel._on_step_complete(StepResult(step, 0.0, 0, 1))

        self.assertEqual(panel._avg_moment.text(), "4.000e+00 emu")
        self.assertEqual(panel._avg_dec.text(), "45.0 deg")
        self.assertEqual(panel._avg_inc.text(), "35.0 deg")
        self.assertEqual(panel._delta_x.text(), "0.000e+00")
        self.assertEqual(panel._ratio_y.text(), "0.00")
        self.assertEqual(panel._sig_drift.text(), "stable")
        self.assertEqual(panel._sig_holder.text(), "N/A")
        self.assertEqual(panel._sig_induced.text(), "N/A")
        panel.deleteLater()
