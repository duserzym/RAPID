from __future__ import annotations

import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.specimen_writer import write_header
from rapid_main.specimen_metadata import capture_specimen_metadata, restore_specimen_metadata
from rapid_main.dialogs.plots import PlotsDialog
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.measurement_worker import StepResult
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.config import AppConfig
from rapid_main.io.sample_index import read_sample_index_registrations
from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations
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
    @staticmethod
    def _dispose_window(window):
        window.deleteLater()
        QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
        QtWidgets.QApplication.instance().processEvents()

    def test_queue_uses_original_index_metadata_and_separate_outputs_for_duplicate_names(self):
        window = QtWidgets.QMainWindow()
        self.addCleanup(self._dispose_window, window)
        panel = MeasurementPanel(window)
        window._sequence_labels = ['IRM100']
        window.queue_measurement_labels = lambda name: ['NRM']
        window.measurement_backend = lambda: NoCommBackend()
        window.set_status = lambda message: None
        window.sample_registrations = SampleIndexRegistrations([SampleIndexRegistration('SAME', formation='WrongUnit', location='WrongSite')])
        folders = []
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            window.config = AppConfig()
            window.config.general.data_dir = str(root / 'output')
            window.config.general.sample_dir = str(root / 'wrong')
            for folder, unit in (('one', 'OriginalUnit'), ('two', 'OtherUnit')):
                index = root / folder / 'index.sam'
                index.parent.mkdir()
                index.write_text('SAME 2 ' + unit + ' Site\n', encoding='latin-1')
                registrations = read_sample_index_registrations(index)
                write_header(index.parent / 'SAME', SpecimenMeta('SAME', comment='Header', volume=8.2, core_plate_strike=123))
                snapshot = capture_specimen_metadata('SAME', sample_dir=index.parent, registrations=registrations)
                window.queue_measurement_source = lambda name, index=index, registrations=registrations: (index, registrations)
                window.queue_measurement_metadata = lambda name, snapshot=snapshot: restore_specimen_metadata(snapshot)
                provenance = dict(schema='rapidpy.queue_specimen_source.v1', source_file=str(index),
                    file_id=str(index), index_sha256='a' * 64, row_id='b' * 32, specimen=snapshot)
                window.queue_measurement_provenance = lambda name, provenance=provenance: provenance
                with patch.object(MeasurementWorker, 'start', lambda worker: None):
                    self.assertTrue(panel.start_measurement_for_sample('SAME', queue_run=True))
                self.assertEqual(panel._worker._meta.site, unit)
                self.assertEqual(panel._worker._meta.location, 'Site')
                self.assertEqual((panel._worker._meta.volume, panel._worker._meta.core_plate_strike), (8.2, 123))
                self.assertEqual(panel._worker._base_provenance()['specimen_source'], provenance)
                folders.append(panel._worker._output_dir)
                worker = panel._worker
                (index.parent / 'SAME').write_bytes(b'Changed header\n')
                self.assertFalse(panel.start_measurement_for_sample('SAME', queue_run=True))
                self.assertIs(panel._worker, worker)
            self.assertNotEqual(folders[0], folders[1])
    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_measurement_backend_uses_window_provider(self) -> None:
        backend = NoCommBackend()
        window = _WindowWithBackend(backend)

        resolved = _resolve_measurement_backend(window, NoCommBackend())

        self.assertIs(resolved, backend)

    def test_queue_worker_receives_file_steps_instead_of_global_sequence(self):
        window = QtWidgets.QMainWindow()
        panel = MeasurementPanel(window)
        window._sequence_labels = ['IRM100']
        window.queue_measurement_labels = lambda sample: ['NRM', 'SUSC']
        window.measurement_backend = lambda: NoCommBackend()
        with tempfile.TemporaryDirectory() as directory:
            window.config = AppConfig()
            window.config.general.data_dir = directory
            window.config.general.sample_dir = directory
            with patch.object(MeasurementWorker, 'start', lambda worker: None):
                self.assertTrue(panel.start_measurement_for_sample('A', queue_run=True))
            self.assertEqual(panel._worker._labels, ['NRM', 'SUSC'])
            panel._on_step_started(1, 'SUSC')
            self.assertEqual(panel._meas_step.text(), '2 / 2')
            self.assertEqual(window._sequence_labels, ['IRM100'])
        window.deleteLater()

    def test_invalid_queue_handoff_creates_no_worker_or_device_lease(self):
        window = QtWidgets.QMainWindow()
        panel = MeasurementPanel(window)
        window._sequence_labels = ['NRM']
        statuses, leases = [], []
        window.set_status = statuses.append
        window.acquire_measurement_device = lambda owner: leases.append(owner)
        def reject(sample):
            raise ValueError('original file sequence changed')
        window.queue_measurement_labels = reject
        self.assertFalse(panel.start_measurement_for_sample('A', queue_run=True))
        self.assertIsNone(panel._worker)
        self.assertEqual(leases, [])
        self.assertIn('original file sequence changed', statuses[-1])
        window.deleteLater()

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
        self.assertTrue(dialog._demo_lbl.text().startswith("READY —"))
        self.assertIn("Current measurement run", dialog.quicklook_summary()["provenance"]["statement"])
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
        self.assertFalse(payload["provenance"]["simulated"])
        panel.deleteLater()

    def test_simulated_run_quicklook_is_marked_in_dialog_and_sidecar(self) -> None:
        panel = MeasurementPanel()
        panel._current_run_simulated = True
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
            )
        ]
        dialog = PlotsDialog(panel)
        try:
            self.assertTrue(panel._apply_completed_steps_to_plots(dialog))
            self.assertEqual(dialog._demo_lbl.property("status"), "simulated")
            self.assertIn("not hardware evidence", dialog._demo_lbl.text())

            with tempfile.TemporaryDirectory() as td:
                panel._current_output_dir = Path(td)
                path = panel._write_quicklook_sidecar()
                self.assertIsNotNone(path)
                payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertTrue(payload["provenance"]["simulated"])
            self.assertIn("not hardware evidence", payload["provenance"]["statement"])
        finally:
            dialog.deleteLater()
            panel.deleteLater()

    def test_quicklook_publish_reconciles_artifact_index(self) -> None:
        panel = MeasurementPanel()
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
            )
        ]
        with tempfile.TemporaryDirectory() as td:
            output = Path(td)
            panel._current_output_dir = output
            (output / "artifact_index.json").write_text(
                json.dumps(
                    {
                        "schema": "rapidpy.measurement.artifact_index.v1",
                        "artifacts": [
                            {
                                "name": "quicklook_summary",
                                "relative_path": "quicklook.json",
                                "exists": False,
                                "size_bytes": 0,
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            written = panel._write_quicklook_sidecar()
            index = json.loads(
                (output / "artifact_index.json").read_text(encoding="utf-8")
            )

        self.assertIsNotNone(written)
        entry = index["artifacts"][0]
        self.assertTrue(entry["exists"])
        self.assertGreater(entry["size_bytes"], 0)
        self.assertEqual(entry["relative_path"], "quicklook.json")
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


class TestMeasurementPanelHolderStats(unittest.TestCase):
    """Holder identity, magnitude, age, and validity must be visible."""

    @classmethod
    def setUpClass(cls) -> None:
        cls._app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def _panel(self) -> MeasurementPanel:
        return MeasurementPanel()

    def test_absent_holder_is_shown_as_none_with_the_reason(self) -> None:
        from rapid_main.holder_state import HolderStateStore

        panel = self._panel()
        status = HolderStateStore().status()

        panel._update_holder_stats(None, status)

        self.assertEqual(panel._holder_id.text(), "None")
        self.assertIn("No holder measurement", panel._holder_id.toolTip())
        self.assertEqual(panel._sig_holder.text(), "N/A")
        self.assertEqual(panel._sig_induced.text(), "N/A")

    def test_valid_holder_renders_identity_magnitude_and_age(self) -> None:
        from datetime import timedelta, timezone

        from rapid_main.holder_state import HolderCorrection, HolderMetrics, HolderStateStore

        now = datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)
        correction = HolderCorrection(
            holder_id="HOLDER-A",
            positions=((0.1, 0.0, 0.0), (0.0, 0.1, 0.0), (-0.1, 0.0, 0.0), (0.0, -0.1, 0.0)),
            measured_at_iso=(now - timedelta(minutes=30)).isoformat(),
            metrics=HolderMetrics(magnitude_emu=1.5e-6, asymmetry_ratio=0.02),
        )
        store = HolderStateStore(clock=lambda: now)
        store.install(correction)

        panel = self._panel()
        panel._update_holder_stats(None, store.status())

        self.assertEqual(panel._holder_id.text(), "HOLDER-A")
        self.assertEqual(panel._holder_moment.text(), "1.500e-06 emu")
        self.assertEqual(panel._holder_age.text(), "30 min")

    def test_stale_holder_is_labelled_invalid(self) -> None:
        from datetime import timedelta, timezone

        from rapid_main.holder_state import HolderCorrection, HolderStateStore

        now = datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)
        correction = HolderCorrection(
            holder_id="HOLDER-B",
            positions=((0.1, 0.0, 0.0),) * 4,
            measured_at_iso=(now - timedelta(days=3)).isoformat(),
        )
        store = HolderStateStore(clock=lambda: now)
        store.install(correction)

        panel = self._panel()
        panel._update_holder_stats(None, store.status())

        self.assertIn("invalid", panel._holder_id.text())
        self.assertIn("stale", panel._holder_id.toolTip())
        self.assertEqual(panel._holder_age.text(), "72.0 h")

    def test_block_result_supplies_the_signal_ratios(self) -> None:
        from rapid_main.magnetometer import BracketedMeasurementBlock, reduce_bracketed_measurement

        block = BracketedMeasurementBlock(
            zero_before=(0.0, 0.0, 0.0),
            positions=(
                (1.0, 0.0, 0.5),
                (0.0, 1.0, 0.5),
                (-1.0, 0.0, 0.5),
                (0.0, -1.0, 0.5),
            ),
            zero_after=(0.001, 0.0, 0.0),
            holder_positions=((0.01, 0.0, 0.0), (0.0, 0.01, 0.0), (-0.01, 0.0, 0.0), (0.0, -0.01, 0.0)),
        )
        result = reduce_bracketed_measurement(block)

        panel = self._panel()
        panel._update_holder_stats(result, None)

        self.assertEqual(panel._sig_holder.text(), f"{result.sig_holder:.2f}")
        self.assertEqual(panel._sig_induced.text(), f"{result.sig_induced:.2f}")
        self.assertNotEqual(panel._sig_holder.text(), "N/A")
