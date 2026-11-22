import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.queue_compiler import QueueOptions, QueueSample, compile_queue
from rapid_main.queue_pair import QueuePair


class PairBackend(NoCommBackend):
    def __init__(self, vector, *, fail_return=False):
        super().__init__()
        self.calls = []
        self.vector = vector
        self.fail_return = fail_return

    def preflight(self):
        self.calls.append('preflight')
        return super().preflight()

    def read_squid(self):
        self.calls.append('SQUID')
        return self.vector

    def read_susceptibility(self):
        self.calls.append('SUSC')
        return .25

    def return_to_safe_state(self):
        self.calls.append('return')
        if self.fail_return:
            raise RuntimeError('return did not settle')


def queue_sides(label='NRM', *, run='queue-one'):
    commands = compile_queue([QueueSample('A', 'index', 9, do_both=True,
        measurement_labels=(label,), row_id='f' * 32, avg_steps=2)], QueueOptions(), strict=True)
    return tuple(QueuePair.from_command(run, commands, command)
                 for command in commands if command.command_type == 'Meas')


class QueuePairTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.output = Path(folder.name)
        self.simulated = self.output / 'SIMULATED'

    def measure(self, side, backend, *, meta=None):
        worker = MeasurementWorker(meta or SpecimenMeta(name='A'), [side.pair.label], self.output, backend,
            samples_per_position=side.pair.count, queue_side=side)
        finished = []
        results = []
        worker.run_finished.connect(finished.append)
        worker.step_complete.connect(results.append)
        worker.run()
        return worker, finished, results

    def test_up_stages_raw_only_and_down_publishes_combined_positions_and_up_susc(self):
        up, down = queue_sides('SUSC')
        up_backend = PairBackend((2e-5, 0., 1e-5))
        first, finished, _ = self.measure(up, up_backend)
        self.assertEqual(finished, [False], first._error_messages)
        self.assertFalse((self.simulated / 'A').exists())
        self.assertEqual(up_backend.calls.count('SUSC'), 1)
        raw = (first._progress.directory / '00000.UP').read_bytes()
        down_backend = PairBackend((4e-5, 0., 3e-5))
        second, finished, results = self.measure(down, down_backend)
        self.assertEqual(finished, [False], second._error_messages)
        result, = results
        self.assertAlmostEqual(result.step.sdx, 3e-5)
        self.assertAlmostEqual(result.step.sdz, 2e-5)
        self.assertEqual(result.collection_stats.position_count, 16)
        self.assertEqual(result.susceptibility, .25)
        self.assertNotIn('SUSC', down_backend.calls)
        self.assertEqual((first._progress.directory / '00000.UP').read_bytes(), raw)
        self.assertTrue((self.simulated / 'A').is_file())
        provenance = json.loads((self.simulated / 'provenance.json').read_text())
        self.assertEqual(provenance['up_acquisition']['identity_sha256'], first._progress.identity_sha256)
        self.assertEqual(provenance['block_collection_statistics'][0]['position_count'], 16)

    def test_down_requires_exact_queue_and_unmodified_legacy_export_before_preflight(self):
        up, down = queue_sides()
        _, other_down = queue_sides(run='queue-two')
        first, finished, _ = self.measure(up, PairBackend((1e-5, 1e-5, 1e-5)))
        self.assertEqual(finished, [False], first._error_messages)
        for side in (other_down, down):
            if side is down:
                (first._progress.directory / '00000.UP').write_bytes(b'changed')
            backend = PairBackend((1e-5, 1e-5, 1e-5))
            _, finished, _ = self.measure(side, backend)
            self.assertEqual(finished, [True])
            self.assertEqual(backend.calls, [])

    def test_failed_return_leaves_up_ineligible_and_cannot_replay_same_command(self):
        up, down = queue_sides()
        first, finished, _ = self.measure(up, PairBackend((1e-5, 1e-5, 1e-5), fail_return=True))
        self.assertEqual(finished, [True])
        self.assertTrue((first._progress.directory / '00000.UP').exists())
        for side in (up, down):
            backend = PairBackend((1e-5, 1e-5, 1e-5))
            _, finished, _ = self.measure(side, backend)
            self.assertEqual(finished, [True])
            self.assertEqual(backend.calls, [])

    def test_changed_specimen_metadata_blocks_down_hardware(self):
        up, down = queue_sides()
        first, finished, _ = self.measure(up, PairBackend((1e-5, 1e-5, 1e-5)))
        self.assertEqual(finished, [False], first._error_messages)
        backend = PairBackend((1e-5, 1e-5, 1e-5))
        _, finished, _ = self.measure(down, backend, meta=SpecimenMeta(name='A', volume=2))
        self.assertEqual(finished, [True])
        self.assertEqual(backend.calls, [])

    def test_failed_down_return_preserves_previous_result_and_index(self):
        up, down = queue_sides()
        backend = PairBackend((1e-5, 1e-5, 1e-5))
        original = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], self.output, backend)
        original.run()
        files = [self.simulated / name for name in ('A', 'provenance.json', 'artifact_index.json')]
        before = [path.read_bytes() for path in files]
        first, finished, _ = self.measure(up, PairBackend((1e-5, 1e-5, 1e-5)))
        self.assertEqual(finished, [False], first._error_messages)
        failed, finished, _ = self.measure(down, PairBackend((2e-5, 1e-5, 1e-5), fail_return=True))
        self.assertEqual(finished, [True])
        self.assertEqual([path.read_bytes() for path in files], before)
        aborted = json.loads((failed._publish_dir / 'artifact_index.json').read_text())
        self.assertTrue(aborted['aborted'])
        self.assertTrue((failed._progress.directory / '00000.UP').is_file())

    def test_live_pair_retains_two_coherent_holder_audits(self):
        from dataclasses import replace
        from rapid_main.acquisition import BlockContext
        from rapid_main.hardware_contracts import PreflightResult
        from test_acquisition import ARCHIVED_CALIBRATION, _build_service
        from tests.acquisition_fakes import counts_for

        up, down = queue_sides()

        class InjectedStation:
            simulated = False

            def __init__(self, side):
                self.side = side
                self.completed = []

            def preflight(self):
                return PreflightResult.pass_ok()

            def is_available(self):
                return True

            def set_measurement_context(self, **context):
                self.context = context

            def set_demag_step(self, label):
                self.label = label

            def return_to_safe_state(self):
                pass

            def read_squid(self):
                z = 3. if self.side.is_up else -3.
                observations = [counts_for((0., 0., value), ARCHIVED_CALIBRATION)
                                for value in (0., z, z, z, z, 0.)]
                service, *_ = _build_service(observations)
                context = BlockContext(**self.context, treatment_label=self.label,
                    config_hash='station', holder_record_id=f'holder-{self.side.is_up}')
                block = service.acquire(is_up=self.side.is_up, context=context).block
                block = replace(block, audit=replace(block.audit,
                    block_id=f'{self.side.run_id}-{len(self.completed)}'))
                self.completed.append(block)
                return block

        for side in (up, down):
            backend = InjectedStation(side)
            worker, finished, _ = self.measure(side, backend)
            self.assertEqual(finished, [False], worker._error_messages)
            evidence, = worker._progress.accepted_steps()
            self.assertEqual({block['audit']['holder_record_id'] for block in evidence['blocks']},
                             {f'holder-{side.is_up}'})
            self.assertTrue(all(len(block['observations']) == 6 for block in evidence['blocks']))
        data = json.loads((self.output / 'provenance.json').read_text())
        self.assertEqual(data['block_collection_statistics'][0]['position_count'], 16)

    def test_main_window_runs_both_passes_and_publishes_one_result(self):
        from rapid_main.app import MainWindow
        from rapid_main.config import AppConfig

        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        config = AppConfig()
        config.general.nocomm = True
        config.general.data_dir = str(self.output)
        config.general.sample_dir = str(self.output)
        errors = []
        with (patch.object(AppConfig, 'load', return_value=config),
              patch.object(AppConfig, 'save', lambda self: None),
              patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: errors.append(args[-1])),
              patch.object(MainWindow, '_notify_queue_terminal', lambda *args: None)):
            window = MainWindow()
            try:
                backend = PairBackend((1e-5, 2e-5, 3e-5))
                window._measurement_backend = backend
                window._sequence_labels = ['SUSC']
                self.assertTrue(window.start_queue_run([QueueSample('A', 'F1', 1, do_both=True)], QueueOptions()))
                deadline = time.monotonic() + 15
                while window._queue_active and time.monotonic() < deadline:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(window._queue_active)
                self.assertEqual(errors, [])
                result = self.output / 'A' / 'SIMULATED'
                provenance = json.loads((result / 'provenance.json').read_text())
                self.assertEqual(backend.calls.count('SUSC'), 1)
                self.assertEqual(provenance['block_collection_statistics'][0]['position_count'], 8)
                self.assertTrue((Path(provenance['up_acquisition']['directory']) / '00000.UP').is_file())
            finally:
                window._queue_advance_timer.stop()
                worker = window._measurement._worker
                if worker is not None:
                    worker.halt()
                    worker.wait()
                window.deleteLater()
                app.sendPostedEvents(None, QtCore.QEvent.Type.DeferredDelete)
