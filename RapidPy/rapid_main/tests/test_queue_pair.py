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
from rapid_main.queue_pair import LegacyUpSource, QueuePair
from tests.test_legacy_up import FIXTURE as VB6_UP


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

    def set_demag_step(self, label):
        self.calls.append('treat ' + label)
        super().set_demag_step(label)

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

    def legacy_side(self, *, label='NRM', range_factor=2e-5, susceptibility=None):
        path = self.output / 'original.UP'
        path.write_text(VB6_UP, encoding='ascii')
        commands = compile_queue([QueueSample('A', 'index', 9, do_up=False, do_both=True,
            measurement_labels=(label,))], QueueOptions(), strict=True)
        command, = [item for item in commands if item.command_type == 'Meas']
        self.assertTrue(command.legacy_up_required)
        with self.assertRaisesRegex(ValueError, 'legacy Up'):
            QueuePair.from_command('legacy-queue', commands, command)
        source = LegacyUpSource.capture(path, file_id='index', label=label, specimens=['A'],
            range_factor=range_factor, operator='Operator', susceptibility=susceptibility)
        return QueuePair.from_command('legacy-queue', commands, command, legacy_up=source)

    def test_legacy_up_is_scaled_once_archived_and_linked_to_down(self):
        side = self.legacy_side(label='SUSC', susceptibility={'A': .125})
        self.assertFalse(side.is_up)
        self.assertIsNone(side.pair.up_index)
        backend = PairBackend((8e-5, 4e-5, 2e-5))
        worker, finished, results = self.measure(side, backend)
        self.assertEqual(finished, [False], worker._error_messages)
        result, = results
        for actual, expected in zip((result.step.sdx, result.step.sdy, result.step.sdz), (6e-5, 6e-5, 7e-5)):
            self.assertAlmostEqual(actual, expected)
        self.assertEqual(result.collection_stats.position_count, 8)
        self.assertEqual(result.susceptibility, .125)
        self.assertNotIn('SUSC', backend.calls)
        self.assertNotIn('treat SUSC', backend.calls)
        provenance = json.loads((self.simulated / 'provenance.json').read_text())
        reference = provenance['up_acquisition']
        self.assertEqual(reference['schema'], 'rapidpy.legacy_up_import.v1')
        self.assertEqual(reference['operator'], 'Operator')
        self.assertFalse(reference['audited'])
        self.assertEqual(Path(reference['archive']).read_bytes(), Path(side.legacy_up.path).read_bytes())
        self.assertEqual(reference['sha256'], side.legacy_up.sha256)
        index = json.loads((self.simulated / 'artifact_index.json').read_text())
        artifact, = [item for item in index['artifacts'] if item['name'] == 'legacy_up_source']
        self.assertEqual((self.simulated / artifact['relative_path']).read_bytes(), Path(reference['archive']).read_bytes())
        self.assertEqual(artifact['sha256'], reference['sha256'])

    def test_legacy_up_mutation_before_preflight_refuses_down(self):
        side = self.legacy_side()
        Path(side.legacy_up.path).write_bytes(b'changed archive')
        backend = PairBackend((1e-5, 2e-5, 3e-5))
        _, finished, _ = self.measure(side, backend)
        self.assertEqual(finished, [True])
        self.assertEqual(backend.calls, [])

    def test_legacy_susc_requires_original_bridge_value(self):
        with self.assertRaisesRegex(ValueError, 'original susceptibility'):
            self.legacy_side(label='SUSC')

    def test_changed_retained_up_copy_preserves_previous_scientific_output(self):
        side = self.legacy_side()
        original = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], self.output, PairBackend((1e-5, 2e-5, 3e-5)))
        original.run()
        accepted = self.simulated / 'A'
        before = accepted.read_bytes()
        backend = PairBackend((8e-5, 4e-5, 2e-5))

        def alter_archive():
            target = side.evidence_directory(self.simulated) / 'legacy-up.UP'
            target.write_bytes(b'not original')

        backend.return_to_safe_state = alter_archive
        worker, finished, _ = self.measure(side, backend)
        self.assertEqual(finished, [True])
        self.assertIn('archived legacy Up input changed', '\n'.join(worker._error_messages))
        self.assertEqual(accepted.read_bytes(), before)

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
        self.assertNotIn('treat SUSC', down_backend.calls)
        self.assertEqual(second._up_reference['rmg_treatment']['label'], 'SUSC')
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

    def test_down_uses_the_applied_up_field_when_label_has_no_level(self):
        from rapid_main.af_treatment import AfPass, AfTreatmentPlan, AfTreatmentRecord, CalibratedAfRamp
        from rapid_main.io.rmg_reader import read_rmg

        up, down = queue_sides('AFmax')

        class UpStation(PairBackend):
            af_treatment_records = ()

            def set_demag_step(self, label):
                super().set_demag_step(label)
                ramp = CalibratedAfRamp(label, 'axial', 125., 1., 1., 50., 1., 1., 400, 100., 'fixture')
                plan = AfTreatmentPlan(label, -100, (AfPass(0., ramp),))
                self.af_treatment_records = (AfTreatmentRecord('af-' + 'b' * 32, 'A', up.run_id,
                    plan, 1, (), '', (), True, True),)

        first, finished, _ = self.measure(up, UpStation((1e-5, 2e-5, 3e-5)))
        self.assertEqual(finished, [False], first._error_messages)
        backend = PairBackend((1e-5, 2e-5, 3e-5))
        second, finished, _ = self.measure(down, backend)
        self.assertEqual(finished, [False], second._error_messages)
        self.assertNotIn('treat AFmax', backend.calls)
        record, = read_rmg(self.simulated / 'A.rmg')
        self.assertEqual((record.step_type, record.step_value, record.hold_time_s), ('AFmax', 1250., .4))
        self.assertEqual(second._up_reference['rmg_treatment']['evidence_id'], 'af-' + 'b' * 32)

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
                if not self.side.is_up:
                    raise AssertionError('Down must preserve the Up treatment')
                self.context['treatment_label'] = label

            def return_to_safe_state(self):
                pass

            def read_squid(self):
                z = 3. if self.side.is_up else -3.
                observations = [counts_for((0., 0., value), ARCHIVED_CALIBRATION)
                                for value in (0., z, z, z, z, 0.)]
                service, *_ = _build_service(observations)
                context = BlockContext(**self.context,
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

        historical = self.legacy_side(range_factor=evidence['blocks'][0]['range_factor'])
        backend = InjectedStation(historical)
        worker, finished, _ = self.measure(historical, backend)
        self.assertEqual(finished, [False], worker._error_messages)
        evidence, = worker._progress.accepted_steps()
        block, = evidence['blocks']
        self.assertEqual(block['audit']['holder_record_id'], 'holder-False')
        self.assertEqual(len(block['observations']), 6)
        data = json.loads((self.output / 'provenance.json').read_text())
        self.assertFalse(data['up_acquisition']['audited'])
        self.assertEqual(data['block_collection_statistics'][0]['position_count'], 8)

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

    def test_main_window_selects_legacy_up_and_runs_only_down(self):
        from rapid_main.app import MainWindow
        from rapid_main.config import AppConfig

        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        config = AppConfig()
        config.general.nocomm = True
        config.general.operator = 'Legacy Operator'
        config.general.data_dir = str(self.output)
        source = self.output / 'NRM.UP'
        source.write_text(VB6_UP, encoding='ascii')
        errors = []
        with (patch.object(AppConfig, 'load', return_value=config),
              patch.object(AppConfig, 'save', lambda self: None),
              patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: errors.append(args[-1])),
              patch.object(QtWidgets.QFileDialog, 'getOpenFileName', lambda *args: (str(source), '')),
              patch.object(QtWidgets.QInputDialog, 'getText', lambda *args, **kwargs: ('1e-5', True)),
              patch.object(MainWindow, '_notify_queue_terminal', lambda *args: None)):
            window = MainWindow()
            try:
                backend = PairBackend((1e-5, 2e-5, 3e-5))
                window._measurement_backend = backend
                window._sequence_labels = ['NRM']
                self.assertTrue(window.start_queue_run([QueueSample('A', 'F1', 1, do_up=False, do_both=True)],
                                                       QueueOptions()))
                deadline = time.monotonic() + 15
                while window._queue_active and time.monotonic() < deadline:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(window._queue_active)
                self.assertEqual(errors, [])
                self.assertEqual(backend.calls.count('SQUID'), 1)
                result = self.output / 'A' / 'SIMULATED'
                provenance = json.loads((result / 'provenance.json').read_text())
                self.assertEqual(provenance['up_acquisition']['operator'], 'Legacy Operator')
                self.assertEqual(provenance['block_collection_statistics'][0]['position_count'], 8)
            finally:
                window._queue_advance_timer.stop()
                worker = window._measurement._worker
                if worker is not None:
                    worker.halt()
                    worker.wait()
                window.deleteLater()
                app.sendPostedEvents(None, QtCore.QEvent.Type.DeferredDelete)
