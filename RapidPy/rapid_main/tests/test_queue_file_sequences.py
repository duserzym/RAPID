"""Per-file executable plans, dual-orientation eligibility and frozen handoffs."""
import unittest
from types import SimpleNamespace

from rapid_main.app import MainWindow
from rapid_main.queue_compiler import QueueSample, QueueOptions, compile_queue, resolve_queue_samples


class QueueFileSequenceTests(unittest.TestCase):
    def test_runtime_sequence_does_not_replace_loaded_sequence(self):
        window = SimpleNamespace(_sequence_labels=['AF20', 'SUSC'],
            _run_timer=SimpleNamespace(start=lambda: None, stop=lambda: None),
            _update_runtime_display=lambda **kwargs: None, set_flow_state=lambda state: None)
        MainWindow.start_run(window, ['NRM'])
        self.assertEqual(window._run_sequence_labels, ['NRM'])
        self.assertEqual(window._sequence_labels, ['AF20', 'SUSC'])
        MainWindow.stop_run(window)
        self.assertIsNone(window._run_sequence_labels)
        self.assertEqual(window._sequence_labels, ['AF20', 'SUSC'])

    def test_explicit_file_steps_need_no_global_default(self):
        samples = resolve_queue_samples([QueueSample('A', 'F1', 1, measurement_labels=('NRM',))], [])
        commands = compile_queue(samples, QueueOptions(), strict=True)
        self.assertEqual(next(item.measurement_labels for item in commands if item.command_type == 'Meas'), ('NRM',))

    def test_distinct_files_keep_distinct_steps_through_second_pass(self):
        samples = [QueueSample('A', 'F1', 1, do_both=True, measurement_labels=('NRM',)),
                   QueueSample('B', 'F2', 2, measurement_step_count=2, measurement_labels=('NRM', 'AF20'))]
        commands = compile_queue(samples, QueueOptions(), strict=True)
        measurements = [command for command in commands if command.command_type == 'Meas']
        self.assertEqual([(item.file_id, item.measurement_labels) for item in measurements],
                         [('F1', ('NRM',)), ('F2', ('NRM', 'AF20')), ('F1', ('NRM',))])

    def test_default_steps_snapshot_and_control_both_side_eligibility(self):
        source = QueueSample('A', 'F1', 1, do_both=True)
        defaults = ['NRM', 'AF20']
        resolved = resolve_queue_samples([source], defaults)
        commands = compile_queue(resolved, QueueOptions(), strict=True)
        defaults[:] = ['IRM100']
        source.do_both = False
        self.assertEqual(resolved[0].measurement_step_count, 2)
        self.assertTrue(resolved[0].do_both)
        self.assertEqual([item.measurement_labels for item in commands if item.command_type == 'Meas'], [('NRM', 'AF20')])
        self.assertFalse(any(item.command_type in {'InitUp', 'Flip'} for item in commands))

    def test_same_file_conflicting_steps_or_flags_are_rejected(self):
        first = QueueSample('A', 'F1', 1, measurement_labels=('NRM',))
        for second in (QueueSample('B', 'F1', 2, measurement_labels=('AF20',)),
                       QueueSample('B', 'F1', 2, do_up=False, measurement_labels=('NRM',)),
                       QueueSample('B', 'F1', 2, do_both=True, measurement_labels=('NRM',))):
            with self.subTest(second=second), self.assertRaisesRegex(ValueError, 'one file must agree'):
                compile_queue([first, second], QueueOptions(), strict=True)

    def test_malformed_counts_flags_and_labels_are_rejected(self):
        for sample in (QueueSample('A', 'F1', True), QueueSample('A', 'F1', 1, do_both=1),
                       QueueSample('A', 'F1', 1, measurement_step_count=True),
                       QueueSample('A', 'F1', 1, measurement_labels=('NRM', 'AF20')),
                       QueueSample('A', 'F1', 1, measurement_labels=(' ',))):
            with self.subTest(sample=sample), self.assertRaises(ValueError):
                compile_queue([sample], QueueOptions(), strict=True)

    def test_handoff_uses_compiled_file_steps_after_global_sequence_changes(self):
        commands = compile_queue([QueueSample('A', 'F1', 1, measurement_labels=('NRM',))], QueueOptions(), strict=True)
        command = next(item for item in commands if item.command_type == 'Meas')
        window = SimpleNamespace(_queue_active=True, _queue_current_command=command, _queue_plan=commands,
            _queue_native_session=None, _rockmag_routine_plan=None, _thermal_routine_plan=None,
            _sequence_labels=['IRM100'])
        self.assertEqual(MainWindow.queue_measurement_labels(window, 'A'), ['NRM'])
        with self.assertRaises(ValueError):
            MainWindow.queue_measurement_labels(window, 'B')

    def test_handoff_rejects_reviewed_routine_identity_mismatch(self):
        commands = compile_queue([QueueSample('A', 'F1', 1, measurement_labels=('NRM',))], QueueOptions(), strict=True)
        window = SimpleNamespace(_queue_active=True, _queue_current_command=next(item for item in commands if item.command_type == 'Meas'),
            _queue_plan=commands, _queue_native_session=None, _thermal_routine_plan=None,
            _rockmag_routine_plan=SimpleNamespace(to_queue_labels=lambda: ['AF20']))
        with self.assertRaisesRegex(ValueError, 'routine identity'):
            MainWindow.queue_measurement_labels(window, 'A')
