"""Match the append order in VB6 SampleCommands.Preprocess."""
import unittest
from rapid_main.queue_compiler import QueueCommand, QueueSample, QueueOptions, compile_queue, preprocess_queue


class QueuePassOrderTests(unittest.TestCase):
    def samples(self):
        return [QueueSample('S1', 'F1', 1, do_both=True), QueueSample('S2', 'F1', 2, do_both=True)]

    def test_whole_first_pass_precedes_flip_and_repeat_reads(self):
        queue = compile_queue(self.samples(), QueueOptions(load_return=False, do_return=False, repeat_holder=False))
        self.assertEqual([(cmd.command_type, cmd.sample_name) for cmd in queue],
            [('InitUp', ''), ('Holder', ''), ('Meas', 'S1'), ('Meas', 'S2'),
             ('Flip', ''), ('Holder', ''), ('Meas', 'S1'), ('Meas', 'S2')])

    def test_multiple_files_follow_original_collection_append_order(self):
        samples = [QueueSample('A', 'F1', 1, do_both=True), QueueSample('B', 'F2', 2, do_both=True)]
        queue = compile_queue(samples, QueueOptions(load_return=False, do_return=False, repeat_holder=False))
        self.assertEqual([(cmd.command_type, cmd.file_id) for cmd in queue],
            [('InitUp', 'F1'), ('InitUp', 'F2'), ('Holder', ''), ('Meas', 'F1'), ('Meas', 'F2'),
             ('Flip', 'F1'), ('Holder', ''), ('Flip', 'F2'), ('Holder', ''), ('Meas', 'F1'), ('Meas', 'F2')])

    def test_second_pass_retains_periodic_holder_and_non_xy_flip_marker(self):
        queue = compile_queue(self.samples(), QueueOptions(load_return=False, do_return=False,
            use_xy_table=False, repeat_holder=True, samples_between_holder=2))
        index = next(i for i, cmd in enumerate(queue) if cmd.command_type == 'Flip')
        self.assertEqual(queue[index].hole, 0)
        self.assertEqual([cmd.command_type for cmd in queue[index:]], ['Flip', 'Holder', 'Meas', 'Meas', 'Holder'])

    def test_existing_flip_ends_repeat_enrollment_for_later_original_reads(self):
        commands = [QueueCommand('InitUp', file_id='F1'), QueueCommand('Meas', file_id='F1', sample_name='A'),
            QueueCommand('Flip', file_id='F1'), QueueCommand('Meas', file_id='F1', sample_name='B')]
        queue = preprocess_queue(commands, {'F1': self.samples()[0]}, options=QueueOptions(repeat_holder=False))
        self.assertEqual([cmd.sample_name for cmd in queue if cmd.command_type == 'Meas'], ['A', 'B', 'A'])
        self.assertEqual(queue[:len(commands)], commands)

    def test_repeat_command_is_an_independent_copy_and_return_follows_both_passes(self):
        queue = compile_queue(self.samples(), QueueOptions(load_return=False, do_return=True, repeat_holder=False))
        reads = [cmd for cmd in queue if cmd.command_type == 'Meas']
        reads[0].sample_name = 'changed'
        self.assertEqual(reads[2].sample_name, 'S1')
        self.assertEqual(queue[-1].command_type, 'Goto')

    def test_periodic_holder_markers_cannot_select_a_specimen_slot_as_blank(self):
        queue = compile_queue(self.samples(), QueueOptions(load_return=False, do_return=False,
            repeat_holder=True, samples_between_holder=1))
        holders = [cmd for cmd in queue if cmd.command_type == 'Holder']
        self.assertGreaterEqual(len(holders), 5)
        self.assertTrue(all(cmd.hole == 0 for cmd in holders))
