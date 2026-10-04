"""Queue UI completion targets the captured row, not a duplicate specimen name."""
import unittest
from types import SimpleNamespace
from PySide6 import QtWidgets

from rapid_main.app import MainWindow
from rapid_main.panels.sample_queue import SampleQueuePanel, _read_queue_rows
from rapid_main.queue_compiler import QueueSample, QueueOptions, compile_queue


class QueueRowIdentityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_duplicate_names_and_second_side_status_follow_persisted_row_identity(self):
        panel, restored = SampleQueuePanel(), SampleQueuePanel()
        try:
            panel.add_sample(position='A1', name='SAME', sample_set='First', do_both=True)
            panel.add_sample(position='A2', name='SAME', sample_set='Second')
            samples, errors = _read_queue_rows(panel._table)
            self.assertEqual(errors, [])
            commands = compile_queue(samples, QueueOptions(), strict=True)
            measurements = [command for command in commands if command.command_type == 'Meas']
            self.assertEqual(measurements[0].row_id, measurements[2].row_id)
            self.assertNotEqual(measurements[0].row_id, measurements[1].row_id)
            restored.load_rows(panel.row_snapshot())
            self.assertTrue(restored.set_queue_command_status(measurements[1], 'Done'))
            self.assertEqual([restored._safe_cell(row, 5) for row in (0, 1)], ['Pending', 'Done'])
            restored.apply_sequential_positions(10)
            self.assertTrue(restored.set_queue_command_status(measurements[2], 'Running'))
            self.assertEqual([restored._safe_cell(row, 5) for row in (0, 1)], ['Running', 'Done'])
        finally:
            panel.deleteLater()
            restored.deleteLater()

    def test_duplicate_or_malformed_persisted_row_identity_is_rejected(self):
        for row_id in ('bad', 'a' * 32):
            samples = [QueueSample('A', 'F1', 1, row_id=row_id), QueueSample('B', 'F2', 2, row_id=row_id)]
            with self.subTest(row_id=row_id), self.assertRaisesRegex(ValueError, 'row_id'):
                compile_queue(samples, QueueOptions(), strict=True)

    def test_wrong_sample_completion_does_not_advance_original_queue(self):
        commands = compile_queue([QueueSample('A', 'F1', 1)], QueueOptions())
        command = next(item for item in commands if item.command_type == 'Meas')
        statuses, finalizations, advances = [], [], []
        window = SimpleNamespace(_queue_active=True, _queue_current_command=command,
            _set_queue_sample_status=lambda name, status: statuses.append((name, status)),
            _finalize_queue_run=lambda state, **kwargs: finalizations.append(state),
            _run_next_queue_command=lambda: advances.append(True))
        MainWindow._on_queue_sample_finished(window, False, 'OTHER')
        self.assertEqual(statuses, [('A', 'Error')])
        self.assertEqual(finalizations, ['error'])
        self.assertEqual(advances, [])
