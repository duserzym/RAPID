import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.panels.sample_queue import SampleQueuePanel, _read_queue_rows, read_queue_file, write_queue_file
from rapid_main.queue_compiler import QueueSample, QueueOptions, compile_queue, validate_queue_samples


class QueueAveragingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def panel(self):
        panel = SampleQueuePanel()
        self.addCleanup(panel.deleteLater)
        return panel

    def test_counts_are_positive_vb6_integers_and_one_file_must_agree(self):
        for value in (0, -1, 32768, True, 2.5, '2', None):
            sample = QueueSample('A', 'F', 1, avg_steps=value)
            with self.subTest(value=value):
                self.assertFalse(validate_queue_samples([sample]).is_valid())
                with self.assertRaises(ValueError):
                    compile_queue([sample], QueueOptions(), strict=True)
        self.assertTrue(validate_queue_samples([QueueSample('A', 'F', 1, avg_steps=32767)]).is_valid())
        samples = [QueueSample('A', 'F', 1, avg_steps=2), QueueSample('B', 'F', 2, avg_steps=3)]
        with self.assertRaisesRegex(ValueError, 'AvgSteps'):
            compile_queue(samples, QueueOptions(), strict=True)

    def test_meas_keeps_file_count_on_both_sides_and_holder_uses_queue_maximum(self):
        samples = [QueueSample('A', 'F1', 1, do_both=True, avg_steps=2),
                   QueueSample('B', 'F2', 2, avg_steps=5)]
        commands = compile_queue(samples, QueueOptions(samples_between_holder=1), strict=True)
        self.assertEqual([(command.sample_name, command.avg_steps) for command in commands
                          if command.command_type == 'Meas'], [('A', 2), ('B', 5), ('A', 2)])
        holders = [command.avg_steps for command in commands if command.command_type == 'Holder']
        self.assertGreaterEqual(len(holders), 4)
        self.assertEqual(set(holders), {5})

    def test_old_rows_default_to_one_and_new_counts_survive_json_csv_and_session_restore(self):
        panel, restored = self.panel(), self.panel()
        panel.load_rows([dict(position='A1', sample_name='OLD', sample_set='Old', treatment='NRM', status='Pending')])
        self.assertEqual(_read_queue_rows(panel._table)[0][0].avg_steps, 1)
        panel.add_sample(position='A2', name='NEW', sample_set='New', avg_steps=4)
        _read_queue_rows(panel._table)
        rows = panel.row_snapshot()
        with tempfile.TemporaryDirectory() as directory:
            for suffix in ('.json', '.csv'):
                path = write_queue_file(Path(directory) / ('queue' + suffix), rows)
                loaded = read_queue_file(path)
                self.assertEqual(loaded, rows)
                restored.load_rows(loaded)
                samples, errors = _read_queue_rows(restored._table)
                self.assertEqual(errors, [])
                self.assertEqual([sample.avg_steps for sample in samples], [1, 4])

    def test_invalid_saved_count_is_rejected_during_read(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'queue.json'
            for value in (True, 0, -1, 1.5, None, '', '01', '32768', '2garbage'):
                path.write_text(json.dumps(dict(schema='rapidpy.sample_queue.v1', rows=[
                    dict(position='A1', sample_name='A', avg_steps=value)])))
                with self.subTest(value=value), self.assertRaises(ValueError):
                    read_queue_file(path)

    def test_edited_invalid_count_is_not_silently_clamped(self):
        panel = self.panel()
        panel.add_sample(position='A1', name='A')
        for text in ('0', '-1', '', 'True', '1.5', '32768'):
            panel._table.item(0, 9).setText(text)
            samples, errors = _read_queue_rows(panel._table)
            with self.subTest(text=text):
                self.assertEqual(samples, [])
                self.assertIn('AvgSteps', errors[0])
        with self.assertRaises(ValueError):
            panel.add_sample(position='A2', name='B', avg_steps=True)

    def test_file_settings_apply_count_to_same_index_only(self):
        panel = self.panel()
        with tempfile.TemporaryDirectory() as directory:
            for position, name, source in (('A1', 'A', 'one.sam'), ('A2', 'B', 'one.sam'), ('A3', 'C', 'two.sam')):
                panel.add_sample(position=position, name=name, sample_set='Same label',
                                 source_file=str(Path(directory) / source))
            def accept(dialog):
                dialog.findChild(QtWidgets.QSpinBox).setValue(3)
                return QtWidgets.QDialog.Accepted
            with patch.object(QtWidgets.QDialog, 'exec', accept):
                self.assertTrue(panel.edit_file_settings(0))
            samples, errors = _read_queue_rows(panel._table)
            self.assertEqual(errors, [])
            self.assertEqual([sample.avg_steps for sample in samples], [3, 3, 1])


if __name__ == '__main__':
    unittest.main()
