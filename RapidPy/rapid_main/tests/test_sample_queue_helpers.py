from __future__ import annotations

import unittest
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

from PySide6 import QtWidgets

from rapid_main.panels.sample_queue import (
    SampleQueuePanel,
    _normalize_status,
    _parse_hole,
    _parse_step_count,
    _read_queue_rows,
    read_queue_file,
    _status_for_recovery_action,
    write_queue_file,
)
from rapid_main.queue_compiler import compile_queue, QueueOptions


class TestSampleQueueHelpers(unittest.TestCase):
    def test_file_settings_update_all_rows_of_one_index_without_touching_other_indexes(self):
        panel = SampleQueuePanel()
        try:
            with tempfile.TemporaryDirectory() as directory:
                first, second = str(Path(directory) / 'one.sam'), str(Path(directory) / 'two.sam')
                panel.add_sample(position='A1', name='A', sample_set='Same Unit', source_file=first)
                panel.add_sample(position='A2', name='B', sample_set='Same Unit', source_file=first)
                panel.add_sample(position='A3', name='C', sample_set='Same Unit', source_file=second)
                def accept(dialog):
                    dialog.findChild(QtWidgets.QComboBox).setCurrentText('Down')
                    dialog.findChild(QtWidgets.QCheckBox).setChecked(True)
                    return QtWidgets.QDialog.Accepted
                with patch.object(QtWidgets.QDialog, 'exec', accept):
                    self.assertTrue(panel.edit_file_settings(0))
                samples, errors = _read_queue_rows(panel._table)
                self.assertEqual(errors, [])
                self.assertEqual([(item.do_up, item.do_both) for item in samples], [(False, True), (False, True), (True, False)])
        finally:
            panel.deleteLater()
    def test_index_identity_and_orientation_survive_json_csv_and_panel_restore(self):
        panel = SampleQueuePanel()
        restored = SampleQueuePanel()
        try:
            with tempfile.TemporaryDirectory() as directory:
                index = Path(directory) / 'index.sam'
                panel.add_sample(position='A1', name='A', sample_set='Display Unit', treatment='NRM',
                    source_file=str(index), do_up=False, do_both=True)
                _read_queue_rows(panel._table)
                rows = panel.row_snapshot()
                for suffix in ('.json', '.csv'):
                    path = write_queue_file(Path(directory) / ('queue' + suffix), rows)
                    self.assertEqual(read_queue_file(path), rows)
                    restored.load_rows(read_queue_file(path))
                    samples, errors = _read_queue_rows(restored._table)
                    self.assertEqual(errors, [])
                    self.assertEqual(samples[0].file_id, os.path.normcase(samples[0].source_file))
                    self.assertNotEqual(samples[0].file_id, 'Display Unit')
                    self.assertFalse(samples[0].do_up)
                    self.assertTrue(samples[0].do_both)
                    self.assertEqual(restored.row_snapshot(), rows)
        finally:
            panel.deleteLater()
            restored.deleteLater()

    def test_distinct_indexes_with_same_display_name_keep_independent_file_flags(self):
        panel = SampleQueuePanel()
        try:
            with tempfile.TemporaryDirectory() as directory:
                panel.add_sample(position='A1', name='A', sample_set='Same Unit', source_file=str(Path(directory) / 'one.sam'), do_both=True)
                panel.add_sample(position='A2', name='B', sample_set='Same Unit', source_file=str(Path(directory) / 'two.sam'), do_up=False)
                samples, errors = _read_queue_rows(panel._table)
                self.assertEqual(errors, [])
                commands = compile_queue(samples, QueueOptions(), strict=True)
                self.assertNotEqual(samples[0].file_id, samples[1].file_id)
                self.assertEqual([item.file_id for item in commands if item.command_type == 'Flip'], [samples[0].file_id])
        finally:
            panel.deleteLater()

    def test_invalid_orientation_or_relative_source_blocks_row_admission(self):
        panel = SampleQueuePanel()
        try:
            panel.add_sample(position='A1', name='A')
            for column, text in ((7, 'Maybe'), (8, 'True'), (6, 'relative.sam')):
                panel._set_file_settings(0)
                panel._table.item(0, column).setText(text)
                samples, errors = _read_queue_rows(panel._table)
                self.assertEqual(samples, [])
                self.assertTrue(errors)
            with self.assertRaises(ValueError):
                panel.add_sample(position='A2', name='B', do_both=1)
        finally:
            panel.deleteLater()

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

    def test_parse_hole_accepts_standard_position_formats(self) -> None:
        hole, err = _parse_hole("A1")
        self.assertEqual(hole, 1)
        self.assertIsNone(err)

        hole, err = _parse_hole("L-03")
        self.assertEqual(hole, 3)
        self.assertIsNone(err)

        hole, err = _parse_hole("7")
        self.assertEqual(hole, 7)
        self.assertIsNone(err)

    def test_parse_hole_rejects_invalid_position(self) -> None:
        hole, err = _parse_hole("")
        self.assertIsNone(hole)
        self.assertEqual(err, "position is empty")

        hole, err = _parse_hole("NoHole")
        self.assertIsNone(hole)
        self.assertIn("does not contain a valid hole index", err)

    def test_parse_step_count_detects_arrows(self) -> None:
        self.assertEqual(_parse_step_count("NRM"), 1)
        self.assertEqual(_parse_step_count("NRM -> 25mT AF -> 50mT AF"), 3)
        self.assertEqual(_parse_step_count("SAMP1 -> SAMP2 -> SAMP3"), 3)

    def test_row_treatment_labels_are_executable_steps_and_empty_steps_fail(self):
        panel = SampleQueuePanel()
        try:
            panel.add_sample(position='A1', name='A', sample_set='F1', treatment='NRM -> AF20 → SUSC')
            samples, errors = _read_queue_rows(panel._table)
            self.assertEqual(errors, [])
            self.assertEqual(samples[0].measurement_labels, ('NRM', 'AF20', 'SUSC'))
            self.assertEqual(samples[0].measurement_step_count, 3)
            panel._table.item(0, 4).setText('NRM -> -> AF20')
            samples, errors = _read_queue_rows(panel._table)
            self.assertEqual(samples, [])
            self.assertIn('empty step', errors[0])
        finally:
            panel.deleteLater()

    def test_normalize_status(self) -> None:
        self.assertEqual(_normalize_status("done"), "Done")
        self.assertEqual(_normalize_status("pending"), "Pending")
        self.assertEqual(_normalize_status("skipped"), "Skipped")
        self.assertEqual(_normalize_status("error"), "Error")
        self.assertEqual(_normalize_status("running"), "Running")
        self.assertEqual(_normalize_status("interrupted"), "Interrupted")
        self.assertEqual(_normalize_status("resume"), "Resume")
        self.assertEqual(_normalize_status("rerun"), "Rerun")
        self.assertEqual(_normalize_status("aborted"), "Aborted")
        self.assertEqual(_normalize_status("Queued"), "Pending")
        self.assertEqual(_normalize_status(""), "Pending")

    def test_interrupted_recovery_actions_map_to_statuses(self) -> None:
        self.assertEqual(_status_for_recovery_action("resume"), "Resume")
        self.assertEqual(_status_for_recovery_action("re-run"), "Rerun")
        self.assertEqual(_status_for_recovery_action("skip"), "Skipped")
        self.assertEqual(_status_for_recovery_action("abort"), "Aborted")

    def test_interrupted_rows_require_explicit_recovery_choice(self) -> None:
        panel = SampleQueuePanel()
        try:
            panel.load_rows(
                [
                    {
                        "position": "A1",
                        "sample_name": "HBK-INT",
                        "sample_set": "set-A",
                        "treatment": "NRM",
                        "status": "Interrupted",
                    }
                ]
            )

            rows, errors = _read_queue_rows(panel._table)
            self.assertEqual(rows, [])
            self.assertTrue(any("requires Resume" in error for error in errors))

            self.assertEqual(panel.interrupted_sample_names(), ["HBK-INT"])
            self.assertEqual(panel.apply_interrupted_recovery("resume"), 1)
            rows, errors = _read_queue_rows(panel._table)
            self.assertEqual(errors, [])
            self.assertEqual(len(rows), 1)
        finally:
            panel.deleteLater()

    def test_sequential_positions_preserve_samples_and_treatments(self) -> None:
        panel = SampleQueuePanel()
        try:
            panel.load_rows(
                [
                    {
                        "position": "A1",
                        "sample_name": "SPEC-1",
                        "sample_set": "SITE-A",
                        "treatment": "NRM → AF20",
                        "status": "Pending",
                    },
                    {
                        "position": "A2",
                        "sample_name": "SPEC-2",
                        "sample_set": "SITE-A",
                        "treatment": "NRM",
                        "status": "Pending",
                    },
                ]
            )
            names_before = [panel._safe_cell(row, 2) for row in range(panel._table.rowCount())]
            treatments_before = [panel._safe_cell(row, 4) for row in range(panel._table.rowCount())]

            panel.apply_sequential_positions(start=7, prefix="C")

            self.assertEqual(
                [panel._safe_cell(row, 1) for row in range(panel._table.rowCount())],
                ["C7", "C8"],
            )
            self.assertEqual(
                [panel._safe_cell(row, 2) for row in range(panel._table.rowCount())],
                names_before,
            )
            self.assertEqual(
                [panel._safe_cell(row, 4) for row in range(panel._table.rowCount())],
                treatments_before,
            )
        finally:
            panel.deleteLater()

    def test_new_and_persisted_empty_queues_remain_empty(self) -> None:
        panel = SampleQueuePanel()
        try:
            self.assertEqual(panel._table.rowCount(), 0)
            self.assertEqual(panel._count_lbl.text(), "0 samples")

            panel.load_rows([])

            self.assertEqual(panel.row_snapshot(), [])
            self.assertEqual(panel._count_lbl.text(), "0 samples")
        finally:
            panel.deleteLater()

    def test_add_sample_validates_and_appends_operator_selection(self) -> None:
        panel = SampleQueuePanel()
        try:
            row = panel.add_sample(
                position=" A7 ",
                name=" SPEC-7 ",
                sample_set=" Unit A ",
                treatment="NRM → AF20",
            )

            self.assertEqual(row, 0)
            self.assertEqual(
                panel.row_snapshot(),
                [
                    {
                        "position": "A7",
                        "sample_name": "SPEC-7",
                        "sample_set": "Unit A",
                        "treatment": "NRM → AF20",
                        "status": "Pending",
                    }
                ],
            )
            with self.assertRaisesRegex(ValueError, "required"):
                panel.add_sample(position="", name="SPEC-8")
        finally:
            panel.deleteLater()

    def test_load_index_button_requests_real_index_workflow(self) -> None:
        panel = SampleQueuePanel()
        try:
            emitted: list[bool] = []
            panel.sample_index_requested.connect(lambda: emitted.append(True))

            panel._load_index_btn.click()

            self.assertEqual(emitted, [True])
        finally:
            panel.deleteLater()

    def test_queue_json_and_csv_round_trip(self) -> None:
        rows = [
            {
                "position": "A1",
                "sample_name": "SPEC-1",
                "sample_set": "SITE-A",
                "treatment": "NRM → AF20",
                "status": "Pending",
            },
            {
                "position": "A2",
                "sample_name": "SPEC-2",
                "sample_set": "SITE-A",
                "treatment": "NRM",
                "status": "Done",
            },
        ]
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for suffix in (".json", ".csv"):
                with self.subTest(suffix=suffix):
                    path = write_queue_file(root / f"queue{suffix}", rows)
                    self.assertEqual(read_queue_file(path), rows)
                    self.assertFalse(path.with_name(path.name + ".tmp").exists())

    def test_queue_import_rejects_unknown_schema_before_table_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "queue.json"
            path.write_text('{"schema":"unknown","rows":[]}', encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Unsupported queue file schema"):
                read_queue_file(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
