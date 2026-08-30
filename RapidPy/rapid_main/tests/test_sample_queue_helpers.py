from __future__ import annotations

import unittest

from PySide6 import QtWidgets

from rapid_main.panels.sample_queue import (
    SampleQueuePanel,
    _normalize_status,
    _parse_hole,
    _parse_step_count,
    _read_queue_rows,
    _status_for_recovery_action,
)


class TestSampleQueueHelpers(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
