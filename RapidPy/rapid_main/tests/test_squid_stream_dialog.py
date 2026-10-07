"""SQUID dialog read-rate settings and the live stream / spectrum dialog."""
from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace

from PySide6 import QtWidgets

from rapid_main.config import SquidConfig
from rapid_main.diagnostic_services import SquidNoCommBackend
from rapid_main.dialogs.squid_comm import SquidCommDialog
from rapid_main.dialogs.squid_stream import SquidStreamDialog
from rapidpy_common.squid_stream import PositionSample, SquidReadTiming, SquidStreamSample, SquidTrace

from tests.test_squid_stream import FakeStreamClient

_APP = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def _pump(seconds: float) -> None:
    deadline = time.perf_counter() + seconds
    while time.perf_counter() < deadline:
        _APP.processEvents()
        time.sleep(0.01)


class _Owner(QtWidgets.QWidget):
    def __init__(self, cfg):
        super().__init__()
        self.config = SimpleNamespace(squid=cfg, save=lambda: None)


class SquidCommStreamSettingsTests(unittest.TestCase):
    def test_capture_settings_round_trip_only_on_save(self):
        cfg = SquidConfig(stream_interval_ms=250, stream_axes="Z", motion_capture_enabled=False)
        owner = _Owner(cfg)
        dialog = SquidCommDialog(owner, backend=SquidNoCommBackend(cfg))
        self.assertIn("every 250 ms", dialog._stream_summary.text())
        self.assertFalse(dialog._capture_turns.isEnabled())
        dialog._capture_enabled.setChecked(True)
        dialog._capture_turns.setChecked(False)
        dialog._stream_settings.stream_interval_ms = 80
        self.assertFalse(cfg.motion_capture_enabled)  # pending until saved
        dialog._accept_settings()
        self.assertTrue(cfg.motion_capture_enabled)
        self.assertFalse(cfg.motion_capture_turns)
        self.assertEqual(cfg.stream_interval_ms, 80)
        self.assertEqual(cfg.stream_axes, "Z")
        dialog.deleteLater()
        owner.deleteLater()


class SquidStreamDialogTests(unittest.TestCase):
    def test_simulated_stream_runs_at_operator_rate_and_stops_cleanly(self):
        settings = SimpleNamespace(baud=9600, stream_interval_ms=40, stream_axes="Z", stream_counts_every=1,
                                   stream_latch_count_hold_ms=0, stream_latch_data_hold_ms=0)
        dialog = SquidStreamDialog(squid_config=settings, simulated=True)
        self.assertEqual(dialog.timing().interval_s, 0.04)
        dialog._start_stream()
        _pump(0.6)
        dialog._stop_stream()
        trace = dialog._trace
        self.assertGreater(len(trace.samples), 5)
        self.assertLess(trace.achieved_rate_hz, 30.0)
        self.assertEqual(trace.label, "live-simulated")
        self.assertIn("Z:", dialog._analysis.toPlainText())
        self.assertTrue(dialog._save.isEnabled())
        dialog._store_config()
        self.assertEqual(settings.stream_interval_ms, 40)
        dialog.deleteLater()

    def test_real_mode_requires_a_client(self):
        dialog = SquidStreamDialog(client_provider=lambda: None, squid_config=None, simulated=False)
        dialog._start_stream()
        self.assertIsNone(dialog._capture)
        self.assertIn("Cannot start", dialog._status.text())
        dialog.deleteLater()

    def test_live_stream_uses_provided_client(self):
        client = FakeStreamClient()
        dialog = SquidStreamDialog(client_provider=lambda: client, squid_config=None, simulated=False)
        dialog._interval.setValue(0)
        dialog._count_hold.setValue(0)
        dialog._data_hold.setValue(0)
        dialog._start_stream()
        _pump(0.2)
        dialog.done(0)  # closing always releases the link
        self.assertIsNone(dialog._capture)
        self.assertTrue(any(call[0] == "latch" for call in client.calls))

    def test_opening_a_turn_capture_shows_rotation_fit(self):
        import math

        timing = SquidReadTiming(interval_s=0.0, axes=("X", "Y"))
        trace = SquidTrace("02-turn-90deg", timing, motion={"segment": "turn"})
        for i in range(40):
            theta = math.radians(90 * i / 39)
            trace.samples.append(
                SquidStreamSample(i, i * 0.05, 0.01, (math.cos(theta + 0.3), math.sin(theta + 0.3), math.nan),
                                  (0, 0, math.nan), (0, 0, math.nan), True)
            )
        trace.positions.extend([PositionSample(0.0, 0.0), PositionSample(39 * 0.05, 90.0)])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "turn.csv"
            path.write_text(trace.to_csv(), encoding="utf-8")
            dialog = SquidStreamDialog(squid_config=None, simulated=True)
            dialog.load_trace_file(path)
            text = dialog._analysis.toPlainText()
            self.assertIn("Rotation fit over 90°", text)
            self.assertIn("horizontal amplitude 1", text)
            dialog.deleteLater()


if __name__ == "__main__":
    unittest.main()
