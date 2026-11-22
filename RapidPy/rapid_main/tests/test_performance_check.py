"""Check performance evidence without an operator profile or hardware."""
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

from rapid_main.performance_check import timer_summary


class PerformanceCheckTests(unittest.TestCase):
    def test_timer_summary_uses_nearest_rank_and_labels_an_empty_sample(self):
        self.assertEqual(timer_summary([]), dict(samples=0, p95_interval_ms=None, max_interval_ms=None))
        self.assertEqual(timer_summary(list(range(20, 0, -1))),
                         dict(samples=20, p95_interval_ms=19, max_interval_ms=20))

    def test_offline_cli_finishes_a_real_queue_in_an_isolated_profile(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            current_config = root / 'operator.json'
            current_safety = root / 'operator-safety.json'
            current_config.write_bytes(b'{"general":{"nocomm":false}}')
            current_safety.write_bytes(b'operator journal sentinel')
            ready = root / 'ready'
            env = {**os.environ, 'RAPID_CONFIG': str(current_config), 'RAPID_SAFETY_STATE': str(current_safety),
                   'RAPID_PERFORMANCE_READY_FILE': str(ready),
                   'PYTHONPATH': os.pathsep.join(path for path in sys.path if path)}
            result = subprocess.run([sys.executable, '-m', 'rapid_main', '--measure-performance',
                                     '--performance-specimens', '2', '--performance-idle-seconds', '5',
                                     '--performance-queue-timeout-seconds', '30'], env=env, capture_output=True,
                                    text=True, timeout=75)
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            report = json.loads(result.stdout)
            self.assertEqual(report['schema'], 'rapidpy.performance_check.v1')
            self.assertTrue(report['simulated'])
            self.assertTrue(report['simulated_vacuum_pump_on'])
            self.assertEqual(report['unexpected_dialogs'], [])
            self.assertGreater(report['entry_to_first_window_ms'], 0)
            self.assertGreaterEqual(report['idle']['elapsed_ms'], 5000)
            self.assertGreater(report['idle']['timer']['samples'], 0)
            self.assertGreater(report['queue']['timer']['samples'], 0)
            self.assertEqual(report['queue']['published_specimens'], 2)
            self.assertEqual(report['queue']['published_steps'], 4)
            self.assertTrue(report['queue']['queue_settled'])
            self.assertGreater(int(ready.read_text(encoding='ascii')), 0)
            self.assertEqual(current_config.read_bytes(), b'{"general":{"nocomm":false}}')
            self.assertEqual(current_safety.read_bytes(), b'operator journal sentinel')
