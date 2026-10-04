import importlib.util
import io
from pathlib import Path
import threading
import unittest

spec = importlib.util.spec_from_file_location('suite_watchdog_under_test',
    Path(__file__).resolve().parents[3] / 'tools' / 'traceback_watchdog.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class WatchdogTests(unittest.TestCase):
    def test_snapshots_waiting_python_thread_and_stops_while_thread_remains_alive(self):
        snapshot_seen, release, started = threading.Event(), threading.Event(), threading.Event()
        class Stream(io.StringIO):
            def write(self, value):
                result = super().write(value)
                snapshot_seen.set()
                return result
        def owned_waiting_fixture():
            started.set()
            release.wait()
        fixture = threading.Thread(target=owned_waiting_fixture)
        stream = Stream()
        watchdog = module.StackWatchdog(.01, stream=stream)
        fixture.start()
        try:
            self.assertTrue(started.wait(2))
            watchdog.start()
            self.assertTrue(snapshot_seen.wait(2))
            watchdog.stop()
            self.assertFalse(watchdog._thread.is_alive())
            self.assertTrue(fixture.is_alive())
            self.assertIn('owned_waiting_fixture', stream.getvalue())
            self.assertIn(f'Thread {fixture.ident}:', stream.getvalue())
        finally:
            release.set()
            fixture.join()

    def test_stop_wakes_watchdog_without_waiting_for_diagnostic_interval(self):
        watchdog = module.StackWatchdog(60, stream=io.StringIO())
        watchdog.start()
        watchdog.stop()
        self.assertFalse(watchdog._thread.is_alive())

    def test_diagnostic_write_failure_is_reported_to_owner(self):
        class BrokenStream:
            def write(self, value):
                raise OSError('diagnostic destination unavailable')
        watchdog = module.StackWatchdog(.01, stream=BrokenStream())
        watchdog.start()
        watchdog._thread.join(2)
        self.assertFalse(watchdog._thread.is_alive())
        with self.assertRaisesRegex(RuntimeError, 'watchdog failed'):
            watchdog.stop()
