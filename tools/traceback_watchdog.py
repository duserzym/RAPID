"""Owned diagnostic snapshots without asynchronous native frame traversal."""
import sys
import threading
import traceback


class StackWatchdog:
    def __init__(self, interval=60., *, stream=None):
        if interval <= 0:
            raise ValueError('watchdog interval must be positive')
        self.interval = interval
        self.stream = sys.stderr if stream is None else stream
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name='suite-stack-watchdog', daemon=True)
        self._failure = None

    def start(self):
        self._thread.start()

    def stop(self):
        self._stop.set()
        self._thread.join()
        if self._failure is not None:
            raise RuntimeError('test stack watchdog failed') from self._failure

    def _snapshot(self):
        # sys._current_frames obtains strong references under Python's runtime
        # synchronization. Never inspect Python frames from a native timer thread.
        frames = sys._current_frames()
        frame = None
        try:
            output = [f'\nTest watchdog: thread snapshot after {self.interval:g}s\n']
            for identity, frame in sorted(frames.items()):
                output.append(f'Thread {identity}:\n')
                output.extend(traceback.format_stack(frame))
            self.stream.write(''.join(output))
            self.stream.flush()
        finally:
            # Captured worker frames can own fixtures/QObjects; do not retain
            # them between snapshots or into application/interpreter shutdown.
            frame = None
            frames.clear()

    def _run(self):
        try:
            while not self._stop.wait(self.interval):
                self._snapshot()
        except Exception as exc:
            self._failure = exc
