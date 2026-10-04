"""Keep native queue commands and terminal recovery off the GUI thread."""
import threading
from PySide6 import QtCore


class QueueCommandWorker(QtCore.QObject):
    settled = QtCore.Signal()

    def __init__(self, backend, action, *, recover_on_error=True):
        super().__init__()
        self.backend, self.action = backend, action
        self.recover_on_error = recover_on_error
        self.stop_event = threading.Event()
        self.ok = False
        self.error = ''

    def stop(self):
        self.stop_event.set()

    @QtCore.Slot()
    def run(self):
        check = getattr(self.backend, 'set_halt_check', None)
        initialized = False
        try:
            if self.stop_event.is_set():
                raise InterruptedError('Queue command cancelled before initialization.')
            if callable(check):
                check(self.stop_event.is_set)
            if self.stop_event.is_set():
                raise InterruptedError('Queue command cancelled before initialization.')
            initialized = True
            self.action()
            if self.stop_event.is_set():
                raise InterruptedError('Queue command halted; terminal recovery is required.')
            self.ok = True
        except Exception as exc:
            self.error = str(exc) or type(exc).__name__
            if initialized and self.recover_on_error:
                try:
                    self.backend.return_to_safe_state()
                except Exception as cleanup:
                    self.error += '; recovery remains unverified: ' + str(cleanup)
        finally:
            try:
                if callable(check):
                    check(None)
            except Exception as exc:
                self.ok = False
                self.error += '; cancellation hook could not be cleared: ' + str(exc)
            finally:
                self.settled.emit()
