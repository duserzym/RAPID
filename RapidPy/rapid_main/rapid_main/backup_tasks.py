"""Deliver backups in a separate process without holding station resources."""
from collections import deque
import json
from pathlib import Path
import sys

from PySide6 import QtCore

from rapid_main.io.bundle_backup import INDEX_FILE
from rapid_main.io.bundle_publication import _digest


def _backup_command() -> tuple[str, ...]:
    if getattr(sys, "frozen", False):
        return (str(Path(sys.executable).with_name("RapidPyMainConsole.exe")),)
    launcher = Path(__file__).resolve().parents[1] / "main.py"
    return (sys.executable, str(launcher)) if launcher.is_file() else (sys.executable, "-m", "rapid_main")


class MeasurementBackups(QtCore.QObject):
    status_changed = QtCore.Signal(str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pending = deque()
        self._active_sample = None
        self._timed_out = False
        self._process = QtCore.QProcess(self)
        self._process.finished.connect(self._finished)
        self._process.errorOccurred.connect(self._failed_to_start)
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._timeout)

    def submit(self, sample: str, source: Path, destination: str) -> None:
        try:
            index_digest = _digest(source / INDEX_FILE)
            if index_digest is None:
                raise ValueError("The published artifact index is missing.")
        except (OSError, ValueError) as exc:
            self.status_changed.emit(f"Backup failed for {sample}. {exc}", True)
            return
        self._pending.append((sample, str(source), destination, index_digest))
        self._start_next()

    def _start_next(self) -> None:
        if self._active_sample is not None or not self._pending:
            return
        sample, source, destination, digest = self._pending.popleft()
        self._active_sample = sample
        self._timed_out = False
        self.status_changed.emit(f"Backing up {sample}…", False)
        program, *prefix = _backup_command()
        self._timer.start(120_000)
        self._process.start(program, [*prefix, "--backup-measurement", source, destination,
                                      "--backup-index-sha256", digest])

    def _timeout(self) -> None:
        self._timed_out = True
        self._process.kill()

    def _failed_to_start(self, error) -> None:
        if error == QtCore.QProcess.ProcessError.FailedToStart:
            self._report_failure(self._process.errorString())

    def _finished(self, exit_code, exit_status) -> None:
        if self._timed_out:
            self._report_failure("The backup process exceeded two minutes.")
        elif exit_code != 0 or exit_status != QtCore.QProcess.ExitStatus.NormalExit:
            error = bytes(self._process.readAllStandardError()).decode(errors="replace").strip()
            self._report_failure(error or "The backup process stopped before verification.")
        else:
            try:
                result = json.loads(bytes(self._process.readAllStandardOutput()))
                message = f"Backup saved for {self._active_sample} at {result['archive']}"
                if result["omitted"]:
                    message += f". Unverified files omitted: {', '.join(result['omitted'])}"
            except (ValueError, KeyError, TypeError):
                self._report_failure("The backup process returned no verification result.")
                return
            self._report(message, bool(result["omitted"]))

    def _report_failure(self, message: str) -> None:
        self._report(f"Backup failed for {self._active_sample}. {message} The primary measurement is saved.", True)

    def _report(self, message: str, warning: bool) -> None:
        self._timer.stop()
        self._active_sample = None
        self.status_changed.emit(message, warning)
        QtCore.QTimer.singleShot(0, self._start_next)
