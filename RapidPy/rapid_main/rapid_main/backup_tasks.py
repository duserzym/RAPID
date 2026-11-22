"""Deliver backups in a separate process without holding station resources."""
from collections import deque
import json
from pathlib import Path
import sys

from PySide6 import QtCore

from rapid_main.config import AppConfig
from rapid_main.io.backup_jobs import BackupJob, BackupJobs
from rapid_main.io.bundle_backup import INDEX_FILE
from rapid_main.io.bundle_publication import _digest


def _backup_command() -> tuple[str, ...]:
    if getattr(sys, "frozen", False):
        return (str(Path(sys.executable).with_name("RapidPyMainConsole.exe")),)
    launcher = Path(__file__).resolve().parents[1] / "main.py"
    return (sys.executable, str(launcher)) if launcher.is_file() else (sys.executable, "-m", "rapid_main")


class MeasurementBackups(QtCore.QObject):
    status_changed = QtCore.Signal(str, bool)
    pending_changed = QtCore.Signal(int)

    def __init__(self, parent=None, *, spool: Path | None = None):
        super().__init__(parent)
        self._jobs = BackupJobs(spool if spool is not None else AppConfig.default_path().with_name("measurement-backup-queue"))
        self._pending: deque[BackupJob] = deque()
        self._active: BackupJob | None = None
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
            job = BackupJob(sample, str(source.absolute()), destination, index_digest)
            self._jobs.save(job)
        except (OSError, ValueError) as exc:
            self.status_changed.emit(f"Backup failed for {sample}. {exc} The primary measurement is saved.", True)
            return
        self.refresh_pending()
        self._enqueue([job])

    def _read_pending(self) -> list[BackupJob]:
        try:
            jobs, unreadable = self._jobs.pending()
        except (OSError, ValueError) as exc:
            self.status_changed.emit(f"Pending measurement backups could not be read. {exc}", True)
            return []
        self.pending_changed.emit(len(jobs) + len(unreadable))
        if unreadable:
            self.status_changed.emit(f"{len(unreadable)} pending backup records need review in {self._jobs.directory}.", True)
        return jobs

    def refresh_pending(self) -> None:
        self._read_pending()

    def retry_pending(self, destination: str) -> None:
        jobs = self._read_pending()
        eligible = [job for job in jobs if job.matches_destination(destination)]
        if eligible:
            self._enqueue(eligible)
        elif jobs:
            self.status_changed.emit("Set the original backup folder in Settings to retry its pending backups.", True)

    def _enqueue(self, jobs: list[BackupJob]) -> None:
        active = {job.key for job in self._pending}
        if self._active is not None:
            active.add(self._active.key)
        self._pending.extend(job for job in jobs if job.key not in active)
        self._start_next()

    def _start_next(self) -> None:
        if self._active is not None or not self._pending:
            return
        job = self._active = self._pending.popleft()
        self._timed_out = False
        self.status_changed.emit(f"Backing up {job.sample}…", False)
        program, *prefix = _backup_command()
        self._timer.start(120_000)
        self._process.start(program, [*prefix, "--backup-measurement", job.source, job.destination,
                                      "--backup-index-sha256", job.index_sha256])

    def _timeout(self) -> None:
        self._timed_out = True
        self._process.kill()

    def _failed_to_start(self, error) -> None:
        if error == QtCore.QProcess.ProcessError.FailedToStart:
            self._report_failure(self._process.errorString())

    def _finished(self, exit_code, exit_status) -> None:
        job = self._active
        if job is None:
            return
        if self._timed_out:
            self._report_failure("The backup process exceeded two minutes.")
        elif exit_code != 0 or exit_status != QtCore.QProcess.ExitStatus.NormalExit:
            error = bytes(self._process.readAllStandardError()).decode(errors="replace").strip()
            self._report_failure(error or "The backup process stopped before verification.")
        else:
            try:
                result = json.loads(bytes(self._process.readAllStandardOutput()))
                if result["index_sha256"] != job.index_sha256:
                    raise ValueError("The backup process returned a different artifact index.")
                message = f"Backup saved for {job.sample} at {result['archive']}"
                if result["omitted"]:
                    message += f". Unverified files omitted: {', '.join(result['omitted'])}"
            except (ValueError, KeyError, TypeError):
                self._report_failure("The backup process returned no verification result.")
                return
            try:
                self._jobs.complete(job)
            except (OSError, ValueError) as exc:
                self._report(f"{message}. The pending record could not be removed. {exc}", True)
                return
            self._report(message, bool(result["omitted"]))

    def _report_failure(self, message: str) -> None:
        if self._active is not None:
            self._report(f"Backup failed for {self._active.sample}. {message} The primary measurement is saved. "
                         "The backup remains pending.", True)

    def _report(self, message: str, warning: bool) -> None:
        self._timer.stop()
        self._active = None
        self.status_changed.emit(message, warning)
        self.refresh_pending()
        QtCore.QTimer.singleShot(0, self._start_next)
