"""Never lose an error silently.

The packaged app has no console, so exceptions raised inside Qt slots and
native crashes used to vanish.  :func:`install_error_logging` records them:

* unhandled Python exceptions (any thread) -> ``logs/rapid_main_errors.log``
  with the full traceback, and a status-bar note pointing at the file;
* Qt warnings/criticals/fatals -> the same log;
* native crashes (access violation, abort) -> ``logs/rapid_main_crash.log``
  via :mod:`faulthandler`, so the next run can show what happened.
"""

from __future__ import annotations

import faulthandler
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable

from PySide6 import QtCore

_STATE: dict = {"installed": False, "crash_file": None, "notify": None}


def log_directory() -> Path:
    from .config import AppConfig

    return AppConfig.default_path().parent / "logs"


def _append(path: Path, text: str) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(text)
    except OSError:
        pass


def record_exception(exc_type, exc, tb, *, where: str = "") -> Path:
    path = log_directory() / "rapid_main_errors.log"
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    header = f"\n=== {stamp} unhandled {exc_type.__name__}{' in ' + where if where else ''} ===\n"
    _append(path, header + "".join(traceback.format_exception(exc_type, exc, tb)))
    notify: Callable[[str], None] | None = _STATE.get("notify")
    if notify is not None:
        try:
            notify(f"Unexpected error: {exc_type.__name__}: {exc} — details saved to {path}")
        except Exception:  # noqa: BLE001 - reporting must never raise
            pass
    return path


def _qt_message(mode: QtCore.QtMsgType, context: QtCore.QMessageLogContext, message: str) -> None:
    if mode == QtCore.QtMsgType.QtDebugMsg or mode == QtCore.QtMsgType.QtInfoMsg:
        return
    if "QWindowsWindow::setGeometry" in message:  # benign multi-monitor resize chatter
        return
    level = {QtCore.QtMsgType.QtWarningMsg: "warning", QtCore.QtMsgType.QtCriticalMsg: "critical",
             QtCore.QtMsgType.QtFatalMsg: "FATAL"}.get(mode, "message")
    _append(log_directory() / "rapid_main_errors.log",
            f"{datetime.now():%Y-%m-%d %H:%M:%S} Qt {level}: {message}\n")


def previous_crash_summary() -> str:
    """Text of a native crash left by the last run (then cleared), or ''."""
    path = log_directory() / "rapid_main_crash.log"
    try:
        text = path.read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return ""
    if not text:
        return ""
    archive = path.with_name(f"rapid_main_crash_{datetime.now():%Y%m%d_%H%M%S}.log")
    try:
        path.replace(archive)
    except OSError:
        pass
    return text


def install_error_logging(notify: Callable[[str], None] | None = None) -> Path:
    """Install the hooks once; ``notify`` receives a one-line operator message."""
    _STATE["notify"] = notify
    directory = log_directory()
    if _STATE["installed"]:
        return directory
    directory.mkdir(parents=True, exist_ok=True)
    try:
        crash_file = open(directory / "rapid_main_crash.log", "a", encoding="utf-8")
        faulthandler.enable(file=crash_file, all_threads=True)
        _STATE["crash_file"] = crash_file
    except OSError:
        pass
    sys.excepthook = lambda t, e, tb: record_exception(t, e, tb)
    threading.excepthook = lambda args: record_exception(
        args.exc_type, args.exc_value, args.exc_traceback, where=f"thread {getattr(args.thread, 'name', '?')}"
    )
    QtCore.qInstallMessageHandler(_qt_message)
    _STATE["installed"] = True
    return directory
