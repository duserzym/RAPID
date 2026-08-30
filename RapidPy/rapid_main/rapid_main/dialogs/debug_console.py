from __future__ import annotations

import datetime
from collections.abc import Callable, Sequence

from PySide6 import QtCore, QtGui, QtWidgets
from rapidpy_common.ui import clamp_window_geometry

from rapid_main.diagnostic_services import DiagnosticStatusLine


class DebugConsoleDialog(QtWidgets.QDialog):
    """Runtime log viewer — replaces VB6 frmDebug."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        *,
        snapshot_provider: Callable[[], Sequence[DiagnosticStatusLine]] | None = None,
    ) -> None:
        super().__init__(parent)
        self._snapshot_provider = snapshot_provider
        self.setWindowTitle("Debug Console")
        self.resize(740, 460)
        self.setWindowFlags(
            self.windowFlags()
            & ~QtCore.Qt.WindowContextHelpButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
        )
        self._build_ui()

    def showEvent(self, event: QtCore.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            if handle.screen() is not None:
                handle.screen().availableGeometryChanged.connect(self._fit_to_screen)
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    def _fit_to_screen(self, screen: QtCore.QObject | None = None) -> None:
        active_screen = (
            screen
            if isinstance(screen, QtGui.QScreen)
            else (self.screen() or QtWidgets.QApplication.primaryScreen())
        )
        if active_screen is None:
            return
        available = active_screen.availableGeometry()
        max_w, max_h = clamp_window_geometry(available, (self.width(), self.height()))
        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), max_w), min(min_size.height(), max_h))
        self.resize(min(self.width(), max_w), min(self.height(), max_h))
        frame = self.frameGeometry()
        frame.setSize(
            QtCore.QSize(
                min(frame.width(), max_w),
                min(frame.height(), max_h),
            )
        )
        if frame.width() > available.width() or frame.height() > available.height():
            frame.moveCenter(available.center())
        else:
            new_x = max(
                available.left(),
                min(frame.left(), available.right() - frame.width() + 1),
            )
            new_y = max(
                available.top(),
                min(frame.top(), available.bottom() - frame.height() + 1),
            )
            frame.moveTopLeft(QtCore.QPoint(new_x, new_y))
        self.setGeometry(frame)

    # ── Public API ─────────────────────────────────────────────────────────
    def append(self, level: str, text: str) -> None:
        """Append a log line.  level: 'DEBUG' | 'INFO' | 'WARNING' | 'ERROR'"""
        chk = self._level_checks.get(level.upper())
        if chk and not chk.isChecked():
            return
        ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        colors = {
            "DEBUG": "#6b7280",
            "INFO": "#1d4ed8",
            "WARNING": "#b45309",
            "ERROR": "#b91c1c",
        }
        color = colors.get(level.upper(), "#2f2827")
        html = (
            f'<span style="color:#9a8885">{ts}</span> '
            f'<b style="color:{color}">[{level.upper():7s}]</b> '
            f'<span style="color:#2f2827">{text}</span>'
        )
        self._console.appendHtml(html)

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(10, 10, 10, 10)
        vl.setSpacing(6)

        # Toolbar row
        tb = QtWidgets.QHBoxLayout()
        tb.setSpacing(6)

        level_lbl = QtWidgets.QLabel("Show:")
        level_lbl.setStyleSheet("color: #9a8885; font-size: 11px;")
        tb.addWidget(level_lbl)

        self._level_checks: dict[str, QtWidgets.QCheckBox] = {}
        for level in ("DEBUG", "INFO", "WARNING", "ERROR"):
            chk = QtWidgets.QCheckBox(level.capitalize())
            chk.setChecked(True)
            self._level_checks[level] = chk
            tb.addWidget(chk)

        tb.addStretch()

        clear_btn = QtWidgets.QPushButton("Clear")
        clear_btn.clicked.connect(self._console.clear if hasattr(self, "_console") else lambda: None)
        tb.addWidget(clear_btn)

        copy_btn = QtWidgets.QPushButton("Copy All")
        copy_btn.clicked.connect(self._copy_all)
        tb.addWidget(copy_btn)

        refresh_btn = QtWidgets.QPushButton("Refresh Snapshot")
        refresh_btn.clicked.connect(self._refresh_snapshot)
        tb.addWidget(refresh_btn)

        vl.addLayout(tb)

        # Console
        self._console = QtWidgets.QTextEdit()
        self._console.setObjectName("console")
        self._console.setReadOnly(True)
        self._console.setFont(_mono_font())
        vl.addWidget(self._console, 1)

        # Wire clear after console is created
        clear_btn.clicked.disconnect()
        clear_btn.clicked.connect(self._console.clear)

        # Buttons
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

        # Seed with a startup message
        self.append("INFO", "Debug console opened.")
        self._refresh_snapshot()

    def _copy_all(self) -> None:
        QtWidgets.QApplication.clipboard().setText(self._console.toPlainText())

    def _refresh_snapshot(self) -> None:
        if self._snapshot_provider is None:
            self.append("DEBUG", "No diagnostic snapshot provider is registered.")
            return
        try:
            lines = list(self._snapshot_provider())
        except Exception as exc:
            self.append("ERROR", f"Diagnostic snapshot failed: {exc}")
            return
        if not lines:
            self.append("WARNING", "Diagnostic snapshot returned no backend status lines.")
            return
        self.append("INFO", "Diagnostic snapshot:")
        for line in lines:
            self.append(line.level, line.format_for_console())


def _mono_font() -> "QtGui.QFont":
    f = QtGui.QFont("Courier New")
    f.setPointSize(10)
    return f
