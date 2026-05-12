from __future__ import annotations

import datetime

from PySide6 import QtCore, QtWidgets


class DebugConsoleDialog(QtWidgets.QDialog):
    """Runtime log viewer — replaces VB6 frmDebug."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Debug Console")
        self.resize(740, 460)
        self.setWindowFlags(
            self.windowFlags()
            & ~QtCore.Qt.WindowContextHelpButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
        )
        self._build_ui()

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

    def _copy_all(self) -> None:
        QtWidgets.QApplication.clipboard().setText(self._console.toPlainText())


def _mono_font() -> "QtCore.QFont":
    from PySide6 import QtGui
    f = QtGui.QFont("Courier New")
    f.setPointSize(10)
    return f
