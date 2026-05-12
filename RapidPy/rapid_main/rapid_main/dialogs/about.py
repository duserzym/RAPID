from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets


class AboutDialog(QtWidgets.QDialog):
    """About box — replaces VB6 frmAbout."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("About RAPID v4")
        self.setMinimumWidth(440)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setSpacing(0)
        vl.setContentsMargins(0, 0, 0, 0)

        # ── Header stripe ──────────────────────────────────────────────────
        header = QtWidgets.QFrame()
        header.setFixedHeight(96)
        header.setStyleSheet(
            "background: qlineargradient(x1:0,y1:0,x2:1,y2:1,"
            "stop:0 #4a0111, stop:1 #7A0219); border: none;"
        )
        hl = QtWidgets.QVBoxLayout(header)
        hl.setContentsMargins(24, 14, 24, 14)

        title_lbl = QtWidgets.QLabel("RAPID v4")
        title_lbl.setStyleSheet(
            "font-size: 28px; font-weight: 800; color: #ffffff; background: transparent;"
        )
        sub_lbl = QtWidgets.QLabel("Paleomagnetics Control System")
        sub_lbl.setStyleSheet(
            "font-size: 12px; color: rgba(255,255,255,0.68); background: transparent;"
        )
        hl.addWidget(title_lbl)
        hl.addWidget(sub_lbl)
        vl.addWidget(header)

        # ── Body ───────────────────────────────────────────────────────────
        body = QtWidgets.QWidget()
        bl = QtWidgets.QVBoxLayout(body)
        bl.setContentsMargins(24, 20, 24, 16)
        bl.setSpacing(8)

        for key, val in [
            ("Version",      "4.0  (Phase 2 — Dialogs)"),
            ("Runtime",      "Python 3.13  ·  PySide6 / Qt6"),
            ("Institution",  "IRM — University of Minnesota"),
            ("License",      "GPLv3 open-source"),
            ("VB6 origin",   "RAPID v3 · Sourceforge"),
        ]:
            row = QtWidgets.QHBoxLayout()
            k = QtWidgets.QLabel(f"{key}:")
            k.setFixedWidth(110)
            k.setStyleSheet("color: #9a8885; font-size: 12px;")
            v = QtWidgets.QLabel(val)
            v.setStyleSheet("color: #2f2827; font-size: 12px;")
            row.addWidget(k)
            row.addWidget(v, 1)
            bl.addLayout(row)

        bl.addSpacing(8)
        desc = QtWidgets.QLabel(
            "Python rewrite of the RAPID palaeomagnetic instrument control system. "
            "Supports SQUID magnetometers, AF demagnetizers, DC motor sample changers, "
            "and ADwin field-control hardware. Staged module-by-module replacement of "
            "the original VB6 application."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #6b7280; font-size: 11px;")
        bl.addWidget(desc)
        vl.addWidget(body)

        # ── Buttons ────────────────────────────────────────────────────────
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setContentsMargins(24, 4, 24, 18)

        gh_btn = QtWidgets.QPushButton("GitHub ↗")
        gh_btn.clicked.connect(
            lambda: QtGui.QDesktopServices.openUrl(
                QtCore.QUrl("https://github.com/duserzym/RAPID")
            )
        )
        ok_btn = QtWidgets.QPushButton("OK")
        ok_btn.setObjectName("accent")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.accept)

        btn_row.addWidget(gh_btn)
        btn_row.addStretch()
        btn_row.addWidget(ok_btn)
        vl.addLayout(btn_row)
