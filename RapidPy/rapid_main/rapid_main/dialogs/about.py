from __future__ import annotations

import platform

from PySide6 import QtCore, QtGui, QtWidgets


class AboutDialog(QtWidgets.QDialog):
    """About box — replaces VB6 frmAbout."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("About RAPID v4")
        self.setAccessibleName("About RAPID version 4")
        self.setMinimumWidth(340)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setSpacing(0)
        vl.setContentsMargins(0, 0, 0, 0)

        # ── Header stripe ──────────────────────────────────────────────────
        header = QtWidgets.QFrame()
        header.setObjectName("dialogHero")
        header.setMinimumHeight(96)
        hl = QtWidgets.QVBoxLayout(header)
        hl.setContentsMargins(24, 14, 24, 14)

        title_lbl = QtWidgets.QLabel("RAPID v4")
        title_lbl.setObjectName("dialogHeroTitle")
        title_lbl.setAccessibleName("RAPID version 4")
        sub_lbl = QtWidgets.QLabel("Paleomagnetics Control System")
        sub_lbl.setObjectName("dialogHeroSubtitle")
        hl.addWidget(title_lbl)
        hl.addWidget(sub_lbl)
        vl.addWidget(header)

        # ── Body ───────────────────────────────────────────────────────────
        body = QtWidgets.QWidget()
        body.setObjectName("dialogBody")
        bl = QtWidgets.QVBoxLayout(body)
        bl.setContentsMargins(24, 20, 24, 16)
        bl.setSpacing(8)

        details = QtWidgets.QFormLayout()
        details.setSpacing(7)
        details.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)
        for key, val in [
            ("Version",      "4.0  (transition build)"),
            ("Runtime",      f"Python {platform.python_version()}  ·  Qt {QtCore.qVersion()}"),
            ("Institution",  "IRM — University of Minnesota"),
            ("License",      "GPLv3 open-source"),
            ("VB6 origin",   "RAPID v3 · Sourceforge"),
        ]:
            k = QtWidgets.QLabel(f"{key}:")
            k.setObjectName("metaKey")
            v = QtWidgets.QLabel(val)
            v.setObjectName("metaValue")
            v.setWordWrap(True)
            details.addRow(k, v)
        bl.addLayout(details)

        bl.addSpacing(8)
        desc = QtWidgets.QLabel(
            "Python rewrite of the RAPID palaeomagnetic instrument control system. "
            "Supports SQUID magnetometers, AF demagnetizers, DC motor sample changers, "
            "and ADwin field-control hardware. Staged module-by-module replacement of "
            "the original VB6 application."
        )
        desc.setWordWrap(True)
        desc.setObjectName("guidanceText")
        desc.setAccessibleName("RAPID application description")
        bl.addWidget(desc)
        vl.addWidget(body)

        # ── Buttons ────────────────────────────────────────────────────────
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setContentsMargins(24, 4, 24, 18)

        self._github_btn = QtWidgets.QPushButton("GitHub ↗")
        self._github_btn.setAccessibleName("Open RAPID repository on GitHub")
        self._github_btn.clicked.connect(
            lambda: QtGui.QDesktopServices.openUrl(
                QtCore.QUrl("https://github.com/duserzym/RAPID")
            )
        )
        self._ok_btn = QtWidgets.QPushButton("OK")
        self._ok_btn.setObjectName("accent")
        self._ok_btn.setAccessibleName("Close About RAPID")
        self._ok_btn.setDefault(True)
        self._ok_btn.clicked.connect(self.accept)

        btn_row.addWidget(self._github_btn)
        btn_row.addStretch()
        btn_row.addWidget(self._ok_btn)
        vl.addLayout(btn_row)
