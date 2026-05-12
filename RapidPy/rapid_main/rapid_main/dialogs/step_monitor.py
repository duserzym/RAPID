from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class StepMonitorDialog(QtWidgets.QDialog):
    """Live step execution monitor — replaces VB6 frmStepMonitor."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Step Monitor")
        self.resize(520, 400)
        self.setWindowFlags(
            self.windowFlags()
            & ~QtCore.Qt.WindowContextHelpButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
        )
        self._build_ui()

    # ── Public API ─────────────────────────────────────────────────────────
    def update_step(
        self,
        step_name: str = "—",
        treatment: str = "—",
        position: str = "—",
        progress: int = 0,
        total: int = 1,
    ) -> None:
        self._step_lbl.setText(step_name)
        self._treatment_lbl.setText(treatment)
        self._pos_lbl.setText(position)
        self._progress.setMaximum(max(total, 1))
        self._progress.setValue(max(0, min(progress, total)))
        self._count_lbl.setText(f"{progress} / {total}")

    def log(self, text: str) -> None:
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.appendPlainText(f"[{ts}]  {text}")

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(16, 14, 16, 14)
        vl.setSpacing(10)

        hdr = QtWidgets.QLabel("Step Monitor")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        # Current step info grid
        info_frame = QtWidgets.QFrame()
        info_frame.setStyleSheet(
            "QFrame { background: rgba(122,2,25,0.04); border: 1px solid rgba(122,2,25,0.12);"
            " border-radius: 8px; }"
        )
        gl = QtWidgets.QGridLayout(info_frame)
        gl.setContentsMargins(14, 10, 14, 10)
        gl.setSpacing(8)

        def _field(label: str) -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(label)
            lbl.setStyleSheet("color: #9a8885; font-size: 11px;")
            return lbl

        def _value() -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel("—")
            lbl.setStyleSheet("color: #2f2827; font-size: 13px; font-weight: 600;")
            return lbl

        self._step_lbl = _value()
        self._treatment_lbl = _value()
        self._pos_lbl = _value()

        gl.addWidget(_field("Step"), 0, 0)
        gl.addWidget(self._step_lbl, 0, 1)
        gl.addWidget(_field("Treatment"), 1, 0)
        gl.addWidget(self._treatment_lbl, 1, 1)
        gl.addWidget(_field("Position"), 2, 0)
        gl.addWidget(self._pos_lbl, 2, 1)

        vl.addWidget(info_frame)

        # Progress bar
        prog_row = QtWidgets.QHBoxLayout()
        self._progress = QtWidgets.QProgressBar()
        self._progress.setRange(0, 1)
        self._progress.setValue(0)
        self._count_lbl = QtWidgets.QLabel("0 / 0")
        self._count_lbl.setFixedWidth(60)
        self._count_lbl.setStyleSheet("color: #4d3a39; font-size: 12px;")
        prog_row.addWidget(self._progress, 1)
        prog_row.addWidget(self._count_lbl)
        vl.addLayout(prog_row)

        # Log
        log_hdr = QtWidgets.QLabel("Step Log")
        log_hdr.setObjectName("sectionHdr")
        vl.addWidget(log_hdr)

        self._log = QtWidgets.QPlainTextEdit()
        self._log.setObjectName("console")
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(140)
        vl.addWidget(self._log, 1)

        # Buttons
        btn_row = QtWidgets.QHBoxLayout()
        clear_btn = QtWidgets.QPushButton("Clear Log")
        clear_btn.clicked.connect(self._log.clear)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

        # Seed log
        self.log("Step monitor ready.")
