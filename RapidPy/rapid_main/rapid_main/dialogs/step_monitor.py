from __future__ import annotations

from PySide6 import QtCore, QtWidgets
from rapidpy_common.ui import clamp_window_geometry


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
            if isinstance(screen, QtCore.QScreen)
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
