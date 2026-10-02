from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets
from rapidpy_common.ui import clamp_window_geometry

from rapid_main.glass_theme import set_semantic_status


class StepMonitorDialog(QtWidgets.QDialog):
    """Live step execution monitor — replaces VB6 frmStepMonitor."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Step Monitor")
        self.setAccessibleName("Live queue step monitor")
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
    def update_step(
        self,
        step_name: str = "—",
        treatment: str = "—",
        position: str = "—",
        progress: int = 0,
        total: int = 1,
    ) -> None:
        normalized_total = max(int(total), 1)
        normalized_progress = max(0, min(int(progress), normalized_total))
        self._step_lbl.setText(str(step_name))
        self._treatment_lbl.setText(str(treatment))
        self._pos_lbl.setText(str(position))
        self._progress.setMaximum(normalized_total)
        self._progress.setValue(normalized_progress)
        self._count_lbl.setText(f"{normalized_progress} / {normalized_total}")
        self._progress.setAccessibleDescription(
            f"Queue progress {normalized_progress} of {normalized_total}."
        )
        if str(step_name).strip() in {"", "—", "-"}:
            set_semantic_status(
                self._state_lbl,
                "Waiting for a queue step",
                "neutral",
                accessible_name="Step execution status",
            )
        elif normalized_progress >= normalized_total:
            set_semantic_status(
                self._state_lbl,
                f"Step complete: {step_name}",
                "ready",
                accessible_name="Step execution status",
            )
        else:
            set_semantic_status(
                self._state_lbl,
                f"Executing step: {step_name}",
                "active",
                accessible_name="Step execution status",
            )

    def log(self, text: str) -> None:
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.appendPlainText(f"[{ts}]  {text}")

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(12, 10, 12, 10)
        vl.setSpacing(6)

        hdr = QtWidgets.QLabel("Step Monitor")
        hdr.setObjectName("dialogTitle")
        vl.addWidget(hdr)

        self._state_lbl = QtWidgets.QLabel()
        set_semantic_status(
            self._state_lbl,
            "Waiting for a queue step",
            "neutral",
            accessible_name="Step execution status",
        )
        vl.addWidget(self._state_lbl)

        # Current step info grid
        info_frame = QtWidgets.QFrame()
        info_frame.setObjectName("dialogCard")
        gl = QtWidgets.QGridLayout(info_frame)
        gl.setContentsMargins(10, 8, 10, 8)
        gl.setSpacing(5)

        def _field(label: str) -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(label)
            lbl.setObjectName("dialogSubtitle")
            return lbl

        def _value() -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel("—")
            lbl.setObjectName("valuePill")
            return lbl

        self._step_lbl = _value()
        self._treatment_lbl = _value()
        self._pos_lbl = _value()
        self._step_lbl.setAccessibleName("Current queue step")
        self._treatment_lbl.setAccessibleName("Current treatment")
        self._pos_lbl.setAccessibleName("Current instrument position")

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
        self._progress.setAccessibleName("Queue step progress")
        self._count_lbl = QtWidgets.QLabel("0 / 0")
        self._count_lbl.setObjectName("dialogSubtitle")
        self._count_lbl.setMinimumWidth(60)
        self._count_lbl.setAccessibleName("Queue step progress count")
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
        self._log.setMinimumHeight(80)
        self._log.setAccessibleName("Queue step event log")
        vl.addWidget(self._log, 1)

        # Buttons
        btn_row = QtWidgets.QHBoxLayout()
        self._clear_btn = QtWidgets.QPushButton("Clear Log")
        self._clear_btn.setAccessibleName("Clear queue step log")
        self._clear_btn.clicked.connect(self._log.clear)
        self._close_btn = QtWidgets.QPushButton("Close")
        self._close_btn.setAccessibleName("Close step monitor")
        self._close_btn.clicked.connect(self.close)
        btn_row.addWidget(self._clear_btn)
        btn_row.addStretch()
        btn_row.addWidget(self._close_btn)
        vl.addLayout(btn_row)

        # Seed log
        self.log("Step monitor ready.")
