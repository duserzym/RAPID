from __future__ import annotations

from collections.abc import Sequence

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import DiagnosticStatusLine


class DashboardPanel(QtWidgets.QWidget):
    """Home panel — instrument status, run state, quick actions, event log.

    Maps to: frmProgram status info + flow state (VB6 MDI parent).
    """

    refresh_diagnostics_requested = QtCore.Signal()
    load_sample_requested = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._instrument_status_labels: dict[str, QtWidgets.QLabel] = {}
        self._instrument_cards: list[QtWidgets.QFrame] = []
        self._dashboard_compact: bool | None = None
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        inner = QtWidgets.QWidget()
        inner.setMinimumWidth(0)
        inner.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        vl = QtWidgets.QVBoxLayout(inner)
        vl.setContentsMargins(24, 20, 24, 24)
        vl.setSpacing(16)

        vl.addLayout(self._build_instrument_section())
        vl.addLayout(self._build_mid_row())
        vl.addWidget(self._build_log_card())
        vl.addStretch()

        scroll.setWidget(inner)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(scroll)

    def resizeEvent(self, event: QtCore.QEvent) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._apply_responsive_layout(self.width() < 980)

    def _apply_responsive_layout(self, compact: bool) -> None:
        if self._dashboard_compact is compact:
            return
        self._dashboard_compact = compact
        columns = 2 if compact else 5
        for card in self._instrument_cards:
            self._instrument_grid.removeWidget(card)
        for index, card in enumerate(self._instrument_cards):
            self._instrument_grid.addWidget(card, index // columns, index % columns)
        for column in range(5):
            self._instrument_grid.setColumnStretch(column, 1 if column < columns else 0)

        self._mid_grid.removeWidget(self._run_card)
        self._mid_grid.removeWidget(self._actions_card)
        if compact:
            self._mid_grid.addWidget(self._run_card, 0, 0)
            self._mid_grid.addWidget(self._actions_card, 1, 0)
            self._mid_grid.setColumnStretch(0, 1)
            self._mid_grid.setColumnStretch(1, 0)
        else:
            self._mid_grid.addWidget(self._run_card, 0, 0)
            self._mid_grid.addWidget(self._actions_card, 0, 1)
            self._mid_grid.setColumnStretch(0, 3)
            self._mid_grid.setColumnStretch(1, 2)

    # ── Instrument status row ─────────────────────────────────────────────────
    def _build_instrument_section(self) -> QtWidgets.QVBoxLayout:
        section = QtWidgets.QVBoxLayout()
        section.setSpacing(8)

        heading = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("SYSTEM READINESS")
        title.setObjectName("sectionHdr")
        heading.addWidget(title)
        heading.addStretch()
        self._diagnostic_refresh_label = QtWidgets.QLabel("Not refreshed")
        self._diagnostic_refresh_label.setObjectName("readLbl")
        heading.addWidget(self._diagnostic_refresh_label)
        refresh = QtWidgets.QPushButton("↻  Refresh")
        refresh.setToolTip("Read the current state of every in-process hardware backend")
        refresh.clicked.connect(self.refresh_diagnostics_requested.emit)
        heading.addWidget(refresh)
        section.addLayout(heading)

        self._instrument_grid = QtWidgets.QGridLayout()
        self._instrument_grid.setHorizontalSpacing(12)
        self._instrument_grid.setVerticalSpacing(10)
        instruments = [
            ("SQUID",     "2G Enterprises 755", "instUnk"),
            ("Vacuum",    "Chamber controller", "instUnk"),
            ("DC Motors", "XY / lift / turn",    "instUnk"),
            ("AF Demag",  "ADwin controlled",    "instUnk"),
            ("IRM/ARM",   "ADwin + DAC",         "instUnk"),
        ]
        for index, (name, model, state) in enumerate(instruments):
            card = self._inst_card(name, model, state)
            self._instrument_cards.append(card)
            self._instrument_grid.addWidget(card, index // 2, index % 2)
        section.addLayout(self._instrument_grid)
        return section

    def _inst_card(self, name: str, model: str, state_name: str) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        card.setMinimumWidth(0)
        card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 12, 14, 12)
        cl.setSpacing(4)

        title = QtWidgets.QLabel(name)
        title.setMinimumWidth(0)
        title.setStyleSheet("font-weight: 700; font-size: 13px; color: #2f2827;")
        subtitle = QtWidgets.QLabel(model)
        subtitle.setMinimumWidth(0)
        subtitle.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        subtitle.setStyleSheet("font-size: 11px; color: #9a8885;")

        status = QtWidgets.QLabel("● Not refreshed")
        status.setObjectName(state_name)
        status.setWordWrap(True)
        status.setMinimumWidth(0)
        status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        status.setMinimumHeight(38)
        self._instrument_status_labels[name] = status

        cl.addWidget(title)
        cl.addWidget(subtitle)
        cl.addSpacing(4)
        cl.addWidget(status)
        return card

    # ── Middle row: run state + quick actions ─────────────────────────────────
    def _build_mid_row(self) -> QtWidgets.QGridLayout:
        self._mid_grid = QtWidgets.QGridLayout()
        self._mid_grid.setHorizontalSpacing(12)
        self._mid_grid.setVerticalSpacing(12)
        self._run_card = self._build_run_card()
        self._actions_card = self._build_actions_card()
        self._mid_grid.addWidget(self._run_card, 0, 0)
        self._mid_grid.addWidget(self._actions_card, 1, 0)
        return self._mid_grid

    def _build_run_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(10)

        hdr = QtWidgets.QLabel("CURRENT RUN")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        grid = QtWidgets.QGridLayout()
        grid.setSpacing(8)
        grid.setColumnMinimumWidth(1, 160)

        def _row(r: int, label: str, val_name: str, default: str = "—") -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(label)
            lbl.setObjectName("readLbl")
            val = QtWidgets.QLabel(default)
            val.setObjectName("valuePill")
            grid.addWidget(lbl, r, 0)
            grid.addWidget(val, r, 1)
            setattr(self, val_name, val)
            return val

        _row(0, "Flow State", "_run_flow",    "Halted")
        _row(1, "Sample",     "_run_sample",  "—")
        _row(2, "Step",       "_run_step",    "—")
        _row(3, "Treatment",  "_run_treat",   "—")
        _row(4, "Elapsed",    "_run_elapsed", "00:00:00")

        cl.addLayout(grid)
        cl.addStretch()

        goto_btn = QtWidgets.QPushButton("→  Go to Live Measurement")
        goto_btn.clicked.connect(lambda: self._goto("measure"))
        cl.addWidget(goto_btn)
        return card

    def _build_actions_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("QUICK ACTIONS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        actions = [
            ("🔬  Set Up Sequence",     "sequence"),
            ("▶   Start New Run",       "measure"),
            ("📊  View Sample Queue",   "queue"),
            ("⚙️  Settings",           "settings"),
        ]
        load_sample = QtWidgets.QPushButton("📂  Load Sample")
        load_sample.clicked.connect(self.load_sample_requested.emit)
        cl.addWidget(load_sample)
        for label, dest in actions:
            btn = QtWidgets.QPushButton(label)
            btn.clicked.connect(lambda _c=False, d=dest: self._goto(d))
            cl.addWidget(btn)

        cl.addStretch()
        return card

    # ── Event log card ────────────────────────────────────────────────────────
    def _build_log_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(8)

        top = QtWidgets.QHBoxLayout()
        hdr = QtWidgets.QLabel("EVENT LOG")
        hdr.setObjectName("sectionHdr")
        top.addWidget(hdr)
        top.addStretch()
        clear_btn = QtWidgets.QPushButton("Clear")
        clear_btn.setFixedWidth(64)
        top.addWidget(clear_btn)
        cl.addLayout(top)

        self._log = QtWidgets.QPlainTextEdit()
        self._log.setObjectName("console")
        self._log.setReadOnly(True)
        self._log.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.WidgetWidth)
        self._log.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._log.setFixedHeight(160)
        self._log.setPlaceholderText("System events will appear here…")
        clear_btn.clicked.connect(self._log.clear)
        cl.addWidget(self._log)
        return card

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _goto(self, key: str) -> None:
        mw = self.window()
        if hasattr(mw, "navigate_to"):
            mw.navigate_to(key)

    def append_log(self, text: str) -> None:
        ts = QtCore.QDateTime.currentDateTime().toString("hh:mm:ss")
        self._log.appendPlainText(f"[{ts}]  {text}")

    def update_diagnostics(
        self,
        lines: Sequence[DiagnosticStatusLine],
        *,
        refreshed_at: QtCore.QDateTime | None = None,
    ) -> None:
        """Render an authoritative backend snapshot without inventing readiness."""
        by_name = {line.name: line for line in lines}
        for name, label in self._instrument_status_labels.items():
            line = by_name.get(name)
            if line is None:
                text = "● Status unavailable"
                object_name = "instErr"
                detail = f"No diagnostic backend was registered for {name}."
            elif line.fault:
                text = f"● Fault — {line.status}"
                object_name = "instErr"
                detail = line.fault_reason or line.status
            elif line.simulated:
                text = f"◆ SIMULATED — {line.status}"
                object_name = "instSim"
                detail = (
                    f"{line.name} is using a no-communication simulator. "
                    "This is not live hardware evidence."
                )
            elif line.connected:
                text = f"● Connected — {line.status}"
                object_name = "instOk"
                detail = f"Live hardware backend: {line.status}"
            else:
                unavailable = "unavailable" in line.status.lower()
                text = f"● {'Unavailable' if unavailable else 'Disconnected'} — {line.status}"
                object_name = "instErr" if unavailable else "instUnk"
                detail = line.status

            label.setText(text)
            label.setToolTip(detail)
            label.setObjectName(object_name)
            label.style().unpolish(label)
            label.style().polish(label)

        timestamp = refreshed_at or QtCore.QDateTime.currentDateTime()
        self._diagnostic_refresh_label.setText(
            f"Refreshed {timestamp.toString('hh:mm:ss AP')}"
        )

    def update_run_state(self, flow: str, sample: str, step: str,
                         treatment: str, elapsed: str) -> None:
        self._run_flow.setText(flow)
        self._run_sample.setText(sample)
        self._run_step.setText(step)
        self._run_treat.setText(treatment)
        self._run_elapsed.setText(elapsed)
