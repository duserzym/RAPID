from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class DashboardPanel(QtWidgets.QWidget):
    """Home panel — instrument status, run state, quick actions, event log.

    Maps to: frmProgram status info + flow state (VB6 MDI parent).
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        inner = QtWidgets.QWidget()
        vl = QtWidgets.QVBoxLayout(inner)
        vl.setContentsMargins(24, 20, 24, 24)
        vl.setSpacing(16)

        vl.addLayout(self._build_instrument_row())
        vl.addLayout(self._build_mid_row())
        vl.addWidget(self._build_log_card())
        vl.addStretch()

        scroll.setWidget(inner)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(scroll)

    # ── Instrument status row ─────────────────────────────────────────────────
    def _build_instrument_row(self) -> QtWidgets.QHBoxLayout:
        row = QtWidgets.QHBoxLayout()
        row.setSpacing(12)
        instruments = [
            ("SQUID",       "2G Enterprises 755",  "instUnk"),
            ("ADwin",       "ADwin Light-16",       "instUnk"),
            ("Changer",     "XY Sample Changer",    "instUnk"),
            ("Gaussmeter",  "F.W. Bell 908A",       "instUnk"),
            ("AF Coil",     "ADwin controlled",     "instUnk"),
        ]
        for name, model, state in instruments:
            row.addWidget(self._inst_card(name, model, state))
        return row

    def _inst_card(self, name: str, model: str, state_name: str) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 12, 14, 12)
        cl.setSpacing(4)

        title = QtWidgets.QLabel(name)
        title.setStyleSheet("font-weight: 700; font-size: 13px; color: #2f2827;")
        subtitle = QtWidgets.QLabel(model)
        subtitle.setStyleSheet("font-size: 11px; color: #9a8885;")

        status = QtWidgets.QLabel("● Not connected")
        status.setObjectName(state_name)
        setattr(self, f"_inst_{name.lower().replace(' ', '_')}", status)

        cl.addWidget(title)
        cl.addWidget(subtitle)
        cl.addSpacing(4)
        cl.addWidget(status)
        return card

    # ── Middle row: run state + quick actions ─────────────────────────────────
    def _build_mid_row(self) -> QtWidgets.QHBoxLayout:
        row = QtWidgets.QHBoxLayout()
        row.setSpacing(12)
        row.addWidget(self._build_run_card(), 3)
        row.addWidget(self._build_actions_card(), 2)
        return row

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
            ("📂  Load Sample",        "queue"),
            ("🔬  Set Up Sequence",     "sequence"),
            ("▶   Start New Run",       "measure"),
            ("📊  View Sample Queue",   "queue"),
            ("⚙️  Settings",           "settings"),
        ]
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

    def update_run_state(self, flow: str, sample: str, step: str,
                         treatment: str, elapsed: str) -> None:
        self._run_flow.setText(flow)
        self._run_sample.setText(sample)
        self._run_step.setText(step)
        self._run_treat.setText(treatment)
        self._run_elapsed.setText(elapsed)
