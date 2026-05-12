from __future__ import annotations

from PySide6 import QtCore, QtWidgets


# Sample table column definitions
_COLS = ["#", "Position", "Sample Name", "Sample Set", "Treatment Steps", "Status"]


class SampleQueuePanel(QtWidgets.QWidget):
    """Sample changer list and run options panel.

    Maps to: frmChanger (Hole Sample List) in VB6.
    Columns mirror the MSHFlexGrid with added Status column.
    Options sidebar mirrors the four VB6 FrameXxx option groups.
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_toolbar())

        body = QtWidgets.QHBoxLayout()
        body.setContentsMargins(16, 12, 16, 16)
        body.setSpacing(12)
        body.addWidget(self._build_table_card(), 3)
        body.addWidget(self._build_options_card(), 1)
        root.addLayout(body)

    # ── Toolbar ───────────────────────────────────────────────────────────────
    def _build_toolbar(self) -> QtWidgets.QFrame:
        bar = QtWidgets.QFrame()
        bar.setObjectName("header")
        bar.setFixedHeight(48)
        hl = QtWidgets.QHBoxLayout(bar)
        hl.setContentsMargins(16, 0, 16, 0)
        hl.setSpacing(8)

        run_btn = QtWidgets.QPushButton("▶  Run Queue")
        run_btn.setObjectName("accent")
        hl.addWidget(run_btn)

        self._pause_btn = QtWidgets.QPushButton("⏸  Pause")
        self._halt_btn  = QtWidgets.QPushButton("■  Halt")
        hl.addWidget(self._pause_btn)
        hl.addWidget(self._halt_btn)

        hl.addWidget(_vline())

        add_btn  = QtWidgets.QPushButton("＋  Add Sample")
        seq_btn  = QtWidgets.QPushButton("↻  Sequential")
        clr_btn  = QtWidgets.QPushButton("✕  Clear")
        exp_btn  = QtWidgets.QPushButton("↑  Export")
        imp_btn  = QtWidgets.QPushButton("↓  Import")
        clr_btn.clicked.connect(self._clear_table)

        for btn in (add_btn, seq_btn, clr_btn, exp_btn, imp_btn):
            hl.addWidget(btn)

        hl.addStretch()

        self._count_lbl = QtWidgets.QLabel("0 samples")
        self._count_lbl.setStyleSheet("color: #7a6f6e; font-size: 12px; padding-right: 8px;")
        hl.addWidget(self._count_lbl)
        return bar

    # ── Main sample table ─────────────────────────────────────────────────────
    def _build_table_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("SAMPLE QUEUE")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        self._table = QtWidgets.QTableWidget(0, len(_COLS))
        self._table.setHorizontalHeaderLabels(_COLS)
        self._table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        self._table.horizontalHeader().setSectionResizeMode(
            2, QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self._table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked)
        self._table.setStyleSheet(
            "QTableWidget { border: none; border-radius: 10px; }"
            "QTableWidget::item:selected { background: rgba(122,2,25,0.12); color: #2f2827; }"
            "QTableWidget { alternate-background-color: rgba(122,2,25,0.04); }"
        )
        self._table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._context_menu)
        self._table.model().rowsInserted.connect(self._update_count)
        self._table.model().rowsRemoved.connect(self._update_count)
        cl.addWidget(self._table)

        # Seed with example placeholder rows
        self._add_sample_row(1, "A1", "HBK-01", "Hole A", "NRM → 25mT AF → 50mT AF", "Pending")
        self._add_sample_row(2, "A2", "HBK-02", "Hole A", "NRM → 25mT AF → 50mT AF", "Pending")
        self._add_sample_row(3, "B1", "HBK-03", "Hole B", "Rockmag the Works",        "Pending")
        return card

    def _add_sample_row(self, num: int, pos: str, name: str,
                        sample_set: str, treatment: str, status: str) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for col, text in enumerate([str(num), pos, name, sample_set, treatment, status]):
            item = QtWidgets.QTableWidgetItem(text)
            if col in (0, 5):
                item.setTextAlignment(QtCore.Qt.AlignCenter)
            self._table.setItem(row, col, item)
        # Colour the status cell
        status_colors = {
            "Pending":   "rgba(107,114,128,0.12)",
            "Running":   "rgba(34,197,94,0.15)",
            "Done":      "rgba(34,197,94,0.08)",
            "Error":     "rgba(220,38,38,0.12)",
            "Skipped":   "rgba(251,191,36,0.15)",
        }
        if status in status_colors:
            self._table.item(row, 5).setBackground(
                QtWidgets.QApplication.palette().base()
            )

    def _clear_table(self) -> None:
        self._table.setRowCount(0)

    def _update_count(self) -> None:
        n = self._table.rowCount()
        self._count_lbl.setText(f"{n} sample{'s' if n != 1 else ''}")

    def _context_menu(self, pos: QtCore.QPoint) -> None:
        row = self._table.rowAt(pos.y())
        if row < 0:
            return
        menu = QtWidgets.QMenu(self)
        menu.addAction("Sample Info")
        menu.addSeparator()
        menu.addAction("Insert Sample Above")
        menu.addAction("Delete",              lambda: self._table.removeRow(row))
        menu.addAction("Delete without Gap",  lambda: self._table.removeRow(row))
        menu.addSeparator()
        menu.addAction("Delete next 9 samples")
        menu.exec(self._table.viewport().mapToGlobal(pos))

    # ── Options sidebar ───────────────────────────────────────────────────────
    def _build_options_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(14)

        hdr = QtWidgets.QLabel("RUN OPTIONS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        def _option_group(title: str, options: list[str], default: int = 0
                          ) -> tuple[QtWidgets.QFrame, list[QtWidgets.QRadioButton]]:
            grp = QtWidgets.QFrame()
            grp.setObjectName("card")
            grp.setStyleSheet("QFrame#card { border-radius: 10px; }")
            gl = QtWidgets.QVBoxLayout(grp)
            gl.setContentsMargins(10, 8, 10, 10)
            gl.setSpacing(4)
            t = QtWidgets.QLabel(title)
            t.setObjectName("readLbl")
            gl.addWidget(t)
            radios = []
            for i, opt in enumerate(options):
                rb = QtWidgets.QRadioButton(opt)
                if i == default:
                    rb.setChecked(True)
                gl.addWidget(rb)
                radios.append(rb)
            return grp, radios

        grp1, self._order_radios = _option_group(
            "Sample Order", ["Ascending", "Descending"], default=0
        )
        grp2, self._reload_radios = _option_group(
            "Reload Position", ["Return to start", "Leave at end"], default=0
        )
        grp3, self._final_radios = _option_group(
            "Final Position", ["Return to start", "Leave at end"], default=0
        )
        grp4, self._holder_radios = _option_group(
            "Multiple Holder Measurements",
            ["Repeat  (weak samples)", "Skip  (strong samples)"],
            default=0
        )
        for grp in (grp1, grp2, grp3, grp4):
            cl.addWidget(grp)

        cl.addStretch()

        print_btn = QtWidgets.QPushButton("🖨  Print List")
        cl.addWidget(print_btn)
        return card


def _vline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.VLine)
    f.setStyleSheet("color: rgba(122,2,25,0.18); margin: 8px 2px;")
    return f
