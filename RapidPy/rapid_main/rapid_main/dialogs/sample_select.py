from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtWidgets
from rapidpy_common.ui import clamp_window_geometry

try:
    from rapid_main.io.sam_reader import specimen_path
    from rapid_main.io.sample_index import read_sample_index_registrations
    from rapid_main.io.specimen_reader import read_specimen
    _IO_AVAILABLE = True
except ImportError:
    _IO_AVAILABLE = False

_DEMO_SAMPLES = [
    ("BK-001", "10.25", "Basalt flow A", "Berkeley"),
    ("BK-002", "18.50", "Basalt flow A", "Berkeley"),
    ("BK-003", "24.00", "Basalt flow B", "Berkeley"),
    ("IC-012", "5.10",  "Ignimbrite",   "Iceland"),
    ("IC-013", "6.40",  "Ignimbrite",   "Iceland"),
    ("MN-047", "312.5", "Red sandstone", "Minnesota"),
    ("MN-048", "318.0", "Red sandstone", "Minnesota"),
    ("MN-049", "324.0", "Red sandstone", "Minnesota"),
]

_COLS = ("Sample Name", "Depth (cm)", "Formation", "Location")


class SampleSelectDialog(QtWidgets.QDialog):
    """Filterable sample browser — replaces VB6 frmSampleSelect."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Select Sample")
        self.resize(560, 380)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._load_demo()

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

    # ── Public ─────────────────────────────────────────────────────────────
    @property
    def selected_sample(self) -> str | None:
        rows = self._table.selectedItems()
        if not rows:
            return None
        return self._table.item(rows[0].row(), 0).text()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(16, 14, 16, 14)
        vl.setSpacing(8)

        hdr = QtWidgets.QLabel("Sample Index")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        # Search
        search_row = QtWidgets.QHBoxLayout()
        self._search = QtWidgets.QLineEdit()
        self._search.setPlaceholderText("Filter by name, formation, or location…")
        self._search.textChanged.connect(self._filter)
        load_btn = QtWidgets.QPushButton("Load File…")
        load_btn.clicked.connect(self._load_file)
        search_row.addWidget(self._search, 1)
        search_row.addWidget(load_btn)
        vl.addLayout(search_row)

        # Table
        self._table = QtWidgets.QTableWidget(0, len(_COLS))
        self._table.setHorizontalHeaderLabels(list(_COLS))
        self._table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.verticalHeader().setVisible(False)
        self._table.doubleClicked.connect(self._on_double_click)
        vl.addWidget(self._table, 1)

        # Buttons
        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        btns.button(QtWidgets.QDialogButtonBox.Ok).setText("Select")
        btns.accepted.connect(self._on_accept)
        btns.rejected.connect(self.reject)
        vl.addWidget(btns)

    def _load_demo(self) -> None:
        for row_data in _DEMO_SAMPLES:
            row = self._table.rowCount()
            self._table.insertRow(row)
            for col, val in enumerate(row_data):
                item = QtWidgets.QTableWidgetItem(val)
                self._table.setItem(row, col, item)

    def _filter(self, text: str) -> None:
        text = text.lower()
        for row in range(self._table.rowCount()):
            match = any(
                text in (self._table.item(row, col).text().lower() or "")
                for col in range(self._table.columnCount())
                if self._table.item(row, col)
            )
            self._table.setRowHidden(row, not match)

    def _load_file(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Sample Index",
            "",
            "SAM index (*.sam);;CSV files (*.csv);;All files (*)",
        )
        if not path:
            return
        p = Path(path)
        if p.suffix.lower() == ".sam":
            self._load_sam(p)
        elif p.suffix.lower() == ".csv":
            self._load_csv(p)
        else:
            # Try to guess from content
            self._load_sam(p)

    def _load_sam(self, sam_path: Path) -> None:
        """Parse a .sam specimen index and populate the table."""
        if not _IO_AVAILABLE:
            QtWidgets.QMessageBox.warning(
                self, "Load SAM", "rapid_main.io not available."
            )
            return
        try:
            registrations = read_sample_index_registrations(sam_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load SAM", f"Could not read file:\n{exc}")
            return
        if not registrations.names:
            QtWidgets.QMessageBox.information(
                self, "Load SAM", "No specimen names found in file."
            )
            return
        self._table.setRowCount(0)
        for reg in registrations.entries:
            name = reg.specimen_name
            sp = specimen_path(sam_path, name)
            depth_or_volume = reg.depth_cm
            comment_str = reg.formation
            location = reg.location
            if sp.exists():
                try:
                    meta, _ = read_specimen(sp, specimen_name=name)
                    depth_or_volume = depth_or_volume or (f"{meta.volume:.2f}" if meta.volume else "")
                    comment_str = meta.comment or reg.formation
                except Exception:
                    pass
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QtWidgets.QTableWidgetItem(name))
            self._table.setItem(row, 1, QtWidgets.QTableWidgetItem(depth_or_volume))
            self._table.setItem(row, 2, QtWidgets.QTableWidgetItem(comment_str))
            self._table.setItem(row, 3, QtWidgets.QTableWidgetItem(location or str(sam_path.parent)))

    def _load_csv(self, csv_path: Path) -> None:
        """Parse a CSV file and populate the table."""
        import csv
        try:
            with csv_path.open(newline="", encoding="utf-8-sig") as fh:
                reader = csv.reader(fh)
                rows = list(reader)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load CSV", f"Could not read file:\n{exc}")
            return
        if not rows:
            return
        # Skip header row if first cell matches a column name
        start = 1 if rows and rows[0][0].strip().lower() in ("sample name", "name", "specimen") else 0
        self._table.setRowCount(0)
        for csv_row in rows[start:]:
            if not csv_row or not csv_row[0].strip():
                continue
            row = self._table.rowCount()
            self._table.insertRow(row)
            for col in range(min(len(_COLS), len(csv_row))):
                self._table.setItem(row, col, QtWidgets.QTableWidgetItem(csv_row[col].strip()))

    def _on_accept(self) -> None:
        if not self.selected_sample:
            QtWidgets.QMessageBox.warning(self, "Select Sample", "Please select a sample row.")
            return
        self.accept()

    def _on_double_click(self) -> None:
        if self.selected_sample:
            self.accept()
