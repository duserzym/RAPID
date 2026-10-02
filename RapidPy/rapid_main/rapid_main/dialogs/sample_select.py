from __future__ import annotations

from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets
from rapidpy_common.ui import clamp_window_geometry
from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations
from rapid_main.glass_theme import set_semantic_status

try:
    from rapid_main.io.sam_reader import specimen_path
    from rapid_main.io.sample_index import read_sample_index_registrations
    from rapid_main.io.specimen_reader import read_specimen
    _IO_AVAILABLE = True
except ImportError:
    _IO_AVAILABLE = False

_COLS = ("Sample Name", "Depth (cm)", "Formation", "Location")


class SampleSelectDialog(QtWidgets.QDialog):
    """Filterable sample browser — replaces VB6 frmSampleSelect."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Select Sample")
        self.setAccessibleName("Select specimen from a sample index")
        self.resize(560, 380)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._source_path: Path | None = None
        self._registrations: SampleIndexRegistrations | None = None
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

    # ── Public ─────────────────────────────────────────────────────────────
    @property
    def selected_sample(self) -> str | None:
        rows = self._table.selectedItems()
        if not rows:
            return None
        return self._table.item(rows[0].row(), 0).text()

    @property
    def selected_record(self) -> dict[str, str] | None:
        """Return the selected index row without inventing missing metadata."""

        selected = self._table.selectionModel().selectedRows()
        if not selected:
            return None
        row = selected[0].row()
        values = [
            self._table.item(row, column).text()
            if self._table.item(row, column) is not None
            else ""
            for column in range(self._table.columnCount())
        ]
        return {
            "sample_name": values[0],
            "depth_cm": values[1],
            "formation": values[2],
            "location": values[3],
        }

    @property
    def source_path(self) -> Path | None:
        return self._source_path

    @property
    def registrations(self) -> SampleIndexRegistrations | None:
        return self._registrations

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(16, 14, 16, 14)
        vl.setSpacing(8)

        hdr = QtWidgets.QLabel("Sample Index")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("Sample index title")
        vl.addWidget(hdr)

        # Search
        search_row = QtWidgets.QHBoxLayout()
        self._search = QtWidgets.QLineEdit()
        self._search.setPlaceholderText("Filter by name, formation, or location…")
        self._search.setAccessibleName("Filter sample index")
        self._search.textChanged.connect(self._filter)
        self._load_btn = QtWidgets.QPushButton("Load File…")
        self._load_btn.setAccessibleName("Load a SAM or CSV sample index")
        self._load_btn.clicked.connect(self._load_file)
        search_row.addWidget(self._search, 1)
        search_row.addWidget(self._load_btn)
        vl.addLayout(search_row)

        self._source_lbl = QtWidgets.QLabel()
        set_semantic_status(
            self._source_lbl,
            "No sample index loaded. Choose Load File to open a real .sam or .csv index.",
            "neutral",
            accessible_name="Sample index source status",
        )
        self._source_lbl.setWordWrap(True)
        vl.addWidget(self._source_lbl)

        # Table
        self._table = QtWidgets.QTableWidget(0, len(_COLS))
        self._table.setHorizontalHeaderLabels(list(_COLS))
        self._table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._table.setAccessibleName("Samples in the loaded index")
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.verticalHeader().setVisible(False)
        self._table.doubleClicked.connect(self._on_double_click)
        self._table.itemSelectionChanged.connect(self._update_select_enabled)
        vl.addWidget(self._table, 1)

        # Buttons
        self._buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self._buttons.setAccessibleName("Sample selection actions")
        self._select_btn = self._buttons.button(QtWidgets.QDialogButtonBox.StandardButton.Ok)
        self._cancel_btn = self._buttons.button(
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        if self._select_btn is not None:
            self._select_btn.setText("Select")
            self._select_btn.setAccessibleName("Select highlighted sample")
            self._select_btn.setEnabled(False)
        if self._cancel_btn is not None:
            self._cancel_btn.setAccessibleName("Cancel sample selection")
        self._buttons.accepted.connect(self._on_accept)
        self._buttons.rejected.connect(self.reject)
        vl.addWidget(self._buttons)

    def _update_select_enabled(self) -> None:
        if self._select_btn is not None:
            self._select_btn.setEnabled(self.selected_sample is not None)

    def _filter(self, text: str) -> None:
        text = text.lower()
        for row in range(self._table.rowCount()):
            match = any(
                text in (self._table.item(row, col).text().lower() or "")
                for col in range(self._table.columnCount())
                if self._table.item(row, col)
            )
            self._table.setRowHidden(row, not match)
        selected = self._table.selectionModel().selectedRows()
        if selected and self._table.isRowHidden(selected[0].row()):
            self._table.clearSelection()
        self._update_select_enabled()

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
            loaded = self._load_sam(p)
        elif p.suffix.lower() == ".csv":
            loaded = self._load_csv(p)
        else:
            # Try to guess from content
            loaded = self._load_sam(p)
        if loaded:
            self._source_path = p
            set_semantic_status(
                self._source_lbl,
                f"Loaded {self._table.rowCount()} specimen"
                f"{'s' if self._table.rowCount() != 1 else ''} from {p.name}",
                "ready",
                accessible_name="Sample index source status",
            )
            self._source_lbl.setToolTip(str(p))

    def _load_sam(self, sam_path: Path) -> bool:
        """Parse a .sam specimen index and populate the table."""
        if not _IO_AVAILABLE:
            QtWidgets.QMessageBox.warning(
                self, "Load SAM", "rapid_main.io not available."
            )
            return False
        try:
            registrations = read_sample_index_registrations(sam_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load SAM", f"Could not read file:\n{exc}")
            return False
        if not registrations.names:
            QtWidgets.QMessageBox.information(
                self, "Load SAM", "No specimen names found in file."
            )
            return False
        self._registrations = registrations
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
        return True

    def _load_csv(self, csv_path: Path) -> bool:
        """Parse a CSV file and populate the table."""
        import csv
        try:
            with csv_path.open(newline="", encoding="utf-8-sig") as fh:
                reader = csv.reader(fh)
                rows = list(reader)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load CSV", f"Could not read file:\n{exc}")
            return False
        if not rows:
            QtWidgets.QMessageBox.information(self, "Load CSV", "The file contains no samples.")
            return False
        # Skip header row if first cell matches a column name
        start = 1 if rows and rows[0][0].strip().lower() in ("sample name", "name", "specimen") else 0
        parsed_rows: list[list[str]] = []
        registrations: list[SampleIndexRegistration] = []
        for csv_row in rows[start:]:
            if not csv_row or not csv_row[0].strip():
                continue
            values = [
                csv_row[col].strip() if col < len(csv_row) else ""
                for col in range(len(_COLS))
            ]
            parsed_rows.append(values)
            registrations.append(
                SampleIndexRegistration(
                    specimen_name=values[0],
                    sample_set=csv_path.stem,
                    depth_cm=values[1],
                    formation=values[2],
                    location=values[3],
                    order=len(registrations) + 1,
                )
            )
        if not parsed_rows:
            QtWidgets.QMessageBox.information(self, "Load CSV", "The file contains no samples.")
            return False
        self._table.setRowCount(0)
        for values in parsed_rows:
            row = self._table.rowCount()
            self._table.insertRow(row)
            for col, value in enumerate(values):
                self._table.setItem(row, col, QtWidgets.QTableWidgetItem(value))
        self._registrations = SampleIndexRegistrations(registrations)
        return True

    def _on_accept(self) -> None:
        if not self.selected_sample:
            QtWidgets.QMessageBox.warning(self, "Select Sample", "Please select a sample row.")
            return
        self.accept()

    def _on_double_click(self) -> None:
        if self.selected_sample:
            self.accept()
