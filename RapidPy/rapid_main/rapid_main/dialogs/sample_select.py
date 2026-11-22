from __future__ import annotations

from csv import Error as CsvError
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets
from rapidpy_common.flow_layout import FlowLayout
from rapidpy_common.glass import fit_workspace_window
from rapid_main.data_model import SampleIndexRegistrations
from rapid_main.glass_theme import set_semantic_status
from rapid_main.io.sample_index import read_sample_index_registrations
from rapid_main.io.specimen_reader import read_specimen
from rapid_main.sample_index_registry import (IndexQueueDefaults, read_registered_indexes,
                                              remove_registered_index, save_registered_index)
from rapid_main.specimen_paths import contained_path

_COLS = ("Sample Name", "Depth (cm)", "Formation", "Location")
_DIRECTIONS = (
    ("Up only", True, False), ("Down only", False, False),
    ("Up then Down", True, True), ("Down with saved Up", False, True),
)


class SampleSelectDialog(QtWidgets.QDialog):
    """Filterable sample browser — replaces VB6 frmSampleSelect."""

    def __init__(self, parent: QtWidgets.QWidget | None = None, *, settings: QtCore.QSettings | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Select Sample")
        self.setAccessibleName("Select specimen from a sample index")
        self.resize(640, 480)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._source_path: Path | None = None
        self._registrations: SampleIndexRegistrations | None = None
        self._registry_settings = settings if settings is not None else getattr(parent, "_settings", None)
        self._registered: dict[Path, IndexQueueDefaults] = {}
        self._build_ui()
        self._refresh_registry()

    def showEvent(self, event: QtCore.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        if getattr(QtWidgets.QApplication.instance(), "_rapidpy_window_guard", None) is not None:
            return
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            if handle.screen() is not None:
                handle.screen().availableGeometryChanged.connect(self._fit_to_screen)
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    def _fit_to_screen(self, screen: QtCore.QObject | None = None) -> None:
        fit_workspace_window(self, (self.width(), self.height()),
                             screen=screen if isinstance(screen, QtGui.QScreen) else None)

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

    @property
    def queue_defaults(self) -> IndexQueueDefaults:
        _, do_up, do_both = _DIRECTIONS[self._direction.currentIndex()]
        return IndexQueueDefaults(self._treatment.text().strip(), do_up, do_both, self._blocks.value())

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(14, 9, 14, 9)
        vl.setSpacing(6)

        hdr = QtWidgets.QLabel("Sample Index")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("Sample index title")
        vl.addWidget(hdr)

        registry = QtWidgets.QWidget()
        registry_row = FlowLayout(registry, margin=0, h_spacing=6, v_spacing=4)
        registry_row.setContentsMargins(0, 0, 0, 0)
        self._indexes = QtWidgets.QComboBox()
        self._indexes.setAccessibleName("Registered sample indexes")
        self._indexes.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self._indexes.setMinimumContentsLength(18)
        self._indexes.activated.connect(self._open_registration)
        self._register_btn = QtWidgets.QPushButton("Save to Registry")
        self._register_btn.clicked.connect(self._save_registration)
        self._remove_btn = QtWidgets.QPushButton("Remove")
        self._remove_btn.setAccessibleName("Remove selected registry entry")
        self._remove_btn.clicked.connect(self._remove_registration)
        registry_row.addWidget(self._indexes)
        registry_row.addWidget(self._register_btn)
        registry_row.addWidget(self._remove_btn)
        vl.addWidget(registry)
        registry.setVisible(self._registry_settings is not None)
        self._registry_status = QtWidgets.QLabel()
        self._registry_status.setTextFormat(QtCore.Qt.TextFormat.PlainText)
        self._registry_status.setWordWrap(True)
        self._registry_status.hide()
        vl.addWidget(self._registry_status)

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
        self._source_lbl.setTextFormat(QtCore.Qt.TextFormat.PlainText)
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
        self._table.setMinimumHeight(68)
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

        self._defaults = QtWidgets.QWidget()
        defaults = QtWidgets.QFormLayout(self._defaults)
        defaults.setContentsMargins(0, 0, 0, 0)
        defaults.setVerticalSpacing(4)
        defaults.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        defaults.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)
        self._treatment = QtWidgets.QLineEdit()
        self._treatment.setPlaceholderText("Use the current Sequence when blank")
        self._treatment.setAccessibleName("Treatment sequence for this index")
        defaults.addRow("Treatment", self._treatment)
        self._direction = QtWidgets.QComboBox()
        self._direction.setAccessibleName("Measurement direction for this index")
        self._direction.addItems([label for label, _, _ in _DIRECTIONS])
        defaults.addRow(self._direction)
        self._blocks = QtWidgets.QSpinBox()
        self._blocks.setAccessibleName("Acquisition blocks per treatment for this index")
        self._blocks.setRange(1, 32767)
        defaults.addRow("Blocks per treatment", self._blocks)
        self._defaults.setEnabled(False)
        vl.addWidget(self._defaults)

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
        self._open_source(Path(path))

    def _open_source(self, path: Path) -> bool:
        try:
            path = path.resolve(strict=True)
        except (OSError, RuntimeError, ValueError) as exc:
            QtWidgets.QMessageBox.warning(self, "Load Sample Index", str(exc))
            return False
        loader = self._load_csv if path.suffix.lower() == ".csv" else self._load_sam
        if not loader(path):
            return False
        self._source_path = path
        self._apply_defaults(self._registered.get(path, IndexQueueDefaults()))
        count = self._table.rowCount()
        specimens = "specimen" if count == 1 else "specimens"
        set_semantic_status(self._source_lbl, f"Loaded {count} {specimens} from {path.name}",
                            "ready", accessible_name="Sample index source status")
        self._source_lbl.setToolTip(str(path))
        self._defaults.setEnabled(True)
        self._filter(self._search.text())
        index = self._indexes.findData(str(path))
        self._indexes.setCurrentIndex(max(0, index))
        self._update_registry_actions()
        return True

    def _apply_defaults(self, defaults: IndexQueueDefaults) -> None:
        self._treatment.setText(defaults.treatment)
        self._blocks.setValue(defaults.avg_steps)
        for index, (_, do_up, do_both) in enumerate(_DIRECTIONS):
            if (do_up, do_both) == (defaults.do_up, defaults.do_both):
                self._direction.setCurrentIndex(index)
                break

    def _registry_message(self, message: str, level: str) -> None:
        set_semantic_status(self._registry_status, message, level, accessible_name="Sample index registry status")
        self._registry_status.setVisible(bool(message))

    def _refresh_registry(self) -> None:
        if self._registry_settings is not None:
            try:
                self._registered, invalid = read_registered_indexes(self._registry_settings)
                if invalid:
                    self._registry_message(f"Skipped {invalid} unreadable registry entries.", "warning")
            except OSError as exc:
                self._registry_message(str(exc), "error")
        self._indexes.clear()
        self._indexes.addItem("Registered indexes…")
        for path in sorted(self._registered, key=lambda item: (item.name.casefold(), str(item).casefold())):
            self._indexes.addItem(f"{path.name}  ({path.parent})", str(path))
        self._indexes.setCurrentIndex(max(0, self._indexes.findData(str(self._source_path))))
        self._update_registry_actions()

    def _update_registry_actions(self) -> None:
        self._register_btn.setEnabled(self._registry_settings is not None and self._source_path is not None)
        self._remove_btn.setEnabled(self._indexes.currentData() is not None)

    def _open_registration(self, index: int) -> None:
        path = self._indexes.itemData(index)
        if path and not self._open_source(Path(path)):
            self._source_path = None
            self._registrations = None
            self._table.setRowCount(0)
            self._defaults.setEnabled(False)
            self._source_lbl.setToolTip(path)
            set_semantic_status(self._source_lbl, "The registered index could not be loaded.", "error",
                                accessible_name="Sample index source status")
        self._update_registry_actions()

    def _save_registration(self) -> None:
        if self._source_path is None or self._registry_settings is None:
            return
        try:
            save_registered_index(self._registry_settings, self._source_path, self.queue_defaults)
        except (OSError, ValueError) as exc:
            self._registry_message(str(exc), "error")
            return
        self._registry_message(f"Saved defaults for {self._source_path.name}.", "ready")
        self._refresh_registry()

    def _remove_registration(self) -> None:
        path = self._indexes.currentData()
        if not path or self._registry_settings is None:
            return
        try:
            remove_registered_index(self._registry_settings, Path(path))
        except OSError as exc:
            self._registry_message(str(exc), "error")
            return
        self._apply_defaults(IndexQueueDefaults())
        self._registry_message(f"Removed {Path(path).name} from the registry.", "ready")
        self._refresh_registry()

    def _load_sam(self, sam_path: Path) -> bool:
        return self._load_index(sam_path, preview_headers=True)

    def _load_csv(self, csv_path: Path) -> bool:
        return self._load_index(csv_path, preview_headers=False)

    def _load_index(self, path: Path, *, preview_headers: bool) -> bool:
        try:
            registrations = read_sample_index_registrations(path)
            specimens = [contained_path(path.parent, entry.specimen_name) for entry in registrations.entries]
            if len({str(item).casefold() for item in specimens}) != len(specimens):
                raise ValueError("The sample index contains duplicate specimen names.")
        except (OSError, RuntimeError, ValueError, CsvError) as exc:
            QtWidgets.QMessageBox.warning(self, "Load Sample Index", str(exc))
            return False
        if not registrations.entries:
            QtWidgets.QMessageBox.information(self, "Load Sample Index", "The file contains no samples.")
            return False

        self._table.setRowCount(0)
        self._table.setRowCount(len(specimens))
        for row, (entry, specimen) in enumerate(zip(registrations.entries, specimens, strict=True)):
            depth, formation, location = entry.depth_cm, entry.formation, entry.location
            if preview_headers:
                location = location or str(path.parent)
                if specimen.is_file():
                    try:
                        meta, _ = read_specimen(specimen, specimen_name=entry.specimen_name)
                        depth = depth or (f"{meta.volume:.2f}" if meta.volume else "")
                        formation = meta.comment or formation
                    except (OSError, ValueError):
                        pass
            for column, value in enumerate((entry.specimen_name, depth, formation, location)):
                self._table.setItem(row, column, QtWidgets.QTableWidgetItem(value))
        self._registrations = registrations
        return True

    def _on_accept(self) -> None:
        if not self.selected_sample:
            QtWidgets.QMessageBox.warning(self, "Select Sample", "Please select a sample row.")
            return
        try:
            self.queue_defaults
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Sample Index Defaults", str(exc))
            return
        self.accept()

    def _on_double_click(self) -> None:
        if self.selected_sample:
            self._on_accept()
