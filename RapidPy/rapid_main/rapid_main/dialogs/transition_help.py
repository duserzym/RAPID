from __future__ import annotations

from dataclasses import dataclass

from PySide6 import QtCore, QtWidgets


@dataclass(frozen=True)
class TransitionEntry:
    vb6: str
    task: str
    rapidpy: str
    destination: str
    status: str


TRANSITION_ENTRIES: tuple[TransitionEntry, ...] = (
    TransitionEntry("frmMagnetometerControl / modFlow", "View system state and control flow", "Dashboard", "dashboard", "Software-ready; hardware acceptance pending"),
    TransitionEntry("frmChanger* / frmSampleSelect", "Load specimens and manage the changer queue", "Sample Queue", "queue", "Integrated; physical changer acceptance pending"),
    TransitionEntry("frmProgram / frmRockmagRoutine", "Create, import, and save treatment sequences", "Sequence", "sequence", "Software-ready; treatment hardware acceptance pending"),
    TransitionEntry("frmMeasure / frmStats / frmPlots", "Run measurements and review live statistics/plots", "Live Measurement", "measure", "Code-complete core; physical acceptance pending"),
    TransitionEntry("frmSettings* / frmOptions / frmINIConverter", "Configure ports, paths, timing, and import VB6 INI", "Settings", "settings", "Integrated with backup/restore"),
    TransitionEntry("frm908AGaussmeter / calibration forms", "Record calibration evidence and plan thermal routines", "Calibration Center", "calibration", "Approval/version/expiry/rollback integrated; hardware acceptance pending"),
    TransitionEntry("frmDCMotors", "Inspect and control changer/lift/turn motors", "Diagnostics → DC Motors", "dc_motors", "Integrated; live motion acceptance pending"),
    TransitionEntry("frmVacuum", "Read pressure and control the vacuum path", "Diagnostics → Vacuum", "vacuum", "Integrated; live fault acceptance pending"),
    TransitionEntry("frmSquid", "Inspect SQUID communication and readings", "Diagnostics → SQUID Comm", "squid", "Integrated; serial robustness acceptance pending"),
    TransitionEntry("frmIRMARM / frmIRM_VoltageCalibration", "Control IRM/ARM and record voltage calibration", "Diagnostics → IRM / ARM", "irm", "Integrated foundation; live DAC acceptance pending"),
    TransitionEntry("frmAF_2G / AF treatment forms", "Set up an AF sequence or open AF tools", "Diagnostics → AF Demagnetizer", "af", "Integrated foundation; live rig acceptance pending"),
    TransitionEntry("frmDebug / modListenAndLog", "Review readiness and communication evidence", "View → Debug Console", "debug", "Integrated foundation"),
    TransitionEntry("frmStepMonitor", "Monitor or recover queue steps", "View → Step Monitor", "step_monitor", "Integrated; restart field acceptance pending"),
    TransitionEntry("frmVRM", "Launch VRM data collection with run handoff", "Diagnostics → VRM Data Collection", "vrm", "Integrated handoff; live acquisition acceptance pending"),
    TransitionEntry("frmWebcam", "Open the specimen webcam monitor", "View → Webcam Monitor", "webcam", "Standalone viewer integrated"),
    TransitionEntry("frmLogin", "Change the active operator session", "File → Log Out", "login", "Integrated foundation"),
    TransitionEntry("frmSplash / frmTip / frmAbout", "Open startup guidance or version information", "Help", "quick_start", "Integrated"),
)


class TransitionHelpDialog(QtWidgets.QDialog):
    """Searchable operator map from legacy VB6 tasks to RapidPy destinations."""

    destination_requested = QtCore.Signal(str)

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Where did this VB6 control go?")
        self.resize(820, 500)
        self.setMinimumSize(520, 320)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._populate()

    def _build_ui(self) -> None:
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        title = QtWidgets.QLabel("VB6 → RapidPy Task Map")
        title.setObjectName("sectionHdr")
        layout.addWidget(title)
        note = QtWidgets.QLabel(
            "Search by an old form name or operator task. Hardware-gated entries are "
            "available for software review but are not production-validated."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self._search = QtWidgets.QLineEdit()
        self._search.setPlaceholderText("Search frmMeasure, vacuum, sequence, calibration…")
        self._search.setClearButtonEnabled(True)
        self._search.textChanged.connect(self._apply_filter)
        layout.addWidget(self._search)

        self._table = QtWidgets.QTableWidget(0, 4)
        self._table.setHorizontalHeaderLabels(
            ["Legacy VB6", "Operator task", "RapidPy destination", "Readiness"]
        )
        self._table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self._table.doubleClicked.connect(self._open_selected)
        layout.addWidget(self._table, 1)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        self._open_button = buttons.addButton(
            "Open Destination", QtWidgets.QDialogButtonBox.ActionRole
        )
        self._open_button.clicked.connect(self._open_selected)
        buttons.rejected.connect(self.close)
        layout.addWidget(buttons)

    def _populate(self) -> None:
        self._table.setRowCount(0)
        for entry in TRANSITION_ENTRIES:
            row = self._table.rowCount()
            self._table.insertRow(row)
            values = (entry.vb6, entry.task, entry.rapidpy, entry.status)
            for column, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                item.setData(QtCore.Qt.UserRole, entry.destination)
                self._table.setItem(row, column, item)
        if self._table.rowCount():
            self._table.selectRow(0)

    def _apply_filter(self, text: str) -> None:
        query = text.casefold().strip()
        for row in range(self._table.rowCount()):
            haystack = " ".join(
                self._table.item(row, column).text()
                for column in range(self._table.columnCount())
            ).casefold()
            self._table.setRowHidden(row, bool(query) and query not in haystack)
        visible = next(
            (row for row in range(self._table.rowCount()) if not self._table.isRowHidden(row)),
            None,
        )
        if visible is not None:
            self._table.selectRow(visible)

    def _open_selected(self) -> None:
        selected = self._table.selectionModel().selectedRows()
        if not selected:
            return
        row = selected[0].row()
        if self._table.isRowHidden(row):
            return
        destination = self._table.item(row, 0).data(QtCore.Qt.UserRole)
        if destination:
            self.destination_requested.emit(str(destination))

