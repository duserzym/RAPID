from __future__ import annotations

import csv
import json
import os
import re
import uuid
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.printing import print_widget_snapshot
from rapid_main.queue_compiler import QueueOptions, QueueSample, validate_queue_samples


# Sample table column definitions
_COLS = ["#", "Position", "Sample Name", "Sample Set", "Treatment Steps", "Status", "Index File", "Orientation", "Both Sides"]
_QUEUE_FILE_SCHEMA = "rapidpy.sample_queue.v1"


class SampleQueuePanel(QtWidgets.QWidget):
    """Sample changer list and run options panel.

    Maps to: frmChanger (Hole Sample List) in VB6.
    Columns mirror the MSHFlexGrid with added Status column.
    Options sidebar mirrors the four VB6 FrameXxx option groups.
    """

    sample_index_requested = QtCore.Signal()

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

    # Toolbar
    def _build_toolbar(self) -> QtWidgets.QFrame:
        bar = QtWidgets.QFrame()
        bar.setObjectName("header")
        bar.setFixedHeight(48)
        hl = QtWidgets.QHBoxLayout(bar)
        hl.setContentsMargins(16, 0, 16, 0)
        hl.setSpacing(8)

        self._run_btn = QtWidgets.QPushButton("[▶]  Run Queue")
        self._run_btn.setObjectName("accent")
        self._run_btn.clicked.connect(self._on_run_queue)
        hl.addWidget(self._run_btn)

        self._pause_btn = QtWidgets.QPushButton("[||]  Pause")
        self._halt_btn = QtWidgets.QPushButton("[■]  Halt")
        self._pause_btn.clicked.connect(self._on_pause_queue)
        self._halt_btn.clicked.connect(self._on_halt_queue)
        hl.addWidget(self._pause_btn)
        hl.addWidget(self._halt_btn)

        hl.addWidget(_vline())

        self._rerun_btn = QtWidgets.QPushButton("[↻]  Rerun Failed")
        self._rerun_btn.clicked.connect(self._on_rerun_failed)
        hl.addWidget(self._rerun_btn)

        self._recovery_btn = QtWidgets.QPushButton("[R]  Recovery")
        recovery_menu = QtWidgets.QMenu(self._recovery_btn)
        recovery_menu.addAction("Resume interrupted", lambda: self._on_recovery_action("resume"))
        recovery_menu.addAction("Re-run interrupted", lambda: self._on_recovery_action("rerun"))
        recovery_menu.addAction("Skip interrupted", lambda: self._on_recovery_action("skip"))
        recovery_menu.addAction("Abort interrupted", lambda: self._on_recovery_action("abort"))
        self._recovery_btn.setMenu(recovery_menu)
        hl.addWidget(self._recovery_btn)
        hl.addWidget(_vline())

        self._add_btn = QtWidgets.QPushButton("[+]  Add Sample")
        self._load_index_btn = QtWidgets.QPushButton("[📂]  Load Index")
        self._seq_btn = QtWidgets.QPushButton("[→]  Sequential")
        clr_btn = QtWidgets.QPushButton("[x]  Clear")
        self._export_btn = QtWidgets.QPushButton("[↑]  Export")
        self._import_btn = QtWidgets.QPushButton("[↓]  Import")
        self._add_btn.clicked.connect(self._add_sample_dialog)
        self._load_index_btn.clicked.connect(self.sample_index_requested.emit)
        self._seq_btn.clicked.connect(self._make_positions_sequential)
        clr_btn.clicked.connect(self._clear_table)
        self._export_btn.clicked.connect(self._export_queue)
        self._import_btn.clicked.connect(self._import_queue)

        for btn in (
            self._add_btn,
            self._load_index_btn,
            self._seq_btn,
            clr_btn,
            self._export_btn,
            self._import_btn,
        ):
            hl.addWidget(btn)

        hl.addStretch()

        self._count_lbl = QtWidgets.QLabel("0 samples")
        self._count_lbl.setStyleSheet("color: #7a6f6e; font-size: 12px; padding-right: 8px;")
        hl.addWidget(self._count_lbl)
        return bar

    # Main sample table
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
        self._table.horizontalHeader().setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(6, 220)
        self._table.horizontalHeaderItem(8).setToolTip('Adds a Down pass after Up for a one-step file. Starting Down runs only Down.')
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

        self._update_count()
        return card

    def _add_sample_row(
        self,
        num: int,
        pos: str,
        name: str,
        sample_set: str,
        treatment: str,
        status: str,
    ) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for col, text in enumerate([str(num), pos, name, sample_set, treatment, status]):
            item = QtWidgets.QTableWidgetItem(text)
            if col in (0, 5):
                item.setTextAlignment(QtCore.Qt.AlignCenter)
            self._table.setItem(row, col, item)
        self._set_status(row, status)
        self._set_file_settings(row)

    def _set_file_settings(self, row: int, source_file: str = '', orientation: str = 'Up', both_sides: str = 'No') -> None:
        for column, value in zip((6, 7, 8), (source_file, orientation, both_sides)):
            item = QtWidgets.QTableWidgetItem(value)
            item.setToolTip(value if column == 6 else ('Up or Down' if column == 7 else 'Yes or No'))
            self._table.setItem(row, column, item)

    def _insert_sample_row(
        self,
        row: int,
        *,
        position: str,
        name: str,
        sample_set: str,
        treatment: str,
        status: str = "Pending",
        source_file: str = '',
        do_up: bool = True,
        do_both: bool = False,
    ) -> None:
        row = max(0, min(int(row), self._table.rowCount()))
        self._table.insertRow(row)
        values = [str(row + 1), position, name, sample_set, treatment, status]
        for col, text in enumerate(values):
            item = QtWidgets.QTableWidgetItem(text)
            if col in (0, 5):
                item.setTextAlignment(QtCore.Qt.AlignCenter)
            self._table.setItem(row, col, item)
        self._set_status(row, status)
        self._set_file_settings(row, source_file, 'Up' if do_up else 'Down', 'Yes' if do_both else 'No')
        self._renumber_rows()

    def _add_sample_dialog(self, *, insert_at: int | None = None) -> bool:
        row = self._table.rowCount() if insert_at is None else int(insert_at)
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Add queue sample")
        form = QtWidgets.QFormLayout(dialog)
        position = QtWidgets.QLineEdit(f"A{row + 1}")
        name = QtWidgets.QLineEdit()
        sample_set = QtWidgets.QLineEdit()
        treatment = QtWidgets.QLineEdit("NRM")
        source_file = QtWidgets.QLineEdit()
        orientation = QtWidgets.QComboBox()
        orientation.addItems(['Up', 'Down'])
        both_sides = QtWidgets.QCheckBox('Add a Down pass after Up (one step)')
        form.addRow("Changer position", position)
        form.addRow("Sample name", name)
        form.addRow("Sample set", sample_set)
        form.addRow("Treatment sequence", treatment)
        form.addRow('Index file (optional)', source_file)
        form.addRow('Initial orientation', orientation)
        form.addRow('', both_sides)
        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Ok
            | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        name.setFocus()
        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return False
        if not position.text().strip() or not name.text().strip():
            QtWidgets.QMessageBox.warning(
                self,
                "Add queue sample",
                "Changer position and sample name are required.",
            )
            return False
        self._insert_sample_row(
            row,
            position=position.text().strip(),
            name=name.text().strip(),
            sample_set=sample_set.text().strip(),
            treatment=treatment.text().strip() or "NRM",
            source_file=source_file.text().strip(),
            do_up=orientation.currentText() == 'Up',
            do_both=both_sides.isChecked(),
        )
        return True

    def add_sample(
        self,
        *,
        position: str,
        name: str,
        sample_set: str = "",
        treatment: str = "NRM",
        insert_at: int | None = None,
        source_file: str = '',
        do_up: bool = True,
        do_both: bool = False,
    ) -> int:
        """Add a validated operator-selected specimen and return its row."""

        clean_position = position.strip()
        clean_name = name.strip()
        if type(do_up) is not bool or type(do_both) is not bool:
            raise ValueError('Orientation settings must be explicit booleans.')
        if not clean_position or not clean_name:
            raise ValueError("Changer position and sample name are required.")
        row = self._table.rowCount() if insert_at is None else int(insert_at)
        self._insert_sample_row(
            row,
            position=clean_position,
            name=clean_name,
            sample_set=sample_set.strip(),
            treatment=treatment.strip() or "NRM",
            source_file=source_file.strip(), do_up=do_up, do_both=do_both,
        )
        return max(0, min(row, self._table.rowCount() - 1))

    def _make_positions_sequential(self) -> bool:
        if self._table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, "Sequential positions", "The queue is empty.")
            return False
        start, accepted = QtWidgets.QInputDialog.getInt(
            self,
            "Sequential positions",
            "First changer hole",
            1,
            1,
            9999,
        )
        if not accepted:
            return False
        self.apply_sequential_positions(start)
        return True

    def apply_sequential_positions(self, start: int = 1, prefix: str = "A") -> None:
        """Assign deterministic changer positions while preserving queue content."""
        for row in range(self._table.rowCount()):
            self._table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{prefix}{start + row}"))
        self._renumber_rows()

    def _export_queue(self) -> Path | None:
        filename, _selected = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export sample queue",
            "rapid_queue.json",
            "RAPID queue (*.json);;CSV queue (*.csv)",
        )
        if not filename:
            return None
        path = Path(filename)
        try:
            write_queue_file(path, self.row_snapshot())
        except (OSError, ValueError) as exc:
            QtWidgets.QMessageBox.critical(self, "Export sample queue", str(exc))
            return None
        return path

    def _import_queue(self) -> Path | None:
        filename, _selected = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Import sample queue",
            "",
            "RAPID queue (*.json *.csv);;All files (*)",
        )
        if not filename:
            return None
        path = Path(filename)
        try:
            rows = read_queue_file(path)
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            QtWidgets.QMessageBox.critical(self, "Import sample queue", str(exc))
            return None
        self.load_rows(rows)
        return path

    def _clear_table(self) -> None:
        self._table.setRowCount(0)
        self._update_count()

    def _update_count(self) -> None:
        n = self._table.rowCount()
        self._count_lbl.setText(f"{n} sample{'s' if n != 1 else ''}")

    def _context_menu(self, pos: QtCore.QPoint) -> None:
        row = self._table.rowAt(pos.y())
        if row < 0:
            return
        menu = QtWidgets.QMenu(self)
        menu.addAction("Sample Info", lambda: self._show_sample_info(row))
        menu.addSeparator()
        menu.addAction("Insert Sample Above", lambda: self._add_sample_dialog(insert_at=row))
        menu.addAction("Delete", lambda: self._table.removeRow(row))
        menu.addAction("Delete without Gap", lambda: self._table.removeRow(row))
        menu.addSeparator()
        menu.addAction("Mark Pending", lambda: self._set_status(row, "Pending"))
        menu.addAction('File settings…', lambda: self.edit_file_settings(row))
        menu.addAction("Mark Done", lambda: self._set_status(row, "Done"))
        menu.addAction("Mark Skipped", lambda: self._set_status(row, "Skipped"))
        menu.addAction("Mark Error", lambda: self._set_status(row, "Error"))
        if self._safe_cell(row, 5) == "Interrupted":
            menu.addSeparator()
            menu.addAction("Resume interrupted sample", lambda: self.apply_recovery_to_row(row, "resume"))
            menu.addAction("Re-run interrupted sample", lambda: self.apply_recovery_to_row(row, "rerun"))
            menu.addAction("Skip interrupted sample", lambda: self.apply_recovery_to_row(row, "skip"))
            menu.addAction("Abort interrupted sample", lambda: self.apply_recovery_to_row(row, "abort"))
        menu.addAction("Delete next 9 samples", lambda: self._delete_row_span(row, 10))
        menu.exec(self._table.viewport().mapToGlobal(pos))

    def _show_sample_info(self, row: int) -> None:
        QtWidgets.QMessageBox.information(
            self,
            "Queue sample",
            "\n".join(
                f"{column}: {self._safe_cell(row, index)}"
                for index, column in enumerate(_COLS[1:], start=1)
            ),
        )

    def edit_file_settings(self, row: int) -> bool:
        """Apply orientation settings to every queued row of the selected file."""
        if not 0 <= row < self._table.rowCount():
            return False
        def identity(index):
            source = self._safe_cell(index, 6)
            return ('index', os.path.normcase(str(Path(source).resolve()))) if source else ('manual', self._safe_cell(index, 3) or 'SampleSet')
        key = identity(row)
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('Queue file settings')
        form = QtWidgets.QFormLayout(dialog)
        description = QtWidgets.QLabel(self._safe_cell(row, 6) or self._safe_cell(row, 3) or 'SampleSet')
        description.setWordWrap(True)
        form.addRow('File', description)
        orientation = QtWidgets.QComboBox(dialog)
        orientation.addItems(['Up', 'Down'])
        orientation.setCurrentText(self._safe_cell(row, 7) or 'Up')
        both = QtWidgets.QCheckBox('Add a Down pass after Up (one step)', dialog)
        both.setChecked(self._safe_cell(row, 8) == 'Yes')
        form.addRow('Initial orientation', orientation)
        form.addRow('', both)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return False
        for index in range(self._table.rowCount()):
            if identity(index) == key:
                self._set_file_settings(index, self._safe_cell(index, 6), orientation.currentText(), 'Yes' if both.isChecked() else 'No')
        return True

    def _delete_row_span(self, start: int, count: int) -> None:
        for _ in range(max(0, int(count))):
            if start >= self._table.rowCount():
                break
            self._table.removeRow(start)
        self._renumber_rows()

    def _on_run_queue(self) -> None:
        rows, errors = _read_queue_rows(self._table)
        validation = validate_queue_samples(rows)
        validation_errors = list(errors) + list(validation.errors)
        if validation_errors:
            QtWidgets.QMessageBox.critical(
                self,
                "Invalid queue",
                "Cannot start queue because:\n\n"
                + "\n".join(f"• {e}" for e in validation_errors),
            )
            return
        if validation.warnings:
            if (
                QtWidgets.QMessageBox.question(
                    self,
                    "Queue validation warnings",
                    "The queue has warnings:\n\n"
                    + "\n".join(f"• {w}" for w in validation.warnings)
                    + "\n\nContinue anyway?",
                    QtWidgets.QMessageBox.StandardButton.Yes
                    | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No,
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return

        options = _queue_options_from_ui(
            self._order_radios,
            self._reload_radios,
            self._final_radios,
            self._holder_radios,
        )
        owner = self.window()
        if hasattr(owner, "start_queue_run"):
            owner.start_queue_run(rows, options)
        else:
            QtWidgets.QMessageBox.critical(
                self,
                "Cannot run queue",
                "Window controller missing.",
            )

    def _on_pause_queue(self) -> None:
        owner = self.window()
        if hasattr(owner, "toggle_queue_pause"):
            owner.toggle_queue_pause()
        elif hasattr(owner, "toggle_measurement_pause"):
            owner.toggle_measurement_pause()
        else:
            QtWidgets.QMessageBox.information(self, "Pause", "No active measurement to pause.")

    def _on_halt_queue(self) -> None:
        owner = self.window()
        if hasattr(owner, "halt_measurement"):
            owner.halt_measurement()
        elif hasattr(owner, "cancel_queue_run"):
            owner.cancel_queue_run("Queue halt requested.")
        else:
            QtWidgets.QMessageBox.information(self, "Halt", "No active run to halt.")

    def _on_rerun_failed(self) -> None:
        count = self.rerun_failed_samples()
        if count == 0:
            QtWidgets.QMessageBox.information(self, "Rerun Failed", "No failed samples to rerun.")
            return
        QtWidgets.QMessageBox.information(
            self,
            "Rerun Failed",
            f"{count} sample{'s' if count != 1 else ''} set to Pending.",
        )

    def _on_recovery_action(self, action: str) -> None:
        count = self.apply_interrupted_recovery(action)
        label = _recovery_label(action)
        if count == 0:
            QtWidgets.QMessageBox.information(
                self,
                "Interrupted Queue Recovery",
                "No interrupted samples require recovery.",
            )
            return
        QtWidgets.QMessageBox.information(
            self,
            "Interrupted Queue Recovery",
            f"{count} interrupted sample{'s' if count != 1 else ''} set to {label}.",
        )

    def start_queue_sample(self, sample_name: str) -> None:
        """Set the selected sample to Running state."""
        for row in range(self._table.rowCount()):
            if self._table.item(row, 2).text() == sample_name:
                self._set_status(row, "Running")
                return

    def set_queue_command_status(self, command, status: str) -> bool:
        """Update only the original compiled row, including repeated side passes."""
        for row in range(self._table.rowCount()):
            number = self._table.item(row, 0)
            if command.row_id:
                matches = number is not None and number.data(QtCore.Qt.UserRole) == command.row_id
            else:
                source = self._safe_cell(row, 6)
                identity = os.path.normcase(str(Path(source).resolve())) if source else self._safe_cell(row, 3) or 'SampleSet'
                hole, error = _parse_hole(self._safe_cell(row, 1))
                matches = error is None and hole == command.hole and identity == command.file_id and self._safe_cell(row, 2) == command.sample_name
            if matches:
                self._set_status(row, status)
                return True
        return False

    def mark_queue_sample_done(self, sample_name: str) -> None:
        """Set the selected sample to Done state."""
        for row in range(self._table.rowCount()):
            if self._table.item(row, 2).text() == sample_name:
                self._set_status(row, "Done")
                return

    def set_queue_sample_failed(self, sample_name: str) -> None:
        """Set the selected sample to Error state."""
        for row in range(self._table.rowCount()):
            if self._table.item(row, 2).text() == sample_name:
                self._set_status(row, "Error")
                return

    def rerun_failed_samples(self) -> int:
        """Reset all Error rows to Pending and return count changed."""
        changed = 0
        for row in range(self._table.rowCount()):
            if self._table.item(row, 5) is None:
                continue
            status = self._table.item(row, 5).text()
            if status == "Error":
                self._set_status(row, "Pending")
                changed += 1
        return changed

    def recover_interrupted_samples(self) -> int:
        """Mark lingering Running rows as Interrupted and return count changed."""
        recovered = 0
        for row in range(self._table.rowCount()):
            if self._table.item(row, 5) is None:
                continue
            status = self._table.item(row, 5).text()
            if status == "Running":
                self._set_status(row, "Interrupted")
                recovered += 1
        return recovered

    def interrupted_sample_names(self) -> list[str]:
        """Return sample names still waiting for an explicit recovery decision."""
        names: list[str] = []
        for row in range(self._table.rowCount()):
            if self._safe_cell(row, 5) == "Interrupted":
                names.append(self._safe_cell(row, 2))
        return names

    def apply_interrupted_recovery(self, action: str) -> int:
        """Apply a recovery action to all interrupted rows."""
        changed = 0
        for row in range(self._table.rowCount()):
            if self.apply_recovery_to_row(row, action):
                changed += 1
        return changed

    def apply_recovery_to_row(self, row: int, action: str) -> bool:
        """Apply an explicit resume/rerun/skip/abort choice to one interrupted row."""
        if row < 0 or row >= self._table.rowCount():
            return False
        if self._safe_cell(row, 5) != "Interrupted":
            return False
        self._set_status(row, _status_for_recovery_action(action))
        return True

    def skip_selected_sample(self, sample_name: str) -> None:
        """Set the selected sample to Skipped."""
        for row in range(self._table.rowCount()):
            if self._table.item(row, 2).text() == sample_name:
                self._set_status(row, "Skipped")
                return

    def row_snapshot(self) -> list[dict[str, str]]:
        """Serialize the queue table for persistence."""
        rows: list[dict[str, str]] = []
        for row in range(self._table.rowCount()):
            rows.append({
                "position": self._safe_cell(row, 1),
                "sample_name": self._safe_cell(row, 2),
                "sample_set": self._safe_cell(row, 3),
                "treatment": self._safe_cell(row, 4),
                "status": self._safe_cell(row, 5),
            })
            source_file, orientation, both_sides = (self._safe_cell(row, column) for column in (6, 7, 8))
            if source_file or orientation not in {'', 'Up'} or both_sides not in {'', 'No'}:
                rows[-1].update(source_file=source_file, orientation=orientation, both_sides=both_sides)
            number = self._table.item(row, 0)
            if number is not None and number.data(QtCore.Qt.UserRole):
                rows[-1]['row_id'] = number.data(QtCore.Qt.UserRole)
        return rows

    def load_rows(self, rows: list[dict[str, str]]) -> None:
        """Load queue table rows from persisted data."""
        if not rows:
            self._clear_table()
            return
        self._table.setRowCount(0)
        for row in rows:
            num = self._safe_cell_from_map(row, "position", "")
            sample_name = self._safe_cell_from_map(row, "sample_name", "")
            sample_set = self._safe_cell_from_map(row, "sample_set", "")
            treatment = self._safe_cell_from_map(row, "treatment", "")
            status = self._safe_cell_from_map(row, "status", "Pending")
            self._append_row(
                num,
                sample_name,
                sample_set,
                treatment,
                status,
            )
            self._set_file_settings(self._table.rowCount() - 1,
                self._safe_cell_from_map(row, 'source_file', ''),
                self._safe_cell_from_map(row, 'orientation', 'Up') or 'Up',
                self._safe_cell_from_map(row, 'both_sides', 'No') or 'No')
            if row.get('row_id'):
                self._table.item(self._table.rowCount() - 1, 0).setData(QtCore.Qt.UserRole, row['row_id'])
        self._renumber_rows()

    def _append_row(
        self,
        pos: str,
        name: str,
        sample_set: str,
        treatment: str,
        status: str,
    ) -> None:
        self._add_sample_row(
            0,
            pos,
            name,
            sample_set,
            treatment,
            _normalize_status(status),
        )

    def _renumber_rows(self) -> None:
        for row in range(self._table.rowCount()):
            if self._table.item(row, 0) is not None:
                self._table.item(row, 0).setText(str(row + 1))

    def _safe_cell_from_map(self, row: dict[str, str], key: str, default: str) -> str:
        value = row.get(key, default)
        return str(value) if value is not None else default

    def _safe_cell(self, row: int, col: int) -> str:
        item = self._table.item(row, col)
        if item is None:
            return ""
        return item.text().strip()

    def _set_status(self, row: int, status: str) -> None:
        item = self._table.item(row, 5)
        if item is None:
            return
        normalized = _normalize_status(status)
        item.setText(normalized)
        if normalized == "Pending":
            item.setForeground(QtGui.QBrush(QtCore.Qt.GlobalColor.darkGray))
        elif normalized == "Running":
            item.setForeground(
                QtGui.QBrush(QtCore.Qt.GlobalColor.darkGreen)
            )
        elif normalized == "Done":
            item.setForeground(
                QtGui.QBrush(QtCore.Qt.GlobalColor.darkGreen)
            )
        elif normalized == "Error":
            item.setForeground(
                QtGui.QBrush(QtCore.Qt.GlobalColor.red)
            )
        elif normalized == "Skipped":
            item.setForeground(
                QtGui.QBrush(QtCore.Qt.GlobalColor.darkYellow)
            )
        elif normalized == "Interrupted":
            item.setForeground(
                QtGui.QBrush(QtGui.QColor("#b45309"))
            )
        elif normalized == "Resume":
            item.setForeground(
                QtGui.QBrush(QtGui.QColor("#0369a1"))
            )
        elif normalized == "Rerun":
            item.setForeground(
                QtGui.QBrush(QtGui.QColor("#7c3aed"))
            )
        elif normalized == "Aborted":
            item.setForeground(
                QtGui.QBrush(QtCore.Qt.GlobalColor.red)
            )
        else:
            item.setForeground(QtGui.QBrush(QtCore.Qt.GlobalColor.black))

    # Options sidebar
    def _build_options_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(14)

        hdr = QtWidgets.QLabel("RUN OPTIONS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        def _option_group(title: str, options: list[str], default: int = 0) -> tuple[
            QtWidgets.QFrame, list[QtWidgets.QRadioButton]
        ]:
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

        grp1, self._order_radios = _option_group("Sample Order", ["Ascending", "Descending"], default=0)
        grp2, self._reload_radios = _option_group("Reload Position", ["Return to start", "Leave at end"], default=0)
        grp3, self._final_radios = _option_group("Final Position", ["Return to start", "Leave at end"], default=0)
        grp4, self._holder_radios = _option_group(
            "Multiple Holder Measurements",
            ["Repeat  (weak samples)", "Skip  (strong samples)"],
            default=0,
        )
        for grp in (grp1, grp2, grp3, grp4):
            cl.addWidget(grp)

        cl.addStretch()

        self._print_btn = QtWidgets.QPushButton("[🖨]  Print List")
        self._print_btn.clicked.connect(self._print_queue)
        cl.addWidget(self._print_btn)
        return card

    def _print_queue(self) -> bool:
        return print_widget_snapshot(
            self,
            self._table,
            title="RAPID sample queue",
        )


def write_queue_file(path: Path | str, rows: list[dict[str, str]]) -> Path:
    """Write a portable queue document using an atomic same-directory replace."""

    destination = Path(path)
    if destination.suffix.lower() not in {".json", ".csv"}:
        destination = destination.with_suffix(".json")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(destination.name + ".tmp")
    fields = ("position", "sample_name", "sample_set", "treatment", "status", "source_file", "orientation", "both_sides", "row_id")
    try:
        if destination.suffix.lower() == ".csv":
            with temporary.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fields)
                writer.writeheader()
                for row in rows:
                    writer.writerow({field: str(row.get(field, "")) for field in fields})
        else:
            temporary.write_text(
                json.dumps(
                    {"schema": _QUEUE_FILE_SCHEMA, "rows": rows},
                    indent=2,
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    return destination


def read_queue_file(path: Path | str) -> list[dict[str, str]]:
    """Read and validate JSON/CSV queue rows without mutating the active table."""

    source = Path(path)
    fields = ("position", "sample_name", "sample_set", "treatment", "status")
    if source.suffix.lower() == ".csv":
        with source.open("r", encoding="utf-8-sig", newline="") as handle:
            raw_rows: object = list(csv.DictReader(handle))
    else:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict) or payload.get("schema") != _QUEUE_FILE_SCHEMA:
            raise ValueError(f"Unsupported queue file schema in {source.name}.")
        raw_rows = payload.get("rows")
    if not isinstance(raw_rows, list):
        raise ValueError(f"Queue file {source.name} does not contain a row list.")

    rows: list[dict[str, str]] = []
    for index, raw in enumerate(raw_rows, start=1):
        if not isinstance(raw, dict):
            raise ValueError(f"Queue row {index} is not an object.")
        row = {field: str(raw.get(field, "") or "").strip() for field in fields}
        if not row["position"] or not row["sample_name"]:
            raise ValueError(f"Queue row {index} requires position and sample_name.")
        row["treatment"] = row["treatment"] or "NRM"
        row["status"] = _normalize_status(row["status"])
        for field in ('source_file', 'orientation', 'both_sides', 'row_id'):
            value = str(raw.get(field, '') or '').strip()
            if value:
                row[field] = value
        rows.append(row)
    return rows


def _vline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.VLine)
    f.setStyleSheet("color: rgba(122,2,25,0.18); margin: 8px 2px;")
    return f


def _parse_hole(position_text: str) -> tuple[int | None, str | None]:
    if not position_text:
        return None, "position is empty"
    match = re.search(r"(\d+)", position_text.strip())
    if not match:
        return None, f"position '{position_text}' does not contain a valid hole index"
    try:
        hole = int(match.group(1))
    except ValueError:
        return None, f"position '{position_text}' could not be parsed"
    if hole < 1:
        return None, f"position '{position_text}' must be >= 1"
    return hole, None


def _parse_step_count(treatment_text: str) -> int:
    if not treatment_text:
        return 1
    return len(re.split(r'→|->', treatment_text))


def _normalize_status(status: str) -> str:
    status = (status or "Pending").strip().title()
    if status in {"Pending", "Running", "Done", "Error", "Skipped", "Interrupted", "Resume", "Rerun", "Aborted"}:
        return status
    if status in {"Queued", "Queued (Ready)"}:
        return "Pending"
    if status in {"Running...", "Running (Queued)"}:
        return "Running"
    return "Pending"


def _status_for_recovery_action(action: str) -> str:
    normalized = (action or "").strip().lower().replace("-", "").replace("_", "")
    if normalized == "resume":
        return "Resume"
    if normalized in {"rerun", "replay"}:
        return "Rerun"
    if normalized == "skip":
        return "Skipped"
    if normalized == "abort":
        return "Aborted"
    raise ValueError(f"Unknown interrupted recovery action '{action}'.")


def _recovery_label(action: str) -> str:
    return _status_for_recovery_action(action)


def _read_queue_rows(table: QtWidgets.QTableWidget) -> tuple[list[QueueSample], list[str]]:
    samples: list[QueueSample] = []
    errors: list[str] = []

    for row in range(table.rowCount()):
        def _cell(i: int) -> str:
            item = table.item(row, i)
            return item.text().strip() if item is not None else ""

        row_no = row + 1
        status = _cell(5)
        if status in {"Done", "Skipped", "Aborted"}:
            continue
        if status == "Interrupted":
            errors.append(
                f"row {row_no}: interrupted sample requires Resume, Re-run, Skip, or Abort before queue start"
            )
            continue

        position = _cell(1)
        sample_name = _cell(2)
        sample_set = _cell(3) or "SampleSet"
        treatment = _cell(4)
        source_file, orientation, both_sides = _cell(6), _cell(7) or 'Up', _cell(8) or 'No'
        if orientation not in {'Up', 'Down'} or both_sides not in {'Yes', 'No'}:
            errors.append(f'row {row_no}: orientation must be Up/Down and Both Sides must be Yes/No')
            continue
        if source_file:
            if not Path(source_file).is_absolute():
                errors.append(f'row {row_no}: index file must be an absolute path')
                continue
            file_id = os.path.normcase(str(Path(source_file).resolve()))
        else:
            file_id = sample_set

        hole, hole_error = _parse_hole(position)
        if hole_error is not None:
            errors.append(f"row {row_no}: {hole_error}")
            continue

        if not sample_name:
            # defer to queue validator for consistent formatting
            sample_name = ""

        labels = tuple(part.strip() for part in re.split(r'→|->', treatment or 'NRM'))
        if any(not label for label in labels):
            errors.append(f'row {row_no}: treatment sequence contains an empty step')
            continue
        number = table.item(row, 0)
        row_id = number.data(QtCore.Qt.UserRole) if number is not None else None
        if not row_id:
            row_id = uuid.uuid4().hex
            if number is not None: number.setData(QtCore.Qt.UserRole, row_id)
        samples.append(
            QueueSample(
                sample_name=sample_name,
                file_id=file_id,
                hole=hole,
                do_up=orientation == 'Up',
                do_both=both_sides == 'Yes',
                measurement_step_count=_parse_step_count(treatment),
                measurement_labels=labels,
                source_file=str(Path(source_file).resolve()) if source_file else '',
                row_id=row_id,
            )
        )

    return samples, errors


def _queue_options_from_ui(
    order_radios: list[QtWidgets.QRadioButton],
    reload_radios: list[QtWidgets.QRadioButton],
    final_radios: list[QtWidgets.QRadioButton],
    holder_radios: list[QtWidgets.QRadioButton],
) -> QueueOptions:
    return QueueOptions(
        ascending=order_radios[0].isChecked(),
        load_return=reload_radios[0].isChecked(),
        do_return=final_radios[0].isChecked(),
        repeat_holder=holder_radios[0].isChecked(),
    )
