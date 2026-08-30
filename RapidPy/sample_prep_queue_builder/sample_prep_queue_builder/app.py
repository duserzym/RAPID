from __future__ import annotations

import json
import sys
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.io.cit_sam import CitSamEntry, CitSamHeader, read_cit_sam, write_cit_sam
from rapid_main.queue_compiler import QueueOptions, QueueSample, compile_queue
from rapidpy_common.ui import apply_liquid_glass_theme, apply_window_bounds_guard, set_app_icon


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Sample Prep and Queue Builder")
        self.resize(1360, 860)

        host = QtWidgets.QWidget()
        self.setCentralWidget(host)
        root = QtWidgets.QVBoxLayout(host)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        root.addWidget(self._build_inputs())

        split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        split.addWidget(self._build_samples_card())
        split.addWidget(self._build_queue_card())
        split.setSizes([760, 580])
        root.addWidget(split, 1)

        root.addWidget(self._build_log_card())

    def _build_inputs(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        gl = QtWidgets.QGridLayout(card)
        gl.setContentsMargins(12, 12, 12, 12)
        gl.setHorizontalSpacing(8)
        gl.setVerticalSpacing(8)

        hdr = QtWidgets.QLabel("CIT SAM Inputs")
        hdr.setObjectName("sectionHdr")
        gl.addWidget(hdr, 0, 0, 1, 6)

        gl.addWidget(QtWidgets.QLabel("SAM file"), 1, 0)
        self.sam_path = QtWidgets.QLineEdit()
        self.sam_path.setPlaceholderText("Select a CIT-format .sam file")
        gl.addWidget(self.sam_path, 1, 1, 1, 3)
        browse_sam = QtWidgets.QPushButton("Browse")
        browse_sam.clicked.connect(self._browse_sam)
        gl.addWidget(browse_sam, 1, 4)
        load_sam = QtWidgets.QPushButton("Load")
        load_sam.setObjectName("accent")
        load_sam.clicked.connect(self._load_sam)
        gl.addWidget(load_sam, 1, 5)

        gl.addWidget(QtWidgets.QLabel("Output folder"), 2, 0)
        self.output_dir = QtWidgets.QLineEdit(str(Path.cwd() / "sample_prep_output"))
        gl.addWidget(self.output_dir, 2, 1, 1, 3)
        browse_out = QtWidgets.QPushButton("Browse")
        browse_out.clicked.connect(self._browse_output)
        gl.addWidget(browse_out, 2, 4)

        self.opt_ascending = QtWidgets.QCheckBox("Ascending order")
        self.opt_ascending.setChecked(True)
        self.opt_load_return = QtWidgets.QCheckBox("Load return")
        self.opt_load_return.setChecked(True)
        self.opt_do_return = QtWidgets.QCheckBox("Final return")
        self.opt_do_return.setChecked(True)
        self.opt_repeat_holder = QtWidgets.QCheckBox("Repeat holder")
        self.opt_repeat_holder.setChecked(True)
        self.opt_xy = QtWidgets.QCheckBox("Use XY table load position")
        self.opt_xy.setChecked(True)

        gl.addWidget(self.opt_ascending, 3, 1)
        gl.addWidget(self.opt_load_return, 3, 2)
        gl.addWidget(self.opt_do_return, 3, 3)
        gl.addWidget(self.opt_repeat_holder, 3, 4)
        gl.addWidget(self.opt_xy, 3, 5)

        gl.addWidget(QtWidgets.QLabel("Samples between holder"), 4, 0)
        self.samples_between_holder = QtWidgets.QSpinBox()
        self.samples_between_holder.setRange(1, 100)
        self.samples_between_holder.setValue(8)
        gl.addWidget(self.samples_between_holder, 4, 1)

        compile_btn = QtWidgets.QPushButton("Compile Queue")
        compile_btn.setObjectName("accent")
        compile_btn.clicked.connect(self._compile_queue)
        save_btn = QtWidgets.QPushButton("Export Queue JSON")
        save_btn.clicked.connect(self._save_queue_json)
        save_sam_btn = QtWidgets.QPushButton("Save SAM Copy")
        save_sam_btn.clicked.connect(self._save_sam_copy)
        gl.addWidget(compile_btn, 4, 3)
        gl.addWidget(save_btn, 4, 4)
        gl.addWidget(save_sam_btn, 4, 5)

        return card

    def _build_samples_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        vl = QtWidgets.QVBoxLayout(card)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        lbl = QtWidgets.QLabel("Sample Assignments")
        lbl.setObjectName("sectionHdr")
        vl.addWidget(lbl)

        self.samples = QtWidgets.QTableWidget(0, 6)
        self.samples.setHorizontalHeaderLabels(["Hole", "Sample", "File ID", "Do Up", "Do Both", "Steps"])
        self.samples.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.samples.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.samples.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.samples.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        self.samples.horizontalHeader().setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.samples.horizontalHeader().setSectionResizeMode(5, QtWidgets.QHeaderView.ResizeToContents)
        self.samples.verticalHeader().setVisible(False)
        self.samples.setAlternatingRowColors(True)
        vl.addWidget(self.samples, 1)

        return card

    def _build_queue_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        vl = QtWidgets.QVBoxLayout(card)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        lbl = QtWidgets.QLabel("Compiled Queue")
        lbl.setObjectName("sectionHdr")
        vl.addWidget(lbl)

        self.queue = QtWidgets.QTableWidget(0, 5)
        self.queue.setHorizontalHeaderLabels(["#", "Command", "Hole", "File ID", "Sample"])
        self.queue.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.queue.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        self.queue.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.queue.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        self.queue.horizontalHeader().setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        self.queue.verticalHeader().setVisible(False)
        self.queue.setAlternatingRowColors(True)
        self.queue.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        vl.addWidget(self.queue, 1)

        return card

    def _build_log_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        vl = QtWidgets.QVBoxLayout(card)
        vl.setContentsMargins(12, 10, 12, 10)
        vl.setSpacing(6)
        lbl = QtWidgets.QLabel("Run log")
        lbl.setObjectName("sectionHdr")
        vl.addWidget(lbl)
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(120)
        vl.addWidget(self.log)
        return card

    def _append(self, text: str) -> None:
        self.log.appendPlainText(text)

    def _browse_sam(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select .sam file", "", "SAM files (*.sam);;All files (*)")
        if path:
            self.sam_path.setText(path)

    def _browse_output(self) -> None:
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select output folder", self.output_dir.text().strip())
        if path:
            self.output_dir.setText(path)

    def _load_sam(self) -> None:
        path = Path(self.sam_path.text().strip())
        if not path.exists():
            QtWidgets.QMessageBox.warning(self, "Load SAM", "Select a valid .sam file first.")
            return
        try:
            header, entries = read_cit_sam(path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load SAM", str(exc))
            return

        self.samples.setRowCount(0)
        file_id = path.stem
        for idx, entry in enumerate(entries, start=1):
            row = self.samples.rowCount()
            self.samples.insertRow(row)
            self.samples.setItem(row, 0, QtWidgets.QTableWidgetItem(str(idx)))
            self.samples.setItem(row, 1, QtWidgets.QTableWidgetItem(entry.specimen_name))
            self.samples.setItem(row, 2, QtWidgets.QTableWidgetItem(file_id))

            do_up = QtWidgets.QTableWidgetItem()
            do_up.setFlags(do_up.flags() | QtCore.Qt.ItemIsUserCheckable)
            do_up.setCheckState(QtCore.Qt.Checked)
            self.samples.setItem(row, 3, do_up)

            do_both = QtWidgets.QTableWidgetItem()
            do_both.setFlags(do_both.flags() | QtCore.Qt.ItemIsUserCheckable)
            do_both.setCheckState(QtCore.Qt.Unchecked)
            self.samples.setItem(row, 4, do_both)

            self.samples.setItem(row, 5, QtWidgets.QTableWidgetItem("1"))

        self._sam_header = header
        self._sam_entries = entries
        self._append(f"Loaded {len(entries)} samples from {path.name} (format={header.format_id}).")

    def _collect_samples(self) -> list[QueueSample]:
        rows: list[QueueSample] = []
        for row in range(self.samples.rowCount()):
            try:
                hole = int(self.samples.item(row, 0).text())
                sample = (self.samples.item(row, 1).text() if self.samples.item(row, 1) else "").strip()
                file_id = (self.samples.item(row, 2).text() if self.samples.item(row, 2) else "").strip()
                do_up = self.samples.item(row, 3).checkState() == QtCore.Qt.Checked
                do_both = self.samples.item(row, 4).checkState() == QtCore.Qt.Checked
                step_count = int((self.samples.item(row, 5).text() if self.samples.item(row, 5) else "1").strip() or "1")
            except Exception:
                continue
            if sample:
                rows.append(
                    QueueSample(
                        sample_name=sample,
                        file_id=file_id,
                        hole=hole,
                        do_up=do_up,
                        do_both=do_both,
                        measurement_step_count=step_count,
                    )
                )
        return rows

    def _compile_queue(self) -> None:
        samples = self._collect_samples()
        if not samples:
            QtWidgets.QMessageBox.warning(self, "Compile Queue", "No valid sample rows to compile.")
            return

        options = QueueOptions(
            ascending=self.opt_ascending.isChecked(),
            load_return=self.opt_load_return.isChecked(),
            do_return=self.opt_do_return.isChecked(),
            repeat_holder=self.opt_repeat_holder.isChecked(),
            samples_between_holder=int(self.samples_between_holder.value()),
            use_xy_table=self.opt_xy.isChecked(),
        )
        queue = compile_queue(samples, options)
        self._compiled_queue = queue

        self.queue.setRowCount(0)
        for idx, cmd in enumerate(queue, start=1):
            row = self.queue.rowCount()
            self.queue.insertRow(row)
            self.queue.setItem(row, 0, QtWidgets.QTableWidgetItem(str(idx)))
            self.queue.setItem(row, 1, QtWidgets.QTableWidgetItem(cmd.command_type))
            self.queue.setItem(row, 2, QtWidgets.QTableWidgetItem(str(cmd.hole)))
            self.queue.setItem(row, 3, QtWidgets.QTableWidgetItem(cmd.file_id))
            self.queue.setItem(row, 4, QtWidgets.QTableWidgetItem(cmd.sample_name))

        self._append(f"Compiled queue with {len(queue)} commands from {len(samples)} samples.")

    def _save_queue_json(self) -> None:
        queue = getattr(self, "_compiled_queue", None)
        if not queue:
            QtWidgets.QMessageBox.warning(self, "Export Queue", "Compile a queue first.")
            return

        out_dir = Path(self.output_dir.text().strip() or ".")
        out_dir.mkdir(parents=True, exist_ok=True)
        path = out_dir / "compiled_queue.json"
        payload = [
            {
                "command_type": cmd.command_type,
                "hole": cmd.hole,
                "file_id": cmd.file_id,
                "sample_name": cmd.sample_name,
            }
            for cmd in queue
        ]
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        self._append(f"Wrote {path}")

    def _save_sam_copy(self) -> None:
        if not hasattr(self, "_sam_header"):
            QtWidgets.QMessageBox.warning(self, "Save SAM", "Load a SAM file first.")
            return

        out_dir = Path(self.output_dir.text().strip() or ".")
        out_dir.mkdir(parents=True, exist_ok=True)
        path = out_dir / "updated_locality.sam"

        entries: list[CitSamEntry] = []
        for row in range(self.samples.rowCount()):
            specimen = (self.samples.item(row, 1).text() if self.samples.item(row, 1) else "").strip()
            if specimen:
                entries.append(CitSamEntry(specimen_name=specimen))

        header: CitSamHeader = self._sam_header
        write_cit_sam(path, header, entries, include_format_line=True)
        self._append(f"Wrote {path}")


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "sample_prep_queue_builder_icon.png", assets_dir)
    window = MainWindow()
    set_app_icon(window, "sample_prep_queue_builder_icon.png", assets_dir)
    window.show()
    return app.exec()
