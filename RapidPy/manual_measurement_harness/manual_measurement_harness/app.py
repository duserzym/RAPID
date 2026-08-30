from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PySide6 import QtCore, QtWidgets

from rapid_main.data_model import SpecimenMeta
from rapid_main.measurement_worker import MeasurementWorker, NoCommBackend, StepResult
from rapidpy_common.ui import apply_liquid_glass_theme, apply_window_bounds_guard, set_app_icon


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Manual Measurement and Holder Workflow Harness")
        self.resize(980, 700)
        self._worker: MeasurementWorker | None = None

        host = QtWidgets.QWidget()
        self.setCentralWidget(host)
        root = QtWidgets.QVBoxLayout(host)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        root.addWidget(self._build_config_card())
        root.addWidget(self._build_actions_card())
        root.addWidget(self._build_log_card(), 1)

    def _build_config_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        gl = QtWidgets.QGridLayout(card)
        gl.setContentsMargins(12, 12, 12, 12)
        gl.setHorizontalSpacing(8)
        gl.setVerticalSpacing(8)

        hdr = QtWidgets.QLabel("Manual Workflow Inputs")
        hdr.setObjectName("sectionHdr")
        gl.addWidget(hdr, 0, 0, 1, 6)

        gl.addWidget(QtWidgets.QLabel("Specimen name"), 1, 0)
        self.specimen_name = QtWidgets.QLineEdit("MANUAL_TEST_01")
        gl.addWidget(self.specimen_name, 1, 1)

        gl.addWidget(QtWidgets.QLabel("Sample"), 1, 2)
        self.sample_name = QtWidgets.QLineEdit("MANUAL_SAMPLE")
        gl.addWidget(self.sample_name, 1, 3)

        gl.addWidget(QtWidgets.QLabel("Operator"), 1, 4)
        self.operator = QtWidgets.QLineEdit("operator")
        gl.addWidget(self.operator, 1, 5)

        gl.addWidget(QtWidgets.QLabel("Site"), 2, 0)
        self.site = QtWidgets.QLineEdit("SITE01")
        gl.addWidget(self.site, 2, 1)

        gl.addWidget(QtWidgets.QLabel("Location"), 2, 2)
        self.location = QtWidgets.QLineEdit("RAPID-LAB")
        gl.addWidget(self.location, 2, 3)

        gl.addWidget(QtWidgets.QLabel("Output folder"), 3, 0)
        self.output_dir = QtWidgets.QLineEdit(str(Path.cwd() / "manual_measure_output"))
        gl.addWidget(self.output_dir, 3, 1, 1, 4)
        browse = QtWidgets.QPushButton("Browse")
        browse.clicked.connect(self._browse_output)
        gl.addWidget(browse, 3, 5)

        gl.addWidget(QtWidgets.QLabel("Step labels (comma-separated)"), 4, 0)
        self.labels = QtWidgets.QLineEdit("NRM,AF20,AF50,TT300")
        gl.addWidget(self.labels, 4, 1, 1, 5)

        return card

    def _build_actions_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        hl = QtWidgets.QHBoxLayout(card)
        hl.setContentsMargins(12, 10, 12, 10)
        hl.setSpacing(10)

        self.measure_holder_btn = QtWidgets.QPushButton("Measure Holder")
        self.measure_holder_btn.setObjectName("accent")
        self.measure_holder_btn.clicked.connect(self._measure_holder)

        self.measure_sample_btn = QtWidgets.QPushButton("Run Manual Sample Workflow")
        self.measure_sample_btn.setObjectName("accent")
        self.measure_sample_btn.clicked.connect(self._measure_sample)

        self.pause_btn = QtWidgets.QPushButton("Pause")
        self.pause_btn.clicked.connect(self._pause)
        self.resume_btn = QtWidgets.QPushButton("Resume")
        self.resume_btn.clicked.connect(self._resume)
        self.halt_btn = QtWidgets.QPushButton("Halt")
        self.halt_btn.clicked.connect(self._halt)

        hl.addWidget(self.measure_holder_btn)
        hl.addWidget(self.measure_sample_btn)
        hl.addWidget(self.pause_btn)
        hl.addWidget(self.resume_btn)
        hl.addWidget(self.halt_btn)
        hl.addStretch(1)

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
        self.log.setMinimumHeight(220)
        vl.addWidget(self.log, 1)

        return card

    def _append(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{stamp}] {text}")

    def _browse_output(self) -> None:
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select output folder", self.output_dir.text().strip())
        if path:
            self.output_dir.setText(path)

    def _meta(self, specimen_name: str) -> SpecimenMeta:
        return SpecimenMeta(
            name=specimen_name,
            comment="Generated by Manual Measurement Harness",
            core_plate_strike=0.0,
            core_plate_dip=90.0,
            bedding_strike=0.0,
            bedding_dip=0.0,
            volume=1.0,
            sample=self.sample_name.text().strip(),
            site=self.site.text().strip(),
            location=self.location.text().strip(),
        )

    def _start_worker(self, meta: SpecimenMeta, labels: list[str]) -> None:
        if self._worker is not None and self._worker.isRunning():
            QtWidgets.QMessageBox.warning(self, "Workflow Busy", "A run is already active.")
            return

        output_dir = Path(self.output_dir.text().strip() or ".")
        output_dir.mkdir(parents=True, exist_ok=True)

        self._worker = MeasurementWorker(
            meta=meta,
            labels=labels,
            output_dir=output_dir,
            backend=NoCommBackend(),
            operator=self.operator.text().strip(),
            parent=self,
        )
        self._worker.step_started.connect(self._on_step_started)
        self._worker.step_complete.connect(self._on_step_complete)
        self._worker.error_occurred.connect(self._on_error)
        self._worker.run_finished.connect(self._on_finished)
        self._worker.start()
        self._append(f"Started run for {meta.name} with {len(labels)} steps.")

    def _measure_holder(self) -> None:
        self._start_worker(self._meta("Holder"), ["NRM"])

    def _measure_sample(self) -> None:
        specimen_name = self.specimen_name.text().strip()
        if not specimen_name:
            QtWidgets.QMessageBox.warning(self, "Missing Specimen", "Enter a specimen name.")
            return
        labels = [part.strip().upper() for part in self.labels.text().split(",") if part.strip()]
        if not labels:
            QtWidgets.QMessageBox.warning(self, "Missing Steps", "Provide at least one step label.")
            return
        self._start_worker(self._meta(specimen_name), labels)

    def _pause(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.pause()
            self._append("Paused run.")

    def _resume(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.resume()
            self._append("Resumed run.")

    def _halt(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.halt()
            self._append("Halt requested.")

    @QtCore.Slot(int, str)
    def _on_step_started(self, index: int, label: str) -> None:
        self._append(f"Step {index + 1} started: {label}")

    @QtCore.Slot(object)
    def _on_step_complete(self, result: StepResult) -> None:
        self._append(
            f"Step complete: {result.step.demag_label}; M={result.step.moment:.3e} emu; susc={result.susceptibility:.3e}"
        )

    @QtCore.Slot(str)
    def _on_error(self, message: str) -> None:
        self._append(f"ERROR: {message}")

    @QtCore.Slot(bool)
    def _on_finished(self, aborted: bool) -> None:
        status = "aborted" if aborted else "completed"
        self._append(f"Run {status}.")


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "manual_measurement_harness_icon.png", assets_dir)
    window = MainWindow()
    set_app_icon(window, "manual_measurement_harness_icon.png", assets_dir)
    window.show()
    return app.exec()
