from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
import sys

from PySide6 import QtCore, QtGui, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.gaussmeter import gaussmeter_driver_status  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402

from .probe import PortProbeResult, sweep_ports  # noqa: E402


class SweepWorker(QtCore.QObject):
    result_ready = QtCore.Signal(object)
    progress = QtCore.Signal(str)
    finished = QtCore.Signal()

    def __init__(self, *, enhanced_only: bool) -> None:
        super().__init__()
        self._enhanced_only = enhanced_only
        self._stop_requested = False

    @QtCore.Slot()
    def run(self) -> None:
        try:
            results = sweep_ports(
                enhanced_only=self._enhanced_only,
                progress=self._on_progress,
                stop_requested=lambda: self._stop_requested,
            )
            for result in results:
                self.result_ready.emit(result)
        finally:
            self.finished.emit()

    def stop(self) -> None:
        self._stop_requested = True

    def _on_progress(self, index: int, total: int, port: str) -> None:
        self.progress.emit(f"Scanning {port} ({index}/{total})")


class MainWindow(QtWidgets.QMainWindow):
    COLUMN_HEADERS = [
        "Port",
        "Description",
        "Adapter",
        "Legacy Hints",
        "Detected Device",
        "Confidence",
        "Protocol",
        "Notes",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy COM Port Mapper")
        self.resize(1320, 760)
        self._worker_thread: QtCore.QThread | None = None
        self._worker: SweepWorker | None = None
        self._results_by_port: dict[str, PortProbeResult] = {}
        self._build_ui()
        self._wire_events()
        self._update_probe_capabilities()
        QtCore.QTimer.singleShot(150, self.start_sweep)

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)

        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        left_card = QtWidgets.QFrame()
        left_card.setObjectName("card")
        left_card.setMinimumWidth(360)
        left = QtWidgets.QVBoxLayout(left_card)
        left.setContentsMargins(18, 18, 18, 18)
        left.setSpacing(12)

        title = QtWidgets.QLabel("COM Port Mapper")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Sweep RAPID serial devices and identify likely roles before launching subsystem apps.")
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        left.addWidget(title)
        left.addWidget(subtitle)

        self.enhanced_only = QtWidgets.QCheckBox("Probe only enhanced / PCI serial adapters (skips USB gaussmeter search)")
        self.enhanced_only.setChecked(False)
        left.addWidget(self.enhanced_only)

        self.capability_label = QtWidgets.QLabel()
        self.capability_label.setObjectName("valuePill")
        self.capability_label.setWordWrap(True)
        left.addWidget(self.capability_label)

        button_row = QtWidgets.QHBoxLayout()
        self.sweep_btn = QtWidgets.QPushButton("Sweep Ports")
        self.sweep_btn.setObjectName("accent")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.copy_btn = QtWidgets.QPushButton("Copy Selected")
        button_row.addWidget(self.sweep_btn)
        button_row.addWidget(self.stop_btn)
        button_row.addWidget(self.copy_btn)
        left.addLayout(button_row)

        self.status_label = QtWidgets.QLabel("Ready")
        self.status_label.setObjectName("valuePill")
        left.addWidget(self.status_label)

        legacy = QtWidgets.QLabel(
            "VB6 default map\n"
            "COM3 Vacuum\n"
            "COM4 Up/Down\n"
            "COM5 Turning\n"
            "COM6 X / Changer\n"
            "COM7 Y\n"
            "COM8 Susceptibility\n"
            "COM9 AF\n"
            "COM10 SQUID"
        )
        legacy.setObjectName("valuePill")
        legacy.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        left.addWidget(legacy)

        note = QtWidgets.QLabel(
            "High-confidence auto-identification covers motor controllers, SQUID, and gaussmeter ports when the "
            "legacy gm0.dll driver is available. Vacuum, susceptibility, and AF still show adapter metadata plus "
            "legacy role hints."
        )
        note.setWordWrap(True)
        left.addWidget(note)

        console_title = QtWidgets.QLabel("Console")
        console_title.setObjectName("subtitle")
        left.addWidget(console_title)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        left.addWidget(self.console, stretch=1)

        right_card = QtWidgets.QFrame()
        right_card.setObjectName("card")
        right = QtWidgets.QVBoxLayout(right_card)
        right.setContentsMargins(14, 14, 14, 14)
        right.setSpacing(10)

        self.table = QtWidgets.QTableWidget(0, len(self.COLUMN_HEADERS))
        self.table.setHorizontalHeaderLabels(self.COLUMN_HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(7, QtWidgets.QHeaderView.Stretch)
        right.addWidget(self.table, stretch=1)

        details_title = QtWidgets.QLabel("Selected Port Details")
        details_title.setObjectName("subtitle")
        right.addWidget(details_title)

        self.details = QtWidgets.QPlainTextEdit()
        self.details.setObjectName("console")
        self.details.setReadOnly(True)
        self.details.setMaximumHeight(170)
        right.addWidget(self.details)

        layout.addWidget(left_card)
        layout.addWidget(right_card, stretch=1)

        apply_card_shadow(left_card)
        apply_card_shadow(right_card)

    def _wire_events(self) -> None:
        self.sweep_btn.clicked.connect(self.start_sweep)
        self.stop_btn.clicked.connect(self.stop_sweep)
        self.copy_btn.clicked.connect(self.copy_selected)
        self.table.itemSelectionChanged.connect(self._update_details)
        self.enhanced_only.toggled.connect(self._update_probe_capabilities)

    def start_sweep(self) -> None:
        if self._worker_thread is not None:
            return
        self._results_by_port.clear()
        self.table.setRowCount(0)
        self.details.clear()
        self.console.clear()
        self._update_probe_capabilities()
        self._set_status("Scanning COM ports...")
        self.sweep_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self._worker_thread = QtCore.QThread(self)
        self._worker = SweepWorker(enhanced_only=self.enhanced_only.isChecked())
        self._worker.moveToThread(self._worker_thread)
        self._worker_thread.started.connect(self._worker.run)
        self._worker.result_ready.connect(self._append_result)
        self._worker.progress.connect(self._log_progress)
        self._worker.finished.connect(self._finish_sweep)
        self._worker.finished.connect(self._worker_thread.quit)
        self._worker_thread.finished.connect(self._worker_thread.deleteLater)
        self._worker_thread.finished.connect(self._clear_worker_refs)
        self._worker_thread.start()

    def stop_sweep(self) -> None:
        if self._worker is None:
            return
        self._worker.stop()
        self._set_status("Stopping sweep...")
        self.console.appendPlainText("Stop requested. Finishing the current probe step.")

    def copy_selected(self) -> None:
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            QtWidgets.QMessageBox.information(self, "No Selection", "Select a row first.")
            return
        row = selected[0].row()
        port = self.table.item(row, 0).text()
        result = self._results_by_port.get(port)
        if result is None:
            return
        payload = json.dumps(asdict(result), indent=2)
        QtWidgets.QApplication.clipboard().setText(payload)
        self.console.appendPlainText(f"Copied details for {port} to the clipboard.")

    def _append_result(self, result: PortProbeResult) -> None:
        self._results_by_port[result.port] = result
        row = self.table.rowCount()
        self.table.insertRow(row)
        values = [
            result.port,
            result.description or "-",
            result.adapter_family,
            "; ".join(result.legacy_hints) or "-",
            result.detected_device,
            result.confidence,
            result.protocol or "-",
            result.notes or "-",
        ]
        for column, value in enumerate(values):
            item = QtWidgets.QTableWidgetItem(value)
            if column in (0, 4, 5, 6):
                item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table.setItem(row, column, item)
        if result.confidence == "High":
            for column in range(len(values)):
                self.table.item(row, column).setBackground(QtGui.QColor(255, 247, 211))
        elif result.confidence == "Blocked":
            for column in range(len(values)):
                self.table.item(row, column).setBackground(QtGui.QColor(255, 231, 231))

        summary = f"{result.port}: {result.detected_device}"
        if result.protocol:
            summary += f" [{result.protocol}]"
        self.console.appendPlainText(summary)

    def _log_progress(self, message: str) -> None:
        self._set_status(message)

    def _finish_sweep(self) -> None:
        self.sweep_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        high_confidence = sum(1 for result in self._results_by_port.values() if result.confidence == "High")
        self._set_status(f"Sweep complete. {high_confidence} high-confidence match(es).")
        self.console.appendPlainText(self.status_label.text())

    def _clear_worker_refs(self) -> None:
        self._worker = None
        self._worker_thread = None

    def _set_status(self, message: str) -> None:
        self.status_label.setText(message)

    def _update_details(self) -> None:
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            self.details.clear()
            return
        row = selected[0].row()
        port = self.table.item(row, 0).text()
        result = self._results_by_port.get(port)
        if result is None:
            self.details.clear()
            return
        payload = [
            f"Port: {result.port}",
            f"Description: {result.description or '-'}",
            f"Manufacturer: {result.manufacturer or '-'}",
            f"Adapter: {result.adapter_family}",
            f"HWID: {result.hwid or '-'}",
            f"Location: {result.location or '-'}",
            f"Legacy hints: {'; '.join(result.legacy_hints) or '-'}",
            f"Detected device: {result.detected_device}",
            f"Confidence: {result.confidence}",
            f"Protocol: {result.protocol or '-'}",
            f"Notes: {result.notes or '-'}",
            f"Raw response: {result.raw_response or '-'}",
        ]
        self.details.setPlainText("\n".join(payload))

    def _update_probe_capabilities(self) -> None:
        available, detail = gaussmeter_driver_status()
        if available:
            message = f"Gaussmeter probing enabled via {detail}."
        else:
            message = f"Gaussmeter probing unavailable: {detail}."
        if self.enhanced_only.isChecked():
            message += " PCI-only filter is on, so USB gaussmeter ports will be skipped."
        self.capability_label.setText(message)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()