"""Operator dialog for transverse AF peak scanning."""
from __future__ import annotations

import threading

from PySide6 import QtCore, QtWidgets

from ..transverse import plan_transverse_scan
from ..transverse_scan import TransverseScanPause


class _ScanThread(QtCore.QThread):
    progress = QtCore.Signal(str)
    sampled = QtCore.Signal(object, str)

    def __init__(self, operation, parent, *, scanning=False):
        super().__init__(parent)
        self.operation = operation
        self.cancel = threading.Event()
        self.pause = TransverseScanPause() if scanning else None
        self.output = None
        self.error = ""

    def run(self):
        try:
            options = {"cancelled": self.cancel.is_set, "progress": self.progress.emit}
            if self.pause is not None:
                options.update(pause=self.pause, sampled=self.sampled.emit)
            self.output = self.operation(**options)
        except Exception as exc:
            self.error = str(exc)


class TransverseScanDialog(QtWidgets.QDialog):
    shutdown_blocked = QtCore.Signal(str)

    def __init__(self, parent, station, acquire, *, simulated):
        super().__init__(parent)
        self._main_window = parent
        self.setWindowTitle("Transverse Probe Scan")
        self.resize(600, 520)
        self.station = station
        self.acquire = acquire
        self.simulated = simulated
        self.result = None
        self.worker = None
        self.lease = None
        self.close_when_stopped = False
        self.action = ""
        layout = QtWidgets.QVBoxLayout(self)
        text = ("No-Comm · Simulated angle and field readings." if simulated else
                "Mount the transverse probe by hand and place it at the calibrated AF height. "
                "Use the current turning encoder reference and a cleared interval that contains the present angle. "
                "Verify degausser cooling before starting.")
        intro = QtWidgets.QLabel(text)
        intro.setWordWrap(True)
        layout.addWidget(intro)
        self.cleared = QtWidgets.QCheckBox("Probe and turning travel verified, AF cooling active")
        if not simulated:
            layout.addWidget(self.cleared)
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)
        self.start_angle = self._number(form, "Start angle", -360, 360, 0, " °")
        self.end_angle = self._number(form, "End angle", -360, 360, 90, " °")
        self.pitch = self._number(form, "Pitch", 1, 180, 15, " °")
        self.field = self._number(form, "Transverse AF field", 0, 10000, 10 if simulated else 0, " mT", decimals=3)
        self.port = QtWidgets.QSpinBox()
        self.port.setRange(0, 256)
        self.port.setPrefix("COM")
        self.port.setSpecialValueText("Select port")
        if not simulated:
            form.addRow("908A gaussmeter", self.port)

        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(("Requested (°)", "Encoder (°)", "Signed DC peak (mT)"))
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        self.message = QtWidgets.QLabel("Ready to scan.")
        self.message.setWordWrap(True)
        layout.addWidget(self.message)
        self.evidence = QtWidgets.QLabel()
        self.evidence.setWordWrap(True)
        self.evidence.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.evidence)

        actions = QtWidgets.QHBoxLayout()
        layout.addLayout(actions)
        self.start_button = QtWidgets.QPushButton("Start Scan", clicked=self._start)
        self.apply_button = QtWidgets.QPushButton("Move to Best Angle", clicked=self._apply)
        self.recover_button = QtWidgets.QPushButton("Verify Safe State", clicked=self._recover)
        self.pause_button = QtWidgets.QPushButton("Pause", clicked=self._toggle_pause)
        self.pause_button.setToolTip("Pause between readings with the AF outputs off.")
        self.stop_button = QtWidgets.QPushButton("Stop", clicked=self.stop)
        self.close_button = QtWidgets.QPushButton("Close", clicked=self.reject)
        for button in (self.start_button, self.apply_button, self.recover_button,
                       self.pause_button, self.stop_button, self.close_button):
            button.setAutoDefault(False)
            actions.addWidget(button)
        self._inputs = self.start_angle, self.end_angle, self.pitch, self.field, self.port
        self.cleared.toggled.connect(self._refresh)
        for field in self._inputs:
            field.valueChanged.connect(self._invalidate_result)
        self._refresh()

    @staticmethod
    def _number(form, label, minimum, maximum, value, suffix, *, decimals=2):
        field = QtWidgets.QDoubleSpinBox()
        field.setRange(minimum, maximum)
        field.setDecimals(decimals)
        field.setValue(value)
        field.setSuffix(suffix)
        form.addRow(label, field)
        return field

    @property
    def active(self):
        return self.worker is not None

    def _invalidate_result(self):
        if self.result is not None:
            self.result = None
            self.message.setText("Settings changed. Start a new scan before applying an angle.")
        self._refresh()

    def _refresh(self):
        active = self.active
        recovery = False
        readable = True
        if not active:
            try:
                recovery = self.station.pending_recovery()
            except Exception as exc:
                readable = False
                self.message.setText(str(exc))
        cleared = self.simulated or self.cleared.isChecked()
        self.start_button.setEnabled(not active and not recovery and readable and cleared)
        self.apply_button.setEnabled(self.start_button.isEnabled() and self.result is not None)
        self.recover_button.setVisible(recovery or active and self.action == "recovery")
        self.recover_button.setEnabled(recovery and not active)
        self.stop_button.setEnabled(active and self.action != "recovery")
        pause = self.worker.pause if active else None
        self.pause_button.setEnabled(pause is not None and not self.worker.cancel.is_set())
        self.pause_button.setText("Resume" if pause is not None and pause.requested else "Pause")
        for field in (*self._inputs, self.cleared):
            field.setEnabled(not active and not recovery)
        return recovery

    def _start(self):
        try:
            plan = plan_transverse_scan(
                self.start_angle.value(), self.end_angle.value(), self.pitch.value(), self.field.value(),
                1 if self.simulated else self.port.value(),
            )
        except ValueError as exc:
            self.message.setText(str(exc))
            return
        self.result = None
        self.table.setRowCount(0)
        self.evidence.clear()
        self._launch("scan", lambda **options: self.station.scan(plan, **options))

    def _apply(self):
        result = self.result
        self.result = None
        self._launch("apply", lambda **options: self.station.apply(result, **options))

    def _recover(self):
        self.result = None
        self._launch("recovery", self.station.recover)

    def _launch(self, action, operation):
        if self.active:
            return
        try:
            self.lease = self.acquire(recovery=action == "recovery")
        except Exception as exc:
            self.message.setText(str(exc))
            self._refresh()
            return
        self.action = action
        self.worker = _ScanThread(operation, self, scanning=action == "scan")
        self.worker.progress.connect(self.message.setText)
        self.worker.sampled.connect(self._append_sample)
        self.worker.finished.connect(self._finish)
        self._refresh()
        self.message.setText("Starting…" if action != "recovery" else "Verifying safe state…")
        self.worker.start()

    @QtCore.Slot(object, str)
    def _append_sample(self, sample, checkpoint):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for column, value in enumerate((sample.requested_deg, sample.actual_deg, sample.peak_mT)):
            self.table.setItem(row, column, QtWidgets.QTableWidgetItem(f"{value:.6g}"))
        if checkpoint:
            self.evidence.setText(checkpoint)

    @QtCore.Slot()
    def _toggle_pause(self):
        if not self.active or self.worker.pause is None or self.worker.cancel.is_set():
            return
        if self.worker.pause.requested:
            self.worker.pause.resume()
        else:
            self.worker.pause.request()
            self.message.setText("Pause requested. Waiting for a safe point between readings…")
        self._refresh()

    @QtCore.Slot()
    def _finish(self):
        worker = self.worker
        self.worker = None
        if self.lease is not None:
            self.lease.release()
            self.lease = None
        worker.deleteLater()
        if worker.error:
            self.message.setText(worker.error)
        elif self.action == "scan":
            self.result = result = worker.output
            self.message.setText(
                f"{'Simulated peak' if self.simulated else 'Measured peak'} at {result.best.actual_deg:.3f}° "
                f"({result.best.peak_mT:+.6g} mT)."
            )
            self.evidence.setText(result.evidence_path)
        elif self.action == "apply":
            self.message.setText(f"{'Simulated angle' if self.simulated else 'Verified encoder angle'} "
                                 f"{worker.output:.3f}°.")
        else:
            self.message.setText("Verified AF outputs off and both motors stopped.")
        recovery = self._refresh()
        if self.close_when_stopped:
            self.close_when_stopped = False
            if recovery:
                self.shutdown_blocked.emit("Verify Safe State in Transverse Probe Scan.")
            else:
                self.reject()

    @QtCore.Slot()
    def stop(self):
        if self.active and self.action != "recovery":
            self.worker.cancel.set()
            if self.worker.pause is not None:
                self.worker.pause.resume()
            self._refresh()
            self.message.setText("Stopping and verifying safe state…")

    def done(self, result):
        if self.active:
            self.close_when_stopped = True
            self.stop()
            return
        super().done(result)

    def closeEvent(self, event):
        if self.active:
            event.ignore()
            self.close_when_stopped = True
            self.stop()
            return
        if getattr(self._main_window, "_shutdown_cleanup_requested", False) and self._refresh():
            event.ignore()
            self.shutdown_blocked.emit("Verify Safe State in Transverse Probe Scan.")
            return
        super().closeEvent(event)
