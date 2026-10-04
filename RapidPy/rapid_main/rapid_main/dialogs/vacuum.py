from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.diagnostic_services import (
    VacuumBackend,
    VacuumNoCommBackend,
    read_vacuum_snapshot,
)
from rapid_main.glass_theme import set_semantic_status


class VacuumCommandWorker(QtCore.QObject):
    succeeded = QtCore.Signal(str)
    failed = QtCore.Signal(str, str)
    settled = QtCore.Signal()

    def __init__(self, label, action):
        super().__init__()
        self.label, self.action = label, action

    @QtCore.Slot()
    def run(self):
        try:
            self.action()
            self.succeeded.emit(self.label)
        except Exception as exc:
            self.failed.emit(self.label, str(exc))
        finally:
            self.settled.emit()


class VacuumDialog(QtWidgets.QDialog):
    """Vacuum pressure monitor — replaces VB6 frmVacuum."""
    shutdown_blocked = QtCore.Signal(str)

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: VacuumBackend | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend: VacuumBackend = backend or VacuumNoCommBackend()
        self._command_thread = self._command_worker = None
        self._closing = False
        self._close_release_attempted = self._close_disconnect_attempted = False
        self._last_error = ''
        self._close_timer = QtCore.QTimer(self)
        self._close_timer.setSingleShot(True)
        self._close_timer.timeout.connect(self._resume_close)
        self._settle_timer = QtCore.QTimer(self)
        self._settle_timer.setSingleShot(True)
        self._settle_timer.timeout.connect(self._command_settled)
        self._layout_timer = QtCore.QTimer(self)
        self._layout_timer.setSingleShot(True)
        self._layout_timer.timeout.connect(self._fit_status_labels)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Vacuum Monitor")
        self.setAccessibleName("Vacuum pressure monitor")
        self.setMinimumWidth(340)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._sync_from_backend()

        # Periodic pressure refresh from backend.
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._refresh)
        self._timer.start(2000)

    def set_pressure(self, mtorr: float) -> None:
        warn = mtorr > float(self._warn_spin.value())
        set_semantic_status(
            self._pressure_lbl,
            f"{mtorr:.3f}",
            "warning" if warn else "ready",
            accessible_name="Vacuum pressure in millitorr",
            show_prefix=False,
        )
        self._set_status(
            "Pressure is above the warning threshold"
            if warn
            else "Pressure is within the configured range",
            "warning" if warn else "ready",
        )

    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(8)

        hdr = QtWidgets.QLabel("Vacuum Pressure")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("Vacuum pressure monitor title")
        vl.addWidget(hdr)
        root = vl
        body = QtWidgets.QWidget()
        vl = QtWidgets.QVBoxLayout(body)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(10)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll.setWidget(body)
        root.addWidget(scroll, 1)
        connection = QtWidgets.QHBoxLayout()
        self._connect_btn = QtWidgets.QPushButton('Connect')
        self._connect_btn.setAccessibleName('Connect configured vacuum transport')
        self._connect_btn.setAccessibleDescription('Connects without enabling or resetting pump and valve outputs.')
        self._connect_btn.clicked.connect(self._connect)
        self._disconnect_btn = QtWidgets.QPushButton('Disconnect')
        self._disconnect_btn.setAccessibleName('Disconnect verified vacuum transport')
        self._disconnect_btn.clicked.connect(self._disconnect)
        connection.addWidget(self._connect_btn)
        connection.addWidget(self._disconnect_btn)
        vl.addLayout(connection)
        self._binding_lbl = QtWidgets.QLabel()
        self._binding_lbl.setWordWrap(True)
        cfg = getattr(self._backend, '_cfg', None)
        self._binding_lbl.setText(f'Configured vacuum: {cfg.port} / {cfg.baud} baud' if cfg and getattr(cfg, 'port', '') else 'Vacuum connection follows the configured backend.')
        vl.addWidget(self._binding_lbl)

        read_frame = QtWidgets.QFrame()
        read_frame.setObjectName("dialogCard")
        rl = QtWidgets.QVBoxLayout(read_frame)
        rl.setAlignment(QtCore.Qt.AlignCenter)
        rl.setContentsMargins(16, 14, 16, 14)

        self._pressure_lbl = QtWidgets.QLabel("—")
        self._pressure_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._pressure_lbl.setObjectName("readingDisplay")
        self._pressure_lbl.setAccessibleName("Vacuum pressure in millitorr")

        unit_lbl = QtWidgets.QLabel("mTorr")
        unit_lbl.setAlignment(QtCore.Qt.AlignCenter)
        unit_lbl.setObjectName("unitLabel")

        self._status_lbl = QtWidgets.QLabel("Initializing...")
        self._status_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._status_lbl.setWordWrap(True)
        set_semantic_status(
            self._status_lbl,
            "Waiting for the first pressure reading",
            "neutral",
            accessible_name="Vacuum system status",
        )
        rl.addWidget(self._pressure_lbl)
        rl.addWidget(unit_lbl)
        rl.addWidget(self._status_lbl)
        self._pump_state_lbl = QtWidgets.QLabel("Pump: unknown")
        self._pump_state_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._pump_state_lbl.setWordWrap(True)
        rl.addWidget(self._pump_state_lbl)
        vl.addWidget(read_frame)

        grp = QtWidgets.QGroupBox("Threshold")
        fl = QtWidgets.QFormLayout(grp)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._target_spin = QtWidgets.QDoubleSpinBox()
        self._target_spin.setRange(0.0, 100.0)
        self._target_spin.setValue(5.0)
        self._target_spin.setSuffix(" mTorr")
        self._target_spin.setAccessibleName("Target vacuum pressure in millitorr")
        fl.addRow("Target pressure:", self._target_spin)

        self._warn_spin = QtWidgets.QDoubleSpinBox()
        self._warn_spin.setRange(0.0, 200.0)
        self._warn_spin.setValue(20.0)
        self._warn_spin.setSuffix(" mTorr")
        self._warn_spin.setAccessibleName("Vacuum warning threshold in millitorr")
        fl.addRow("Warning threshold:", self._warn_spin)

        vl.addWidget(grp)

        pump_row = QtWidgets.QHBoxLayout()
        self._pump_btn = QtWidgets.QPushButton("▶  Pump On")
        self._pump_btn.setCheckable(True)
        self._pump_btn.setAccessibleName("Toggle vacuum pump")
        self._pump_btn.setAccessibleDescription(
            "Starts or stops the configured vacuum pump backend."
        )
        self._pump_btn.toggled.connect(self._on_pump_toggle)
        pump_row.addWidget(self._pump_btn)
        pump_row.addStretch()
        root.addLayout(pump_row)
        self._release_btn = QtWidgets.QPushButton('Release / Verify Off')
        self._release_btn.setAccessibleName('Release vacuum and verify off commands')
        self._release_btn.setAccessibleDescription('Recovers the original vacuum operation by independently acknowledging valve-off and pump-off commands. No pressure or grip telemetry is available.')
        self._release_btn.clicked.connect(self._release)
        root.addWidget(self._release_btn)
        evidence = QtWidgets.QLabel('Live legacy vacuum uses command acknowledgements. Physical pressure and specimen grip are not measured. Saved AutoPump never enables live outputs on startup.')
        evidence.setWordWrap(True)
        vl.addWidget(evidence)

        self._close_btn = QtWidgets.QPushButton("Close")
        self._close_btn.setAccessibleName("Close vacuum monitor")
        self._close_btn.clicked.connect(self.close)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self._close_btn)
        root.addLayout(btn_row)

    def _annotate_status(self, text: str) -> str:
        status = text.strip()
        if self._backend.simulated and "sim" not in status.lower():
            status = f"{status} (simulated)"
        return status

    def _set_status(self, text: str, level: str = "neutral") -> None:
        if self._backend.simulated and level in {"neutral", "ready", "active"}:
            level = "simulated"
        set_semantic_status(
            self._status_lbl,
            self._annotate_status(text),
            level,
            accessible_name="Vacuum system status",
        )
        self._layout_timer.start(0)

    @QtCore.Slot()
    def _fit_status_labels(self):
        for name in ('_status_lbl', '_pump_state_lbl'):
            label = getattr(self, name, None)
            if label is not None:
                height = label.heightForWidth(max(1, label.width()))
                if height > 0:
                    label.setMinimumHeight(height)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, '_pump_state_lbl'):
            self._layout_timer.start(0)

    def _set_pump_state(self, on: bool | None) -> None:
        if on is None:
            set_semantic_status(self._pump_state_lbl, 'Pump / valve state is unverified', 'warning', accessible_name='Vacuum output state')
            self._pump_btn.setText('Pump On')
            self._pump_btn.setAccessibleDescription('Output state is unverified. Pump On explicitly requests enable acknowledgements; Release / Verify Off is available for recovery.')
            return
        state = "On" if on else "Off"
        level = "simulated" if self._backend.simulated else ("active" if on else "neutral")
        set_semantic_status(
            self._pump_state_lbl,
            f"Pump / valve commanded {state.lower()}" if getattr(self._backend, 'pressure_telemetry_available', True) is False else f"Pump is {state.lower()}",
            level,
            accessible_name="Vacuum pump state",
        )
        self._pump_btn.setText("Pump Off" if on else "Pump On")
        self._pump_btn.setAccessibleDescription(
            f"The pump is currently {state.lower()}. Activate to turn it "
            f"{'off' if on else 'on'}."
        )

    def _sync_from_backend(self) -> None:
        cfg = getattr(self._backend, "_cfg", None)
        if cfg is not None:
            self._target_spin.setValue(getattr(cfg, "target_pressure", self._target_spin.value()))
            self._warn_spin.setValue(getattr(cfg, "warn_threshold", self._warn_spin.value()))
        try:
            pump_on = bool(self._backend.is_pump_on())
        except Exception:
            pump_on = None
        self._pump_btn.blockSignals(True)
        try:
            self._pump_btn.setChecked(pump_on is True)
        finally:
            self._pump_btn.blockSignals(False)

        try:
            connected = bool(self._backend.is_connected())
        except Exception:
            connected = False
        if connected:
            try:
                status = self._backend.status()
            except Exception as exc:
                status = f"Backend status read failed: {exc}"
            self._set_status(status, "ready")
            self._pump_btn.setEnabled(True)
        else:
            try:
                status = self._backend.status()
            except Exception:
                status = "Vacuum backend is disconnected"
            self._set_status(status, "unavailable")
            self._pump_btn.setEnabled(False)

        self._set_pump_state(pump_on)
        busy = self._command_thread is not None or self._closing
        try:
            recovery = getattr(self._backend, 'can_recover_pending', False) is True
        except Exception as exc:
            recovery = True
            self._last_error = 'Hardware safety state is unverified: ' + str(exc)
        available = getattr(self._backend, 'is_available', lambda: True)()
        self._connect_btn.setEnabled(bool(available and not connected and not busy and callable(getattr(self._backend, 'connect', None))))
        self._disconnect_btn.setEnabled(bool(connected and not busy and not getattr(self._backend, 'outputs_held', False) and callable(getattr(self._backend, 'disconnect', None))))
        self._pump_btn.setEnabled(connected and not busy and not recovery)
        self._release_btn.setEnabled(connected and not busy)
        if self._last_error:
            self._set_status(self._last_error, 'error')

    def _start_command(self, label, action, *, closing=False):
        if self._command_thread is not None or (self._closing and not closing):
            return
        self._last_error = ''
        thread = QtCore.QThread(self)
        worker = VacuumCommandWorker(label, action)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.succeeded.connect(self._command_succeeded)
        worker.failed.connect(self._command_failed)
        worker.settled.connect(worker.deleteLater)
        worker.settled.connect(thread.quit)
        thread.finished.connect(self._command_settled)
        self._command_thread, self._command_worker = thread, worker
        self._sync_from_backend()
        self._set_status(label + ' in progress; closing waits for output cleanup.', 'active')
        thread.start()

    @QtCore.Slot(str)
    def _command_succeeded(self, label):
        self._last_error = ''

    @QtCore.Slot(str, str)
    def _command_failed(self, label, message):
        self._last_error = label + ' failed: ' + message

    @QtCore.Slot()
    def _command_settled(self):
        if self._command_thread is not None and self._command_thread.isRunning():
            self._settle_timer.start(10)
            return
        self._settle_timer.stop()
        if self._command_thread is not None:
            self._command_thread.deleteLater()
        self._command_thread = self._command_worker = None
        self._sync_from_backend()
        self._resume_close()

    def _connect(self):
        action = getattr(self._backend, 'connect', None)
        if callable(action):
            self._start_command('Connect', action)

    def _disconnect(self):
        action = getattr(self._backend, 'disconnect', None)
        if callable(action):
            self._start_command('Disconnect', action)

    def _release(self):
        self._start_command('Release / Verify Off', lambda: self._backend.set_pump(False))

    def _on_pump_toggle(self, on: bool) -> None:
        self._start_command('Pump On' if on else 'Release / Verify Off', lambda: self._backend.set_pump(on))

    def _refresh(self) -> None:
        if self._command_thread is not None or self._closing:
            return
        snapshot = read_vacuum_snapshot(
            self._backend,
            warn_threshold=float(self._warn_spin.value()),
        )
        self._set_pump_state(snapshot.pump_on if getattr(self._backend, 'output_state_known', True) else None)
        if self._last_error:
            self._set_status(self._last_error, 'error')
            return
        if snapshot.pressure_mtorr is None:
            self._pressure_lbl.setText("—")
            self._set_status(
                snapshot.fault_reason or snapshot.status or "Pressure reading unavailable",
                "error" if snapshot.fault else "warning",
            )
            return

        self.set_pressure(snapshot.pressure_mtorr)

        if snapshot.fault:
            self._set_status(snapshot.fault_reason, "error")
        elif snapshot.connected:
            self._set_status(snapshot.status, "ready")
        else:
            self._set_status("Vacuum backend is disconnected", "unavailable")

    @QtCore.Slot()
    def _resume_close(self):
        if not self._closing:
            return
        if self._command_thread is not None or getattr(self._backend, 'operation_active', False) is True:
            self._close_timer.start(100)
            return
        connected = self._backend.is_connected()
        if connected and not self._close_release_attempted:
            self._close_release_attempted = True
            self._start_command('Release before close', lambda: self._backend.set_pump(False), closing=True)
            return
        if getattr(self._backend, 'outputs_held', False) is True:
            self._abort_close('Vacuum release is unverified. The window and connection remain available for recovery.')
            return
        disconnect = getattr(self._backend, 'disconnect', None)
        if connected and callable(disconnect):
            if not self._close_disconnect_attempted:
                self._close_disconnect_attempted = True
                self._start_command('Disconnect before close', disconnect, closing=True)
                return
            self._abort_close('Vacuum disconnect failed. Retry after checking the original transport.')
            return
        self.close()

    def _abort_close(self, reason):
        self._closing = False
        self._close_release_attempted = self._close_disconnect_attempted = False
        self._last_error = (self._last_error + '; ' + reason).strip('; ')
        self._timer.start(2000)
        self._sync_from_backend()
        self.shutdown_blocked.emit(reason)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._timer.stop()
        if self._backend.simulated is True and self._command_thread is None:
            self._backend.set_pump(False)
            self._close_release_attempted = True
        connected = self._backend.is_connected()
        if self._command_thread is not None or (connected and not self._close_release_attempted) or getattr(self._backend, 'outputs_held', False) is True:
            self._closing = True
            event.ignore()
            self._sync_from_backend()
            self._resume_close()
            return
        self._close_timer.stop()
        super().closeEvent(event)
