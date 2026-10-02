from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import (
    VacuumBackend,
    VacuumNoCommBackend,
    read_vacuum_snapshot,
)
from rapid_main.glass_theme import set_semantic_status


class VacuumDialog(QtWidgets.QDialog):
    """Vacuum pressure monitor — replaces VB6 frmVacuum."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: VacuumBackend | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend: VacuumBackend = backend or VacuumNoCommBackend()
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
        vl.setSpacing(14)

        hdr = QtWidgets.QLabel("Vacuum Pressure")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("Vacuum pressure monitor title")
        vl.addWidget(hdr)

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
        vl.addLayout(pump_row)

        self._close_btn = QtWidgets.QPushButton("Close")
        self._close_btn.setAccessibleName("Close vacuum monitor")
        self._close_btn.clicked.connect(self.close)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self._close_btn)
        vl.addLayout(btn_row)

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

    def _set_pump_state(self, on: bool) -> None:
        state = "On" if on else "Off"
        level = "simulated" if self._backend.simulated else ("active" if on else "neutral")
        set_semantic_status(
            self._pump_state_lbl,
            f"Pump is {state.lower()}",
            level,
            accessible_name="Vacuum pump state",
        )
        self._pump_btn.setText("⏹  Pump Off" if on else "▶  Pump On")
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
            pump_on = False
        self._pump_btn.blockSignals(True)
        try:
            self._pump_btn.setChecked(pump_on)
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

    def _on_pump_toggle(self, on: bool) -> None:
        try:
            self._backend.set_pump(on)
            self._set_status(f"Pump command completed: {'on' if on else 'off'}", "active")
        except Exception as exc:
            self._set_status(f"Pump command failed: {exc}", "error")
            try:
                actual = bool(self._backend.is_pump_on())
            except Exception:
                actual = not on
            self._pump_btn.blockSignals(True)
            self._pump_btn.setChecked(actual)
            self._pump_btn.blockSignals(False)
            self._set_pump_state(actual)
            return
        else:
            self._set_pump_state(on)
        self._refresh()

    def _refresh(self) -> None:
        snapshot = read_vacuum_snapshot(
            self._backend,
            warn_threshold=float(self._warn_spin.value()),
        )
        self._set_pump_state(snapshot.pump_on)
        if snapshot.pressure_mtorr is None:
            self._set_status(snapshot.fault_reason or "Pressure reading unavailable", "error")
            return

        self.set_pressure(snapshot.pressure_mtorr)

        if snapshot.fault:
            self._set_status(snapshot.fault_reason, "error")
        elif snapshot.connected:
            self._set_status(snapshot.status, "ready")
        else:
            self._set_status("Vacuum backend is disconnected", "unavailable")

    def closeEvent(self, event: "QtCore.QEvent") -> None:  # type: ignore[override]
        self._timer.stop()
        super().closeEvent(event)
