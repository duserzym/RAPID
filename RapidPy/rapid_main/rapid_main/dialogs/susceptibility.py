from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import (
    SusceptibilityBackend,
    SusceptibilityNoCommBackend,
)
from rapid_main.glass_theme import set_semantic_status


class SusceptibilityDialog(QtWidgets.QDialog):
    """Bartington bridge diagnostic surface replacing the misrouted SQUID dialog."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: SusceptibilityBackend | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend = backend or SusceptibilityNoCommBackend()
        self.setObjectName("glassDialog")
        self.setWindowTitle("Susceptibility Bridge")
        self.setAccessibleName("Bartington susceptibility bridge diagnostics")
        self.setMinimumWidth(380)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._refresh_status()

    def _build_ui(self) -> None:
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(14)

        title = QtWidgets.QLabel("Bartington Susceptibility Bridge")
        title.setObjectName("dialogTitle")
        title.setAccessibleName("Susceptibility bridge dialog title")
        layout.addWidget(title)

        note = QtWidgets.QLabel(
            "Diagnostic zero and measure commands use the legacy Z/M serial protocol. "
            "This window does not move the sample, apply holder correction, or prove a "
            "production SUSC queue step."
        )
        note.setObjectName("dialogSubtitle")
        note.setWordWrap(True)
        layout.addWidget(note)

        card = QtWidgets.QFrame()
        card.setObjectName("dialogCard")
        card_layout = QtWidgets.QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(8)

        self._reading = QtWidgets.QLabel("—")
        self._reading.setObjectName("readingDisplay")
        self._reading.setAlignment(QtCore.Qt.AlignCenter)
        self._reading.setAccessibleName("Latest susceptibility bridge reading")
        card_layout.addWidget(self._reading)

        unit = QtWidgets.QLabel("bridge units after configured scale factor")
        unit.setObjectName("unitLabel")
        unit.setAlignment(QtCore.Qt.AlignCenter)
        card_layout.addWidget(unit)

        self._status_lbl = QtWidgets.QLabel()
        self._status_lbl.setWordWrap(True)
        self._status_lbl.setAlignment(QtCore.Qt.AlignCenter)
        card_layout.addWidget(self._status_lbl)
        layout.addWidget(card)

        actions = QtWidgets.QHBoxLayout()
        self._connect_btn = QtWidgets.QPushButton("Connect")
        self._connect_btn.setAccessibleName("Connect susceptibility bridge")
        self._connect_btn.setAccessibleDescription(
            "Opens the configured Bartington bridge while this diagnostic owns the device."
        )
        self._connect_btn.clicked.connect(self._connect)
        actions.addWidget(self._connect_btn)

        self._zero_btn = QtWidgets.QPushButton("Zero Bridge")
        self._zero_btn.setAccessibleName("Zero susceptibility bridge")
        self._zero_btn.setAccessibleDescription(
            "Sends the legacy Z command. Ensure the sample is outside the coil."
        )
        self._zero_btn.clicked.connect(self._zero)
        actions.addWidget(self._zero_btn)

        self._measure_btn = QtWidgets.QPushButton("Measure")
        self._measure_btn.setObjectName("accent")
        self._measure_btn.setAccessibleName("Measure susceptibility bridge")
        self._measure_btn.setAccessibleDescription(
            "Sends the legacy M command without moving the sample."
        )
        self._measure_btn.clicked.connect(self._measure)
        actions.addWidget(self._measure_btn)
        actions.addStretch()
        layout.addLayout(actions)

        close_btn = QtWidgets.QPushButton("Close")
        close_btn.setAccessibleName("Close susceptibility bridge diagnostics")
        close_btn.clicked.connect(self.close)
        close_row = QtWidgets.QHBoxLayout()
        close_row.addStretch()
        close_row.addWidget(close_btn)
        layout.addLayout(close_row)

    def _set_status(self, text: str, level: str) -> None:
        if bool(getattr(self._backend, "simulated", False)) and level in {
            "neutral",
            "ready",
            "active",
        }:
            level = "simulated"
            if "simulat" not in text.lower():
                text = f"{text} (simulated)"
        set_semantic_status(
            self._status_lbl,
            text,
            level,
            accessible_name="Susceptibility bridge status",
        )

    def _refresh_status(self) -> None:
        try:
            connected = bool(self._backend.is_connected())
            status = str(self._backend.status())
        except Exception as exc:
            connected = False
            status = f"Susceptibility backend status failed: {exc}"
        self._zero_btn.setEnabled(connected)
        self._measure_btn.setEnabled(connected)
        self._connect_btn.setEnabled(not connected)
        self._set_status(status, "ready" if connected else "unavailable")

    @QtCore.Slot()
    def _connect(self) -> None:
        try:
            self._backend.test_connection()
        except Exception as exc:
            self._set_status(f"Bridge connection failed: {exc}", "error")
            return
        self._refresh_status()

    @QtCore.Slot()
    def _zero(self) -> None:
        try:
            self._backend.zero()
        except Exception as exc:
            self._set_status(f"Bridge zero failed: {exc}", "error")
            return
        self._set_status("Bridge zero acknowledged", "active")

    @QtCore.Slot()
    def _measure(self) -> None:
        try:
            value = float(self._backend.measure())
        except Exception as exc:
            self._set_status(f"Bridge measurement failed: {exc}", "error")
            return
        self._reading.setText(f"{value:.9g}")
        self._set_status("Bridge measurement received", "ready")
