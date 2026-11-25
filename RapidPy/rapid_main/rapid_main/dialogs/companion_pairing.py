"""Pair iPhones and iPads with RAPID Companion by QR code or pairing key."""

from __future__ import annotations

from collections.abc import Callable

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.companion import CompanionBridge, normalize_endpoint, pairing_link
from rapid_main.glass_theme import set_semantic_status

QR_PIXELS = 264


def qr_image(text: str, size: int = QR_PIXELS) -> QtGui.QImage:
    """Render ``text`` as a black-on-white QR code with the standard quiet zone."""
    import segno

    rows = [list(row) for row in segno.make(text, error="m").matrix_iter(scale=1, border=4)]
    modules = len(rows)
    scale = max(1, size // modules)
    image = QtGui.QImage(modules * scale, modules * scale, QtGui.QImage.Format.Format_RGB32)
    image.fill(QtGui.QColor("white"))
    painter = QtGui.QPainter(image)
    painter.setPen(QtCore.Qt.PenStyle.NoPen)
    painter.setBrush(QtGui.QColor("black"))
    for y, row in enumerate(rows):
        for x, dark in enumerate(row):
            if dark:
                painter.drawRect(x * scale, y * scale, scale, scale)
    painter.end()
    return image


class CompanionPairingDialog(QtWidgets.QDialog):
    """Show the pairing code while sharing stays active; closing does not stop sharing."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None,
        bridge: CompanionBridge,
        *,
        address: str,
        save_address: Callable[[str], None],
        control_action: QtGui.QAction,
        stop_sharing: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Mobile Companion")
        self.setAccessibleName("Mobile Companion pairing")
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowType.WindowContextHelpButtonHint)
        self.setMinimumWidth(420)
        self._bridge = bridge
        self._save_address = save_address
        self._saved_address = address
        self._control_action = control_action
        self._stop_sharing = stop_sharing
        self._build_ui()
        self.address.setText(address)
        self._refresh_code()
        self._refresh_status()
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._refresh_status)
        self._timer.start()

    def _build_ui(self) -> None:
        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(20, 16, 20, 16)
        outer.setSpacing(12)
        title = QtWidgets.QLabel("Pair an iPhone or iPad")
        title.setObjectName("dialogTitle")
        outer.addWidget(title)
        intro = QtWidgets.QLabel(
            "Scan the code with the device camera or with Scan Pairing Code in RAPID Companion, "
            "then confirm Connect on the device. Sharing continues after this window closes."
        )
        intro.setObjectName("guidanceText")
        intro.setWordWrap(True)
        outer.addWidget(intro)

        form = QtWidgets.QFormLayout()
        self.address = QtWidgets.QLineEdit()
        self.address.setPlaceholderText("station.example.org")
        self.address.setAccessibleName("Station secure address for phones")
        self.address.setToolTip(
            "The wss:// address of the station's private TLS proxy, which forwards /rapid/v1 "
            "to this PC's local bridge."
        )
        self.address.textChanged.connect(self._refresh_code)
        self.address.editingFinished.connect(self._commit_address)
        form.addRow("Phone address", self.address)

        key_row = QtWidgets.QHBoxLayout()
        self.key = QtWidgets.QLineEdit(self._bridge.token)
        self.key.setReadOnly(True)
        self.key.setCursorPosition(0)
        self.key.setAccessibleName("Pairing key")
        self.copy_key = QtWidgets.QPushButton("Copy Key")
        self.copy_key.clicked.connect(self._copy_key)
        key_row.addWidget(self.key, 1)
        key_row.addWidget(self.copy_key)
        form.addRow("Pairing key", key_row)
        outer.addLayout(form)

        self.code = QtWidgets.QLabel()
        self.code.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.code.setMinimumSize(QR_PIXELS + 24, QR_PIXELS + 24)
        self.code.setWordWrap(True)
        self.code.setAccessibleName("Pairing code")
        self.code.setStyleSheet("background: white; color: #1d1d1f; border-radius: 14px; padding: 12px;")
        outer.addWidget(self.code, 0, QtCore.Qt.AlignmentFlag.AlignHCenter)

        self.control = QtWidgets.QCheckBox("Allow remote workflow control (pause, resume, halt)")
        self.control.setChecked(self._control_action.isChecked())
        self.control.toggled.connect(self._control_action.setChecked)
        self._control_action.toggled.connect(self._sync_control)
        outer.addWidget(self.control)

        self.status = QtWidgets.QLabel()
        self.status.setWordWrap(True)
        outer.addWidget(self.status)

        buttons = QtWidgets.QDialogButtonBox()
        self.stop = buttons.addButton("Stop Sharing", QtWidgets.QDialogButtonBox.ButtonRole.DestructiveRole)
        self.stop.clicked.connect(self._stop_sharing)
        close = buttons.addButton(QtWidgets.QDialogButtonBox.StandardButton.Close)
        close.clicked.connect(self.close)
        outer.addWidget(buttons)

    def _sync_control(self, checked: bool) -> None:
        if self.control.isChecked() != checked:
            self.control.blockSignals(True)
            self.control.setChecked(checked)
            self.control.blockSignals(False)
        self._refresh_status()

    def _copy_key(self) -> None:
        QtGui.QGuiApplication.clipboard().setText(self._bridge.token)
        self.copy_key.setText("Copied")
        QtCore.QTimer.singleShot(1500, lambda: self.copy_key.setText("Copy Key"))

    def _endpoint(self) -> str | None:
        try:
            return normalize_endpoint(self.address.text()) if self.address.text().strip() else None
        except ValueError:
            return None

    def _commit_address(self) -> None:
        endpoint = self._endpoint()
        if endpoint and endpoint != self._saved_address:
            self._saved_address = endpoint
            self._save_address(endpoint)

    def _refresh_code(self) -> None:
        text = self.address.text().strip()
        if not text:
            self.code.setPixmap(QtGui.QPixmap())
            self.code.setText(
                "Enter the secure address phones use to reach this station to show a pairing code. "
                "Without one, enter the address and key on the device."
            )
            return
        try:
            link = pairing_link(text, self._bridge.token)
        except ValueError as exc:
            self.code.setPixmap(QtGui.QPixmap())
            self.code.setText(str(exc))
            return
        try:
            image = qr_image(link)
        except ImportError:
            self.code.setPixmap(QtGui.QPixmap())
            self.code.setText("QR codes need the segno package. Enter the address and key on the device.")
            return
        self.code.setText("")
        self.code.setPixmap(QtGui.QPixmap.fromImage(image))
        self.code.setToolTip(normalize_endpoint(text))

    def _refresh_status(self) -> None:
        count = self._bridge.client_count
        devices = "No devices connected" if count == 0 else f"{count} device{'s' if count != 1 else ''} connected"
        control = "remote control allowed" if self._control_action.isChecked() else "read-only"
        set_semantic_status(
            self.status,
            f"Sharing on local port {self._bridge.port}. {devices}; {control}.",
            "active" if count else "ready",
            accessible_name="Mobile Companion sharing status",
        )
