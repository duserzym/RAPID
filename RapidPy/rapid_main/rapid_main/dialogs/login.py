from __future__ import annotations

from PySide6 import QtCore, QtWidgets

_PRESET_OPERATORS = [
    "Operator",
    "Lab Technician",
    "Graduate Student",
    "PI",
]


class LoginDialog(QtWidgets.QDialog):
    """Operator login — replaces VB6 frmLogin.

    Returns operator name via `operator_name` property after `exec()`.
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Operator Login")
        self.setFixedWidth(360)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    # ── Public ─────────────────────────────────────────────────────────────
    @property
    def operator_name(self) -> str:
        return self._name_edit.currentText().strip()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(24, 20, 24, 20)
        vl.setSpacing(14)

        hdr = QtWidgets.QLabel("Sign in to begin session")
        hdr.setStyleSheet("font-size: 15px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        fl = QtWidgets.QFormLayout()
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._name_edit = QtWidgets.QComboBox()
        self._name_edit.setEditable(True)
        self._name_edit.addItems(_PRESET_OPERATORS)
        self._name_edit.setCurrentIndex(-1)
        self._name_edit.lineEdit().setPlaceholderText("Your name or initials")
        fl.addRow("Operator:", self._name_edit)

        self._lab_lbl = QtWidgets.QLabel("IRM — University of Minnesota")
        self._lab_lbl.setStyleSheet("color: #6b7280; font-size: 11px;")
        fl.addRow("Laboratory:", self._lab_lbl)

        self._nocomm_chk = QtWidgets.QCheckBox("Start in No-Comm mode (no hardware)")
        fl.addRow("", self._nocomm_chk)

        vl.addLayout(fl)

        # ── Buttons ──────────────────────────────────────────────────────
        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        btns.accepted.connect(self._on_accept)
        btns.rejected.connect(self.reject)
        vl.addWidget(btns)

        self._name_edit.lineEdit().returnPressed.connect(self._on_accept)

    def _on_accept(self) -> None:
        if not self.operator_name:
            QtWidgets.QMessageBox.warning(self, "Login", "Please enter an operator name.")
            return
        self.accept()

    @property
    def nocomm(self) -> bool:
        return self._nocomm_chk.isChecked()
