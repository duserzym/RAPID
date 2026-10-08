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
        self.setObjectName("glassDialog")
        self.setWindowTitle("Operator Login")
        self.setAccessibleName("Operator session login")
        self.setMinimumWidth(320)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self.resize(360, self.sizeHint().height())

    # ── Public ─────────────────────────────────────────────────────────────
    @property
    def operator_name(self) -> str:
        return self._name_edit.currentText().strip()

    @property
    def operator_email(self) -> str:
        return self._email_edit.text().strip()

    def set_operator_email(self, address: str) -> None:
        self._email_edit.setText(str(address or ""))

    def set_operator_name(self, name: str) -> None:
        if str(name or "").strip():
            self._name_edit.setCurrentText(str(name).strip())

    def set_nocomm(self, enabled: bool) -> None:
        self._nocomm_chk.setChecked(bool(enabled))

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(24, 20, 24, 20)
        vl.setSpacing(14)

        from rapid_main.branding import brand_label

        title_row = QtWidgets.QHBoxLayout()
        title_row.setSpacing(12)
        title_row.addWidget(brand_label(44, self), 0, QtCore.Qt.AlignmentFlag.AlignVCenter)
        hdr = QtWidgets.QLabel("Sign in to begin session")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("Operator login title")
        title_row.addWidget(hdr, 1)
        vl.addLayout(title_row)

        fl = QtWidgets.QFormLayout()
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._name_edit = QtWidgets.QComboBox()
        self._name_edit.setEditable(True)
        self._name_edit.addItems(_PRESET_OPERATORS)
        self._name_edit.setCurrentIndex(-1)
        self._name_edit.lineEdit().setPlaceholderText("Your name or initials")
        self._name_edit.setAccessibleName("Operator name or initials")
        fl.addRow("Operator:", self._name_edit)

        self._email_edit = QtWidgets.QLineEdit()
        self._email_edit.setPlaceholderText("Optional — for run complete / error emails")
        self._email_edit.setAccessibleName("Operator email for notices")
        fl.addRow("Email:", self._email_edit)

        self._lab_lbl = QtWidgets.QLabel("IRM — University of Minnesota")
        self._lab_lbl.setObjectName("dialogSubtitle")
        self._lab_lbl.setWordWrap(True)
        self._lab_lbl.setAccessibleName("Laboratory name")
        fl.addRow("Laboratory:", self._lab_lbl)

        self._nocomm_chk = QtWidgets.QCheckBox("Start in No-Comm mode (no hardware)")
        self._nocomm_chk.setAccessibleName("Start session without hardware communication")
        self._nocomm_chk.setAccessibleDescription(
            "Uses simulation-only diagnostic backends and does not provide hardware evidence."
        )
        fl.addRow("", self._nocomm_chk)

        vl.addLayout(fl)

        # ── Buttons ──────────────────────────────────────────────────────
        self._buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self._buttons.setAccessibleName("Operator login actions")
        self._ok_button = self._buttons.button(QtWidgets.QDialogButtonBox.StandardButton.Ok)
        self._cancel_button = self._buttons.button(
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        if self._ok_button is not None:
            self._ok_button.setText("Start Session")
            self._ok_button.setAccessibleName("Start operator session")
        if self._cancel_button is not None:
            self._cancel_button.setAccessibleName("Cancel operator login")
        self._buttons.accepted.connect(self._on_accept)
        self._buttons.rejected.connect(self.reject)
        vl.addWidget(self._buttons)

        self._name_edit.lineEdit().returnPressed.connect(self._on_accept)

    def _on_accept(self) -> None:
        if not self.operator_name:
            QtWidgets.QMessageBox.warning(self, "Login", "Please enter an operator name.")
            return
        email = self.operator_email
        if email and ("@" not in email or " " in email or email.startswith("@") or email.endswith("@")):
            QtWidgets.QMessageBox.warning(self, "Login", "Please enter a valid email address or leave it blank.")
            return
        self.accept()

    @property
    def nocomm(self) -> bool:
        return self._nocomm_chk.isChecked()
