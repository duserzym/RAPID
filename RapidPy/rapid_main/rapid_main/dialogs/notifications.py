"""Email notification settings (replaces VB6 frmSendMail settings)."""

from __future__ import annotations

import threading
from pathlib import Path

from PySide6 import QtCore, QtWidgets

from rapid_main.glass_theme import set_semantic_status
from rapidpy_common.notify import (
    EmailSettings,
    NotificationSettings,
    Notifier,
    email_settings_from_vb6_ini,
)


class _TestBridge(QtCore.QObject):
    finished = QtCore.Signal(str)


class NotificationsDialog(QtWidgets.QDialog):
    """Edit who gets notices, which events, and how mail leaves this PC."""

    def __init__(self, parent: QtWidgets.QWidget | None, settings: NotificationSettings, *,
                 settings_path: Path, log_path: Path | None = None, operator_email: str = "") -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Email Notifications")
        self.setAccessibleName("Email notification settings")
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self.setMinimumWidth(420)
        self._settings = settings
        self._settings_path = Path(settings_path)
        self._log_path = log_path
        self._operator_email = operator_email
        self._bridge = _TestBridge(self)
        self._bridge.finished.connect(self._on_test_finished)
        self._build_ui()
        self._load(settings)

    # ── layout ───────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(20, 16, 20, 16)
        outer.setSpacing(12)
        title = QtWidgets.QLabel("Email notifications")
        title.setObjectName("dialogTitle")
        outer.addWidget(title)
        intro = QtWidgets.QLabel(
            "RAPID emails the logged-in operator and the CC list when a run completes, stops or fails. "
            "Messages are sent in the background and never pause a measurement."
        )
        intro.setObjectName("guidanceText")
        intro.setWordWrap(True)
        outer.addWidget(intro)

        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("dialogScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QtWidgets.QWidget()
        body = QtWidgets.QVBoxLayout(host)
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(10)

        self._enabled = QtWidgets.QCheckBox("Send email notices")
        self._enabled.setAccessibleName("Enable email notices")
        body.addWidget(self._enabled)

        who = QtWidgets.QGroupBox("Recipients")
        who_form = QtWidgets.QFormLayout(who)
        who_form.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)
        self._include_operator = QtWidgets.QCheckBox("Logged-in operator (email given at login)")
        who_form.addRow(self._include_operator)
        self._recipients = QtWidgets.QLineEdit()
        self._recipients.setPlaceholderText("lab-manager@umn.edu; status-monitor@umn.edu")
        self._recipients.setAccessibleName("Always copy these addresses")
        who_form.addRow("Always copy:", self._recipients)
        body.addWidget(who)

        what = QtWidgets.QGroupBox("Which notices")
        what_layout = QtWidgets.QVBoxLayout(what)
        self._codes = {
            "red": QtWidgets.QCheckBox("Red — errors and emergencies (run failed, hardware fault)"),
            "orange": QtWidgets.QCheckBox("Orange — attention (sample done: remove it, run halted)"),
            "yellow": QtWidgets.QCheckBox("Yellow — recoverable issues (re-measure, flux jump)"),
            "green": QtWidgets.QCheckBox("Green — run complete"),
        }
        for box in self._codes.values():
            what_layout.addWidget(box)
        body.addWidget(what)

        how = QtWidgets.QGroupBox("Sending")
        how_form = QtWidgets.QFormLayout(how)
        how_form.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)
        self._method = QtWidgets.QComboBox()
        self._method.addItem("SMTP server", "smtp")
        self._method.addItem("Classic Outlook on this PC (no password stored)", "outlook")
        self._method.currentIndexChanged.connect(self._update_method)
        how_form.addRow("Send with:", self._method)
        self._host = QtWidgets.QLineEdit()
        how_form.addRow("SMTP host:", self._host)
        self._port = QtWidgets.QSpinBox()
        self._port.setRange(1, 65535)
        how_form.addRow("Port:", self._port)
        self._security = QtWidgets.QComboBox()
        for label, value in (("STARTTLS (587)", "starttls"), ("SSL/TLS (465)", "ssl"), ("None", "none")):
            self._security.addItem(label, value)
        how_form.addRow("Security:", self._security)
        self._username = QtWidgets.QLineEdit()
        how_form.addRow("Username:", self._username)
        self._password = QtWidgets.QLineEdit()
        self._password.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self._password.setAccessibleName("SMTP password")
        how_form.addRow("Password:", self._password)
        self._sender = QtWidgets.QLineEdit()
        self._sender.setPlaceholderText("Defaults to the username")
        how_form.addRow("From address:", self._sender)
        self._sender_name = QtWidgets.QLineEdit()
        how_form.addRow("From name:", self._sender_name)
        self._smtp_widgets = (self._host, self._port, self._security, self._username, self._password, self._sender)
        body.addWidget(how)
        body.addStretch(1)
        scroll.setWidget(host)
        outer.addWidget(scroll, 1)

        actions = QtWidgets.QHBoxLayout()
        self._import_btn = QtWidgets.QPushButton("Import from VB6 INI…")
        self._import_btn.clicked.connect(self._import_vb6)
        self._test_btn = QtWidgets.QPushButton("Send Test Email")
        self._test_btn.clicked.connect(self._send_test)
        actions.addWidget(self._import_btn)
        actions.addWidget(self._test_btn)
        actions.addStretch(1)
        outer.addLayout(actions)
        self._status = QtWidgets.QLabel()
        self._status.setWordWrap(True)
        outer.addWidget(self._status)
        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)

    # ── settings <-> widgets ─────────────────────────────────────────────
    def _load(self, settings: NotificationSettings) -> None:
        self._enabled.setChecked(settings.email_enabled)
        self._include_operator.setChecked(settings.include_operator)
        for code, box in self._codes.items():
            box.setChecked(bool(getattr(settings, f"notify_{code}")))
        self._load_email(settings.email)
        stored = "a password is stored (encrypted for this Windows user)" if settings.email.password_protected else "no password stored"
        self._password.setPlaceholderText(f"Leave blank to keep: {stored}")
        set_semantic_status(self._status, "Edit settings, then Save or send a test.", "neutral",
                            accessible_name="Notification status")

    def _load_email(self, email: EmailSettings) -> None:
        index = self._method.findData(email.method)
        self._method.setCurrentIndex(max(0, index))
        self._recipients.setText("; ".join(email.recipients))
        self._host.setText(email.smtp_host)
        self._port.setValue(int(email.smtp_port))
        self._security.setCurrentIndex(max(0, self._security.findData(email.security)))
        self._username.setText(email.username)
        self._sender.setText(email.sender)
        self._sender_name.setText(email.sender_name)
        self._update_method()

    def _update_method(self) -> None:
        smtp = self._method.currentData() == "smtp"
        for widget in self._smtp_widgets:
            widget.setEnabled(smtp)

    def settings(self) -> NotificationSettings:
        """The edited settings (the stored password is kept unless a new one is typed)."""
        current = self._settings
        email = EmailSettings(
            recipients=[a.strip() for a in self._recipients.text().replace(",", ";").split(";") if a.strip()],
            method=str(self._method.currentData()),
            smtp_host=self._host.text().strip(),
            smtp_port=int(self._port.value()),
            security=str(self._security.currentData()),
            username=self._username.text().strip(),
            password_protected=current.email.password_protected,
            sender=self._sender.text().strip(),
            sender_name=self._sender_name.text().strip(),
        )
        if self._password.text():
            email.set_password(self._password.text())
        return NotificationSettings(
            email_enabled=self._enabled.isChecked(),
            email=email,
            include_operator=self._include_operator.isChecked(),
            notify_red=self._codes["red"].isChecked(),
            notify_orange=self._codes["orange"].isChecked(),
            notify_yellow=self._codes["yellow"].isChecked(),
            notify_green=self._codes["green"].isChecked(),
        )

    # ── actions ──────────────────────────────────────────────────────────
    def _import_vb6(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Import VB6 email settings", "", "INI files (*.ini *.INI)")
        if path:
            self.import_vb6_ini(Path(path))

    def import_vb6_ini(self, path: Path) -> None:
        email, notes = email_settings_from_vb6_ini(path)
        self._load_email(email)
        message = "Imported the VB6 [Email] settings." + (" " + " ".join(notes) if notes else "")
        set_semantic_status(self._status, message, "warning" if notes else "ready",
                            accessible_name="Notification status")

    def _send_test(self) -> None:
        settings = self.settings()
        notifier = Notifier(settings, log_path=self._log_path, retries=0)
        self._test_btn.setEnabled(False)
        set_semantic_status(self._status, "Sending test email…", "active", accessible_name="Notification status")

        def run() -> None:
            self._bridge.finished.emit(notifier.send_test(self._operator_email))

        threading.Thread(target=run, daemon=True, name="rapid-notify-test").start()

    def _on_test_finished(self, result: str) -> None:
        self._test_btn.setEnabled(True)
        set_semantic_status(self._status, result, "ready" if "sent" in result and "failed" not in result else "error",
                            accessible_name="Notification status")

    def _save(self) -> None:
        updated = self.settings()
        try:
            updated.save(self._settings_path)
        except OSError as exc:
            set_semantic_status(self._status, f"Unable to save: {exc}", "error", accessible_name="Notification status")
            return
        self._settings = updated
        self.accept()

    @property
    def saved_settings(self) -> NotificationSettings:
        return self._settings
