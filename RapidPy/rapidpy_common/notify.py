"""Email notices from RapidPy instruments (measurement complete, errors, ...).

Ported from the email path of ``asc_oven_control.infrastructure.notify``
(Institute-for-Rock-Magnetism/ASC_oven_control) and shaped after VB6
``frmSendMail.MailNotification``:

* Every notice carries a VB6 status code -- Red (emergency), Orange
  (attention required), Yellow (oops), Green (everything running) -- and the
  operator, sample, step, orientation and time in its footer.
* It goes to the logged-in operator's address plus a CC list (VB6
  ``LoginEmail`` + ``MailCCList``/``MailStatusMonitor``).
* Email leaves either over SMTP -- the password is stored encrypted with
  Windows DPAPI, readable only by this Windows user on this PC -- or through
  classic Outlook on this PC (no password in the app).

Sending never blocks the caller and never raises: each notice is delivered
on a daemon thread with a timeout and retries, and every attempt is appended
to a log file.  Settings live next to the user's RapidPy config
(``notifications.json``), never in the repository.  VB6 kept the SMTP
password in its INI in plain text; :func:`email_settings_from_vb6_ini`
deliberately does not import it.
"""

from __future__ import annotations

import configparser
import json
import os
import threading
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

CODE_RED = "Red"  # EMERGENCY!
CODE_ORANGE = "Orange"  # Attention required
CODE_YELLOW = "Yellow"  # Oops!
CODE_GREEN = "Green"  # Everything running
CODES = (CODE_RED, CODE_ORANGE, CODE_YELLOW, CODE_GREEN)


# ── secrets (Windows DPAPI) ─────────────────────────────────────────────────


def protect_secret(secret: str) -> str:
    """Encrypt for this Windows user (DPAPI); base64 text. Plain prefix elsewhere."""
    if not secret:
        return ""
    if os.name != "nt":
        return "plain:" + secret
    import base64

    blob = _dpapi(secret.encode("utf-8"), protect=True)
    return "dpapi:" + base64.b64encode(blob).decode("ascii")


def reveal_secret(stored: str) -> str:
    if not stored:
        return ""
    if stored.startswith("plain:"):
        return stored[6:]
    if stored.startswith("dpapi:") and os.name == "nt":
        import base64

        try:
            return _dpapi(base64.b64decode(stored[6:]), protect=False).decode("utf-8")
        except OSError:
            return ""
    return ""


def _dpapi(data: bytes, protect: bool) -> bytes:
    import ctypes
    from ctypes import wintypes

    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buffer = ctypes.create_string_buffer(data, len(data))
    source = Blob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
    target = Blob()
    crypt32 = ctypes.windll.crypt32
    function = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    ok = function(ctypes.byref(source), None, None, None, None, 0, ctypes.byref(target))
    if not ok:
        raise OSError("DPAPI failed")
    try:
        return ctypes.string_at(target.pbData, target.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(target.pbData)


# ── settings ────────────────────────────────────────────────────────────────


def _addresses(values) -> list[str]:
    if isinstance(values, str):
        values = values.replace(",", ";").split(";")
    return [str(v).strip() for v in values or [] if str(v).strip()]


@dataclass
class EmailSettings:
    """``method``: "smtp" or "outlook" (classic Outlook on this PC)."""

    recipients: list[str] = field(default_factory=list)  # always copied (VB6 MailCCList)
    method: str = "smtp"
    smtp_host: str = "smtp.gmail.com"
    smtp_port: int = 587
    security: str = "starttls"  # "starttls", "ssl" or "none"
    username: str = ""
    password_protected: str = ""  # see protect_secret(); never plain text on Windows
    sender: str = ""
    sender_name: str = "RAPID Magnetometer"

    def configured(self, extra_recipients: list[str] | None = None) -> bool:
        if not (self.recipients or _addresses(extra_recipients or [])):
            return False
        if self.method == "outlook":
            return True
        return bool(self.smtp_host and (self.sender or self.username))

    def set_password(self, password: str) -> None:
        self.password_protected = protect_secret(password)

    def password(self) -> str:
        return reveal_secret(self.password_protected)


@dataclass
class NotificationSettings:
    email_enabled: bool = False
    email: EmailSettings = field(default_factory=EmailSettings)
    include_operator: bool = True  # also send to the logged-in operator's address
    notify_red: bool = True  # errors / emergencies
    notify_orange: bool = True  # attention required (sample done, run halted)
    notify_yellow: bool = False  # recoverable oops (re-measure, flux jump)
    notify_green: bool = True  # run complete, calibration done

    def wants(self, code: str) -> bool:
        return {
            CODE_RED: self.notify_red,
            CODE_ORANGE: self.notify_orange,
            CODE_YELLOW: self.notify_yellow,
            CODE_GREEN: self.notify_green,
        }.get(code, True)

    @classmethod
    def load(cls, path: Path | str) -> "NotificationSettings":
        path = Path(path)
        if not path.exists():
            return cls()
        try:
            return cls.from_dict(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError, TypeError):
            return cls()

    def save(self, path: Path | str) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        data = asdict(self)
        data["about"] = (
            "RapidPy email notices. The SMTP password is stored DPAPI-encrypted for this Windows "
            "user (or use classic Outlook on this PC). Keep this file out of version control."
        )
        temporary = path.with_name(path.name + ".tmp")
        temporary.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
        os.replace(temporary, path)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "NotificationSettings":
        settings = cls()
        settings.email_enabled = bool(data.get("email_enabled", False))
        settings.include_operator = bool(data.get("include_operator", True))
        for code in ("red", "orange", "yellow", "green"):
            key = f"notify_{code}"
            setattr(settings, key, bool(data.get(key, getattr(settings, key))))
        email = data.get("email") or {}
        settings.email = EmailSettings(
            recipients=_addresses(email.get("recipients", [])),
            method=str(email.get("method", "smtp")),
            smtp_host=str(email.get("smtp_host", "smtp.gmail.com")),
            smtp_port=int(email.get("smtp_port", 587)),
            security=str(email.get("security", "starttls")),
            username=str(email.get("username", "")),
            password_protected=str(email.get("password_protected", "")),
            sender=str(email.get("sender", "")),
            sender_name=str(email.get("sender_name", "RAPID Magnetometer")),
        )
        return settings


def email_settings_from_vb6_ini(path: Path | str) -> tuple[EmailSettings, list[str]]:
    """Map the VB6 ``[Email]`` section; the plain-text password is never imported."""

    parser = configparser.ConfigParser(interpolation=None)
    parser.optionxform = str
    parser.read(str(path), encoding="latin-1")
    notes: list[str] = []
    if not parser.has_section("Email"):
        return EmailSettings(), ["The INI has no [Email] section."]
    section = parser["Email"]

    def value(key: str, default: str = "") -> str:
        return str(section.get(key, default)).strip()

    use_ssl = value("MailUseSSLEncryption").lower() in ("true", "1", "yes")
    port = int(value("MailSMTPPort", "465" if use_ssl else "587") or 0) or (465 if use_ssl else 587)
    settings = EmailSettings(
        recipients=_addresses(value("MailCCList")) + _addresses(value("MailStatusMonitor")),
        method="smtp",
        smtp_host=value("MailSMTPHost", "smtp.gmail.com"),
        smtp_port=port,
        security="ssl" if use_ssl else "starttls",
        username=value("MailSMTPUsername") or value("MailFrom"),
        sender=value("MailFrom"),
        sender_name=value("MailFromName", "RAPID Magnetometer"),
    )
    if value("MailFromPassword") or value("MailSMTPPassword"):
        notes.append(
            "VB6 stored the SMTP password in the INI in plain text; it was not imported. "
            "Enter it here (it is stored encrypted for this Windows user) and change the "
            "password of that mail account, because the INI copy is readable by anyone."
        )
    return settings, notes


# ── message text ────────────────────────────────────────────────────────────


def compose(subject: str, body: str, code: str, context: dict | None = None) -> tuple[str, str]:
    """Subject and VB6-style body (operator header, sample/step/time/code footer)."""

    context = dict(context or {})
    operator = str(context.get("operator") or "").strip()
    head = f"{operator}:\n\n" if operator else ""
    lines = []
    for label, key in (("Sample", "sample"), ("Step", "step"), ("Orientation", "orientation"),
                       ("Run", "run"), ("Station", "station")):
        if context.get(key):
            lines.append(f"{label}: {context[key]}")
    lines.append(f"Time: {datetime.now():%Y-%m-%d %H:%M:%S}")
    lines.append(f"Code: {code}")
    text = f"{head}{body}\n\n" + "\n".join(lines) + "\n\n-- RAPID (automatic notice)"
    prefix = {CODE_RED: "ALERT - ", CODE_ORANGE: "Attention - "}.get(code, "")
    return f"RAPID: {prefix}{subject}", text


_OUTLOOK_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = $env:RAPID_MAIL_TO
$mail.Subject = $env:RAPID_MAIL_SUBJECT
$mail.Body = $env:RAPID_MAIL_BODY
$mail.Send()
"""


def send_with_outlook(recipients: list[str], subject: str, text: str, timeout_s: float = 60.0) -> None:
    """Send through classic Outlook on this PC (its signed-in account).

    The text travels in environment variables, never on the command line, so
    nothing in a message can be interpreted as a command.
    """
    import subprocess

    env = dict(os.environ, RAPID_MAIL_TO="; ".join(recipients), RAPID_MAIL_SUBJECT=subject, RAPID_MAIL_BODY=text)
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    done = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", _OUTLOOK_SCRIPT],
        env=env, capture_output=True, text=True, timeout=timeout_s, creationflags=flags,
    )
    if done.returncode != 0:
        raise OSError((done.stderr or done.stdout or "Outlook send failed").strip().splitlines()[-1][:200])


# ── delivery ────────────────────────────────────────────────────────────────


class Notifier:
    """Fire-and-forget email delivery; never raises into the caller, never blocks."""

    def __init__(self, settings: NotificationSettings, log_path: Path | str | None = None, *,
                 retries: int = 2, timeout_s: float = 20.0, smtp_factory=None,
                 outlook_sender: Callable[..., None] | None = None,
                 sleep: Callable[[float], None] = time.sleep) -> None:
        self.settings = settings
        self.log_path = Path(log_path) if log_path else None
        self.retries = retries
        self.timeout_s = timeout_s
        self.smtp_factory = smtp_factory  # (host, port, timeout, ssl) -> smtplib-like client
        self.outlook_sender = outlook_sender or send_with_outlook
        self._sleep = sleep
        self._lock = threading.Lock()
        self.threads: list[threading.Thread] = []

    def recipients(self, operator_email: str = "") -> list[str]:
        targets = list(self.settings.email.recipients)
        if self.settings.include_operator and operator_email.strip():
            targets.insert(0, operator_email.strip())
        seen: set[str] = set()
        return [t for t in targets if not (t.lower() in seen or seen.add(t.lower()))]

    def active(self, operator_email: str = "") -> bool:
        return self.settings.email_enabled and self.settings.email.configured(self.recipients(operator_email))

    def notify(self, subject: str, body: str, code: str = CODE_GREEN, *, context: dict | None = None,
               operator_email: str = "") -> bool:
        """Queue a notice; returns whether one was sent off."""
        if not self.active(operator_email) or not self.settings.wants(code):
            return False
        full_subject, text = compose(subject, body, code, context)
        targets = self.recipients(operator_email)
        thread = threading.Thread(target=self._deliver, args=(targets, full_subject, text), daemon=True,
                                  name="rapid-notify-email")
        thread.start()
        self.threads = [t for t in self.threads if t.is_alive()] + [thread]
        return True

    def send_test(self, operator_email: str = "") -> str:
        """Deliver a test message now (blocking; run it off the UI thread)."""
        targets = self.recipients(operator_email)
        if not targets:
            return "No recipients: add a CC address or log in with an email address."
        subject, text = compose("Test notice", "This is a test of RAPID email notifications.", CODE_GREEN,
                                {"operator": "RAPID"})
        return self._deliver(targets, subject, text, retries=0)

    def wait(self, timeout_s: float = 60.0) -> None:
        deadline = time.monotonic() + timeout_s
        for thread in list(self.threads):
            thread.join(max(deadline - time.monotonic(), 0.0))

    def _deliver(self, targets: list[str], subject: str, text: str, retries: int | None = None) -> str:
        email = self.settings.email
        attempts = (self.retries if retries is None else retries) + 1
        shown = ", ".join(targets)
        result = ""
        for attempt in range(attempts):
            try:
                if email.method == "outlook":
                    self.outlook_sender(targets, subject, text, self.timeout_s * 3)
                else:
                    self._send_smtp(email, targets, subject, text)
                result = f"email sent to {shown} via {email.method}"
                break
            except Exception as exc:  # noqa: BLE001 - never let a notice break a run
                result = f"email to {shown} via {email.method} failed: {exc}"
                if attempt + 1 < attempts:
                    self._sleep(min(10.0 * (attempt + 1), 30.0))
        self._log(f"{result} | {subject!r}")
        return result

    def _send_smtp(self, email: EmailSettings, targets: list[str], subject: str, text: str) -> None:
        import smtplib
        import ssl
        from email.message import EmailMessage
        from email.utils import formataddr

        message = EmailMessage()
        sender = email.sender or email.username
        message["From"] = formataddr((email.sender_name, sender)) if email.sender_name else sender
        message["To"] = ", ".join(targets)
        message["Subject"] = subject
        message.set_content(text)
        context = ssl.create_default_context()
        if self.smtp_factory is not None:
            client = self.smtp_factory(email.smtp_host, email.smtp_port, self.timeout_s, email.security == "ssl")
        elif email.security == "ssl":
            client = smtplib.SMTP_SSL(email.smtp_host, email.smtp_port, timeout=self.timeout_s, context=context)
        else:
            client = smtplib.SMTP(email.smtp_host, email.smtp_port, timeout=self.timeout_s)
        with client:
            if email.security == "starttls":
                client.starttls(context=context)
            if email.username:
                client.login(email.username, email.password())
            client.send_message(message)

    def _log(self, line: str) -> None:
        if self.log_path is None:
            return
        with self._lock:
            try:
                self.log_path.parent.mkdir(parents=True, exist_ok=True)
                with open(self.log_path, "a", encoding="utf-8") as handle:
                    handle.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {line}\n")
            except OSError:
                pass
