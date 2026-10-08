"""Email notices: VB6-style composition, routing, delivery, secrets and app hooks."""
from __future__ import annotations

import os
import tempfile
import threading
import unittest
from pathlib import Path

from PySide6 import QtWidgets

from rapidpy_common.notify import (
    CODE_GREEN,
    CODE_ORANGE,
    CODE_RED,
    CODE_YELLOW,
    EmailSettings,
    NotificationSettings,
    Notifier,
    compose,
    email_settings_from_vb6_ini,
    protect_secret,
    reveal_secret,
)

_APP = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


class FakeSMTP:
    instances: list["FakeSMTP"] = []

    def __init__(self, host, port, timeout, use_ssl, *, fail_times: int = 0):
        self.host, self.port, self.use_ssl = host, port, use_ssl
        self.started_tls = False
        self.login_args = None
        self.messages = []
        FakeSMTP.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def starttls(self, context=None):
        self.started_tls = True

    def login(self, user, password):
        self.login_args = (user, password)

    def send_message(self, message):
        self.messages.append(message)


def _settings(**overrides) -> NotificationSettings:
    email = EmailSettings(recipients=["lab@umn.edu"], smtp_host="smtp.umn.edu", smtp_port=587,
                          username="rapid@umn.edu", sender="rapid@umn.edu")
    email.set_password("s3cret")
    values = dict(email_enabled=True, email=email)
    values.update(overrides)
    return NotificationSettings(**values)


class ComposeTests(unittest.TestCase):
    def test_vb6_style_body_and_subject(self):
        subject, text = compose("Measurement run error", "Queue stopped.", CODE_RED,
                                {"operator": "YZ", "sample": "BG01-1", "step": "AF20", "station": "IRM"})
        self.assertEqual(subject, "RAPID: ALERT - Measurement run error")
        self.assertTrue(text.startswith("YZ:\n\nQueue stopped."))
        for line in ("Sample: BG01-1", "Step: AF20", "Station: IRM", "Code: Red"):
            self.assertIn(line, text)
        self.assertEqual(compose("Run complete", "ok", CODE_GREEN)[0], "RAPID: Run complete")
        self.assertIn("Attention", compose("Sample done", "x", CODE_ORANGE)[0])


class NotifierTests(unittest.TestCase):
    def setUp(self):
        FakeSMTP.instances = []

    def test_routes_to_operator_and_cc_without_duplicates(self):
        notifier = Notifier(_settings())
        self.assertEqual(notifier.recipients("yz@umn.edu"), ["yz@umn.edu", "lab@umn.edu"])
        self.assertEqual(notifier.recipients("LAB@umn.edu"), ["LAB@umn.edu"])
        notifier.settings.include_operator = False
        self.assertEqual(notifier.recipients("yz@umn.edu"), ["lab@umn.edu"])

    def test_disabled_or_filtered_codes_send_nothing(self):
        sent = []
        factory = lambda *a: sent.append(a) or FakeSMTP(*a)  # noqa: E731
        self.assertFalse(Notifier(_settings(email_enabled=False), smtp_factory=factory).notify("x", "y"))
        self.assertFalse(Notifier(_settings(), smtp_factory=factory).notify("x", "y", CODE_YELLOW))  # off by default
        self.assertFalse(Notifier(_settings(notify_red=False), smtp_factory=factory).notify("x", "y", CODE_RED))
        self.assertEqual(sent, [])

    def test_smtp_delivery_uses_starttls_and_decrypted_password(self):
        notifier = Notifier(_settings(), smtp_factory=FakeSMTP)
        self.assertTrue(notifier.notify("Run complete", "All done.", CODE_GREEN, operator_email="yz@umn.edu",
                                        context={"operator": "YZ"}))
        notifier.wait(5)
        client = FakeSMTP.instances[0]
        self.assertTrue(client.started_tls)
        self.assertEqual(client.login_args, ("rapid@umn.edu", "s3cret"))
        message = client.messages[0]
        self.assertEqual(message["Subject"], "RAPID: Run complete")
        self.assertIn("yz@umn.edu", message["To"])
        self.assertIn("RAPID Magnetometer", message["From"])

    def test_failure_is_retried_logged_and_never_raises(self):
        attempts = []

        def flaky(*args):
            attempts.append(args)
            if len(attempts) == 1:
                raise OSError("network down")
            return FakeSMTP(*args)

        with tempfile.TemporaryDirectory() as tmp:
            log = Path(tmp) / "notifications.log"
            notifier = Notifier(_settings(), log_path=log, smtp_factory=flaky, retries=1, sleep=lambda _s: None)
            notifier.notify("Run error", "boom", CODE_RED)
            notifier.wait(5)
            text = log.read_text(encoding="utf-8")
        self.assertEqual(len(attempts), 2)
        self.assertIn("email sent to lab@umn.edu via smtp", text)
        self.assertNotIn("s3cret", text)

    def test_outlook_method_needs_no_smtp(self):
        calls = []
        settings = _settings()
        settings.email.method = "outlook"
        notifier = Notifier(settings, outlook_sender=lambda *a: calls.append(a))
        notifier.notify("Sample done", "Remove sample.", CODE_ORANGE)
        notifier.wait(5)
        self.assertEqual(calls[0][0], ["lab@umn.edu"])
        self.assertIn("Sample done", calls[0][1])

    def test_notify_returns_immediately(self):
        gate = threading.Event()

        def slow(*args):
            gate.wait(5)
            return FakeSMTP(*args)

        notifier = Notifier(_settings(), smtp_factory=slow)
        self.assertTrue(notifier.notify("x", "y"))  # returns while delivery is still blocked
        gate.set()
        notifier.wait(5)


class SecretsAndSettingsTests(unittest.TestCase):
    def test_password_round_trip_never_plain_on_windows(self):
        stored = protect_secret("hunter2")
        self.assertEqual(reveal_secret(stored), "hunter2")
        if os.name == "nt":
            self.assertTrue(stored.startswith("dpapi:"))
            self.assertNotIn("hunter2", stored)

    def test_settings_file_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "notifications.json"
            _settings(notify_yellow=True).save(path)
            loaded = NotificationSettings.load(path)
            raw = path.read_text(encoding="utf-8")
        self.assertTrue(loaded.email_enabled and loaded.notify_yellow)
        self.assertEqual(loaded.email.recipients, ["lab@umn.edu"])
        self.assertEqual(loaded.email.password(), "s3cret")
        if os.name == "nt":
            self.assertNotIn("s3cret", raw)
        self.assertFalse(NotificationSettings.load(Path(tmp) / "missing.json").email_enabled)

    def test_vb6_import_maps_settings_but_never_the_password(self):
        with tempfile.TemporaryDirectory() as tmp:
            ini = Path(tmp) / "Paleomag.INI"
            ini.write_text(
                "[Email]\nMailFromName=Hargraves Magnetometer\nMailFromPassword=plaintext\n"
                "MailFrom=station@example.org\nMailSMTPHost=smtp.example.org\nMailCCList=pi@example.org\n"
                "MailSMTPPort=465\nMailSMTPPassword=plaintext\nMailSMTPUsername=station@example.org\n"
                "MailUseSSLEncryption=True\nMailStatusMonitor=monitor@example.org\n",
                encoding="utf-8",
            )
            email, notes = email_settings_from_vb6_ini(ini)
        self.assertEqual((email.smtp_host, email.smtp_port, email.security), ("smtp.example.org", 465, "ssl"))
        self.assertEqual(email.recipients, ["pi@example.org", "monitor@example.org"])
        self.assertEqual(email.sender_name, "Hargraves Magnetometer")
        self.assertEqual(email.password_protected, "")
        self.assertIn("plain text", notes[0])


class _RecordingNotifier:
    def __init__(self):
        self.calls = []
        self.settings = NotificationSettings()

    def notify(self, subject, body, code=CODE_GREEN, *, context=None, operator_email=""):
        self.calls.append((subject, body, code, dict(context or {}), operator_email))
        return True


class MainWindowNoticeTests(unittest.TestCase):
    def _window(self):
        from rapid_main.app import MainWindow

        window = MainWindow()
        window._notifier = _RecordingNotifier()
        window.config.general.operator = "YZ"
        window.config.general.operator_email = "yz@umn.edu"
        return window

    def _dispose(self, window):
        window._confirm_shutdown = lambda **kwargs: True
        window.close()
        window.deleteLater()

    def test_queue_terminal_states_send_coded_notices(self):
        window = self._window()
        try:
            window._notify_queue_terminal("complete", "Queue finished: 12 samples.")
            window._notify_queue_terminal("halted", None)
            window._notify_queue_terminal("error", "Lift did not reach the measurement position.")
            window._notify_queue_terminal("idle", None)
            codes = [call[2] for call in window._notifier.calls]
            self.assertEqual(codes, [CODE_GREEN, CODE_ORANGE, CODE_RED])
            subject, body, _code, context, address = window._notifier.calls[2]
            self.assertEqual(subject, "Measurement run error")
            self.assertIn("Lift did not reach", body)
            self.assertEqual((context["operator"], address), ("YZ", "yz@umn.edu"))
        finally:
            self._dispose(window)

    def test_finalize_queue_run_reports_its_terminal_state(self):
        window = self._window()
        try:
            window._release_queue_lease = lambda: True
            window._finalize_queue_run("complete", reason="Queue complete.", return_to_safe=False)
            self.assertEqual(window._notifier.calls[-1][:3], ("Measurement run complete", "Queue complete.", CODE_GREEN))
        finally:
            self._dispose(window)

    def test_manual_sample_done_asks_operator_to_remove_sample(self):
        window = self._window()
        try:
            window._queue_active = False
            window._on_queue_sample_finished(False, "BG01-1")
            subject, body, code, _context, _address = window._notifier.calls[-1]
            self.assertEqual((subject, code), ("Sample done", CODE_ORANGE))
            self.assertIn("Please remove sample", body)
        finally:
            self._dispose(window)

    def test_notice_failures_never_reach_the_run(self):
        window = self._window()
        try:
            class Broken:
                def notify(self, *a, **k):
                    raise RuntimeError("smtp exploded")

            window._notifier = Broken()
            self.assertFalse(window.notify("x", "y", CODE_RED))
        finally:
            self._dispose(window)


class NotificationsDialogTests(unittest.TestCase):
    def test_blank_password_keeps_stored_one_and_save_writes_file(self):
        from rapid_main.dialogs.notifications import NotificationsDialog

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "notifications.json"
            dialog = NotificationsDialog(None, _settings(), settings_path=path)
            edited = dialog.settings()
            self.assertEqual(edited.email.password(), "s3cret")
            dialog._password.setText("n3w")
            self.assertEqual(dialog.settings().email.password(), "n3w")
            dialog._save()
            self.assertTrue(path.exists())
            self.assertEqual(NotificationSettings.load(path).email.password(), "n3w")
            dialog.deleteLater()


if __name__ == "__main__":
    unittest.main()
