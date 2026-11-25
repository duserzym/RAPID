from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from unittest.mock import patch
from uuid import uuid4

from PySide6 import QtCore, QtNetwork, QtTest, QtWebSockets, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.companion import (
    PATH, CompanionBridge, load_station_address, normalize_endpoint, pairing_link, save_station_address,
)
from rapid_main.config import AppConfig


class CompanionBridgeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def setUp(self):
        self.phase = "measuring"
        self.run_epoch = 1
        self.actions = []
        self.bridge = CompanionBridge(self._read_state, self._dispatch, self.app)
        self.addCleanup(self.bridge.deleteLater)
        self.addCleanup(self.bridge.stop)
        self.bridge.start(0)

    def _read_state(self):
        allowed = {"measuring": ["pause", "halt"], "paused": ["resume", "halt"]}
        return self.run_epoch, {
            "station": "RAPID Test",
            "mode": "simulation",
            "phase": self.phase,
            "sample": "DEMO",
            "step": "Measure",
            "treatment": "NRM",
            "elapsed": "00:00:00",
            "allowed_commands": allowed.get(self.phase, []),
            "diagnostics": [{"name": "SQUID", "state": "simulated"}],
        }

    def _dispatch(self, action):
        self.actions.append(action)
        self.phase = {"pause": "paused", "resume": "measuring", "halt": "halted"}[action]

    def _until(self, predicate):
        deadline = QtCore.QDeadlineTimer(1500)
        while not predicate() and not deadline.hasExpired():
            QtTest.QTest.qWait(10)
        self.assertTrue(predicate())

    def _connect(self):
        socket = QtWebSockets.QWebSocket()
        self.addCleanup(socket.deleteLater)
        self.addCleanup(socket.abort)
        messages = []
        socket.textMessageReceived.connect(lambda text: messages.append(json.loads(text)))
        socket.open(QtCore.QUrl(f"ws://127.0.0.1:{self.bridge.port}{PATH}"))
        self._until(lambda: socket.state() == QtNetwork.QAbstractSocket.SocketState.ConnectedState)
        return socket, messages

    @staticmethod
    def _send(socket, message_type, **fields):
        socket.sendTextMessage(json.dumps(dict(v=1, type=message_type, **fields)))

    def _authenticate(self):
        socket, messages = self._connect()
        self._send(socket, "hello", token=self.bridge.token)
        self._until(lambda: messages and messages[-1]["type"] == "snapshot")
        return socket, messages

    def _command(self, socket, messages, action, revision):
        request_id = str(uuid4())
        self._send(socket, "command", id=request_id, action=action, revision=revision)
        self._until(lambda: any(event.get("id") == request_id for event in messages))
        return next(event for event in messages if event.get("id") == request_id)

    def test_bad_key_closes_without_releasing_status(self):
        socket, messages = self._connect()
        self._send(socket, "hello", token="invalid")
        self._until(lambda: socket.state() == QtNetwork.QAbstractSocket.SocketState.UnconnectedState)
        self.assertEqual(messages, [])

    def test_status_and_existing_workflow_handlers_stay_in_sync(self):
        first_client, first_messages = self._authenticate()
        second_client, second_messages = self._authenticate()
        initial = first_messages[-1]["snapshot"]
        self.assertEqual(initial["allowed_commands"], [])
        self.assertFalse(self._command(first_client, first_messages, "pause", initial["revision"])["accepted"])

        self.bridge.control_enabled = True
        self.bridge.publish()
        self._until(lambda: first_messages[-1]["snapshot"]["control_enabled"])
        current = first_messages[-1]["snapshot"]
        old_revision = current["revision"]
        self.assertEqual(current["allowed_commands"], ["pause", "halt"])

        paused = self._command(first_client, first_messages, "pause", old_revision)
        self.assertTrue(paused["accepted"])
        self.assertEqual(paused["snapshot"]["phase"], "paused")
        self._until(lambda: second_messages[-1]["snapshot"]["phase"] == "paused")
        self.assertFalse(
            self._command(second_client, second_messages, "halt", old_revision)["accepted"]
        )

        resumed = self._command(first_client, first_messages, "resume", paused["snapshot"]["revision"])
        self.assertTrue(resumed["accepted"])
        self.assertEqual(resumed["snapshot"]["phase"], "measuring")

        self.run_epoch += 1
        stale_run = self._command(first_client, first_messages, "halt", resumed["snapshot"]["revision"])
        self.assertFalse(stale_run["accepted"])
        halted = self._command(first_client, first_messages, "halt", stale_run["snapshot"]["revision"])
        self.assertTrue(halted["accepted"])
        self.assertEqual(halted["snapshot"]["phase"], "halted")
        self.assertEqual(self.actions, ["pause", "resume", "halt"])

        self.bridge.stop()
        self._until(lambda: first_client.state() == QtNetwork.QAbstractSocket.SocketState.UnconnectedState)
        self._until(lambda: second_client.state() == QtNetwork.QAbstractSocket.SocketState.UnconnectedState)

    def test_desktop_pairing_dialog_starts_read_only_and_uses_workflow_handlers(self):
        self.bridge.stop()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        settings = Path(folder.name) / "companion.json"
        config = AppConfig()
        config.general.nocomm = True
        with (
            patch.object(AppConfig, "load", return_value=config),
            patch.object(AppConfig, "save"),
            patch.object(MainWindow, "_operating_mode_change_blocker", lambda self, on=None: ""),
        ):
            window = MainWindow()
        self.addCleanup(window.deleteLater)
        self.addCleanup(window._stop_companion_sharing)
        start_on_port = CompanionBridge.start
        with (
            patch.object(MainWindow, "_companion_settings_path", staticmethod(lambda: settings)),
            patch.object(CompanionBridge, "start", lambda bridge: start_on_port(bridge, 0)),
        ):
            window._start_companion_sharing()
            bridge = window._companion
            dialog = window._companion_dialog
            self.assertIsNotNone(bridge)
            self.assertTrue(dialog.isVisible())
            self.assertFalse(dialog.isModal())
            self.assertIsNone(QtWidgets.QApplication.activeModalWidget())
            self.assertEqual(dialog.key.text(), bridge.token)
            self.assertFalse(bridge.control_enabled)
            self.assertFalse(dialog.control.isChecked())
            self.assertTrue(dialog.code.pixmap().isNull())

            dialog.address.setText("ws://station.example.org")
            self.assertTrue(dialog.code.pixmap().isNull())
            self.assertIn("wss://", dialog.code.text())
            dialog.address.setText("station.example.org")
            self.assertFalse(dialog.code.pixmap().isNull())
            dialog.address.editingFinished.emit()
            self.assertEqual(load_station_address(settings), "wss://station.example.org/rapid/v1")

            dialog.copy_key.click()
            self.assertEqual(QtWidgets.QApplication.clipboard().text(), bridge.token)

            window._start_companion_sharing()
            self.assertIs(window._companion_dialog, dialog, "Reopening shows the existing dialog")

        window._queue_active = True
        window.set_flow_state("running")
        window.set_sample("DEMO-01")
        window.set_step("AF 20 mT")
        dialog.control.setChecked(True)
        self.assertTrue(window._companion_control_action.isChecked())
        self.assertTrue(bridge.control_enabled)
        self.bridge = bridge
        client, messages = self._authenticate()
        self.assertEqual(bridge.client_count, 1)
        dialog._refresh_status()
        self.assertIn("1 device connected; remote control allowed", dialog.status.text())
        initial = messages[-1]["snapshot"]
        self.assertEqual(initial["mode"], "simulation")
        self.assertEqual(initial["sample"], "DEMO-01")
        self.assertEqual(initial["step"], "AF 20 mT")
        self.assertEqual(initial["allowed_commands"], ["pause", "halt"],
                         "An open pairing dialog must not block remote commands")
        response = self._command(client, messages, "pause", initial["revision"])
        self.assertTrue(response["accepted"])
        self.assertTrue(window._queue_paused)
        self.assertEqual(window._dashboard._run_flow.text(), "Paused")

        window._companion_control_action.setChecked(False)
        self.assertFalse(dialog.control.isChecked())

        window._queue_active = False
        dialog.stop.click()
        self.assertIsNone(window._companion)
        self.assertIsNone(window._companion_dialog)
        self.assertFalse(dialog.isVisible())
        self.assertFalse(window._companion_control_action.isEnabled())
        self._until(lambda: client.state() == QtNetwork.QAbstractSocket.SocketState.UnconnectedState)


class PairingLinkTests(unittest.TestCase):
    def test_station_addresses_follow_the_mobile_app_rules(self):
        for text, expected in [
            ("station.example.org", "wss://station.example.org/rapid/v1"),
            ("  wss://Station.Example.org/  ", "wss://station.example.org/rapid/v1"),
            ("WSS://station.example.org:8443/rapid/v1", "wss://station.example.org:8443/rapid/v1"),
            ("wss://[fd00::5]:9443", "wss://[fd00::5]:9443/rapid/v1"),
        ]:
            self.assertEqual(normalize_endpoint(text), expected, text)
        for text in [
            "", "ws://station.example.org", "https://station.example.org/rapid/v1", "wss:///rapid/v1",
            "wss://user:secret@station.example.org", "wss://station.example.org/rapid/v1?key=1",
            "wss://station.example.org/rapid/v1#frag", "wss://station.example.org/other",
            "wss://station.example.org:0", "wss://station.example.org:99999",
        ]:
            with self.assertRaises(ValueError, msg=text):
                normalize_endpoint(text)

    def test_pairing_link_carries_version_address_and_key(self):
        link = pairing_link("station.example.org:8443", "k3y-_Value")
        parts = urlsplit(link)
        self.assertEqual((parts.scheme, parts.netloc, parts.path), ("rapid-companion", "pair", ""))
        self.assertEqual(parse_qs(parts.query, strict_parsing=True), {
            "v": ["1"], "url": ["wss://station.example.org:8443/rapid/v1"], "key": ["k3y-_Value"],
        })

    def test_station_address_round_trips_and_ignores_bad_files(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        folder = Path(temporary.name)
        path = folder / "companion.json"
        self.assertEqual(load_station_address(path), "")
        save_station_address(path, "wss://station.example.org/rapid/v1")
        self.assertEqual(load_station_address(path), "wss://station.example.org/rapid/v1")
        self.assertEqual(list(folder.iterdir()), [path])
        path.write_text('{"station_address": "ws://insecure"}', encoding="utf-8")
        self.assertEqual(load_station_address(path), "")
        path.write_text("[]", encoding="utf-8")
        self.assertEqual(load_station_address(path), "")
