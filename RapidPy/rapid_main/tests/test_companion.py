from __future__ import annotations

import json
import unittest
from unittest.mock import patch
from uuid import uuid4

from PySide6 import QtCore, QtNetwork, QtTest, QtWebSockets, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.companion import PATH, CompanionBridge
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

    def test_desktop_menu_starts_read_only_and_uses_workflow_handlers(self):
        self.bridge.stop()
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
        dialogs = []
        def record_dialog(*args, **kwargs):
            dialogs.append((args, kwargs))
            return '', False
        start_on_port = CompanionBridge.start
        with (
            patch.object(QtWidgets.QInputDialog, "getText", record_dialog),
            patch.object(CompanionBridge, "start", lambda bridge: start_on_port(bridge, 0)),
        ):
            window._start_companion_sharing()
        bridge = window._companion
        self.assertIsNotNone(bridge)
        self.assertEqual(len(dialogs), 1)
        self.assertEqual(dialogs[0][1]["text"], bridge.token)
        self.assertFalse(bridge.control_enabled)

        window._queue_active = True
        window.set_flow_state("running")
        window.set_sample("DEMO-01")
        window.set_step("AF 20 mT")
        window._companion_control_action.setChecked(True)
        self.bridge = bridge
        client, messages = self._authenticate()
        initial = messages[-1]["snapshot"]
        self.assertEqual(initial["mode"], "simulation")
        self.assertEqual(initial["sample"], "DEMO-01")
        self.assertEqual(initial["step"], "AF 20 mT")
        response = self._command(client, messages, "pause", initial["revision"])
        self.assertTrue(response["accepted"])
        self.assertTrue(window._queue_paused)
        self.assertEqual(window._dashboard._run_flow.text(), "Paused")

        window._queue_active = False
        window._stop_companion_sharing()
        self.assertIsNone(window._companion)
        self.assertFalse(window._companion_control_action.isEnabled())
