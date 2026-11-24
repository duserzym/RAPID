"""Loopback WebSocket bridge for a private TLS proxy and RAPID Companion."""

from __future__ import annotations

import json
import logging
import secrets
import time
from collections.abc import Callable
from functools import partial
from uuid import UUID

from PySide6 import QtCore, QtNetwork, QtWebSockets

PORT = 8765
PATH = "/rapid/v1"
VERSION = 1
_MAX_CLIENTS = 8
_MAX_MESSAGE = 4096
_log = logging.getLogger(__name__)


class CompanionBridge(QtCore.QObject):
    """Run on the desktop GUI thread so snapshots and commands agree with the UI."""

    def __init__(
        self,
        read_state: Callable[[], tuple[int, dict]],
        dispatch: Callable[[str], None],
        parent: QtCore.QObject,
    ):
        super().__init__(parent)
        self.token = secrets.token_urlsafe(24)
        self.control_enabled = False
        self._read_state = read_state
        self._dispatch = dispatch
        self._revision = 0
        self._command_context = None
        self._dispatching = False
        self._peers: dict[QtWebSockets.QWebSocket, float | None] = {}
        self._server = QtWebSockets.QWebSocketServer(
            "RAPID Companion", QtWebSockets.QWebSocketServer.SslMode.NonSecureMode, self
        )
        self._server.setMaxPendingConnections(_MAX_CLIENTS)
        self._server.newConnection.connect(self._accept)
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self.publish)

    def start(self, port: int = PORT) -> None:
        if not self._server.listen(QtNetwork.QHostAddress("127.0.0.1"), port):
            raise OSError(self._server.errorString())
        self._timer.start()

    @property
    def port(self) -> int:
        return self._server.serverPort()

    def stop(self) -> None:
        self._timer.stop()
        self._server.close()
        for peer in tuple(self._peers):
            self._discard(peer)

    def _discard(self, peer: QtWebSockets.QWebSocket) -> None:
        if peer in self._peers:
            del self._peers[peer]
            peer.abort()
            peer.deleteLater()

    def _accept(self) -> None:
        while self._server.hasPendingConnections():
            peer = self._server.nextPendingConnection()
            if len(self._peers) >= _MAX_CLIENTS or peer.requestUrl().path() != PATH:
                peer.abort()
                peer.deleteLater()
                continue
            self._peers[peer] = time.monotonic() + 5
            peer.setMaxAllowedIncomingFrameSize(_MAX_MESSAGE)
            peer.setMaxAllowedIncomingMessageSize(_MAX_MESSAGE)
            peer.textMessageReceived.connect(partial(self._receive, peer))
            peer.binaryMessageReceived.connect(lambda _, socket=peer: self._discard(socket))
            peer.disconnected.connect(partial(self._discard, peer))

    def _snapshot(self) -> dict:
        run_epoch, state = self._read_state()
        context = {
            key: value for key, value in state.items() if key not in ("elapsed", "diagnostics")
        }
        context["control_enabled"] = self.control_enabled
        if self._command_context != (run_epoch, context):
            self._command_context = (run_epoch, context)
            self._revision += 1
        allowed = state["allowed_commands"] if self.control_enabled and not self._dispatching else []
        return dict(
            state, revision=self._revision, control_enabled=self.control_enabled, allowed_commands=allowed
        )

    @staticmethod
    def _encode(message_type: str, **fields) -> str:
        return json.dumps(dict(v=VERSION, type=message_type, **fields), separators=(",", ":"))

    def _send(self, peer: QtWebSockets.QWebSocket, payload: str) -> None:
        if peer in self._peers:
            if peer.bytesToWrite() > 65536 or peer.sendTextMessage(payload) <= 0:
                self._discard(peer)

    def publish(self) -> None:
        now = time.monotonic()
        clients = []
        for peer, deadline in tuple(self._peers.items()):
            if deadline is None:
                clients.append(peer)
            elif now >= deadline:
                self._discard(peer)
        if clients:
            payload = self._encode("snapshot", snapshot=self._snapshot())
            for peer in clients:
                self._send(peer, payload)

    def _receive(self, peer: QtWebSockets.QWebSocket, text: str) -> None:
        if peer not in self._peers:
            return
        try:
            message = json.loads(text)
            if not isinstance(message, dict) or type(message.get("v")) is not int or message["v"] != VERSION:
                raise ValueError("Incompatible protocol")
            deadline = self._peers[peer]
            if deadline is not None:
                token = message.get("token")
                if (time.monotonic() >= deadline or message.get("type") != "hello"
                        or not isinstance(token, str) or not token.isascii()
                        or not secrets.compare_digest(self.token, token)):
                    raise ValueError("Unauthorized")
                self._peers[peer] = None
                self._send(peer, self._encode("snapshot", snapshot=self._snapshot()))
                return
            if message.get("type") != "command" or type(message.get("revision")) is not int:
                raise ValueError("Invalid command")
            request_id = str(UUID(str(message.get("id"))))
        except (ValueError, RecursionError):
            self._discard(peer)
            return
        self._command(peer, request_id, message.get("action"), message["revision"])

    def _command(self, peer, request_id: str, action: object, revision: int) -> None:
        snapshot = self._snapshot()
        accepted = False
        if revision != snapshot["revision"]:
            detail = "Workflow changed. Review current status and try again."
        elif action not in snapshot["allowed_commands"]:
            detail = "This action is unavailable at the station."
        else:
            self._revision += 1
            self._dispatching = True
            try:
                self._dispatch(action)
                accepted = True
                detail = f"{action.capitalize()} requested."
            except Exception:
                _log.exception("RAPID mobile workflow request failed")
                detail = "Workflow request failed. Check the desktop before trying again."
            finally:
                self._dispatching = False
            snapshot = self._snapshot()
        self._send(peer, self._encode("ack", id=request_id, accepted=accepted, message=detail, snapshot=snapshot))
        self.publish()
