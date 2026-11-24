"""Record real RAPID Companion protocol messages for the RAPID Mobile contract tests.

Drives the desktop ``CompanionBridge`` over a loopback WebSocket with a scripted
workflow and writes each server message verbatim, one JSON file per case, so the
iPhone/iPad decoder is tested against exactly what the desktop sends.

    .venv\\Scripts\\python.exe tools\\export_companion_fixtures.py <output-dir>
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "RapidPy" / "rapid_main"))
sys.path.insert(0, str(ROOT / "RapidPy"))

from PySide6 import QtCore, QtNetwork, QtTest, QtWebSockets  # noqa: E402

from rapid_main.companion import PATH, CompanionBridge  # noqa: E402


class _Station:
    def __init__(self) -> None:
        self.phase = "measuring"

    def read(self):
        allowed = {"measuring": ["pause", "halt"], "paused": ["resume", "halt"]}
        return 1, {
            "station": "IRM RAPID",
            "mode": "hardware",
            "phase": self.phase,
            "sample": "BC-12A",
            "step": "Step 3 of 8",
            "treatment": "AF 20 mT",
            "elapsed": "00:01:23",
            "allowed_commands": allowed.get(self.phase, []),
            "diagnostics": [
                {"name": "SQUID", "state": "connected"},
                {"name": "Vacuum", "state": "connected"},
                {"name": "AF Demag", "state": "fault"},
                {"name": "DC Motors", "state": "disconnected"},
            ],
        }

    def dispatch(self, action: str) -> None:
        self.phase = {"pause": "paused", "resume": "measuring", "halt": "returning"}[action]


def _until(predicate, timeout_ms: int = 3000) -> None:
    deadline = QtCore.QDeadlineTimer(timeout_ms)
    while not predicate():
        if deadline.hasExpired():
            raise TimeoutError("RAPID Companion bridge did not respond")
        QtTest.QTest.qWait(10)


def record(output: Path) -> list[Path]:
    app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])
    station = _Station()
    bridge = CompanionBridge(station.read, station.dispatch, app)
    bridge.start(0)
    socket = QtWebSockets.QWebSocket()
    raw: list[str] = []
    socket.textMessageReceived.connect(raw.append)
    socket.open(QtCore.QUrl(f"ws://127.0.0.1:{bridge.port}{PATH}"))
    _until(lambda: socket.state() == QtNetwork.QAbstractSocket.SocketState.ConnectedState)

    def send(**fields) -> None:
        socket.sendTextMessage(json.dumps(dict(v=1, **fields)))

    def latest(kind: str) -> dict:
        return next(json.loads(text) for text in reversed(raw) if json.loads(text)["type"] == kind)

    cases: dict[str, str] = {}
    send(type="hello", token=bridge.token)
    _until(lambda: raw)
    cases["snapshot-read-only"] = raw[-1]

    bridge.control_enabled = True
    bridge.publish()
    _until(lambda: latest("snapshot")["snapshot"]["control_enabled"])
    cases["snapshot-control"] = next(t for t in reversed(raw) if '"control_enabled":true' in t)
    revision = latest("snapshot")["snapshot"]["revision"]

    def command(action: str, at: int) -> str:
        request = str(uuid4())
        send(type="command", id=request, action=action, revision=at)
        _until(lambda: any(json.loads(t).get("id") == request for t in raw))
        return next(t for t in raw if json.loads(t).get("id") == request)

    cases["ack-accepted"] = command("pause", revision)
    cases["ack-stale-revision"] = command("halt", revision)

    socket.abort()
    bridge.stop()
    output.mkdir(parents=True, exist_ok=True)
    written = []
    for name, text in cases.items():
        path = output / f"{name}.json"
        path.write_text(text + "\n", encoding="utf-8")
        written.append(path)
    return written


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("output", type=Path, help="Directory for the recorded JSON messages")
    for path in record(parser.parse_args().output):
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
