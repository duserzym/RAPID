"""Local bench server for the aux output test.

Serves control.html and proxies its buttons to the Arduino over pyserial.

The browser never touches the serial port. Web Serial only exists in Chrome and
Edge, and is unavailable from file:// and from sandboxed frames such as an
editor preview pane -- which is most of the ways someone will actually open the
page. Doing the serial work here makes the page a plain fetch() client that
works in any browser, and it matches how the real integration will run anyway:
the host owns the port, not a browser tab.

    python hardware/arduino-aux-test/bench_server.py            # autodetect
    python hardware/arduino-aux-test/bench_server.py --port COM11
    python hardware/arduino-aux-test/bench_server.py --list

Then open http://127.0.0.1:8000/

The port is opened once at startup and held. Opening it asserts DTR, which
resets the board, so startup waits out the bootloader before talking.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import serial
from serial.tools import list_ports

HERE = Path(__file__).resolve().parent
PAGE = HERE / "control.html"

BAUD = 115200
BOOTLOADER_WAIT = 2.2       # seconds; opening the port resets the board
REPLY_TIMEOUT = 1.5
MANAGED_PINS = (7, 8, 12, 13)

STATUS_RE = re.compile(r"^\d+=[01](\s+\d+=[01])*$")


class WrongDevice(Exception):
    """The port opened, but whatever answered is not the bench board.

    Worth failing on rather than warning about: two devices can be assigned the
    same COM number, in which case this port belongs to something else on the
    bench -- on this system, a serial card wired to lab instruments. Writing
    pin commands into that is not acceptable.
    """

    def __init__(self, port_name: str, reply: str) -> None:
        self.port_name = port_name
        self.reply = reply
        super().__init__(
            f"{port_name} answered {reply!r} rather than RAPID-AUX-TEST"
            if reply else f"{port_name} did not answer *IDN?"
        )


class Board:
    """One serial connection, guarded so concurrent requests cannot interleave."""

    def __init__(self, port_name: str) -> None:
        self.port_name = port_name
        self._lock = threading.Lock()
        self._serial = serial.Serial(
            port=port_name, baudrate=BAUD, timeout=REPLY_TIMEOUT
        )
        self.log: list[str] = []
        self.idn = ""

        self._note(f"opened {port_name} at {BAUD}")
        self._note(f"waiting {BOOTLOADER_WAIT:.1f}s for the bootloader")
        time.sleep(BOOTLOADER_WAIT)
        self._serial.reset_input_buffer()

        self.idn = self.command("*IDN?")
        if not self.idn.startswith("RAPID-AUX-TEST"):
            self._serial.close()
            raise WrongDevice(port_name, self.idn)
        self.command("ALL 0")

    def _note(self, text: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        self.log.append(f"{stamp}  {text}")
        del self.log[:-400]

    def command(self, text: str) -> str:
        """Send one line, return the first reply line. Blank if nothing came back."""
        with self._lock:
            self._note(f"> {text}")
            self._serial.reset_input_buffer()
            self._serial.write((text + "\n").encode("ascii"))
            self._serial.flush()

            deadline = time.monotonic() + REPLY_TIMEOUT
            while time.monotonic() < deadline:
                raw = self._serial.readline()
                if not raw:
                    break
                reply = raw.decode("ascii", errors="replace").strip()
                if not reply:
                    continue
                # The sketch greets on reset; skip it rather than mistake it
                # for the answer to whatever we just asked.
                if reply.endswith("ready") and not text.startswith("*IDN"):
                    self._note(f"< {reply}")
                    continue
                self._note(f"< {reply}")
                return reply

            self._note("< (no reply)")
            return ""

    def read_state(self) -> dict[int, str]:
        reply = self.command("STAT?")
        state: dict[int, str] = {p: "unknown" for p in MANAGED_PINS}
        if STATUS_RE.match(reply):
            for pair in reply.split():
                pin_text, _, value = pair.partition("=")
                pin = int(pin_text)
                if pin in state:
                    state[pin] = "high" if value == "1" else "low"
        return state

    def close(self) -> None:
        try:
            self.command("ALL 0")
        finally:
            self._serial.close()


def find_arduino() -> str | None:
    """Prefer something that names itself an Arduino, else the only USB port."""
    candidates = list(list_ports.comports())
    for info in candidates:
        haystack = f"{info.description} {info.manufacturer or ''}".lower()
        if "arduino" in haystack:
            return info.device
    usb = [c for c in candidates if c.vid is not None]
    return usb[0].device if len(usb) == 1 else None


class Handler(BaseHTTPRequestHandler):
    board: Board | None = None
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt: str, *args) -> None:  # quieter console
        pass

    # -- helpers -------------------------------------------------------

    def _send(self, code: int, body: bytes, content_type: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, payload: dict, code: int = 200) -> None:
        self._send(code, json.dumps(payload).encode("utf-8"), "application/json")

    def _snapshot(self, **extra) -> dict:
        board = Handler.board
        payload = {
            "connected": board is not None,
            "port": board.port_name if board else None,
            "idn": board.idn if board else "",
            "pins": {str(p): "unknown" for p in MANAGED_PINS},
            "log": board.log[-60:] if board else [],
        }
        if board is not None:
            payload["pins"] = {str(k): v for k, v in board.read_state().items()}
        payload.update(extra)
        return payload

    # -- routes --------------------------------------------------------

    def do_GET(self) -> None:
        route = urlparse(self.path)
        query = parse_qs(route.query)
        board = Handler.board

        if route.path in ("/", "/control.html"):
            try:
                body = PAGE.read_bytes()
            except OSError as exc:
                self._send(500, str(exc).encode(), "text/plain; charset=utf-8")
                return
            self._send(200, body, "text/html; charset=utf-8")
            return

        if route.path == "/api/state":
            self._json(self._snapshot())
            return

        if route.path == "/api/pin":
            if board is None:
                self._json({"error": "no board"}, 503)
                return
            try:
                pin = int(query.get("n", [""])[0])
                value = int(query.get("v", [""])[0])
            except ValueError:
                self._json({"error": "bad n or v"}, 400)
                return
            if pin not in MANAGED_PINS or value not in (0, 1):
                self._json({"error": f"pin must be one of {MANAGED_PINS}, v 0 or 1"}, 400)
                return
            reply = board.command(f"PIN {pin} {value}")
            self._json(self._snapshot(reply=reply))
            return

        if route.path == "/api/all_low":
            if board is None:
                self._json({"error": "no board"}, 503)
                return
            reply = board.command("ALL 0")
            self._json(self._snapshot(reply=reply))
            return

        self._send(404, b"not found", "text/plain; charset=utf-8")


def _serialcomm() -> dict[str, str]:
    r"""HKLM\HARDWARE\DEVICEMAP\SERIALCOMM -- the ports that really exist."""
    mapping: dict[str, str] = {}
    try:
        import winreg
    except ImportError:
        return mapping
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DEVICEMAP\SERIALCOMM"
        ) as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                name, value, _ = winreg.EnumValue(key, i)
                mapping[name] = str(value)
    except OSError:
        pass
    return mapping


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", help="serial port, e.g. COM11")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--http-port", type=int, default=8000)
    parser.add_argument("--list", action="store_true", help="list serial ports and exit")
    args = parser.parse_args()

    if args.list:
        for info in list_ports.comports():
            print(f"{info.device:8} {info.description}")
        return 0

    port_name = args.port or find_arduino()
    if not port_name:
        print("Could not identify the Arduino. Pass --port COMnn.", file=sys.stderr)
        print("Available ports:", file=sys.stderr)
        for info in list_ports.comports():
            print(f"  {info.device:8} {info.description}", file=sys.stderr)
        return 2

    claimants = [dev for dev, com in _serialcomm().items()
                 if com.upper() == port_name.upper()]
    if len(claimants) > 1:
        print(f"{port_name} is claimed by more than one device:", file=sys.stderr)
        for dev in claimants:
            print(f"    {dev}", file=sys.stderr)
        print(file=sys.stderr)
        print("Whichever registered first wins, so this port may not be the", file=sys.stderr)
        print("board at all. Give each device its own number in Device Manager", file=sys.stderr)
        print("before going further.", file=sys.stderr)
        return 2

    try:
        Handler.board = Board(port_name)
    except WrongDevice as exc:
        print(f"Refusing to drive {port_name}: {exc}", file=sys.stderr)
        print(file=sys.stderr)
        print("Either aux_test.ino is not loaded, or this COM number belongs to", file=sys.stderr)
        print("a different device. The device map is:", file=sys.stderr)
        print(file=sys.stderr)
        for device, com in _serialcomm().items():
            print(f"    {device:24} {com}", file=sys.stderr)
        return 2
    except serial.SerialException as exc:
        print(f"Could not open {port_name}: {exc}", file=sys.stderr)
        print(file=sys.stderr)
        if "FileNotFoundError" in str(exc) or "cannot find the file" in str(exc):
            # Device Manager can show a COM number that has no symbolic link
            # behind it -- typically because the number collides with another
            # device, or because a USB CDC port has not re-enumerated since it
            # was renumbered. Print the real device map so it is obvious which.
            print("Windows lists that port but has no device behind it.", file=sys.stderr)
            print("Check the real mapping, which is what actually gets opened:", file=sys.stderr)
            print(file=sys.stderr)
            for device, com in _serialcomm().items():
                print(f"    {device:24} {com}", file=sys.stderr)
            print(file=sys.stderr)
            print("A port missing from that list has no symbolic link. If it is a", file=sys.stderr)
            print("USB board, unplug and replug it; if two devices claim the same", file=sys.stderr)
            print("number, renumber one of them in Device Manager first.", file=sys.stderr)
        else:
            print("Close the Arduino IDE's serial monitor and try again.", file=sys.stderr)
        return 2

    print(f"Board:  {port_name}  {Handler.board.idn}")

    server = ThreadingHTTPServer((args.host, args.http_port), Handler)
    print(f"Open:   http://{args.host}:{args.http_port}/")
    print("Ctrl-C to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nstopping, driving all pins low")
    finally:
        server.server_close()
        if Handler.board:
            Handler.board.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
