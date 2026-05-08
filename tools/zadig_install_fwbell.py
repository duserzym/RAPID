from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

from pywinauto import Application
from pywinauto.findbestmatch import MatchError


EXPECTED_DEVICE = "FW Bell 5100"
EXPECTED_DRIVER_TOKEN = "libusb-win32"
WINDOW_TITLE = "Zadig"
INSTALL_TIMEOUT_SECONDS = 180


def wait_for_window(app: Application, title: str, timeout: float = 30.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            window = app.window(title=title)
            if window.exists(timeout=0.5):
                return window
        except Exception:
            pass
        time.sleep(0.5)
    raise RuntimeError(f"Timed out waiting for window: {title}")


def try_click_named_button(title_candidates: tuple[str, ...], button_candidates: tuple[str, ...]) -> bool:
    desktop = Application(backend="win32").connect(path="explorer.exe")
    for window in desktop.windows():
        title = window.window_text().strip()
        if not title:
            continue
        if not any(candidate.lower() in title.lower() for candidate in title_candidates):
            continue
        for button_name in button_candidates:
            try:
                button = window.child_window(title=button_name, class_name="Button")
                if button.exists(timeout=0.2):
                    button.click_input()
                    return True
            except (MatchError, Exception):
                continue
    return False


def main() -> int:
    script_dir = Path(__file__).resolve().parent
    zadig_exe = script_dir / "zadig.exe"
    if not zadig_exe.exists():
        raise FileNotFoundError(zadig_exe)

    subprocess.run(["taskkill", "/F", "/IM", "zadig.exe"], check=False, capture_output=True)
    time.sleep(1.0)

    subprocess.Popen([str(zadig_exe)], cwd=str(script_dir))
    app = Application(backend="win32").connect(path=str(zadig_exe), timeout=30)
    dialog = wait_for_window(app, WINDOW_TITLE)

    device_name = dialog.child_window(control_id=1001, class_name="ComboBox").window_text().strip()
    target_driver = dialog.child_window(control_id=1011, class_name="Edit").window_text().strip()
    current_driver = dialog.child_window(control_id=1008, class_name="Edit").window_text().strip()

    print(f"device={device_name}")
    print(f"current_driver={current_driver}")
    print(f"target_driver={target_driver}")

    if EXPECTED_DEVICE.lower() not in device_name.lower():
        raise RuntimeError(f"Unexpected device selection: {device_name!r}")
    if EXPECTED_DRIVER_TOKEN.lower() not in target_driver.lower():
        raise RuntimeError(f"Unexpected target driver: {target_driver!r}")

    dialog.child_window(control_id=1009, class_name="Button").click_input()
    print("install_clicked=true")

    deadline = time.time() + INSTALL_TIMEOUT_SECONDS
    while time.time() < deadline:
        if try_click_named_button(("certificate", "security warning", "windows security"), ("Install", "Yes", "OK", "Install this driver software anyway")):
            print("security_prompt_handled=true")

        if not dialog.exists(timeout=0.5):
            print("zadig_window_closed=true")
            return 0

        try:
            status = dialog.child_window(control_id=1006).window_text().strip()
        except Exception:
            status = ""
        if status:
            print(f"status={status}")
            if "success" in status.lower():
                return 0
            if "error" in status.lower() or "failed" in status.lower():
                raise RuntimeError(status)

        time.sleep(1.0)

    raise RuntimeError("Timed out waiting for Zadig installation to finish")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"error={exc}", file=sys.stderr)
        raise