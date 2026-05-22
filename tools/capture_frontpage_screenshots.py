from __future__ import annotations

from datetime import datetime, timedelta
import os
from pathlib import Path
import sys
import time
import traceback

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

REPO_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = REPO_ROOT / "RapidPy"
for candidate in (REPO_ROOT, RAPIDPY_ROOT):
    text = str(candidate)
    if text not in sys.path:
        sys.path.insert(0, text)

from PySide6 import QtGui, QtWidgets

from rapidpy_common.gaussmeter import GaussmeterReading
from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon
from RapidPy.changer_xy_control.changer_xy_control.app import MainWindow as ChangerXYWindow
from RapidPy.com_port_mapper.com_port_mapper.app import MainWindow as ComPortMapperWindow
from RapidPy.com_port_mapper.com_port_mapper.probe import PortProbeResult
from RapidPy.gaussmeter_control.gaussmeter_control.app import MainWindow as GaussmeterWindow, SessionSample
from RapidPy.vrm_logger.vrm_logger.app import MainWindow as VrmLoggerWindow
from RapidPy.vrm_logger.vrm_logger.models import MeasurementSample


LOG_PATH = REPO_ROOT / "tools" / "capture_frontpage_screenshots.log"


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message}\n")


def save_pixmap_to_targets(pixmap: QtGui.QPixmap, name: str) -> None:
    for folder in (REPO_ROOT / "docs" / "images", REPO_ROOT / "docs" / "site" / "images"):
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / name
        if not pixmap.save(str(target)):
            raise RuntimeError(f"Failed to save screenshot: {target}")
        log(f"saved {target}")


def capture_gaussmeter(app: QtWidgets.QApplication) -> None:
    log("capture_gaussmeter:start")
    apply_liquid_glass_theme(app)
    assets_dir = REPO_ROOT / "RapidPy" / "gaussmeter_control" / "assets"
    set_app_icon(app, "gaussmeter_icon.png", assets_dir)

    window = GaussmeterWindow()
    window.dll_path_edit.setText(r"C:\Program Files\RapidPy\Gaussmeter\drivers\usb5100.dll")
    window.driver_status.setPlainText(
        "Driver ready (Published bundle):\n"
        r"C:\Program Files\RapidPy\Gaussmeter\drivers\usb5100.dll"
    )
    window._set_connected_state(True)
    window._set_reading(
        GaussmeterReading(
            value=42.315,
            raw_value=42.315,
            units_index=1,
            units_label="mT",
            base_units_label="T",
            mode_index=0,
            mode_label="DC",
            range_index=1,
            timestamp=datetime.now(),
        )
    )
    window._session_samples = [
        SessionSample(
            index=index + 1,
            elapsed_s=float(index * 5),
            captured_at=datetime.now() + timedelta(seconds=index * 5),
            reading=GaussmeterReading(
                value=value,
                raw_value=value,
                units_index=1,
                units_label="mT",
                base_units_label="T",
                mode_index=0,
                mode_label="DC",
                range_index=1,
                timestamp=datetime.now() + timedelta(seconds=index * 5),
            ),
            driver_path=r"C:\Program Files\RapidPy\Gaussmeter\drivers\usb5100.dll",
        )
        for index, value in enumerate((40.8, 41.2, 41.9, 42.4, 42.1, 42.7, 42.3, 42.9))
    ]
    window._update_session_plot()
    window._update_sampling_controls()
    window.console.setPlainText(
        "Connected using USB auto. Driver: usb5100.dll\n"
        "Captured 8-point verification session.\n"
        "Ready for operator sampling run."
    )

    window.show()
    QtWidgets.QApplication.processEvents()
    save_pixmap_to_targets(window.grab(), "gaussmeter_app.png")
    window.close()
    QtWidgets.QApplication.processEvents()
    log("capture_gaussmeter:done")


def capture_vrm_logger(app: QtWidgets.QApplication) -> None:
    log("capture_vrm_logger:start")
    window = VrmLoggerWindow()
    if window.port_combo.findText("COM10") < 0:
        window.port_combo.addItem("COM10")
    window.port_combo.setCurrentText("COM10")
    window.file_edit.setText(r"C:\RapidPy\logs\vrm_decay_2026_05_21.csv")
    window._baseline_raw = (1.2e-4, -8.0e-5, 4.0e-5)
    window._update_baseline_labels()
    window._set_status("Connected to COM10 at 1200,N,8,1")
    window.start_btn.setEnabled(True)
    window._session_start_epoch = time.time() - 180.0
    for sample in (
        MeasurementSample(0.0, 1.20e-4, -8.0e-5, 4.0e-5),
        MeasurementSample(5.0, 1.15e-4, -7.6e-5, 3.8e-5),
        MeasurementSample(15.0, 1.05e-4, -7.1e-5, 3.5e-5),
        MeasurementSample(30.0, 9.1e-5, -6.5e-5, 3.1e-5),
        MeasurementSample(60.0, 8.4e-5, -6.0e-5, 2.8e-5),
        MeasurementSample(90.0, 7.8e-5, -5.4e-5, 2.5e-5),
        MeasurementSample(120.0, 7.1e-5, -4.9e-5, 2.2e-5),
    ):
        window._handle_sample(sample)
    window._append_console("Previewed baseline-subtracted decay trace for published docs screenshot.")

    window.show()
    QtWidgets.QApplication.processEvents()
    save_pixmap_to_targets(window.grab(), "vrm_logger_app.png")
    window.close()
    QtWidgets.QApplication.processEvents()
    log("capture_vrm_logger:done")


def capture_com_port_mapper(app: QtWidgets.QApplication) -> None:
    log("capture_com_port_mapper:start")
    apply_liquid_glass_theme(app)
    assets_dir = REPO_ROOT / "RapidPy" / "com_port_mapper" / "assets"
    set_app_icon(app, "com_port_mapper_icon.png", assets_dir)

    original_start_sweep = ComPortMapperWindow.start_sweep
    ComPortMapperWindow.start_sweep = lambda self: None
    try:
        window = ComPortMapperWindow()
    finally:
        ComPortMapperWindow.start_sweep = original_start_sweep

    results = (
        PortProbeResult(
            port="COM4",
            description="Sunix PCI Serial Port",
            manufacturer="Sunix",
            hwid="PCI\\VEN_9710&DEV_9900",
            location="PCI bus 3",
            adapter_family="Enhanced / PCI serial",
            detected_device="X / changer motor",
            confidence="High",
            protocol="Quicksilver motor",
            notes="Matched controller identity and motion register handshake.",
        ),
        PortProbeResult(
            port="COM5",
            description="Sunix PCI Serial Port",
            manufacturer="Sunix",
            hwid="PCI\\VEN_9710&DEV_9900",
            location="PCI bus 3",
            adapter_family="Enhanced / PCI serial",
            detected_device="SQUID magnetometer",
            confidence="High",
            protocol="1200,N,8,1",
            notes="Matched 3-axis SQUID ASCII stream.",
        ),
        PortProbeResult(
            port="COM7",
            description="USB Serial Adapter",
            manufacturer="FTDI",
            hwid="USB\\VID_0403&PID_6001",
            location="Rear panel USB",
            adapter_family="Other serial",
            detected_device="Gaussmeter",
            confidence="Hint",
            protocol="usb5100 / gm0",
            notes="Legacy VB6 role hint only; no active RAPID protocol probe on this adapter.",
        ),
    )
    for result in results:
        window._append_result(result)
    window._set_status("Sweep complete. 2 high-confidence match(es).")
    window.capability_label.setText(
        "Enhanced/PCI probing enabled. Legacy VB6 COM-role hints remain visible only when no high-confidence match is found."
    )
    window.table.selectRow(0)
    window._update_details()

    window.show()
    QtWidgets.QApplication.processEvents()
    save_pixmap_to_targets(window.grab(), "com_port_mapper_app.png")
    window.close()
    QtWidgets.QApplication.processEvents()
    log("capture_com_port_mapper:done")


def capture_changer_xy(app: QtWidgets.QApplication) -> None:
    log("capture_changer_xy:start")
    apply_liquid_glass_theme(app)
    assets_dir = REPO_ROOT / "RapidPy" / "changer_xy_control" / "assets"
    set_app_icon(app, "changer_xy_control_icon.png", assets_dir)

    original_poll_stage_state = ChangerXYWindow._poll_stage_state
    ChangerXYWindow._poll_stage_state = lambda self: None
    try:
        window = ChangerXYWindow()
    finally:
        ChangerXYWindow._poll_stage_state = original_poll_stage_state

    window._poll_timer.stop()
    window.target_hole.setValue(12)
    window.stage_scene.set_target_hole(12)
    window.stage_scene.set_current_hole(11)
    window.stage_scene.set_current_xy(347_820, 512_440)
    window.x_pos_label.setText("347,820")
    window.y_pos_label.setText("512,440")
    window.z_pos_label.setText("18,240")
    window.current_hole_label.setText("Cup 11")
    window.position_source_label.setText("Published preview")
    window.stage_status.setText("Cup 12 is selected. Double-click a cup to move there, or use LOAD to return to the corner.")
    window.cup_table_summary.setText("100 cups loaded from the VB6 settings profile. Cup 12 is the active target.")
    window.calibration_summary.setText("100 calibrated cups loaded from the current settings profile.")
    window.console.setPlainText(
        "Loaded Paleomag_v3 INI profile.\n"
        "Previewing a published XY changer operator state with Cup 12 selected.\n"
        "XY-only mode remains available when the Z axis is disconnected."
    )
    window._sync_cup_table_selection(12)

    window.show()
    QtWidgets.QApplication.processEvents()
    save_pixmap_to_targets(window.grab(), "changer_xy_app.png")
    window.close()
    QtWidgets.QApplication.processEvents()
    log("capture_changer_xy:done")


def main() -> int:
    LOG_PATH.write_text("", encoding="utf-8")
    log("main:start")
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    try:
        capture_gaussmeter(app)
        capture_vrm_logger(app)
        capture_com_port_mapper(app)
        capture_changer_xy(app)
        log("main:success")
        print("CAPTURED=gaussmeter_app.png,vrm_logger_app.png,com_port_mapper_app.png,changer_xy_app.png")
        return 0
    except Exception as exc:
        log(f"main:error:{exc}")
        log(traceback.format_exc())
        raise
    finally:
        app.quit()


if __name__ == "__main__":
    raise SystemExit(main())