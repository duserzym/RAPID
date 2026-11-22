"""Offline release check with temporary configuration and Qt preferences."""
from __future__ import annotations

from dataclasses import replace
from importlib import import_module
import json
import os
from pathlib import Path
import struct
import tempfile


def _check_transverse_probe(window, app):
    from PySide6 import QtCore, QtGui

    actions = [action for action in window.menuBar().findChildren(QtGui.QAction)
               if action.text() == "Transverse Probe Scan"]
    if len(actions) != 1:
        return False
    actions[0].trigger()
    dialog = window._owned_dialogs.get("transverse_probe")
    if dialog is None:
        return False

    def wait_until(predicate):
        timer = QtCore.QElapsedTimer()
        timer.start()
        while not predicate() and timer.elapsed() < 5000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        return predicate()

    try:
        dialog.start_button.click()
        dialog.pause_button.click()
        if not wait_until(lambda: dialog.message.text() == "Paused between readings."):
            return False
        dialog.pause_button.click()
        if not wait_until(lambda: not dialog.active) or not dialog.apply_button.isEnabled() or dialog.table.rowCount() < 3:
            return False
        dialog.apply_button.click()
        return wait_until(lambda: not dialog.active) and dialog.message.text().startswith("Simulated angle")
    finally:
        dialog.stop()
        wait_until(lambda: not dialog.active)
        dialog.close()
        app.processEvents()


def _write_webcam_fixture(path):
    """Build a short AVI clip with a known RGB color and no codec dependencies."""
    def chunk(tag, data):
        return tag + struct.pack("<I", len(data)) + data + (b"\x00" if len(data) % 2 else b"")

    width = height = 16
    fps, frame_count = 4, 12
    pixels = bytes((38, 100, 220)) * (width * height)
    frame_size = len(pixels)
    main = struct.pack("<14I", 1_000_000 // fps, frame_size * fps, 0, 0, frame_count,
                       0, 1, frame_size, width, height, 0, 0, 0, 0)
    stream = struct.pack("<4s4sIHHIIIIIIIIhhhh", b"vids", b"DIB ", 0, 0, 0,
                         0, 1, fps, 0, frame_count, frame_size, 0xFFFFFFFF, 0, 0, 0, width, height)
    bitmap = struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, frame_size, 0, 0, 0, 0)
    headers = chunk(b"LIST", b"hdrl" + chunk(b"avih", main)
                    + chunk(b"LIST", b"strl" + chunk(b"strh", stream) + chunk(b"strf", bitmap)))
    frames = chunk(b"LIST", b"movi" + chunk(b"00db", pixels) * frame_count)
    path.write_bytes(chunk(b"RIFF", b"AVI " + headers + frames))


def _check_webcam_decode(window, app, root):
    from PySide6 import QtCore, QtGui

    clip = root / "offline-camera.avi"
    _write_webcam_fixture(clip)
    window._launch_webcam()
    dialog = window._webcam_dlg
    widget = getattr(dialog, "_webcam", None)
    if widget is None:
        return False
    widget._snapshot_dir = root / "offline-camera-snapshots"
    try:
        widget._open_stream(QtCore.QUrl.fromLocalFile(str(clip)))
        timer = QtCore.QElapsedTimer()
        timer.start()
        while widget._sink is not None and not widget._connected and timer.elapsed() < 5000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        if not widget._connected:
            return False
        widget._snapshot_btn.click()
        snapshots = list(widget._snapshot_dir.glob("*.png"))
        if len(snapshots) != 1:
            return False
        snapshot = QtGui.QImage(str(snapshots[0]))
        if snapshot.size() != QtCore.QSize(16, 16):
            return False
        color = snapshot.pixelColor(8, 8)
        return all(abs(actual - expected) <= 8 for actual, expected in
                   zip((color.red(), color.green(), color.blue()), (220, 100, 38), strict=True))
    finally:
        widget.disconnect_camera()
        dialog.close()
        app.processEvents()


def _check_measurement_archive(root):
    from data_viewer.data_loading import load_magic_directory
    from rapid_main.data_model import MeasurementStep, SpecimenMeta
    from rapid_main.io.measurement_bundle import MeasurementBundleWriter

    specimen = SpecimenMeta("offline-specimen")
    step = MeasurementStep(demag_label="NRM", gdec=10.0, ginc=30.0, sdec=10.0, sinc=30.0,
                           moment=1e-6, error_angle=2.0, crdec=10.0, crinc=30.0,
                           sdx=0.0, sdy=0.0, sdz=0.0)
    for measurement in (step, replace(step, demag_label="AF20")):
        with MeasurementBundleWriter(root / "data", specimen, simulated=True) as writer:
            writer.append_step(measurement)
    output = writer.publish_dir
    provenance = json.loads((output / "provenance.json").read_text(encoding="utf-8"))
    archive = output / provenance["previous_publication"]["relative_path"]
    previous = load_magic_directory(archive.parent)
    current = load_magic_directory(output)
    return (previous.source_kind == "measurement_archive"
            and [item.demag_label for item in previous.specimens[0].steps] == ["NRM"]
            and [item.demag_label for item in current.specimens[0].steps] == ["AF20"])


def _check_measurement_backup(app, root):
    from PySide6 import QtCore
    from data_viewer.data_loading import load_magic_directory
    from rapid_main.backup_tasks import MeasurementBackups
    from rapid_main.data_model import SpecimenMeta
    from rapid_main.hardware_contracts import NoCommBackend
    from rapid_main.io.bundle_history import _archive_io_path
    from rapid_main.measurement_worker import MeasurementWorker

    with tempfile.TemporaryDirectory(prefix="backup-", dir=_archive_io_path(root)) as directory:
        parent = Path(directory)
        destination = parent / "target"
        destination.mkdir()
        specimen = SpecimenMeta("backup-specimen")
        worker = MeasurementWorker(specimen, ["NRM"], parent / "source", NoCommBackend())
        worker.run()
        if worker._pending_run_result is not False:
            return False
        backups = MeasurementBackups()
        statuses = []
        backups.status_changed.connect(lambda message, warning: statuses.append((message, warning)))
        backups.submit(specimen.name, worker._publish_dir, str(destination))
        timer = QtCore.QElapsedTimer()
        timer.start()
        while len(statuses) < 2 and timer.elapsed() < 15_000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        if len(statuses) < 2 or statuses[-1][1]:
            return False
        archives = list((destination / "RapidPy backups" / "SIMULATED").iterdir())
        return len(archives) == 1 and load_magic_directory(archives[0]).specimens[0].steps[0].demag_label == "NRM"


def _check_sample_index_registry(window, root):
    from rapid_main.dialogs.sample_select import SampleSelectDialog

    index = root / "offline.sam"
    index.write_text("offline-specimen\n", encoding="ascii")
    first = SampleSelectDialog(window)
    second = None
    try:
        if not first._open_source(index):
            return False
        first._treatment.setText("SUSC")
        first._direction.setCurrentText("Up then Down")
        first._blocks.setValue(3)
        first._register_btn.click()
        second = SampleSelectDialog(window)
        entry = second._indexes.findData(str(index.resolve()))
        if entry < 1:
            return False
        second._indexes.setCurrentIndex(entry)
        second._open_registration(entry)
        second._table.selectRow(0)
        defaults = second.queue_defaults
        return (second.selected_sample == "offline-specimen" and defaults.treatment == "SUSC"
                and defaults.do_up and defaults.do_both and defaults.avg_steps == 3)
    finally:
        if second is not None:
            second.deleteLater()
        first.deleteLater()


def run_smoke_test() -> int:
    from PySide6 import QtCore, QtWidgets
    from rapid_main.config import AppConfig
    from rapid_main.package_launch import HELPER_MODULES

    with tempfile.TemporaryDirectory(prefix="rapidpy-smoke-") as directory:
        root = Path(directory)
        previous_config = os.environ.get("RAPID_CONFIG")
        os.environ["RAPID_CONFIG"] = str(root / "config.json")
        config = AppConfig()
        config.general.nocomm = True
        config.general.data_dir = str(root / "data")
        config.save()
        previous_format = QtCore.QSettings.defaultFormat()
        QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
        QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.UserScope, str(root))
        original_settings = QtCore.QSettings
        class IsolatedSettings(original_settings):
            def __init__(self, *args, **kwargs):
                if len(args) == 2 and all(isinstance(arg, str) for arg in args):
                    super().__init__(original_settings.IniFormat, original_settings.UserScope, *args, **kwargs)
                else:
                    super().__init__(*args, **kwargs)
        QtCore.QSettings = IsolatedSettings
        window = None
        try:
            # Constructing this window uses only explicit No-Communication backends.
            from rapid_main.app import MainWindow
            app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
            app.setQuitOnLastWindowClosed(False)
            window = MainWindow()
            window.show()
            app.processEvents()
            helpers = {}
            for name in HELPER_MODULES:
                helpers[name] = callable(import_module(f"{name}.app").main)
            transverse_probe = _check_transverse_probe(window, app)
            webcam_decode = _check_webcam_decode(window, app, root)
            measurement_archive = _check_measurement_archive(root)
            measurement_backup = _check_measurement_backup(app, root)
            sample_index_registry = _check_sample_index_registry(window, root)
            report = {
                "schema": "rapidpy.release_smoke.v1", "simulated": True,
                "ok": (window.config.general.nocomm and all(helpers.values()) and transverse_probe
                       and webcam_decode and measurement_archive and measurement_backup and sample_index_registry),
                "transverse_probe": transverse_probe,
                "webcam_decode": webcam_decode,
                "measurement_archive": measurement_archive,
                "measurement_backup": measurement_backup,
                "sample_index_registry": sample_index_registry,
                "compact_display_check": True,
                "panels": window._stack.count(), "helpers": helpers,
                "width": window.width(), "height": window.height(),
            }
            print(json.dumps(report, indent=2, sort_keys=True))
            return 0 if report["ok"] else 2
        finally:
            if window is not None:
                window._confirm_shutdown(prompt=False)
                window.hide()
                window.deleteLater()
                app.processEvents()
            QtCore.QSettings.setDefaultFormat(previous_format)
            QtCore.QSettings = original_settings
            if previous_config is None:
                os.environ.pop("RAPID_CONFIG", None)
            else:
                os.environ["RAPID_CONFIG"] = previous_config
