"""Offline release check with temporary configuration and Qt preferences."""
from __future__ import annotations

from dataclasses import replace
from importlib import import_module
import json
import math
import os
from pathlib import Path
import struct
import sys
import tempfile

from rapid_main import software_version


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
    from data_viewer.data_loading import load_input, load_magic_directory
    from rapid_main.data_model import MeasurementStep, SpecimenMeta
    from rapid_main.io.measurement_bundle import MeasurementBundleWriter
    from rapid_main.io.rmg_writer import describe_rmg_treatment

    specimen = SpecimenMeta("offline-specimen")
    step = MeasurementStep(demag_label="NRM", gdec=10.0, ginc=30.0, sdec=10.0, sinc=30.0,
                           moment=1e-6, error_angle=2.0, crdec=10.0, crinc=30.0,
                           sdx=0.0, sdy=0.0, sdz=0.0)
    for measurement, susceptibility in ((step, 0.001), (replace(step, demag_label="AF20"), -0.002)):
        with MeasurementBundleWriter(root / "data", specimen, simulated=True) as writer:
            writer.append_step(measurement, susceptibility=susceptibility,
                               rmg_treatment=describe_rmg_treatment(measurement.demag_label))
    output = writer.publish_dir
    provenance = json.loads((output / "provenance.json").read_text(encoding="utf-8"))
    archive = output / provenance["previous_publication"]["relative_path"]
    previous = load_magic_directory(archive.parent)
    current = load_magic_directory(output)
    direct = load_input(output / specimen.name)
    return (previous.source_kind == "measurement_archive"
            and [item.demag_label for item in previous.specimens[0].steps] == ["NRM"]
            and [item.demag_label for item in current.specimens[0].steps] == ["AF20"]
            and direct.source_kind == "cit_specimen"
            and [item.demag_label for item in direct.specimens[0].steps] == ["AF20"]
            and current.specimens[0].steps[0].error_angle == direct.specimens[0].steps[0].error_angle == 2.0
            and direct.specimens[0].susceptibility_points == current.specimens[0].susceptibility_points
            and [item.susceptibility for item in previous.specimens[0].susceptibility_points] == [0.001]
            and [(item.level, item.susceptibility) for item in current.specimens[0].susceptibility_points]
            == [(200.0, -0.002)])


def _check_data_review_projection(root):
    from data_viewer.app import ZijderveldWidget
    from data_viewer.data_loading import load_cit_specimen

    source = root / "data" / "SIMULATED" / "offline-specimen"
    specimen, = load_cit_specimen(source).specimens
    plot = ZijderveldWidget()
    try:
        plot.set_specimen(specimen, "specimen", False, [], False)
        x_ns, y_ns = plot._horiz_line.getData()
        _, z_ns = plot._vert_line.getData()
        plot._projection.setCurrentIndex(plot._projection.findData("ew"))
        x_ew, y_ew = plot._horiz_line.getData()
        _, z_ew = plot._vert_line.getData()
        return (len(x_ns) == len(x_ew) == 1 and z_ns[0] < 0
                and math.isclose(x_ns[0], -y_ew[0]) and math.isclose(y_ns[0], x_ew[0])
                and math.isclose(z_ns[0], z_ew[0]))
    finally:
        plot.deleteLater()


def _check_measurement_backup(app, root):
    from PySide6 import QtCore
    from data_viewer.data_loading import load_magic_directory
    from rapid_main.backup_tasks import MeasurementBackups
    from rapid_main.data_model import SpecimenMeta
    from rapid_main.hardware_contracts import NoCommBackend
    from rapid_main.io.backup_jobs import BackupJob, BackupJobs
    from rapid_main.io.bundle_backup import INDEX_FILE
    from rapid_main.io.bundle_history import _archive_io_path
    from rapid_main.io.bundle_publication import _digest
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
        digest = _digest(worker._publish_dir / INDEX_FILE)
        queue = parent / "pending-backups"
        BackupJobs(queue).save(BackupJob(specimen.name, str(worker._publish_dir), str(destination), digest))
        backups = MeasurementBackups(spool=queue)
        statuses = []
        backups.status_changed.connect(lambda message, warning: statuses.append((message, warning)))
        backups.retry_pending(str(destination))
        timer = QtCore.QElapsedTimer()
        timer.start()
        while len(statuses) < 2 and timer.elapsed() < 15_000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        if len(statuses) < 2 or statuses[-1][1] or BackupJobs(queue).pending() != ([], []):
            return False
        archives = list((destination / "RapidPy backups" / "SIMULATED").iterdir())
        return len(archives) == 1 and load_magic_directory(archives[0]).specimens[0].steps[0].demag_label == "NRM"


def _check_measurement_checkpoint(root):
    from rapid_main.data_model import SpecimenMeta
    from rapid_main.hardware_contracts import NoCommBackend
    from rapid_main.measurement_checkpoint import pending_checkpoint
    from rapid_main.measurement_worker import MeasurementWorker
    from rapidpy_common.paths import absolute_io_path

    with tempfile.TemporaryDirectory(prefix='checkpoint-', dir=absolute_io_path(root)) as directory:
        output = Path(directory)
        plan = ['NRM', 'AF10']
        first = MeasurementWorker(SpecimenMeta('offline'), plan, output, NoCommBackend(),
                                  run_id='offline-checkpoint', checkpoint_profile='offline')
        first.request_checkpoint()
        first.run()
        saved = pending_checkpoint(first._pair_directory)
        if not first.checkpoint_saved or saved is None or (first._pair_directory / 'measurements.txt').exists():
            return False
        resumed = MeasurementWorker(SpecimenMeta('offline'), plan, output, NoCommBackend(),
                                    checkpoint_profile='offline', continue_from=saved)
        resumed.run()
        return (resumed._pending_run_result is False and resumed._restored_count == 1
                and pending_checkpoint(resumed._pair_directory) is None
                and [step['label'] for step in resumed._progress.accepted_steps()] == plan)


def _check_manual_interruption(root):
    from rapid_main.data_model import SpecimenMeta
    from rapid_main.hardware_contracts import NoCommBackend
    from rapid_main.manual_run_recovery import acknowledge_interruption, reconcile_manual_run
    from rapid_main.measurement_worker import MeasurementWorker
    from rapidpy_common.paths import absolute_io_path

    class InterruptedSimulator(NoCommBackend):
        def set_demag_step(self, label):
            raise RuntimeError('Offline interruption fixture')

    with tempfile.TemporaryDirectory(prefix='interruption-', dir=absolute_io_path(root)) as directory:
        output = Path(directory)
        meta = SpecimenMeta('offline')
        original = MeasurementWorker(meta, ['NRM'], output, InterruptedSimulator(), track_manual=True)
        original.run()
        interrupted = reconcile_manual_run(original._pair_directory, meta.name)
        if interrupted is None or 'Treatment may already have occurred' not in interrupted.describe():
            return False
        fresh = MeasurementWorker(meta, ['NRM'], output, NoCommBackend())
        fresh.run()
        if fresh._progress is not None or not fresh._error_messages:
            return False
        receipt = acknowledge_interruption(interrupted, operator='Offline fixture operator',
                                           notes='Resolved the injected offline interruption.')
        return (receipt.is_file() and reconcile_manual_run(original._pair_directory, meta.name) is None
                and not (original._pair_directory / 'measurements.txt').exists())


def _check_vrm_practice(app, root):
    from PySide6 import QtCore
    from rapid_main.vrm import VRM_CONTEXT_ENV, build_vrm_launch_context, write_vrm_launch_context
    from rapidpy_common.paths import absolute_io_path
    from vrm_logger.app import MainWindow
    from vrm_logger.session_manifest import file_sha256

    output = tempfile.TemporaryDirectory(prefix="vrm-", dir=absolute_io_path(root))
    output_dir = Path(output.name)
    context = build_vrm_launch_context(nocomm=True, intended_output_dir=output_dir)
    handoff = write_vrm_launch_context(root, context=context)
    previous_handoff = os.environ.get(VRM_CONTEXT_ENV)
    os.environ[VRM_CONTEXT_ENV] = str(handoff)
    window = None

    def wait_until(predicate):
        timer = QtCore.QElapsedTimer()
        timer.start()
        while not predicate() and timer.elapsed() < 5000:
            app.processEvents()
            QtCore.QThread.msleep(10)
        return predicate()

    try:
        window = MainWindow()
        if window.connect_btn.isEnabled() or not window.baseline_btn.isEnabled():
            return False
        window.interval_spin.setValue(0.05)
        window.baseline_btn.click()
        window.start_btn.click()
        if not wait_until(lambda: len(window._time) >= 2):
            return False
        window.stop_btn.click()
        if not wait_until(lambda: window._thread is None):
            return False
        manifests = list(absolute_io_path(output_dir / "SIMULATED").glob("*/*.vrm.json"))
        if len(manifests) != 1:
            print(window.console_output.toPlainText(), file=sys.stderr)
            return False
        manifest = json.loads(manifests[0].read_bytes())
        csv_path = manifests[0].with_name("vrm_log.csv")
        return (manifest["provenance"]["simulated"] is True and manifest["outcome"] == "operator_stopped"
                and manifest["row_count"] >= 2 and manifest["output_sha256"] == file_sha256(csv_path))
    finally:
        if window is not None:
            window.close()
            window.deleteLater()
        if previous_handoff is None:
            os.environ.pop(VRM_CONTEXT_ENV, None)
        else:
            os.environ[VRM_CONTEXT_ENV] = previous_handoff
        output.cleanup()


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
    from PySide6 import QtWidgets
    from rapid_main.offline_profile import offline_profile
    from rapid_main.package_launch import HELPER_MODULES

    with tempfile.TemporaryDirectory(prefix="rapidpy-smoke-") as directory, offline_profile(Path(directory)):
        root = Path(directory)
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
            data_review_projection = _check_data_review_projection(root)
            measurement_backup = _check_measurement_backup(app, root)
            measurement_checkpoint = _check_measurement_checkpoint(root)
            manual_interruption = _check_manual_interruption(root)
            sample_index_registry = _check_sample_index_registry(window, root)
            vrm_practice = _check_vrm_practice(app, root)
            report = {
                "schema": "rapidpy.release_smoke.v1", "simulated": True,
                "software_version": software_version(),
                "ok": (window.config.general.nocomm and all(helpers.values()) and transverse_probe
                       and webcam_decode and measurement_archive and data_review_projection
                       and measurement_backup and sample_index_registry
                       and vrm_practice and measurement_checkpoint and manual_interruption),
                "transverse_probe": transverse_probe,
                "webcam_decode": webcam_decode,
                "measurement_archive": measurement_archive,
                "data_review_projection": data_review_projection,
                "measurement_backup": measurement_backup,
                "measurement_checkpoint": measurement_checkpoint,
                "manual_interruption": manual_interruption,
                "sample_index_registry": sample_index_registry,
                "vrm_practice": vrm_practice,
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
