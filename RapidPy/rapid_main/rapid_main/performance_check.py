"""Measure one No-Communication queue using the normal Qt workflow."""
import hashlib
import json
import math
import os
from pathlib import Path
import tempfile
import time

from rapid_main import software_version
from rapid_main.offline_profile import offline_profile


_TIMER_INTERVAL_MS = 20
_LABELS = ('NRM', 'AF10')


def timer_summary(gaps):
    """Summarize observed intervals using the nearest-rank 95th percentile."""
    if not gaps:
        return dict(samples=0, p95_interval_ms=None, max_interval_ms=None)
    ordered = sorted(gaps)
    return dict(samples=len(ordered), p95_interval_ms=round(ordered[math.ceil(len(ordered) * .95) - 1], 3),
                max_interval_ms=round(ordered[-1], 3))


def _observe(seconds, finished=None):
    from PySide6 import QtCore

    loop = QtCore.QEventLoop()
    sampler = QtCore.QTimer()
    sampler.setTimerType(QtCore.Qt.TimerType.PreciseTimer)
    sampler.setInterval(_TIMER_INTERVAL_MS)
    deadline = QtCore.QTimer()
    deadline.setTimerType(QtCore.Qt.TimerType.PreciseTimer)
    deadline.setSingleShot(True)
    gaps = []
    wall_start = previous = time.perf_counter_ns()
    wall_deadline = wall_start + seconds * 1_000_000_000
    cpu_start = time.process_time_ns()

    def tick():
        nonlocal previous
        now = time.perf_counter_ns()
        gaps.append((now - previous) / 1e6)
        previous = now
        if finished is not None and finished():
            loop.quit()

    def finish_at_deadline():
        remaining_ns = wall_deadline - time.perf_counter_ns()
        if remaining_ns > 0:
            deadline.start((remaining_ns + 999_999) // 1_000_000)
        else:
            loop.quit()

    sampler.timeout.connect(tick)
    deadline.timeout.connect(finish_at_deadline)
    sampler.start()
    deadline.start(seconds * 1000)
    loop.exec()
    wall_end = time.perf_counter_ns()
    cpu_ms = (time.process_time_ns() - cpu_start) / 1e6
    sampler.stop()
    deadline.stop()
    elapsed_ms = (wall_end - wall_start) / 1e6
    # A delay at the very end can run the deadline before the sampler.
    if (wall_end - previous) / 1e6 >= _TIMER_INTERVAL_MS:
        gaps.append((wall_end - previous) / 1e6)
    return dict(elapsed_ms=round(elapsed_ms, 3), process_cpu_ms=round(cpu_ms, 3),
                cpu_percent_of_one_core=round(100 * cpu_ms / elapsed_ms, 3),
                timer=timer_summary(gaps))


def _ready_signal():
    destination = os.environ.get('RAPID_PERFORMANCE_READY_FILE')
    if destination:
        with Path(destination).open('x', encoding='ascii') as stream:
            stream.write(str(os.getpid()))


def _verified_outputs(data, names):
    count = steps = 0
    for name in names:
        directory = data / name / 'SIMULATED'
        try:
            provenance = json.loads((directory / 'provenance.json').read_bytes())
            index = json.loads((directory / 'artifact_index.json').read_bytes())
            measurements = (directory / 'measurements.txt').read_bytes()
        except (OSError, ValueError):
            continue
        if not isinstance(provenance, dict) or not isinstance(index, dict):
            continue
        entries = index.get('artifacts')
        if not isinstance(entries, list):
            continue
        digest = hashlib.sha256(measurements).hexdigest()
        valid_files = {entry.get('relative_path'): entry.get('sha256') for entry in entries if isinstance(entry, dict)
                       and isinstance(entry.get('relative_path'), str) and entry.get('required') is True
                       and entry.get('exists') is True}
        measurement_count = len(measurements.splitlines()[2:])
        if (provenance.get('simulated') is True and provenance.get('specimen') == name
                and tuple(provenance.get('labels', ())) == _LABELS and index.get('aborted') is False
                and index.get('diagnostics_pending') is not True and (directory / 'SIMULATION.txt').is_file()
                and valid_files.get('measurements.txt') == digest and measurement_count == len(_LABELS)):
            count += 1
            steps += measurement_count
    return count, steps


def run_performance_check(specimens, idle_seconds, queue_timeout_seconds, *, entry_started_ns):
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    from PySide6 import QtCore, QtWidgets
    from rapid_main.startup import collect_startup_environment

    with tempfile.TemporaryDirectory(prefix='rapidpy-perf-') as directory, offline_profile(Path(directory)) as config:
        # The simulated vacuum otherwise relaxes toward its warning threshold
        # while we measure idle time. Keep the normal queue interlock active.
        config.vacuum.auto_pump = True
        config.save()
        startup = collect_startup_environment()
        if not startup.ok:
            print(startup.to_json())
            return 2
        from rapid_main.app import MainWindow

        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        app.setQuitOnLastWindowClosed(False)
        window = MainWindow()
        unexpected_dialogs = []
        dialog_watcher = QtCore.QTimer()

        def reject_unexpected_dialog():
            dialog = app.activeModalWidget()
            if not isinstance(dialog, QtWidgets.QDialog):
                return
            text = dialog.text() if isinstance(dialog, QtWidgets.QMessageBox) else ''
            unexpected_dialogs.append(f'{dialog.windowTitle()} — {text}'.strip(' —'))
            dialog.reject()

        dialog_watcher.timeout.connect(reject_unexpected_dialog)
        try:
            window._set_show_startup_guide(False)
            window.show()
            app.processEvents()
            vacuum = window._vacuum_backend
            if (not window.config.general.nocomm or not window.measurement_backend().simulated
                    or not vacuum.simulated or not vacuum.is_pump_on()):
                raise RuntimeError('The benchmark requires simulated measurement and vacuum with the simulated pump on.')
            ready_ms = (time.perf_counter_ns() - entry_started_ns) / 1e6
            _ready_signal()
            idle = _observe(idle_seconds)

            queue = window._sample_queue
            names = [f'Perf-{number:03d}' for number in range(1, specimens + 1)]
            for position, name in enumerate(names, 1):
                queue.add_sample(position=str(position), name=name, treatment=' -> '.join(_LABELS))
            finished = []
            window._measurement.sample_run_finished.connect(lambda aborted, name: finished.append((name, aborted)))
            dialog_watcher.start(100)
            clicked_at = time.perf_counter_ns()
            queue._run_btn.click()
            click_ms = (time.perf_counter_ns() - clicked_at) / 1e6
            active = _observe(queue_timeout_seconds, lambda: not window._has_active_automation())
            active['run_button_return_ms'] = round(click_ms, 3)
            done = not window._has_active_automation()
            reported = {name for name, aborted in finished if not aborted}
            published, steps = _verified_outputs(Path(config.general.data_dir), names)
            active.update(requested_specimens=specimens, requested_steps=specimens * len(_LABELS),
                          finished_specimens=len(reported), published_specimens=published,
                          published_steps=steps, queue_settled=done)
            ok = (done and not unexpected_dialogs and len(finished) == specimens
                  and reported == set(names) and published == specimens
                  and idle['timer']['samples'] > 0 and active['timer']['samples'] > 0)
            print(json.dumps(dict(schema='rapidpy.performance_check.v1', ok=ok, simulated=True,
                                  software_version=software_version(), qt_platform=app.platformName(),
                                  entry_to_first_window_ms=round(ready_ms, 3), qt_timer_interval_ms=_TIMER_INTERVAL_MS,
                                  idle_seconds_requested=idle_seconds, treatment_labels=_LABELS,
                                  simulated_vacuum_pump_on=config.vacuum.auto_pump,
                                  unexpected_dialogs=unexpected_dialogs, terminal_status=window._sb_status.text(),
                                  idle=idle, queue=active), indent=2, sort_keys=True), flush=True)
            return 0 if ok else 2
        finally:
            if window._has_active_automation():
                window.cancel_queue_run('Performance check ended.')
                _observe(10, lambda: not window._has_active_automation())
            dialog_watcher.stop()
            window._confirm_shutdown(prompt=False)
            window.hide()
            window.deleteLater()
            app.processEvents()
