from __future__ import annotations

import argparse
import json
import sys
from importlib import import_module
from time import perf_counter_ns

from rapid_main.package_launch import HELPER_MODULES


def _bounded_int(minimum, maximum):
    def parse(text):
        value = int(text)
        if not minimum <= value <= maximum:
            raise argparse.ArgumentTypeError(f'expected an integer from {minimum} to {maximum}')
        return value
    return parse


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="rapid-main")
    parser.add_argument(
        "--smoke-test", action="store_true",
        help="construct the main window with isolated simulated settings, then exit",
    )
    parser.add_argument(
        "--tool", choices=HELPER_MODULES,
        help="launch a bundled instrument or review tool in its own process",
    )
    parser.add_argument(
        "--check-startup",
        action="store_true",
        help="print a read-only dependency/resource report without opening the UI",
    )
    parser.add_argument("--check-display", action="store_true", help="check the offline dialog layout and window state")
    parser.add_argument("--check-station-evidence", nargs=2, metavar=("MANIFEST", "CANDIDATE_COMMIT"),
                        help="audit a station evidence folder without opening hardware")
    parser.add_argument("--backup-measurement", nargs=2, metavar=("PUBLICATION", "BACKUP_FOLDER"),
                        help="verify and back up a completed specimen publication")
    parser.add_argument("--backup-index-sha256", help=argparse.SUPPRESS)
    parser.add_argument('--measure-performance', action='store_true',
                        help='measure startup, idle CPU and UI responsiveness under an offline queue')
    parser.add_argument('--performance-specimens', type=_bounded_int(1, 100), default=40,
                        help='number of synthetic specimens in the performance queue (default 40)')
    parser.add_argument('--performance-idle-seconds', type=_bounded_int(1, 60), default=5,
                        help='duration of the idle measurement in seconds (default 5)')
    parser.add_argument('--performance-queue-timeout-seconds', type=_bounded_int(10, 1800), default=240,
                        help='limit for the offline performance queue in seconds (default 240)')
    return parser


def main(argv: list[str] | None = None) -> int:
    entry_started_ns = perf_counter_ns()
    args = _parser().parse_args(argv)
    if args.check_station_evidence:
        from rapid_main.station_evidence import audit_station_evidence

        try:
            result = audit_station_evidence(*args.check_station_evidence)
        except (OSError, ValueError, RuntimeError) as exc:
            print(str(exc), file=sys.stderr)
            return 2
        print(json.dumps(result, indent=2))
        return 0 if result["all_cases_have_passing_evidence"] else 1
    if args.backup_measurement:
        from rapid_main.io.bundle_backup import backup_measurement

        try:
            result = backup_measurement(*args.backup_measurement, expected_index_sha256=args.backup_index_sha256)
        except (OSError, ValueError) as exc:
            print(str(exc), file=sys.stderr)
            return 2
        print(json.dumps(result, sort_keys=True))
        return 0
    if args.smoke_test:
        from rapid_main.release_smoke import run_smoke_test
        return run_smoke_test()
    if args.measure_performance:
        from rapid_main.performance_check import run_performance_check

        return run_performance_check(args.performance_specimens, args.performance_idle_seconds,
                                     args.performance_queue_timeout_seconds, entry_started_ns=entry_started_ns)
    if args.check_display:
        from rapid_main.display_check import run_display_check
        return run_display_check()
    if args.check_startup:
        from rapid_main.startup import collect_startup_environment

        report = collect_startup_environment()
        print(report.to_json())
        return 0 if report.ok else 2

    from rapid_main.startup import collect_startup_environment

    report = collect_startup_environment()
    if not report.ok:
        print(report.to_json(), file=sys.stderr)
        return 2
    if args.tool:
        entry = import_module(f"{args.tool}.app").main
        previous_argv = sys.argv
        try:
            sys.argv = [args.tool]
            result = entry()
            return int(result or 0)
        finally:
            sys.argv = previous_argv
    try:
        from rapid_main.app import main as app_main
    except ModuleNotFoundError as exc:
        print(
            "RAPID main application cannot start because a required dependency "
            f"is missing: {exc.name}. Install the berkeley-rapidpy distribution "
            "and its dependencies, then run `python -m rapid_main --check-startup`.",
            file=sys.stderr,
        )
        return 2
    return int(app_main())


def console_main() -> int:
    """Console-script adapter that preserves command-line startup diagnostics."""

    return main(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())

