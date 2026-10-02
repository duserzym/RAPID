from __future__ import annotations

import argparse
import sys


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="rapid-main")
    parser.add_argument(
        "--check-startup",
        action="store_true",
        help="print a read-only dependency/resource report without opening the UI",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
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

