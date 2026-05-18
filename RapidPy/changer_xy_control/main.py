"""Launcher for the RapidPy changer XY control app.

See changer_xy_control.app for the detailed motion notes, including why the
changer UI keeps raw VB6-compatible velocity inputs while converting them to
estimated physical cm/s for operator feedback.
"""

from changer_xy_control.app import main


if __name__ == "__main__":
    raise SystemExit(main())
