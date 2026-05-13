"""Zijderveld Viewer — standalone palaeomagnetic demagnetisation plot viewer.

A self-contained RapidPy utility for visualising demagnetisation data:
  • Zijderveld diagram  (horizontal + vertical projections)
  • Equal-area stereonet (Lambert equal-area projection)
  • Intensity decay curve

Run from repo root:
    python RapidPy/zijderveld_viewer/main.py
"""
import sys
from pathlib import Path

# Allow imports from RapidPy/ whether installed or run in-place
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from zijderveld_viewer.app import main

if __name__ == "__main__":
    raise SystemExit(main())
