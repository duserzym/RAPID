"""RapidPy DataViewer — lightweight palaeomagnetic data viewer.

A self-contained RapidPy utility for visualising demagnetisation data:
  - 2D component-style directional view
  - Equal-area stereonet (Lambert equal-area projection)
  - Intensity and paleointensity quicklooks

Run from repo root:
    python RapidPy/data_viewer/main.py
"""
import sys
from pathlib import Path

# Allow imports from this package root and the broader RapidPy tree.
PACKAGE_ROOT = Path(__file__).resolve().parent
RAPIDPY_ROOT = PACKAGE_ROOT.parent
sys.path.insert(0, str(PACKAGE_ROOT))
sys.path.insert(0, str(RAPIDPY_ROOT))

from data_viewer.app import main

if __name__ == "__main__":
    raise SystemExit(main())
