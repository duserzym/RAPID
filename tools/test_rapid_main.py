"""Run the main-app suite without changing operator settings or registry state."""
from pathlib import Path
import argparse
import os
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
for module in ("rapid_main", "updown_control", "vrm_logger", "af_tuner", "af_clip_test", "adwin_comms"):
    sys.path.insert(0, str(ROOT / "RapidPy" / module))
sys.path.insert(0, str(ROOT / "RapidPy"))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pattern', default='test*.py', help='Unittest discovery filename pattern')
    args = parser.parse_args()
    from PySide6 import QtCore
    with tempfile.TemporaryDirectory(prefix="rapidpy-tests-") as directory:
        os.environ["RAPID_CONFIG"] = str(Path(directory) / "config.json")
        os.environ["RAPID_SAFETY_STATE"] = str(Path(directory) / "hardware_safety.json")
        QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
        QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.UserScope, directory)
        original_settings = QtCore.QSettings
        class IsolatedSettings(original_settings):
            def __init__(self, *args, **kwargs):
                if len(args) == 2 and all(isinstance(arg, str) for arg in args):
                    super().__init__(original_settings.IniFormat, original_settings.UserScope, *args, **kwargs)
                else:
                    super().__init__(*args, **kwargs)
        QtCore.QSettings = IsolatedSettings
        suite = unittest.defaultTestLoader.discover(str(ROOT / "RapidPy" / "rapid_main" / "tests"), pattern=args.pattern)
        return 0 if unittest.TextTestRunner(verbosity=1).run(suite).wasSuccessful() else 1

if __name__ == "__main__":
    raise SystemExit(main())
