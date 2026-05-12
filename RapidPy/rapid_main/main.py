from __future__ import annotations

import sys
from pathlib import Path

# Ensure RapidPy/ is on the path so rapidpy_common and rapid_main are importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from rapid_main.app import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
