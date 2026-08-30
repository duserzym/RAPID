from __future__ import annotations

import sys
from pathlib import Path

rapidpy_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(rapidpy_root))
sys.path.insert(0, str(rapidpy_root / "rapid_main"))

from sample_prep_queue_builder.app import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())
