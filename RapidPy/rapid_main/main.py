from __future__ import annotations

import sys
from pathlib import Path

# Ensure rapid_main package and RapidPy namespace are importable for repo-relative runs.
package_root = Path(__file__).resolve().parent
rapidpy_root = package_root.parent
sys.path.insert(0, str(package_root))
sys.path.insert(0, str(rapidpy_root))

from rapid_main.app import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
