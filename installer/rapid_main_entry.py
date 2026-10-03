"""Frozen main app and bundled-tool entry point."""
from rapid_main.__main__ import console_main

if __name__ == "__main__":
    raise SystemExit(console_main())
