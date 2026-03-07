#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
python tools/generate_icon.py

ICON_ARG=()
if [[ -f "assets/updown_icon.icns" ]]; then
  ICON_ARG=(--icon "assets/updown_icon.icns")
fi

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name RapidPyUpDown \
  --onefile \
  "${ICON_ARG[@]}" \
  main.py

echo "Build complete: dist/RapidPyUpDown.app and dist/RapidPyUpDown"
