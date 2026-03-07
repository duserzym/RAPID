#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
python tools/generate_icon.py

ICON_ARG=()
if [[ -f "assets/vrm_icon.icns" ]]; then
  ICON_ARG=(--icon "assets/vrm_icon.icns")
fi

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name RapidPyVRM \
  --onefile \
  "${ICON_ARG[@]}" \
  main.py

echo "Build complete: dist/RapidPyVRM.app and dist/RapidPyVRM"
