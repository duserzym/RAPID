from __future__ import annotations

import shutil
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]

ICON_SOURCES = {
    "gaussmeter-control-icon.png": REPO_ROOT / "RapidPy" / "gaussmeter_control" / "assets" / "gaussmeter_icon.png",
    "vrm-decay-logger-icon.png": REPO_ROOT / "RapidPy" / "vrm_logger" / "assets" / "vrm_icon.png",
    "adwin-comms-icon.png": REPO_ROOT / "RapidPy" / "adwin_comms" / "assets" / "adwin_icon.png",
    "com-port-mapper-icon.png": REPO_ROOT / "RapidPy" / "com_port_mapper" / "assets" / "com_port_mapper_icon.png",
    "af-tuner-icon.png": REPO_ROOT / "RapidPy" / "af_tuner" / "assets" / "af_tuner_icon.png",
    "changer-xy-control-icon.png": REPO_ROOT / "RapidPy" / "changer_xy_control" / "assets" / "changer_xy_control_icon.png",
    "updown-control-icon.png": REPO_ROOT / "RapidPy" / "updown_control" / "assets" / "updown_control_icon.png",
}


def copy_icons(target_root: Path) -> list[str]:
    destination = target_root / "app-icons"
    destination.mkdir(parents=True, exist_ok=True)
    copied: list[str] = []
    for name, source in ICON_SOURCES.items():
        if not source.exists():
            raise FileNotFoundError(f"Missing icon asset: {source}")
        shutil.copy2(source, destination / name)
        copied.append(name)
    return copied


def main() -> int:
    copied: list[str] = []
    for root in (REPO_ROOT / "docs" / "images", REPO_ROOT / "docs" / "site" / "images"):
        copied.extend(copy_icons(root))
    print("COPIED=" + ",".join(sorted(set(copied))))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())