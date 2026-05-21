from __future__ import annotations

import math
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

REPO_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = REPO_ROOT / "RapidPy"
for candidate in (REPO_ROOT, RAPIDPY_ROOT):
    text = str(candidate)
    if text not in sys.path:
        sys.path.insert(0, text)

from PySide6 import QtGui, QtWidgets
from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon
from RapidPy.updown_control.updown_control.app import MainWindow, ScanPoint, ScanResult


def save_pixmap_to_targets(pixmap: QtGui.QPixmap, name: str) -> None:
    for folder in (REPO_ROOT / "docs" / "images", REPO_ROOT / "docs" / "site" / "images"):
        folder.mkdir(parents=True, exist_ok=True)
        pixmap.save(str(folder / name))


def enclosing_card(widget: QtWidgets.QWidget) -> QtWidgets.QWidget | None:
    current = widget
    while current is not None and current.objectName() != "card":
        current = current.parentWidget()
    return current


def main() -> int:
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    apply_liquid_glass_theme(app)
    font_family = "Segoe UI"
    for font_path in (Path(r"C:\Windows\Fonts\segoeui.ttf"), Path(r"C:\Windows\Fonts\arial.ttf")):
        if font_path.exists():
            font_id = QtGui.QFontDatabase.addApplicationFont(str(font_path))
            if font_id >= 0:
                families = QtGui.QFontDatabase.applicationFontFamilies(font_id)
                if families:
                    font_family = families[0]
                    break
    app.setFont(QtGui.QFont(font_family, 10))

    assets_dir = REPO_ROOT / "RapidPy" / "updown_control" / "assets"
    set_app_icon(app, "updown_control_icon.png", assets_dir)

    window = MainWindow()
    window.show()
    QtWidgets.QApplication.processEvents()

    counts_per_cm = window._counts_per_cm() or 961
    assumed_target_raw = window._assumed_measurement_target_raw() or window.meas_pos_spin.value()
    current_cm = window._raw_to_cm(assumed_target_raw)

    window._baseline_raw = (0.0, 0.0, 0.0)
    window.baseline_label.setText("Baseline captured: X +0.000e+00, Y +0.000e+00, Z +0.000e+00 emu")
    window.current_position_pill.setText(f"Z {assumed_target_raw:,}")
    window.live_raw_label.setText(f"Raw {assumed_target_raw:,}")
    if current_cm is not None:
        window.live_cm_label.setText(f"Z {current_cm:+.3f} cm")
    window.top_switch_pill.setText("Z TOP OFF")
    window._last_live_raw = assumed_target_raw

    points: list[ScanPoint] = []
    z_samples = [-1.8, -1.2, -0.8, -0.4, 0.0, 0.2, 0.45, 0.8, 1.2, 1.7]
    for index, z_cm in enumerate(z_samples, start=1):
        raw = assumed_target_raw + int(round(z_cm * counts_per_cm))
        moment = 2.6e-6 * math.exp(-((z_cm - 0.22) ** 2) / 0.38)
        points.append(
            ScanPoint(
                index=index,
                raw_position=raw,
                z_cm=z_cm,
                x_emu=moment * 0.15,
                y_emu=moment * -0.08,
                z_emu=moment * 0.92,
                moment_emu=moment,
            )
        )

    suggested_z_cm = 0.22
    suggested_target_raw = assumed_target_raw + int(round(suggested_z_cm * counts_per_cm))
    suggested_meas_pos_raw = suggested_target_raw - int(round(window.sample_height_spin.value() * counts_per_cm / 2.0))
    window._handle_scan_complete(
        ScanResult(
            points=points,
            suggested_z_cm=suggested_z_cm,
            suggested_target_raw=suggested_target_raw,
            suggested_meas_pos_raw=suggested_meas_pos_raw,
            fit_method="quadratic",
        )
    )
    window._refresh_profile_model(assumed_target_raw)
    window.profile_scene._select_target("band:Measurement level")
    window.profile_scene.update()
    QtWidgets.QApplication.processEvents()

    save_pixmap_to_targets(window.grab(), "updown-control-overview.png")
    save_pixmap_to_targets(window.profile_scene.grab(), "updown-control-profile.png")

    scan_card = enclosing_card(window.scan_start_btn)
    if scan_card is not None:
        save_pixmap_to_targets(scan_card.grab(), "updown-control-scan-panel.png")

    print("CAPTURED=updown-control-overview.png,updown-control-profile.png,updown-control-scan-panel.png")
    window.close()
    QtWidgets.QApplication.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
