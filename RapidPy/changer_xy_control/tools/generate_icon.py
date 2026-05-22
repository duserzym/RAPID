from __future__ import annotations

import platform
import subprocess
from pathlib import Path

from PySide6 import QtCore, QtGui


def draw_icon(output_png: Path, output_ico: Path) -> None:
    size = 1024
    image = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(image)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)

    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#163a5f"))
    gradient.setColorAt(1.0, QtGui.QColor("#0b2238"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    title_font = QtGui.QFont("Avenir Next", 128)
    title_font.setBold(True)
    title_font.setLetterSpacing(QtGui.QFont.PercentageSpacing, 104)
    painter.setFont(title_font)
    painter.setPen(QtGui.QColor("#f4ede0"))
    painter.drawText(QtCore.QRectF(0, 124, size, 132), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "XY")

    tray_rect = QtCore.QRectF(238, 286, 548, 548)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(tray_rect, 66, 66)

    inner_rect = tray_rect.adjusted(30, 30, -30, -30)
    plate_gradient = QtGui.QLinearGradient(inner_rect.topLeft(), inner_rect.bottomRight())
    plate_gradient.setColorAt(0.0, QtGui.QColor("#fbf5ea"))
    plate_gradient.setColorAt(1.0, QtGui.QColor("#e7ddd0"))
    painter.setBrush(plate_gradient)
    painter.drawRoundedRect(inner_rect, 42, 42)

    painter.setBrush(QtGui.QColor("#050505"))
    notch_rects = (
        QtCore.QRectF(inner_rect.center().x() - 104, inner_rect.top() - 10, 208, 18),
        QtCore.QRectF(inner_rect.center().x() - 104, inner_rect.bottom() - 8, 208, 18),
        QtCore.QRectF(inner_rect.left() - 10, inner_rect.center().y() - 76, 18, 152),
        QtCore.QRectF(inner_rect.right() - 8, inner_rect.center().y() - 76, 18, 152),
    )
    for rect in notch_rects:
        painter.drawRoundedRect(rect, 9, 9)

    top_y = inner_rect.top() + inner_rect.height() * 0.23
    mid_y = inner_rect.top() + inner_rect.height() * 0.50
    bottom_y = inner_rect.top() + inner_rect.height() * 0.77
    left_x = inner_rect.left() + inner_rect.width() * 0.24
    center_x = inner_rect.left() + inner_rect.width() * 0.50
    right_x = inner_rect.left() + inner_rect.width() * 0.76
    mid_left_x = inner_rect.left() + inner_rect.width() * 0.31
    mid_right_x = inner_rect.left() + inner_rect.width() * 0.69

    cup_positions = [
        (left_x, top_y),
        (center_x, top_y),
        (right_x, top_y),
        (mid_left_x, mid_y),
        (center_x, mid_y),
        (mid_right_x, mid_y),
        (left_x, bottom_y),
        (center_x, bottom_y),
        (right_x, bottom_y),
    ]
    cup_radius = 30
    for index, (x_pos, y_pos) in enumerate(cup_positions):
        if index == 4:
            painter.setBrush(QtGui.QColor("#6a4b35"))
            painter.drawEllipse(QtCore.QPointF(x_pos, y_pos), cup_radius + 6, cup_radius + 6)
            painter.setBrush(QtGui.QColor("#221914"))
            painter.drawEllipse(QtCore.QPointF(x_pos, y_pos), cup_radius - 4, cup_radius - 4)
            continue

        painter.setBrush(QtGui.QColor("#2c2c2f"))
        painter.drawEllipse(QtCore.QPointF(x_pos, y_pos), cup_radius, cup_radius)
        painter.setBrush(QtGui.QColor("#6b6b70"))
        painter.drawEllipse(QtCore.QPointF(x_pos - 8, y_pos - 8), 10, 10)

    target_x, target_y = cup_positions[7]
    painter.setPen(QtGui.QPen(QtGui.QColor("#ffb14d"), 16))
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawEllipse(QtCore.QPointF(target_x, target_y), 48, 48)
    painter.setPen(QtGui.QPen(QtGui.QColor("#7a0219"), 10))
    painter.drawLine(target_x - 22, target_y, target_x + 22, target_y)
    painter.drawLine(target_x, target_y - 22, target_x, target_y + 22)

    painter.end()

    output_png.parent.mkdir(parents=True, exist_ok=True)
    image.save(str(output_png))
    image.save(str(output_ico))


def generate_icns(output_png: Path, output_icns: Path) -> bool:
    if platform.system() != "Darwin":
        return False

    iconutil = subprocess.run(["which", "iconutil"], capture_output=True, text=True)
    sips = subprocess.run(["which", "sips"], capture_output=True, text=True)
    if iconutil.returncode != 0 or sips.returncode != 0:
        return False

    iconset_dir = output_icns.parent / "changer_xy_control.iconset"
    iconset_dir.mkdir(parents=True, exist_ok=True)
    for size in (16, 32, 64, 128, 256, 512):
        subprocess.run(
            ["sips", "-z", str(size), str(size), str(output_png), "--out", str(iconset_dir / f"icon_{size}x{size}.png")],
            check=True,
        )
        subprocess.run(
            ["sips", "-z", str(size * 2), str(size * 2), str(output_png), "--out", str(iconset_dir / f"icon_{size}x{size}@2x.png")],
            check=True,
        )

    subprocess.run(["iconutil", "-c", "icns", str(iconset_dir), "-o", str(output_icns)], check=True)
    return output_icns.exists()


def main() -> int:
    app = QtGui.QGuiApplication([])
    root = Path(__file__).resolve().parent.parent
    output_png = root / "assets" / "changer_xy_control_icon.png"
    output_ico = root / "assets" / "changer_xy_control_icon.ico"
    output_icns = root / "assets" / "changer_xy_control_icon.icns"

    draw_icon(output_png, output_ico)
    generated_icns = generate_icns(output_png, output_icns)

    print(f"Generated {output_png}")
    print(f"Generated {output_ico}")
    if generated_icns:
        print(f"Generated {output_icns}")
    else:
        print("Skipped .icns generation (non-macOS or missing iconutil/sips).")
    app.quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())