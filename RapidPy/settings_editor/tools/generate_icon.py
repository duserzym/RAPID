from __future__ import annotations

import math
import platform
import subprocess
from pathlib import Path

from PySide6 import QtCore, QtGui


def _gear_path(center: QtCore.QPointF, inner_r: float, outer_r: float, tooth_count: int) -> QtGui.QPainterPath:
    path = QtGui.QPainterPath()
    angle_step = math.tau / tooth_count
    tooth_half = angle_step * 0.24

    for index in range(tooth_count):
        angle = -math.pi / 2 + index * angle_step
        points: list[QtCore.QPointF] = []
        points.append(QtCore.QPointF(center.x() + math.cos(angle - tooth_half * 1.45) * inner_r, center.y() + math.sin(angle - tooth_half * 1.45) * inner_r))
        points.append(QtCore.QPointF(center.x() + math.cos(angle - tooth_half) * outer_r, center.y() + math.sin(angle - tooth_half) * outer_r))
        points.append(QtCore.QPointF(center.x() + math.cos(angle + tooth_half) * outer_r, center.y() + math.sin(angle + tooth_half) * outer_r))
        points.append(QtCore.QPointF(center.x() + math.cos(angle + tooth_half * 1.45) * inner_r, center.y() + math.sin(angle + tooth_half * 1.45) * inner_r))
        if index == 0:
            path.moveTo(points[0])
        for point in points[1:]:
            path.lineTo(point)
    path.closeSubpath()

    inner_hole = QtGui.QPainterPath()
    inner_hole.addEllipse(center, inner_r * 0.42, inner_r * 0.42)
    return path.subtracted(inner_hole)


def draw_icon(output_png: Path, output_ico: Path) -> None:
    size = 1024
    image = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(image)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)

    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#435766"))
    gradient.setColorAt(1.0, QtGui.QColor("#243745"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    title_font = QtGui.QFont("Courier New", 80)
    title_font.setBold(True)
    title_font.setLetterSpacing(QtGui.QFont.PercentageSpacing, 108)
    painter.setFont(title_font)
    painter.setPen(QtGui.QColor("#f4ede0"))
    painter.drawText(QtCore.QRectF(0, 94, size, 104), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "INI")

    # Sheet behind the gear.
    sheet = QtCore.QRectF(210, 238, 424, 540)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(sheet, 58, 58)
    inner_sheet = sheet.adjusted(24, 24, -24, -24)
    painter.setBrush(QtGui.QColor("#f7f1e7"))
    painter.drawRoundedRect(inner_sheet, 42, 42)

    line_pen = QtGui.QPen(QtGui.QColor("#c0b39f"), 14)
    line_pen.setCapStyle(QtCore.Qt.RoundCap)
    painter.setPen(line_pen)
    for y_pos in (352, 434, 516, 598):
        painter.drawLine(318, y_pos, 560, y_pos)

    # Classic macOS-style gear in matching RapidPy palette.
    center = QtCore.QPointF(686, 618)
    gear = _gear_path(center, 122, 178, 10)
    metal_gradient = QtGui.QLinearGradient(center.x() - 120, center.y() - 120, center.x() + 120, center.y() + 120)
    metal_gradient.setColorAt(0.0, QtGui.QColor("#efede8"))
    metal_gradient.setColorAt(0.5, QtGui.QColor("#d8d7d2"))
    metal_gradient.setColorAt(1.0, QtGui.QColor("#c4c6ca"))
    painter.setPen(QtGui.QPen(QtGui.QColor("#6a7380"), 9))
    painter.setBrush(metal_gradient)
    painter.drawPath(gear)

    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QColor("#7a0219"))
    painter.drawEllipse(center, 44, 44)
    painter.setBrush(QtGui.QColor("#ffca3a"))
    painter.drawEllipse(center, 17, 17)

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

    iconset_dir = output_icns.parent / "settings_editor.iconset"
    iconset_dir.mkdir(parents=True, exist_ok=True)
    for icon_size in (16, 32, 64, 128, 256, 512):
        subprocess.run(
            ["sips", "-z", str(icon_size), str(icon_size), str(output_png), "--out", str(iconset_dir / f"icon_{icon_size}x{icon_size}.png")],
            check=True,
        )
        subprocess.run(
            ["sips", "-z", str(icon_size * 2), str(icon_size * 2), str(output_png), "--out", str(iconset_dir / f"icon_{icon_size}x{icon_size}@2x.png")],
            check=True,
        )

    subprocess.run(["iconutil", "-c", "icns", str(iconset_dir), "-o", str(output_icns)], check=True)
    return output_icns.exists()


def main() -> int:
    app = QtGui.QGuiApplication([])
    root = Path(__file__).resolve().parent.parent
    output_png = root / "assets" / "settings_editor_icon.png"
    output_ico = root / "assets" / "settings_editor_icon.ico"
    output_icns = root / "assets" / "settings_editor_icon.icns"

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
