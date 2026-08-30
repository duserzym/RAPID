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
    gradient.setColorAt(0.0, QtGui.QColor("#204668"))
    gradient.setColorAt(1.0, QtGui.QColor("#112c44"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    title_font = QtGui.QFont("Avenir Next", 90)
    title_font.setBold(True)
    title_font.setLetterSpacing(QtGui.QFont.PercentageSpacing, 103)
    painter.setFont(title_font)
    painter.setPen(QtGui.QColor("#f4ede0"))
    painter.drawText(QtCore.QRectF(0, 88, size, 104), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "SAM")

    # Left tray / sample-prep card.
    tray = QtCore.QRectF(118, 238, 368, 548)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(tray, 58, 58)
    painter.setBrush(QtGui.QColor("#f7f1e7"))
    painter.drawRoundedRect(tray.adjusted(24, 24, -24, -24), 42, 42)

    cup_positions = [
        QtCore.QPointF(248, 370),
        QtCore.QPointF(386, 370),
        QtCore.QPointF(248, 522),
        QtCore.QPointF(386, 522),
        QtCore.QPointF(248, 674),
        QtCore.QPointF(386, 674),
    ]
    for point in cup_positions:
        painter.setBrush(QtGui.QColor("#2c2c2f"))
        painter.drawEllipse(point, 42, 42)
        painter.setBrush(QtGui.QColor("#6b6b70"))
        painter.drawEllipse(point + QtCore.QPointF(-10, -10), 14, 14)

    painter.setPen(QtGui.QPen(QtGui.QColor("#ffb14d"), 16))
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawEllipse(cup_positions[3], 62, 62)
    painter.setPen(QtGui.QPen(QtGui.QColor("#7a0219"), 10))
    painter.drawLine(cup_positions[3].x() - 20, cup_positions[3].y(), cup_positions[3].x() + 20, cup_positions[3].y())
    painter.drawLine(cup_positions[3].x(), cup_positions[3].y() - 20, cup_positions[3].x(), cup_positions[3].y() + 20)

    # Right queue list card.
    queue_card = QtCore.QRectF(542, 238, 364, 548)
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(queue_card, 58, 58)
    painter.setBrush(QtGui.QColor("#f7f1e7"))
    painter.drawRoundedRect(queue_card.adjusted(24, 24, -24, -24), 42, 42)

    line_pen = QtGui.QPen(QtGui.QColor("#31566d"), 16)
    line_pen.setCapStyle(QtCore.Qt.RoundCap)
    painter.setPen(line_pen)
    rows_y = [344, 442, 540, 638]
    for y_pos in rows_y:
        painter.drawLine(638, y_pos, 812, y_pos)

    # Queue bullets/checks.
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QColor("#ffca3a"))
    for y_pos in rows_y[:3]:
        painter.drawEllipse(QtCore.QPointF(598, y_pos), 18, 18)

    arrow_pen = QtGui.QPen(QtGui.QColor("#7a0219"), 20)
    arrow_pen.setCapStyle(QtCore.Qt.RoundCap)
    arrow_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(arrow_pen)
    painter.drawLine(580, 710, 786, 710)
    arrow = QtGui.QPainterPath(QtCore.QPointF(734, 654))
    arrow.lineTo(804, 710)
    arrow.lineTo(734, 766)
    painter.drawPath(arrow)

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

    iconset_dir = output_icns.parent / "sample_prep_queue_builder.iconset"
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
    output_png = root / "assets" / "sample_prep_queue_builder_icon.png"
    output_ico = root / "assets" / "sample_prep_queue_builder_icon.ico"
    output_icns = root / "assets" / "sample_prep_queue_builder_icon.icns"

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
