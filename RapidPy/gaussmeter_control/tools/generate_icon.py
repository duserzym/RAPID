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

    # --- Background: deep navy blue rounded square ---
    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#0B2545"))
    gradient.setColorAt(1.0, QtGui.QColor("#061328"))

    painter.setBrush(QtGui.QBrush(gradient))
    painter.setPen(QtCore.Qt.NoPen)
    painter.drawRoundedRect(48, 48, 928, 928, 210, 210)

    # --- Gloss highlight (top half) ---
    gloss = QtGui.QLinearGradient(80, 80, 80, 560)
    gloss.setColorAt(0.0, QtGui.QColor(255, 255, 255, 70))
    gloss.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
    painter.setBrush(QtGui.QBrush(gloss))
    painter.drawRoundedRect(88, 88, 848, 420, 170, 170)

    # --- Magnetic field lines (concentric arcs around a pole point) ---
    # Draw field lines radiating upward from a source point at bottom-center.
    # Each arc is a semicircle of increasing radius, offset vertically.
    arc_pen = QtGui.QPen(QtGui.QColor("#38C1FF"), 44)
    arc_pen.setCapStyle(QtCore.Qt.RoundCap)
    arc_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(arc_pen)
    painter.setBrush(QtCore.Qt.NoBrush)

    # --- "GM" text label — top 30% of canvas ---
    font = QtGui.QFont("SF Pro Display", 210)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Avenir Next", 210)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Segoe UI", 210)
    font.setWeight(QtGui.QFont.Black)
    font.setLetterSpacing(QtGui.QFont.AbsoluteSpacing, 8)

    painter.setFont(font)
    painter.setPen(QtGui.QColor("#D8EEFF"))
    # Text rect: y 70‥370 — occupies top ~30% of the 1024-px canvas
    painter.drawText(QtCore.QRect(0, 70, size, 300), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "GM")

    # --- Magnetic field arcs — bottom 45% of canvas (y ≥ 460) ---
    arc_pen = QtGui.QPen(QtGui.QColor("#38C1FF"), 44)
    arc_pen.setCapStyle(QtCore.Qt.RoundCap)
    arc_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(arc_pen)
    painter.setBrush(QtCore.Qt.NoBrush)

    # Source pole at y=900 so the largest arc top ≈ 900 - 340*1.25 = 475
    cx, cy = 512, 900

    for radius in (140, 240, 340):
        path = QtGui.QPainterPath()
        path.moveTo(cx + radius, cy)
        path.cubicTo(
            cx + radius, cy - radius * 1.25,
            cx - radius, cy - radius * 1.25,
            cx - radius, cy,
        )
        painter.drawPath(path)

    # --- Small filled circle at the pole source ---
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QBrush(QtGui.QColor("#38C1FF")))
    painter.drawEllipse(QtCore.QPointF(cx, cy), 36, 36)
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

    iconset_dir = output_icns.parent / "gaussmeter.iconset"
    iconset_dir.mkdir(parents=True, exist_ok=True)

    sizes = [16, 32, 64, 128, 256, 512]
    for s in sizes:
        subprocess.run(
            ["sips", "-z", str(s), str(s), str(output_png), "--out", str(iconset_dir / f"icon_{s}x{s}.png")],
            check=True,
        )
        subprocess.run(
            ["sips", "-z", str(s * 2), str(s * 2), str(output_png), "--out", str(iconset_dir / f"icon_{s}x{s}@2x.png")],
            check=True,
        )

    subprocess.run(["iconutil", "-c", "icns", str(iconset_dir), "-o", str(output_icns)], check=True)
    return output_icns.exists()


def main() -> int:
    app = QtGui.QGuiApplication([])
    root = Path(__file__).resolve().parent.parent
    output_png = root / "assets" / "gaussmeter_icon.png"
    output_ico = root / "assets" / "gaussmeter_icon.ico"
    output_icns = root / "assets" / "gaussmeter_icon.icns"

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
