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
    painter.setRenderHint(QtGui.QPainter.TextAntialiasing)

    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#7A0219"))
    gradient.setColorAt(1.0, QtGui.QColor("#41010D"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 240, 240)

    pen = QtGui.QPen(QtGui.QColor("#FFF5E1"), 54)
    pen.setCapStyle(QtCore.Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(QtCore.Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)

    path = QtGui.QPainterPath()
    path.moveTo(140, 620)
    path.cubicTo(250, 220, 360, 220, 470, 620)
    path.lineTo(540, 620)
    path.cubicTo(650, 220, 760, 220, 884, 620)
    painter.drawPath(path)

    clip_pen = QtGui.QPen(QtGui.QColor("#FFD166"), 42)
    clip_pen.setCapStyle(QtCore.Qt.PenCapStyle.RoundCap)
    painter.setPen(clip_pen)
    painter.drawLine(260, 340, 430, 340)
    painter.drawLine(600, 340, 770, 340)
    painter.drawLine(260, 700, 430, 700)
    painter.drawLine(600, 700, 770, 700)

    marker_pen = QtGui.QPen(QtGui.QColor("#FFF5E1"), 20)
    marker_pen.setCapStyle(QtCore.Qt.PenCapStyle.RoundCap)
    painter.setPen(marker_pen)
    painter.drawLine(510, 180, 510, 844)
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

    iconset_dir = output_icns.parent / "af_clip_test.iconset"
    iconset_dir.mkdir(parents=True, exist_ok=True)
    sizes = [16, 32, 64, 128, 256, 512]
    for size in sizes:
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
    output_png = root / "assets" / "af_clip_test_icon.png"
    output_ico = root / "assets" / "af_clip_test_icon.ico"
    output_icns = root / "assets" / "af_clip_test_icon.icns"

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