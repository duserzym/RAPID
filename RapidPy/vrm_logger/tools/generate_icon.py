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
    gradient.setColorAt(0.0, QtGui.QColor("#7A0219"))
    gradient.setColorAt(1.0, QtGui.QColor("#520012"))

    painter.setBrush(QtGui.QBrush(gradient))
    painter.setPen(QtCore.Qt.NoPen)
    painter.drawRoundedRect(48, 48, 928, 928, 210, 210)

    gloss = QtGui.QLinearGradient(80, 80, 80, 560)
    gloss.setColorAt(0.0, QtGui.QColor(255, 255, 255, 76))
    gloss.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
    painter.setBrush(QtGui.QBrush(gloss))
    painter.drawRoundedRect(88, 88, 848, 420, 170, 170)

    wave_pen = QtGui.QPen(QtGui.QColor("#FFCD34"), 54)
    wave_pen.setCapStyle(QtCore.Qt.RoundCap)
    painter.setPen(wave_pen)
    path = QtGui.QPainterPath(QtCore.QPointF(120, 640))
    path.cubicTo(220, 360, 330, 860, 430, 640)
    path.cubicTo(530, 420, 640, 760, 760, 570)
    path.cubicTo(840, 470, 890, 530, 930, 500)
    painter.drawPath(path)

    font = QtGui.QFont("SF Pro Display", 270)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Avenir Next", 270)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Segoe UI", 270)
    font.setWeight(QtGui.QFont.Black)
    font.setLetterSpacing(QtGui.QFont.AbsoluteSpacing, 12)

    painter.setFont(font)
    painter.setPen(QtGui.QColor("#fff6e4"))
    painter.drawText(QtCore.QRect(0, 140, size, 360), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop, "VRM")
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

    iconset_dir = output_icns.parent / "vrm.iconset"
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
    output_png = root / "assets" / "vrm_icon.png"
    output_ico = root / "assets" / "vrm_icon.ico"
    output_icns = root / "assets" / "vrm_icon.icns"

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
