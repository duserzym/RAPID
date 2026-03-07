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

    gloss = QtGui.QLinearGradient(88, 88, 88, 560)
    gloss.setColorAt(0.0, QtGui.QColor(255, 255, 255, 76))
    gloss.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
    painter.setBrush(QtGui.QBrush(gloss))
    painter.drawRoundedRect(88, 88, 848, 420, 170, 170)

    # Cartoon servo body.
    body = QtGui.QPainterPath()
    body.addRoundedRect(QtCore.QRectF(286, 320, 452, 392), 72, 72)
    painter.fillPath(body, QtGui.QColor("#fff6e4"))
    body_pen = QtGui.QPen(QtGui.QColor("#6F0A1F"), 12)
    painter.setPen(body_pen)
    painter.drawPath(body)

    # Mounting ears.
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QColor("#E8B72B"))
    painter.drawRoundedRect(QtCore.QRectF(230, 434, 72, 170), 26, 26)
    painter.drawRoundedRect(QtCore.QRectF(722, 434, 72, 170), 26, 26)

    # Servo top cap and horn.
    painter.setBrush(QtGui.QColor("#7A0219"))
    painter.drawRoundedRect(QtCore.QRectF(360, 250, 304, 104), 34, 34)
    painter.setBrush(QtGui.QColor("#FFCD34"))
    painter.drawEllipse(QtCore.QRectF(448, 200, 128, 128))
    painter.setBrush(QtGui.QColor("#6F0A1F"))
    painter.drawEllipse(QtCore.QRectF(490, 242, 44, 44))

    # Label band.
    painter.setBrush(QtGui.QColor("#7A0219"))
    painter.drawRoundedRect(QtCore.QRectF(344, 456, 336, 120), 30, 30)

    # Vertical motion arrows.
    accent_pen = QtGui.QPen(QtGui.QColor("#FFCD34"), 24)
    accent_pen.setCapStyle(QtCore.Qt.RoundCap)
    painter.setPen(accent_pen)
    painter.drawLine(178, 396, 178, 632)
    painter.drawLine(846, 396, 846, 632)

    arrow_pen = QtGui.QPen(QtGui.QColor("#FFCD34"), 28)
    arrow_pen.setCapStyle(QtCore.Qt.RoundCap)
    arrow_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(arrow_pen)

    up_arrow = QtGui.QPainterPath(QtCore.QPointF(124, 458))
    up_arrow.lineTo(178, 396)
    up_arrow.lineTo(232, 458)
    painter.drawPath(up_arrow)

    down_arrow = QtGui.QPainterPath(QtCore.QPointF(792, 570))
    down_arrow.lineTo(846, 632)
    down_arrow.lineTo(900, 570)
    painter.drawPath(down_arrow)
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

    iconset_dir = output_icns.parent / "updown.iconset"
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
    output_png = root / "assets" / "updown_icon.png"
    output_ico = root / "assets" / "updown_icon.ico"
    output_icns = root / "assets" / "updown_icon.icns"

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
