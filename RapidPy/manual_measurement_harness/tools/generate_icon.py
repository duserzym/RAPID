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
    gradient.setColorAt(0.0, QtGui.QColor("#7a0219"))
    gradient.setColorAt(1.0, QtGui.QColor("#4d0a16"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    title_font = QtGui.QFont("Avenir Next", 102)
    title_font.setBold(True)
    title_font.setLetterSpacing(QtGui.QFont.PercentageSpacing, 104)
    painter.setFont(title_font)
    painter.setPen(QtGui.QColor("#f4ede0"))
    painter.drawText(QtCore.QRectF(0, 88, size, 108), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "MM")

    # Measurement card.
    card = QtCore.QRectF(184, 244, 656, 540)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(card, 62, 62)
    painter.setBrush(QtGui.QColor("#f7f1e7"))
    painter.drawRoundedRect(card.adjusted(28, 28, -28, -28), 44, 44)

    # Core/sample silhouette.
    painter.setBrush(QtGui.QColor("#2a2c31"))
    painter.drawRoundedRect(QtCore.QRectF(286, 336, 126, 300), 52, 52)
    painter.setBrush(QtGui.QColor("#d3a34c"))
    painter.drawEllipse(QtCore.QRectF(286, 308, 126, 94))
    painter.setBrush(QtGui.QColor("#f5c867"))
    painter.drawEllipse(QtCore.QRectF(310, 330, 78, 46))

    # Holder / sensor rings.
    ring_pen = QtGui.QPen(QtGui.QColor("#31566d"), 16)
    painter.setPen(ring_pen)
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawEllipse(QtCore.QRectF(448, 342, 248, 248))
    painter.drawEllipse(QtCore.QRectF(491, 385, 162, 162))

    # Live trace.
    trace_pen = QtGui.QPen(QtGui.QColor("#ffca3a"), 22)
    trace_pen.setCapStyle(QtCore.Qt.RoundCap)
    trace_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(trace_pen)
    trace = QtGui.QPainterPath(QtCore.QPointF(262, 700))
    trace.lineTo(348, 700)
    trace.lineTo(388, 640)
    trace.lineTo(448, 752)
    trace.lineTo(530, 596)
    trace.lineTo(596, 700)
    trace.lineTo(758, 700)
    painter.drawPath(trace)

    # Accent crosshair.
    painter.setPen(QtGui.QPen(QtGui.QColor("#7a0219"), 12))
    painter.drawLine(572, 424, 572, 506)
    painter.drawLine(532, 465, 612, 465)

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

    iconset_dir = output_icns.parent / "manual_measurement_harness.iconset"
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
    output_png = root / "assets" / "manual_measurement_harness_icon.png"
    output_ico = root / "assets" / "manual_measurement_harness_icon.ico"
    output_icns = root / "assets" / "manual_measurement_harness_icon.icns"

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
