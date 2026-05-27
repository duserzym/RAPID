from __future__ import annotations

import platform
import shutil
import subprocess
from pathlib import Path

from PySide6 import QtCore, QtGui


def draw_icon(output_png: Path, output_ico: Path) -> None:
    size = 1024
    tile_color = QtGui.QColor("#CBB15C")
    ring_color = QtGui.QColor("#E4A83C")
    image = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(image)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)

    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(tile_color)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    panel_rect = QtCore.QRectF(190, 208, 644, 612)
    painter.setBrush(QtGui.QColor("#050505"))
    painter.drawRoundedRect(panel_rect, 74, 74)

    inner_rect = panel_rect.adjusted(28, 28, -28, -28)
    panel_gradient = QtGui.QLinearGradient(inner_rect.topLeft(), inner_rect.bottomRight())
    panel_gradient.setColorAt(0.0, QtGui.QColor("#fbf6eb"))
    panel_gradient.setColorAt(1.0, QtGui.QColor("#e7ddd0"))
    painter.setBrush(panel_gradient)
    painter.drawRoundedRect(inner_rect, 50, 50)

    plot_rect = inner_rect.adjusted(54, 86, -54, -64)
    painter.setPen(QtGui.QPen(QtGui.QColor("#1b1b1d"), 14, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap))
    painter.drawLine(QtCore.QPointF(plot_rect.left(), plot_rect.bottom()), QtCore.QPointF(plot_rect.right(), plot_rect.bottom()))
    painter.drawLine(QtCore.QPointF(plot_rect.left(), plot_rect.bottom()), QtCore.QPointF(plot_rect.left(), plot_rect.top()))

    grid_pen = QtGui.QPen(QtGui.QColor(80, 86, 95, 60), 6)
    painter.setPen(grid_pen)
    for frac in (0.25, 0.5, 0.75):
        x = plot_rect.left() + plot_rect.width() * frac
        y = plot_rect.top() + plot_rect.height() * frac
        painter.drawLine(QtCore.QPointF(x, plot_rect.top()), QtCore.QPointF(x, plot_rect.bottom()))
        painter.drawLine(QtCore.QPointF(plot_rect.left(), y), QtCore.QPointF(plot_rect.right(), y))

    path = QtGui.QPainterPath()
    points = [
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.10, plot_rect.top() + plot_rect.height() * 0.16),
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.26, plot_rect.top() + plot_rect.height() * 0.24),
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.43, plot_rect.top() + plot_rect.height() * 0.42),
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.59, plot_rect.top() + plot_rect.height() * 0.60),
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.76, plot_rect.top() + plot_rect.height() * 0.72),
        QtCore.QPointF(plot_rect.left() + plot_rect.width() * 0.90, plot_rect.top() + plot_rect.height() * 0.82),
    ]
    path.moveTo(points[0])
    for point in points[1:]:
        path.lineTo(point)
    painter.setPen(QtGui.QPen(QtGui.QColor("#7A0219"), 20, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
    painter.drawPath(path)

    filled_brush = QtGui.QBrush(QtGui.QColor("#7A0219"))
    open_pen = QtGui.QPen(QtGui.QColor("#5b6570"), 12)
    for index, point in enumerate(points):
        if index in (1, 3, 5):
            painter.setPen(open_pen)
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawEllipse(point, 20, 20)
        else:
            painter.setPen(QtGui.QPen(QtGui.QColor("#7A0219"), 4))
            painter.setBrush(filled_brush)
            painter.drawEllipse(point, 16, 16)

    net_center = QtCore.QPointF(inner_rect.right() - 128, inner_rect.top() + 130)
    painter.setPen(QtGui.QPen(QtGui.QColor("#5b6570"), 10))
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawEllipse(net_center, 92, 92)
    painter.drawLine(QtCore.QPointF(net_center.x() - 92, net_center.y()), QtCore.QPointF(net_center.x() + 92, net_center.y()))
    painter.drawLine(QtCore.QPointF(net_center.x(), net_center.y() - 92), QtCore.QPointF(net_center.x(), net_center.y() + 92))
    painter.setBrush(QtGui.QBrush(QtGui.QColor("#111111")))
    painter.setPen(QtGui.QPen(QtGui.QColor("#111111"), 4))
    painter.drawEllipse(QtCore.QPointF(net_center.x() - 26, net_center.y() + 18), 14, 14)
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.setPen(QtGui.QPen(QtGui.QColor("#111111"), 8))
    painter.drawEllipse(QtCore.QPointF(net_center.x() + 34, net_center.y() - 28), 16, 16)

    highlight = points[2]
    painter.setPen(QtGui.QPen(ring_color, 14))
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawEllipse(highlight, 28, 28)

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

    iconset_dir = output_icns.parent / "data_viewer.iconset"
    if iconset_dir.exists():
        shutil.rmtree(iconset_dir)
    iconset_dir.mkdir(parents=True, exist_ok=True)

    for size in (16, 32, 64, 128, 256, 512):
        subprocess.run(["sips", "-z", str(size), str(size), str(output_png), "--out", str(iconset_dir / f"icon_{size}x{size}.png")], check=True)
        subprocess.run(["sips", "-z", str(size * 2), str(size * 2), str(output_png), "--out", str(iconset_dir / f"icon_{size}x{size}@2x.png")], check=True)

    subprocess.run(["iconutil", "-c", "icns", str(iconset_dir), "-o", str(output_icns)], check=True)
    shutil.rmtree(iconset_dir, ignore_errors=True)
    return output_icns.exists()


def main() -> int:
    app = QtGui.QGuiApplication([])
    root = Path(__file__).resolve().parent.parent
    output_png = root / "assets" / "data_viewer_icon.png"
    output_ico = root / "assets" / "data_viewer_icon.ico"
    output_icns = root / "assets" / "data_viewer_icon.icns"

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