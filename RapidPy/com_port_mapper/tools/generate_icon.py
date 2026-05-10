from __future__ import annotations

import platform
import subprocess
from pathlib import Path

from PySide6 import QtCore, QtGui


def draw_icon(output_png: Path, output_ico: Path) -> None:
    """Generate a 1024×1024 COM Port Mapper icon with navy gradient and network symbol."""
    size = 1024
    image = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(image)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)

    # Navy-blue-gray gradient background (representing network/connectivity)
    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#1C3A52"))  # Navy-blue
    gradient.setColorAt(1.0, QtGui.QColor("#0D1F2D"))  # Deep navy-gray

    painter.setBrush(QtGui.QBrush(gradient))
    painter.setPen(QtCore.Qt.NoPen)
    painter.drawRoundedRect(48, 48, 928, 928, 210, 210)

    # Gloss overlay (top left, subtle shine)
    gloss = QtGui.QLinearGradient(80, 80, 80, 560)
    gloss.setColorAt(0.0, QtGui.QColor(255, 255, 255, 76))
    gloss.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
    painter.setBrush(QtGui.QBrush(gloss))
    painter.drawRoundedRect(88, 88, 848, 420, 170, 170)

    # --- "CPM" text label — top 30% of canvas ---
    font = QtGui.QFont("SF Pro Display", 185)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Avenir Next", 185)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Segoe UI", 185)
    font.setWeight(QtGui.QFont.Black)
    font.setLetterSpacing(QtGui.QFont.AbsoluteSpacing, 4)

    painter.setFont(font)
    painter.setPen(QtGui.QColor("#E8F0F7"))  # Light cream variant
    # Text rect: y 70‥370 — occupies top ~30% of the 1024-px canvas
    painter.drawText(QtCore.QRect(0, 70, size, 300), QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter, "CPM")

    # --- Network diagram symbol — bottom 45% of canvas (y ≥ 480) ---
    # Three circles connected by lines, suggesting "port mapping" and connectivity
    net_pen = QtGui.QPen(QtGui.QColor("#FFCD34"), 16)  # Gold connecting lines
    net_pen.setCapStyle(QtCore.Qt.RoundCap)
    net_pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(net_pen)

    # Top-left node
    node_tl_x, node_tl_y = 280, 520
    # Top-right node
    node_tr_x, node_tr_y = 740, 520
    # Bottom center node
    node_bc_x, node_bc_y = 510, 780

    # Draw connecting lines
    painter.drawLine(node_tl_x, node_tl_y, node_bc_x, node_bc_y)
    painter.drawLine(node_tr_x, node_tr_y, node_bc_x, node_bc_y)
    painter.drawLine(node_tl_x, node_tl_y, node_tr_x, node_tr_y)

    # Draw node circles (filled with gradient, outlined in gold)
    node_radius = 40
    node_brush = QtGui.QBrush(QtGui.QColor("#FFCD34"))
    node_outline_pen = QtGui.QPen(QtGui.QColor("#FFE07A"), 6)
    node_outline_pen.setCapStyle(QtCore.Qt.RoundCap)

    painter.setBrush(node_brush)
    painter.setPen(node_outline_pen)

    painter.drawEllipse(QtCore.QPointF(node_tl_x, node_tl_y), node_radius, node_radius)
    painter.drawEllipse(QtCore.QPointF(node_tr_x, node_tr_y), node_radius, node_radius)
    painter.drawEllipse(QtCore.QPointF(node_bc_x, node_bc_y), node_radius, node_radius)

    painter.end()

    output_png.parent.mkdir(parents=True, exist_ok=True)
    image.save(str(output_png))
    image.save(str(output_ico))


def generate_icns(output_png: Path, output_icns: Path) -> bool:
    """Generate macOS .icns format (if on macOS with required tools)."""
    if platform.system() != "Darwin":
        return False

    iconutil = subprocess.run(["which", "iconutil"], capture_output=True, text=True)
    sips = subprocess.run(["which", "sips"], capture_output=True, text=True)
    if iconutil.returncode != 0 or sips.returncode != 0:
        return False

    iconset_dir = output_icns.parent / "com_port_mapper.iconset"
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
    output_png = root / "assets" / "com_port_mapper_icon.png"
    output_ico = root / "assets" / "com_port_mapper_icon.ico"
    output_icns = root / "assets" / "com_port_mapper_icon.icns"

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
