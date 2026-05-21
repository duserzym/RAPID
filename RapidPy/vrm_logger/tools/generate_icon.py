from __future__ import annotations

import platform
import subprocess
from pathlib import Path

from PySide6 import QtCore, QtGui


def draw_icon(glyph_png: Path, output_png: Path, output_ico: Path) -> None:
    source = QtGui.QImage(str(glyph_png)).convertToFormat(QtGui.QImage.Format_ARGB32)
    if source.isNull():
        raise FileNotFoundError(f"Missing glyph image: {glyph_png}")

    left = source.width()
    top = source.height()
    right = -1
    bottom = -1
    for y in range(source.height()):
        for x in range(source.width()):
            if source.pixelColor(x, y).alpha() > 0:
                left = min(left, x)
                top = min(top, y)
                right = max(right, x)
                bottom = max(bottom, y)

    if right < left or bottom < top:
        raise RuntimeError(f"No visible glyph pixels found in {glyph_png}")

    glyph = source.copy(left, top, right - left + 1, bottom - top + 1)
    target = 460
    scale = target / max(glyph.width(), glyph.height())
    glyph = glyph.scaled(
        max(1, int(round(glyph.width() * scale))),
        max(1, int(round(glyph.height() * scale))),
        QtCore.Qt.KeepAspectRatio,
        QtCore.Qt.SmoothTransformation,
    )

    size = 1024
    image = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(image)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)

    gradient = QtGui.QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QtGui.QColor("#7A0219"))
    gradient.setColorAt(1.0, QtGui.QColor("#520012"))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(gradient)
    painter.drawRoundedRect(QtCore.QRectF(0, 0, size, size), 256, 256)

    x_pos = (size - glyph.width()) // 2
    y_pos = (size - glyph.height()) // 2
    painter.drawImage(x_pos, y_pos, glyph)
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
    glyph_png = root / "assets" / "vrm_icon_glyph.png"
    output_png = root / "assets" / "vrm_icon.png"
    output_ico = root / "assets" / "vrm_icon.ico"
    output_icns = root / "assets" / "vrm_icon.icns"

    draw_icon(glyph_png, output_png, output_ico)
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
