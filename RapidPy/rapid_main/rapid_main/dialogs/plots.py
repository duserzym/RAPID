from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Sequence

try:
    import numpy as np
except ImportError:
    np = None  # type: ignore[assignment]

from PySide6 import QtCore, QtGui, QtWidgets
from rapidpy_common.ui import clamp_window_geometry

try:
    import pyqtgraph as pg
    _HAS_PG = True
except ImportError:
    _HAS_PG = False


# ── Geometry helpers ──────────────────────────────────────────────────────────
def _equal_area(inc_deg: float, dec_deg: float) -> tuple[float, float]:
    """Lambert equal-area projection.  Returns (x, y) in unit-disk coords."""
    inc = math.radians(abs(inc_deg))
    dec = math.radians(dec_deg)
    r = math.sqrt(2.0) * math.sin((math.pi / 2.0 - inc) / 2.0)
    return r * math.sin(dec), r * math.cos(dec)


def _cart_to_inc_dec(x: float, y: float, z: float) -> tuple[float, float]:
    """Convert Cartesian (N, E, Up) to (Inc, Dec) in degrees."""
    h = math.hypot(x, y)
    inc = math.degrees(math.atan2(-z, h))  # positive = below horizontal
    dec = math.degrees(math.atan2(y, x)) % 360.0
    return inc, dec


def build_quicklook_summary(
    north: Sequence[float],
    east: Sequence[float],
    up: Sequence[float],
    labels: Sequence[str],
) -> dict[str, object]:
    """Build the backend-independent quicklook payload without creating a dialog."""

    n = list(north)
    e = list(east)
    u = list(up)
    step_labels = list(labels)
    lengths = {len(n), len(e), len(u), len(step_labels)}
    if len(lengths) != 1:
        raise ValueError(
            "Quicklook north/east/up vectors and labels must have equal lengths "
            f"(got {len(n)}, {len(e)}, {len(u)}, {len(step_labels)})."
        )
    for axis, values in (("north", n), ("east", e), ("up", u)):
        for index, value in enumerate(values, start=1):
            if not math.isfinite(float(value)):
                raise ValueError(f"Quicklook {axis} value {index} is not finite.")
    down = [-z for z in u]
    intensity = [math.sqrt(x**2 + y**2 + z**2) for x, y, z in zip(n, e, u)]
    inc: list[float] = []
    dec: list[float] = []
    for x, y, z in zip(n, e, u):
        i, d = _cart_to_inc_dec(x, y, z)
        inc.append(i)
        dec.append(d)
    return {
        "step_count": len(step_labels),
        "labels": step_labels,
        "vectors": {
            "north": n,
            "east": e,
            "up": u,
            "down": down,
        },
        "intensity": intensity,
        "inclination": inc,
        "declination": dec,
    }


def write_quicklook_json(path: str | Path, summary: dict[str, object]) -> Path:
    """Write a quicklook summary as a deterministic JSON sidecar artifact."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        json.dumps(summary, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return target


# ── Stereonet widget (custom QPainter, no pyqtgraph dependency) ───────────────
class _StereonetWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(260, 260)
        self._lower: list[tuple[float, float]] = []  # (x, y) projected, lower hemi
        self._upper: list[tuple[float, float]] = []  # upper hemi
        self._labels: list[str] = []

    def set_data(
        self,
        inc: Sequence[float],
        dec: Sequence[float],
        labels: Sequence[str],
    ) -> None:
        self._lower.clear()
        self._upper.clear()
        self._labels = list(labels)
        for i, d in zip(inc, dec):
            x, y = _equal_area(i, d)
            if i >= 0:
                self._lower.append((x, y))
                self._upper.append((None, None))  # type: ignore[arg-type]
            else:
                self._lower.append((None, None))  # type: ignore[arg-type]
                self._upper.append((x, y))
        self.update()

    def paintEvent(self, _event: QtGui.QPaintEvent) -> None:
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        w, h = self.width(), self.height()
        r = min(w, h) / 2.0 - 8
        cx, cy = w / 2.0, h / 2.0

        # Outer circle + cardinal ticks
        p.setPen(QtGui.QPen(QtGui.QColor(180, 160, 155), 1.5))
        p.setBrush(QtGui.QBrush(QtGui.QColor(250, 248, 244)))
        p.drawEllipse(QtCore.QPointF(cx, cy), r, r)

        # Grid lines (two diameters)
        p.setPen(QtGui.QPen(QtGui.QColor(200, 185, 180), 1))
        p.drawLine(QtCore.QPointF(cx - r, cy), QtCore.QPointF(cx + r, cy))
        p.drawLine(QtCore.QPointF(cx, cy - r), QtCore.QPointF(cx, cy + r))

        # Cardinal labels
        p.setPen(QtGui.QPen(QtGui.QColor(155, 135, 130)))
        font = p.font()
        font.setPointSize(8)
        p.setFont(font)
        p.drawText(QtCore.QPointF(cx - 4, cy - r - 4), "N")
        p.drawText(QtCore.QPointF(cx + r + 4, cy + 4), "E")

        # Plot points
        dot_r = 5.0
        for idx, (xl, yl) in enumerate(self._lower):
            if xl is None:
                continue
            px = cx + xl * r
            py = cy - yl * r  # flip y
            p.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25), 1.5))
            p.setBrush(QtGui.QBrush(QtGui.QColor(122, 2, 25)))
            p.drawEllipse(QtCore.QPointF(px, py), dot_r, dot_r)

        for idx, (xu, yu) in enumerate(self._upper):
            if xu is None:
                continue
            px = cx + xu * r
            py = cy - yu * r
            p.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25), 1.5))
            p.setBrush(QtCore.Qt.NoBrush)
            p.drawEllipse(QtCore.QPointF(px, py), dot_r, dot_r)

        p.end()


# ── Zijderveld widget ─────────────────────────────────────────────────────────
class _ZijderveldWidget(QtWidgets.QWidget):
    """Zijderveld demagnetisation diagram using pyqtgraph if available."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(0, 0, 0, 0)

        if _HAS_PG:
            pg.setConfigOption("background", "#faf8f4")
            pg.setConfigOption("foreground", "#2f2827")
            self._plot = pg.PlotWidget()
            self._plot.setAspectLocked(True)
            self._plot.showGrid(x=True, y=True, alpha=0.2)
            self._plot.getAxis("bottom").setLabel("North (N)")
            self._plot.getAxis("left").setLabel("East / Down")
            # Two scatter series
            self._horiz = self._plot.plot(
                [], [],
                pen=pg.mkPen("#7A0219", width=1),
                symbol="o",
                symbolBrush=pg.mkBrush(None),
                symbolPen=pg.mkPen("#7A0219", width=1.5),
                symbolSize=8,
                name="Horizontal (open)",
            )
            self._vert = self._plot.plot(
                [], [],
                pen=pg.mkPen("#7A0219", width=1, style=QtCore.Qt.DashLine),
                symbol="o",
                symbolBrush=pg.mkBrush("#7A0219"),
                symbolPen=pg.mkPen("#7A0219", width=1.5),
                symbolSize=8,
                name="Vertical (filled)",
            )
            vl.addWidget(self._plot)
        else:
            lbl = QtWidgets.QLabel("pyqtgraph not installed — pip install pyqtgraph")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setStyleSheet("color: #9a8885;")
            vl.addWidget(lbl)

    def set_data(
        self,
        north: Sequence[float],
        east: Sequence[float],
        down: Sequence[float],
    ) -> None:
        if not _HAS_PG:
            return
        n = list(north)
        e = list(east)
        d = list(down)
        self._horiz.setData(n, e)   # horizontal: N vs E
        self._vert.setData(n, d)    # vertical: N vs Down


# ── Intensity decay widget ────────────────────────────────────────────────────
class _IntensityWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(0, 0, 0, 0)

        if _HAS_PG:
            pg.setConfigOption("background", "#faf8f4")
            pg.setConfigOption("foreground", "#2f2827")
            self._plot = pg.PlotWidget()
            self._plot.showGrid(x=True, y=True, alpha=0.2)
            self._plot.getAxis("bottom").setLabel("Step")
            self._plot.getAxis("left").setLabel("Intensity (A/m)")
            self._curve = self._plot.plot(
                [], [],
                pen=pg.mkPen("#7A0219", width=2),
                symbol="o",
                symbolBrush=pg.mkBrush("#7A0219"),
                symbolPen=pg.mkPen("#7A0219"),
                symbolSize=7,
            )
            vl.addWidget(self._plot)
        else:
            lbl = QtWidgets.QLabel("pyqtgraph not installed")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            vl.addWidget(lbl)

    def set_data(self, steps: Sequence[int], intensity: Sequence[float]) -> None:
        if not _HAS_PG:
            return
        self._curve.setData(list(steps), list(intensity))


# ── Main dialog ───────────────────────────────────────────────────────────────
class PlotsDialog(QtWidgets.QDialog):
    """Zijderveld + equal-area + intensity plots — replaces VB6 frmPlots."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Demagnetisation Plots")
        self.resize(860, 560)
        self.setWindowFlags(
            self.windowFlags()
            & ~QtCore.Qt.WindowContextHelpButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
        )
        self._last_plot_data: dict[str, object] = {}
        self._build_ui()

    def showEvent(self, event: QtCore.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            if handle.screen() is not None:
                handle.screen().availableGeometryChanged.connect(self._fit_to_screen)
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    def _fit_to_screen(self, screen: QtCore.QObject | None = None) -> None:
        active_screen = (
            screen
            if isinstance(screen, QtGui.QScreen)
            else (self.screen() or QtWidgets.QApplication.primaryScreen())
        )
        if active_screen is None:
            return
        available = active_screen.availableGeometry()
        max_w, max_h = clamp_window_geometry(available, (self.width(), self.height()))
        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), max_w), min(min_size.height(), max_h))
        self.resize(min(self.width(), max_w), min(self.height(), max_h))
        frame = self.frameGeometry()
        frame.setSize(
            QtCore.QSize(
                min(frame.width(), max_w),
                min(frame.height(), max_h),
            )
        )
        if frame.width() > available.width() or frame.height() > available.height():
            frame.moveCenter(available.center())
        else:
            new_x = max(
                available.left(),
                min(frame.left(), available.right() - frame.width() + 1),
            )
            new_y = max(
                available.top(),
                min(frame.top(), available.bottom() - frame.height() + 1),
            )
            frame.moveTopLeft(QtCore.QPoint(new_x, new_y))
        self.setGeometry(frame)

    # ── Public API ─────────────────────────────────────────────────────────
    def set_data(
        self,
        north: Sequence[float],
        east: Sequence[float],
        up: Sequence[float],
        labels: Sequence[str],
    ) -> None:
        """Update all three plots.  Coordinates in A/m (N, E, Up)."""
        summary = build_quicklook_summary(north, east, up, labels)
        vectors = summary["vectors"]
        self._zij.set_data(vectors["north"], vectors["east"], vectors["down"])
        self._int_plot.set_data(
            list(range(len(summary["intensity"]))),
            summary["intensity"],
        )
        self._stereo.set_data(
            summary["inclination"],
            summary["declination"],
            summary["labels"],
        )
        self._last_plot_data = {
            "north": vectors["north"],
            "east": vectors["east"],
            "up": vectors["up"],
            "down": vectors["down"],
            "labels": summary["labels"],
            "intensity": summary["intensity"],
            "inclination": summary["inclination"],
            "declination": summary["declination"],
        }
        self._demo_lbl.setText(
            f"Measurement data — {len(summary['labels'])} step"
            f"{'s' if len(summary['labels']) != 1 else ''}"
        )
        self._demo_lbl.setStyleSheet("color: #4b5563; font-size: 11px;")

    def quicklook_summary(self) -> dict[str, object]:
        """Return a reproducible summary of the current quicklook data."""

        data = getattr(self, "_last_plot_data", {})
        labels = list(data.get("labels", []))
        return {
            "step_count": len(labels),
            "labels": labels,
            "vectors": {
                "north": list(data.get("north", [])),
                "east": list(data.get("east", [])),
                "up": list(data.get("up", [])),
                "down": list(data.get("down", [])),
            },
            "intensity": list(data.get("intensity", [])),
            "inclination": list(data.get("inclination", [])),
            "declination": list(data.get("declination", [])),
        }

    def write_quicklook_json(self, path: str | Path) -> Path:
        """Write the current quicklook summary as a JSON sidecar artifact."""

        return write_quicklook_json(path, self.quicklook_summary())

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        hdr_row = QtWidgets.QHBoxLayout()
        hdr = QtWidgets.QLabel("Demagnetisation Plots")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        hdr_row.addWidget(hdr)
        hdr_row.addStretch()

        self._demo_lbl = QtWidgets.QLabel("No measurement data loaded")
        self._demo_lbl.setStyleSheet("color: #6b7280; font-size: 11px;")
        hdr_row.addWidget(self._demo_lbl)
        vl.addLayout(hdr_row)

        tabs = QtWidgets.QTabWidget()

        # Zijderveld
        zij_wrap = QtWidgets.QWidget()
        self._zij = _ZijderveldWidget()
        QtWidgets.QVBoxLayout(zij_wrap).addWidget(self._zij)
        tabs.addTab(zij_wrap, "Zijderveld")

        # Equal-area stereonet
        stereo_wrap = QtWidgets.QWidget()
        sl = QtWidgets.QHBoxLayout(stereo_wrap)
        sl.setAlignment(QtCore.Qt.AlignCenter)
        self._stereo = _StereonetWidget()
        self._stereo.setFixedSize(320, 320)
        sl.addWidget(self._stereo)
        tabs.addTab(stereo_wrap, "Equal-Area")

        # Intensity decay
        int_wrap = QtWidgets.QWidget()
        self._int_plot = _IntensityWidget()
        QtWidgets.QVBoxLayout(int_wrap).addWidget(self._int_plot)
        tabs.addTab(int_wrap, "Intensity Decay")

        vl.addWidget(tabs, 1)

        # Buttons
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.close)
        demo_btn = QtWidgets.QPushButton("Load SIMULATED example")
        demo_btn.setToolTip("Load synthetic values for UI demonstration only; not hardware evidence")
        demo_btn.clicked.connect(self._load_demo)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addWidget(demo_btn)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

    def _load_demo(self) -> None:
        """Load an explicit, unmistakably simulated UI example."""
        if np is not None:
            rng = np.random.default_rng(42)
            nrm = np.array([0.85, 0.45, -0.12])
            decay = 0.76
            steps = [nrm]
            for _ in range(8):
                prev = steps[-1]
                noise = rng.normal(0, 0.008, 3)
                steps.append(prev * decay + noise)
            steps_arr = np.array(steps)
            labels = ["NRM", "5", "10", "15", "20", "25", "30", "40", "50"]
            self.set_data(
                steps_arr[:, 0].tolist(),
                steps_arr[:, 1].tolist(),
                steps_arr[:, 2].tolist(),
                labels,
            )
            self._mark_simulated_example()
            return

        import random

        rng = random.Random(42)
        nrm = [0.85, 0.45, -0.12]
        decay = 0.76
        steps = [nrm]
        for _ in range(8):
            prev = steps[-1]
            noise = [rng.gauss(0, 0.008) for _ in range(3)]
            steps.append([prev[i] * decay + noise[i] for i in range(3)])
        labels = ["NRM", "5", "10", "15", "20", "25", "30", "40", "50"]
        self.set_data(
            [row[0] for row in steps],
            [row[1] for row in steps],
            [row[2] for row in steps],
            labels,
        )
        self._mark_simulated_example()

    def _mark_simulated_example(self) -> None:
        self._demo_lbl.setText("SIMULATED EXAMPLE — not hardware evidence")
        self._demo_lbl.setStyleSheet(
            "color: #92400e; background: rgba(245,158,11,0.13); "
            "border: 1px solid rgba(217,119,6,0.38); border-radius: 7px; "
            "padding: 3px 8px; font-size: 11px; font-weight: 650;"
        )
