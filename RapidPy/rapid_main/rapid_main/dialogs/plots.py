from __future__ import annotations

import math
from typing import Sequence

import numpy as np
from PySide6 import QtCore, QtGui, QtWidgets

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
        self._build_ui()
        self._load_demo()

    # ── Public API ─────────────────────────────────────────────────────────
    def set_data(
        self,
        north: Sequence[float],
        east: Sequence[float],
        up: Sequence[float],
        labels: Sequence[str],
    ) -> None:
        """Update all three plots.  Coordinates in A/m (N, E, Up)."""
        down = [-z for z in up]
        self._zij.set_data(north, east, down)

        intensity = [math.sqrt(n**2 + e**2 + u**2) for n, e, u in zip(north, east, up)]
        self._int_plot.set_data(list(range(len(intensity))), intensity)

        inc = []
        dec = []
        for n, e, u in zip(north, east, up):
            i, d = _cart_to_inc_dec(n, e, u)
            inc.append(i)
            dec.append(d)
        self._stereo.set_data(inc, dec, labels)

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

        self._demo_lbl = QtWidgets.QLabel("Demo data — load real data in Phase 3")
        self._demo_lbl.setStyleSheet("color: #9a8885; font-size: 11px;")
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
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

    def _load_demo(self) -> None:
        """Synthetic demagnetization sequence for demonstration."""
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
