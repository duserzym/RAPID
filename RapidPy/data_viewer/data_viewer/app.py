"""RapidPy DataViewer — lightweight live demagnetization viewer.

This module is intentionally smaller than pmag_gui / demag_gui. It focuses on:
    - quick opening of live CIT or MagIC data files while an experiment is running
    - rapid directional inspection in 2D or 3D
    - simple PCA fit feedback over a chosen step interval
    - lightweight next-step hints for students during AF, thermal, or IZZI runs
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import numpy as np
from PySide6 import QtCore, QtGui, QtWidgets

try:
    import pyqtgraph as pg
    _HAS_PG = True
except ImportError:
    _HAS_PG = False

from data_viewer.analysis import next_step_suggestion, principal_component_fit, summarize_paleointensity, vector_for_step
from data_viewer.data_loading import MeasurementStep, SpecimenMeta, ViewerDataset, ViewerSpecimen, load_input, load_magic_directory, watch_paths_for_dataset


# ── Demo data ────────────────────────────────────────────────────────────────
def _synthetic_dataset() -> ViewerDataset:
    """Generate a synthetic AF dataset for demo use."""
    rng = np.random.default_rng(7)
    nrm = np.array([0.850, 0.450, -0.120], dtype=float)
    steps_mT = [0, 5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100]
    decay = 0.74
    steps: list[MeasurementStep] = []
    vec = nrm.copy()
    for mT in steps_mT:
        noise = rng.normal(0, 0.004, 3)
        sample = vec + noise
        moment = float(np.linalg.norm(sample))
        dec = math.degrees(math.atan2(float(sample[1]), float(sample[0]))) % 360.0
        inc = math.degrees(math.asin(float(-sample[2]) / max(moment, 1e-15)))
        steps.append(
            MeasurementStep(
                demag_label="NRM" if mT == 0 else f"AF{mT}",
                gdec=dec,
                ginc=inc,
                sdec=dec,
                sinc=inc,
                moment=moment,
                error_angle=max(1.0, mT / 12 if mT else 1.0),
                crdec=dec,
                crinc=inc,
                sdx=float(sample[0]),
                sdy=float(sample[1]),
                sdz=float(sample[2]),
            )
        )
        vec = vec * decay
    specimen = ViewerSpecimen(
        name="DEMO_AF01",
        meta=SpecimenMeta(name="DEMO_AF01", comment="Synthetic AF demo specimen"),
        steps=steps,
        experiment_type="AF",
        source_kind="demo",
        source_path=None,
    )
    return ViewerDataset(source_kind="demo", source_path=None, specimens=[specimen])


# ── Geometry helpers ──────────────────────────────────────────────────────────
def _cart_to_inc_dec(vector: np.ndarray) -> tuple[float, float]:
    n, e, u = float(vector[0]), float(vector[1]), float(vector[2])
    h = math.hypot(n, e)
    inc = math.degrees(math.atan2(-u, h))
    dec = math.degrees(math.atan2(e, n)) % 360.0
    return inc, dec


def _equal_area(inc_deg: float, dec_deg: float) -> tuple[float, float]:
    inc = math.radians(abs(inc_deg))
    dec = math.radians(dec_deg)
    r = math.sqrt(2.0) * math.sin((math.pi / 2.0 - inc) / 2.0)
    return r * math.sin(dec), r * math.cos(dec)


def _orthogonal_projection_components(vector: np.ndarray) -> tuple[float, float, float]:
    north, east, up = float(vector[0]), float(vector[1]), float(vector[2])
    return east, north, -up


def _vector_headers_for_coordinate_system(coordinate_system: str) -> tuple[str, str, str]:
    if coordinate_system == "specimen":
        return ("X", "Y", "Z")
    return ("North", "East", "Up")


def _has_paleointensity_data(specimen: ViewerSpecimen | None) -> bool:
    return bool(specimen and specimen.paleointensity_points)


# ── Stereonet widget ──────────────────────────────────────────────────────────
class StereonetWidget(QtWidgets.QWidget):
    """Lambert equal-area stereonet drawn with QPainter."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(300, 300)
        self._specimen: ViewerSpecimen | None = None
        self._coordinate_system = "specimen"
        self._scale_by_csd = False

    def set_specimen(self, specimen: ViewerSpecimen | None, coordinate_system: str, scale_by_csd: bool) -> None:
        self._specimen = specimen
        self._coordinate_system = coordinate_system
        self._scale_by_csd = scale_by_csd
        self.update()

    def paintEvent(self, _e: QtGui.QPaintEvent) -> None:
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        w, h = self.width(), self.height()
        r = min(w, h) / 2.0 - 16
        cx, cy = w / 2.0, h / 2.0

        # Background circle
        p.setPen(QtGui.QPen(QtGui.QColor("#c9d0d7"), 1.5))
        p.setBrush(QtGui.QBrush(QtGui.QColor("#ffffff")))
        p.drawEllipse(QtCore.QPointF(cx, cy), r, r)

        # Grid lines and cardinal ticks
        p.setPen(QtGui.QPen(QtGui.QColor("#e2e8f0"), 1))
        p.drawLine(QtCore.QPointF(cx - r, cy), QtCore.QPointF(cx + r, cy))
        p.drawLine(QtCore.QPointF(cx, cy - r), QtCore.QPointF(cx, cy + r))
        for ang in range(0, 360, 30):
            rad = math.radians(ang)
            p.drawLine(
                QtCore.QPointF(cx + (r - 8) * math.sin(rad), cy - (r - 8) * math.cos(rad)),
                QtCore.QPointF(cx + r * math.sin(rad), cy - r * math.cos(rad)),
            )

        # Cardinal labels
        font = p.font()
        font.setPointSize(8)
        font.setBold(True)
        p.setFont(font)
        p.setPen(QtGui.QPen(QtGui.QColor("#6b7280")))
        p.drawText(QtCore.QPointF(cx - 4, cy - r - 6), "N")
        p.drawText(QtCore.QPointF(cx + r + 6, cy + 4), "E")
        p.drawText(QtCore.QPointF(cx - 4, cy + r + 14), "S")
        p.drawText(QtCore.QPointF(cx - r - 14, cy + 4), "W")

        if not self._specimen or not self._specimen.steps:
            p.end()
            return

        # Plot path line
        pts_lower: list[tuple[QtCore.QPointF, MeasurementStep]] = []
        pts_upper: list[tuple[QtCore.QPointF, MeasurementStep]] = []
        steps = self._specimen.steps

        for i, step in enumerate(steps):
            vector = vector_for_step(step, self._coordinate_system)
            inc, dec = _cart_to_inc_dec(vector)
            x, y = _equal_area(inc, dec)
            px = cx + x * r
            py = cy - y * r
            pt = QtCore.QPointF(px, py)

            if inc >= 0:
                pts_lower.append((pt, step))
            else:
                pts_upper.append((pt, step))

            # Label first and last
            is_first = i == 0
            is_last = i == len(steps) - 1
            if is_first or is_last:
                font2 = p.font()
                font2.setPointSize(7)
                font2.setBold(False)
                p.setFont(font2)
                p.setPen(QtGui.QPen(QtGui.QColor("#4b5563")))
                p.drawText(pt + QtCore.QPointF(8, -4), step.demag_label)

        # Draw connecting path
        pen_path = QtGui.QPen(QtGui.QColor(122, 2, 25, 100), 1, QtCore.Qt.DashLine)
        p.setPen(pen_path)
        all_pts = [pt for pt, _step in pts_lower + pts_upper]
        for i in range(1, len(all_pts)):
            p.drawLine(all_pts[i - 1], all_pts[i])

        # Lower hemisphere: filled circles
        p.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25), 1.5))
        p.setBrush(QtGui.QBrush(QtGui.QColor(122, 2, 25)))
        for pt, step in pts_lower:
            dot_r = _dot_radius(step.error_angle, self._scale_by_csd)
            p.drawEllipse(pt, dot_r, dot_r)

        # Upper hemisphere: open circles
        p.setBrush(QtCore.Qt.NoBrush)
        for pt, step in pts_upper:
            dot_r = _dot_radius(step.error_angle, self._scale_by_csd)
            p.drawEllipse(pt, dot_r, dot_r)

        # Legend
        font3 = p.font()
        font3.setPointSize(8)
        p.setFont(font3)
        p.setPen(QtGui.QPen(QtGui.QColor("#6b7280")))
        p.setBrush(QtGui.QBrush(QtGui.QColor(122, 2, 25)))
        p.drawEllipse(QtCore.QPointF(cx - r + 10, cy + r - 18), 5, 5)
        p.drawText(QtCore.QPointF(cx - r + 20, cy + r - 14), "Lower hemi")
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawEllipse(QtCore.QPointF(cx - r + 10, cy + r - 4), 5, 5)
        p.drawText(QtCore.QPointF(cx - r + 20, cy + r), "Upper hemi")

        p.end()


# ── Zijderveld widget ─────────────────────────────────────────────────────────
class ZijderveldWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(0, 0, 0, 0)
        self._specimen: ViewerSpecimen | None = None
        self._coordinate_system = "specimen"
        self._scale_by_csd = False
        self._fit_steps: list[MeasurementStep] = []
        self._include_origin = False
        self._label_items: list[pg.TextItem] = [] if _HAS_PG else []

        if _HAS_PG:
            pg.setConfigOptions(antialias=True, background="#ffffff", foreground="#1f2937")
            self._plot = pg.PlotWidget()
            self._plot.setAspectLocked(True)
            self._plot.showGrid(x=True, y=True, alpha=0.14)
            self._plot.getAxis("bottom").setLabel("East →")
            self._plot.getAxis("left").setLabel("North ↑    Down ↓")
            self._plot.addLegend(offset=(10, 10))
            self._plot.setMenuEnabled(False)

            self._horiz_line = self._plot.plot(
                [], [], pen=pg.mkPen("#111111", width=1.5),
                name="Horizontal projection (E/N)",
            )
            self._horiz_pts = self._plot.plot(
                [], [], pen=None,
                symbol="o", symbolSize=9,
                symbolBrush=pg.mkBrush("#111111"),
                symbolPen=pg.mkPen("#111111", width=1.6),
            )
            self._vert_line = self._plot.plot(
                [], [], pen=pg.mkPen("#5c6470", width=1.5, style=QtCore.Qt.DashLine),
                name="Vertical projection (E/Down)",
            )
            self._vert_pts = self._plot.plot(
                [], [], pen=None,
                symbol="o", symbolSize=9,
                symbolBrush=pg.mkBrush(None),
                symbolPen=pg.mkPen("#5c6470", width=1.8),
            )
            self._fit_line = self._plot.plot([], [], pen=pg.mkPen("#0c7a6b", width=2, style=QtCore.Qt.DashLine), name="PCA fit")
            vl.addWidget(self._plot)
        else:
            lbl = QtWidgets.QLabel("pyqtgraph not installed\npip install pyqtgraph")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setStyleSheet("color: #9a8885; font-size: 13px;")
            vl.addWidget(lbl)

    def set_specimen(
        self,
        specimen: ViewerSpecimen | None,
        coordinate_system: str,
        scale_by_csd: bool,
        fit_steps: list[MeasurementStep],
        include_origin: bool,
    ) -> None:
        self._specimen = specimen
        self._coordinate_system = coordinate_system
        self._scale_by_csd = scale_by_csd
        self._fit_steps = fit_steps
        self._include_origin = include_origin
        self._update_plot()

    def _update_plot(self) -> None:
        if not _HAS_PG:
            return
        for item in self._label_items:
            self._plot.removeItem(item)
        self._label_items.clear()
        if not self._specimen:
            self._horiz_line.setData([], [])
            self._horiz_pts.setData([], [])
            self._vert_line.setData([], [])
            self._vert_pts.setData([], [])
            self._fit_line.setData([], [])
            return

        vectors = [vector_for_step(step, self._coordinate_system) for step in self._specimen.steps]
        east_north_down = [_orthogonal_projection_components(vector) for vector in vectors]
        e = [component[0] for component in east_north_down]
        n = [component[1] for component in east_north_down]
        d = [component[2] for component in east_north_down]
        sizes = [_dot_radius(step.error_angle, self._scale_by_csd) * 1.8 for step in self._specimen.steps]
        self._horiz_line.setData(e, n)
        self._horiz_pts.setData(e, n, symbolSize=sizes)
        self._vert_line.setData(e, d)
        self._vert_pts.setData(e, d, symbolSize=sizes)

        fit = principal_component_fit(
            self._fit_steps or self._specimen.steps,
            coordinate_system=self._coordinate_system,
            include_origin=self._include_origin,
        )
        if fit is None:
            self._fit_line.setData([], [])
        else:
            fit_vectors = [vector_for_step(step, self._coordinate_system) for step in (self._fit_steps or self._specimen.steps)]
            projections = [float(np.dot(vector - fit.center, fit.direction)) for vector in fit_vectors]
            half_span = max(abs(min(projections, default=-1.0)), abs(max(projections, default=1.0)), 1.0)
            p1 = fit.center - fit.direction * half_span
            p2 = fit.center + fit.direction * half_span
            fit_e1, fit_n1, _fit_d1 = _orthogonal_projection_components(p1)
            fit_e2, fit_n2, _fit_d2 = _orthogonal_projection_components(p2)
            self._fit_line.setData([fit_e1, fit_e2], [fit_n1, fit_n2])

        for index, step in enumerate(self._specimen.steps):
            if index in (0, len(self._specimen.steps) - 1):
                label = pg.TextItem(text=step.demag_label, color="#5b544f", anchor=(0, 1))
                label.setPos(e[index], n[index])
                self._plot.addItem(label)
                self._label_items.append(label)


class Vector3DWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(320, 320)
        self._specimen: ViewerSpecimen | None = None
        self._coordinate_system = "specimen"
        self._fit_steps: list[MeasurementStep] = []
        self._scale_by_csd = False
        self._include_origin = False
        self._yaw = -35.0
        self._pitch = 22.0
        self._last_mouse_pos: QtCore.QPoint | None = None

    def set_specimen(
        self,
        specimen: ViewerSpecimen | None,
        coordinate_system: str,
        scale_by_csd: bool,
        fit_steps: list[MeasurementStep],
        include_origin: bool,
    ) -> None:
        self._specimen = specimen
        self._coordinate_system = coordinate_system
        self._scale_by_csd = scale_by_csd
        self._fit_steps = fit_steps
        self._include_origin = include_origin
        self.update()

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        self._last_mouse_pos = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._last_mouse_pos is not None:
            current = event.position().toPoint()
            delta = current - self._last_mouse_pos
            self._yaw += delta.x() * 0.5
            self._pitch = max(-89.0, min(89.0, self._pitch + delta.y() * 0.35))
            self._last_mouse_pos = current
            self.update()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        self._last_mouse_pos = None
        super().mouseReleaseEvent(event)

    def paintEvent(self, _event: QtGui.QPaintEvent) -> None:
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.fillRect(self.rect(), QtGui.QColor("#ffffff"))
        margin = 26
        center = QtCore.QPointF(self.width() / 2, self.height() / 2)
        radius = max(60.0, min(self.width(), self.height()) / 2 - margin)

        p.setPen(QtGui.QPen(QtGui.QColor("#d9dde3"), 1))
        p.drawRoundedRect(self.rect().adjusted(8, 8, -8, -8), 12, 12)

        if not self._specimen or not self._specimen.steps:
            p.setPen(QtGui.QColor("#6b7280"))
            p.drawText(self.rect(), QtCore.Qt.AlignCenter, "No specimen loaded")
            p.end()
            return

        vectors = [vector_for_step(step, self._coordinate_system) for step in self._specimen.steps]
        norms = [float(np.linalg.norm(vector)) for vector in vectors]
        scale = radius / max(max(norms, default=1.0), 1e-9)

        axes = {
            "X": np.array([1.0, 0.0, 0.0]),
            "Y": np.array([0.0, 1.0, 0.0]),
            "Z": np.array([0.0, 0.0, 1.0]),
        }
        rotated_axes = {name: self._project(vec * (radius / scale), scale, center) for name, vec in axes.items()}

        axis_pen = QtGui.QPen(QtGui.QColor("#bbb4ac"), 1.2)
        p.setPen(axis_pen)
        for name, endpoint in rotated_axes.items():
            p.drawLine(center, endpoint)
            p.drawText(endpoint + QtCore.QPointF(6, -4), name)

        points = [self._project(vector, scale, center) for vector in vectors]

        path_pen = QtGui.QPen(QtGui.QColor("#111111"), 1.6)
        p.setPen(path_pen)
        for index in range(1, len(points)):
            p.drawLine(points[index - 1], points[index])

        fit = principal_component_fit(
            self._fit_steps or self._specimen.steps,
            coordinate_system=self._coordinate_system,
            include_origin=self._include_origin,
        )
        if fit is not None:
            fit_vectors = [vector_for_step(step, self._coordinate_system) for step in (self._fit_steps or self._specimen.steps)]
            projections = [float(np.dot(vector - fit.center, fit.direction)) for vector in fit_vectors]
            half_span = max(abs(min(projections, default=-1.0)), abs(max(projections, default=1.0)), 1.0)
            p1 = self._project(fit.center - fit.direction * half_span, scale, center)
            p2 = self._project(fit.center + fit.direction * half_span, scale, center)
            p.setPen(QtGui.QPen(QtGui.QColor("#0c7a6b"), 2, QtCore.Qt.DashLine))
            p.drawLine(p1, p2)

        for index, point in enumerate(points):
            step = self._specimen.steps[index]
            radius_px = _dot_radius(step.error_angle, self._scale_by_csd)
            if index in (0, len(points) - 1):
                p.setPen(QtGui.QPen(QtGui.QColor("#111111"), 1.4))
                p.setBrush(QtGui.QBrush(QtGui.QColor("#111111")))
                p.drawEllipse(point, radius_px, radius_px)
                p.drawText(point + QtCore.QPointF(8, -6), step.demag_label)
            else:
                p.setPen(QtGui.QPen(QtGui.QColor("#5c6470"), 1.2))
                p.setBrush(QtGui.QBrush(QtGui.QColor(255, 255, 255, 180)))
                p.drawEllipse(point, radius_px, radius_px)

        p.setPen(QtGui.QColor("#6b7280"))
        p.drawText(self.rect().adjusted(16, 12, -16, -12), QtCore.Qt.AlignTop | QtCore.Qt.AlignRight, "Drag to rotate")
        p.end()

    def _project(self, vector: np.ndarray, scale: float, center: QtCore.QPointF) -> QtCore.QPointF:
        yaw = math.radians(self._yaw)
        pitch = math.radians(self._pitch)
        rot_yaw = np.array([
            [math.cos(yaw), -math.sin(yaw), 0.0],
            [math.sin(yaw), math.cos(yaw), 0.0],
            [0.0, 0.0, 1.0],
        ])
        rot_pitch = np.array([
            [1.0, 0.0, 0.0],
            [0.0, math.cos(pitch), -math.sin(pitch)],
            [0.0, math.sin(pitch), math.cos(pitch)],
        ])
        rotated = rot_pitch @ (rot_yaw @ vector)
        x = center.x() + rotated[0] * scale
        y = center.y() - rotated[1] * scale
        return QtCore.QPointF(float(x), float(y))


# ── Intensity widget ──────────────────────────────────────────────────────────
class IntensityWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(0, 0, 0, 0)

        if _HAS_PG:
            pg.setConfigOptions(antialias=True, background="#ffffff", foreground="#1f2937")
            self._plot = pg.PlotWidget()
            self._plot.showGrid(x=True, y=True, alpha=0.14)
            self._plot.getAxis("bottom").setLabel("Step #")
            self._plot.getAxis("left").setLabel("Moment (emu)")
            self._plot.setMenuEnabled(False)
            self._curve = self._plot.plot(
                [], [],
                pen=pg.mkPen("#7A0219", width=2),
                symbol="o", symbolSize=8,
                symbolBrush=pg.mkBrush("#7A0219"),
                symbolPen=pg.mkPen("#7A0219"),
            )
            vl.addWidget(self._plot)
        else:
            lbl = QtWidgets.QLabel("pyqtgraph not installed")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            vl.addWidget(lbl)

    def set_specimen(self, specimen: ViewerSpecimen | None) -> None:
        if not _HAS_PG:
            return
        if not specimen:
            self._curve.setData([], [])
            return
        xs = list(range(len(specimen.steps)))
        ys = [step.moment for step in specimen.steps]
        self._curve.setData(xs, ys)


class PaleointensityWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self._summary = QtWidgets.QLabel("Load an IZZI specimen to view the Arai-style plot.")
        self._summary.setWordWrap(True)
        self._summary.setStyleSheet("color: #5b544f; padding: 6px 10px;")
        layout.addWidget(self._summary)

        if _HAS_PG:
            pg.setConfigOptions(antialias=True, background="#ffffff", foreground="#1f2937")
            self._plot = pg.PlotWidget()
            self._plot.showGrid(x=True, y=True, alpha=0.14)
            self._plot.getAxis("bottom").setLabel("pTRM gained / NRM")
            self._plot.getAxis("left").setLabel("NRM remaining / NRM")
            self._plot.setMenuEnabled(False)
            self._zf_curve = self._plot.plot([], [], pen=pg.mkPen("#111111", width=1.6), symbol="o", symbolBrush=pg.mkBrush("#111111"), symbolPen=pg.mkPen("#111111"), name="ZF/IF path")
            self._ptrm_curve = self._plot.plot([], [], pen=None, symbol="t", symbolSize=10, symbolBrush=pg.mkBrush("#0c7a6b"), symbolPen=pg.mkPen("#0c7a6b"), name="pTRM checks")
            self._fit_line = self._plot.plot([], [], pen=pg.mkPen("#7A0219", width=2, style=QtCore.Qt.DashLine), name="Linear fit")
            layout.addWidget(self._plot, 1)
        else:
            self._plot = None
            label = QtWidgets.QLabel("pyqtgraph not installed")
            label.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(label, 1)

    def set_specimen(self, specimen: ViewerSpecimen | None) -> None:
        if self._plot is None:
            return
        if not specimen or not specimen.paleointensity_points:
            self._summary.setText("No IZZI/Thellier sequence was detected for the current specimen.")
            self._zf_curve.setData([], [])
            self._ptrm_curve.setData([], [])
            self._fit_line.setData([], [])
            return
        summary = summarize_paleointensity(specimen.paleointensity_points)
        zf_if = [point for point in specimen.paleointensity_points if point.step_kind in {"ZF", "IF"}]
        ptrm = [point for point in specimen.paleointensity_points if point.step_kind == "PTRM"]
        self._zf_curve.setData([point.ptrm_gained for point in zf_if], [point.nrm_remaining for point in zf_if])
        self._ptrm_curve.setData([point.ptrm_gained for point in ptrm], [point.nrm_remaining for point in ptrm])
        if summary.slope is not None and summary.intercept is not None:
            xs = np.array([point.ptrm_gained for point in zf_if], dtype=float)
            lo = float(xs.min()) if xs.size else 0.0
            hi = float(xs.max()) if xs.size else 1.0
            if math.isclose(lo, hi):
                hi = lo + 1.0
            span = np.array([lo, hi])
            self._fit_line.setData(span, summary.slope * span + summary.intercept)
            temp_span = ""
            if summary.temperature_span is not None:
                temp_span = f" across {summary.temperature_span[0]:.0f}-{summary.temperature_span[1]:.0f} C"
            self._summary.setText(
                f"Arai fit slope {summary.slope:.2f} from {summary.point_count} points{temp_span}. Use this as a quick trend check, not a full paleointensity solution."
            )
        else:
            self._fit_line.setData([], [])
            self._summary.setText("There are not enough ZF/IF points yet to draw a paleointensity trend.")


# ── Data table widget ─────────────────────────────────────────────────────────
class DataTableWidget(QtWidgets.QTableWidget):
    _BASE_COLS = ("Step", "Inc (°)", "Dec (°)", "Moment (emu)", "CSD (°)")

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(0, 8)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.verticalHeader().setVisible(False)
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.setAlternatingRowColors(True)
        self.setShowGrid(False)
        self.setWordWrap(False)
        self.setSortingEnabled(False)
        self.horizontalHeader().setHighlightSections(False)
        self.horizontalHeader().setStretchLastSection(False)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.verticalHeader().setDefaultSectionSize(26)
        self.setStyleSheet(
            "QTableWidget { alternate-background-color: #f7f8fa; selection-background-color: rgba(122,2,25,0.10); selection-color: #1f2937; }"
            "QHeaderView::section { background: #f3f4f6; color: #374151; padding: 6px 8px; border: 0; border-bottom: 1px solid #d9dde3; font-weight: 600; }"
        )
        self._numeric_font = QtGui.QFontDatabase.systemFont(QtGui.QFontDatabase.FixedFont)
        self._numeric_font.setPointSize(max(9, self._numeric_font.pointSize()))
        self._apply_headers("specimen")

    def _apply_headers(self, coordinate_system: str) -> None:
        vector_headers = _vector_headers_for_coordinate_system(coordinate_system)
        self.setHorizontalHeaderLabels(["Step", *vector_headers, *self._BASE_COLS[1:]])

    def set_specimen(self, specimen: ViewerSpecimen | None, coordinate_system: str) -> None:
        self.setRowCount(0)
        self._apply_headers(coordinate_system)
        if not specimen:
            return
        for step in specimen.steps:
            row = self.rowCount()
            self.insertRow(row)
            vector = vector_for_step(step, coordinate_system)
            inc, dec = _cart_to_inc_dec(vector)
            for col, val in enumerate([
                step.demag_label,
                f"{vector[0]:.3e}",
                f"{vector[1]:.3e}",
                f"{vector[2]:.3e}",
                f"{inc:.1f}",
                f"{dec:.1f}",
                f"{step.moment:.3e}",
                f"{step.error_angle:.1f}",
            ]):
                item = QtWidgets.QTableWidgetItem(val)
                item.setTextAlignment(
                    QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter
                    if col > 0 else QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
                )
                if col > 0:
                    item.setFont(self._numeric_font)
                self.setItem(row, col, item)
        self.resizeColumnsToContents()
        self.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        for col in range(1, self.columnCount()):
            self.horizontalHeader().setSectionResizeMode(col, QtWidgets.QHeaderView.Stretch)


# ── Main window ───────────────────────────────────────────────────────────────
class ZijderveldWindow(QtWidgets.QMainWindow):
    """Main application window for RapidPy DataViewer."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy DataViewer — RAPID")
        self.resize(1380, 860)
        self.setMinimumSize(980, 620)
        self._dataset: ViewerDataset | None = None
        self._current_specimen: ViewerSpecimen | None = None
        self._load_request: tuple[str, Path] | None = None
        self._loaded_dataset_name = ""
        self._watcher = QtCore.QFileSystemWatcher(self)
        self._watcher.fileChanged.connect(self._schedule_watched_reload)
        self._watcher.directoryChanged.connect(self._schedule_watched_reload)
        self._reload_timer = QtCore.QTimer(self)
        self._reload_timer.setSingleShot(True)
        self._reload_timer.setInterval(800)
        self._reload_timer.timeout.connect(self._reload_watched_source)
        self._build_ui()
        self._build_menu()
        self._load_demo()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        self.setCentralWidget(root)
        vl = QtWidgets.QVBoxLayout(root)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(0)

        # Toolbar strip
        tb_frame = QtWidgets.QFrame()
        tb_frame.setFixedHeight(92)
        tb_frame.setStyleSheet(
            "QFrame { background: #ffffff; border-bottom: 1px solid rgba(122,2,25,0.10); }"
        )
        tb_layout = QtWidgets.QVBoxLayout(tb_frame)
        tb_layout.setContentsMargins(14, 8, 14, 8)
        tb_layout.setSpacing(6)

        # Top row: title, dataset, load buttons
        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(8)
        title = QtWidgets.QLabel("RapidPy DataViewer")
        title.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        top_row.addWidget(title)

        top_row.addWidget(_vline())

        self._dataset_lbl = QtWidgets.QLabel("Demo dataset")
        self._dataset_lbl.setStyleSheet("color: #6b7280; font-size: 12px;")
        top_row.addWidget(self._dataset_lbl)

        top_row.addStretch()

        load_btn = QtWidgets.QPushButton("📂  Open File")
        load_btn.clicked.connect(self._open_file)
        dir_btn = QtWidgets.QPushButton("📁  Open MagIC Directory")
        dir_btn.clicked.connect(self._open_directory)
        demo_btn = QtWidgets.QPushButton("🔄  Load Demo")
        demo_btn.clicked.connect(self._load_demo)
        top_row.addWidget(load_btn)
        top_row.addWidget(dir_btn)
        top_row.addWidget(demo_btn)
        tb_layout.addLayout(top_row)

        # Bottom row: controls and settings
        bottom_row = QtWidgets.QHBoxLayout()
        bottom_row.setSpacing(8)

        bottom_row.addWidget(QtWidgets.QLabel("Coords"))
        self._coord_combo = QtWidgets.QComboBox()
        self._coord_combo.addItem("Specimen", "specimen")
        self._coord_combo.addItem("Geographic", "geographic")
        self._coord_combo.addItem("Tilt-corrected", "tilt-corrected")
        self._coord_combo.currentIndexChanged.connect(self._update_all)
        bottom_row.addWidget(self._coord_combo)

        bottom_row.addWidget(QtWidgets.QLabel("Directional view"))
        self._view_combo = QtWidgets.QComboBox()
        self._view_combo.addItems(["2D component view", "3D vector view"])
        self._view_combo.currentIndexChanged.connect(self._sync_directional_stack)
        bottom_row.addWidget(self._view_combo)

        self._scale_csd = QtWidgets.QCheckBox("Scale point size by CSD")
        self._scale_csd.toggled.connect(self._update_all)
        bottom_row.addWidget(self._scale_csd)

        bottom_row.addWidget(QtWidgets.QLabel("Fit range"))
        self._fit_start = QtWidgets.QComboBox()
        self._fit_end = QtWidgets.QComboBox()
        self._fit_start.currentIndexChanged.connect(self._handle_fit_range_change)
        self._fit_end.currentIndexChanged.connect(self._handle_fit_range_change)
        self._fit_start.setMinimumWidth(120)
        self._fit_end.setMinimumWidth(120)
        bottom_row.addWidget(self._fit_start)
        bottom_row.addWidget(QtWidgets.QLabel("to"))
        bottom_row.addWidget(self._fit_end)

        self._include_origin = QtWidgets.QCheckBox("Include origin in fit")
        self._include_origin.toggled.connect(self._update_all)
        bottom_row.addWidget(self._include_origin)

        self._auto_watch = QtWidgets.QCheckBox("Auto-watch")
        self._auto_watch.setChecked(True)
        self._auto_watch.toggled.connect(self._refresh_watch_paths)
        bottom_row.addWidget(self._auto_watch)

        self._watch_lbl = QtWidgets.QLabel("Watch off")
        self._watch_lbl.setStyleSheet("color: #7a736d; font-size: 11px;")
        bottom_row.addWidget(self._watch_lbl)

        bottom_row.addStretch()
        tb_layout.addLayout(bottom_row)

        vl.addWidget(tb_frame)

        main_split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_split.setChildrenCollapsible(False)

        left_panel = QtWidgets.QFrame()
        left_panel.setMaximumWidth(340)
        left_layout = QtWidgets.QVBoxLayout(left_panel)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(10)

        source_card = QtWidgets.QFrame()
        source_card.setObjectName("viewerCard")
        source_layout = QtWidgets.QVBoxLayout(source_card)
        source_layout.setContentsMargins(12, 12, 12, 12)
        source_layout.setSpacing(6)
        source_layout.addWidget(_section_label("Dataset Summary"))
        self._source_lbl = QtWidgets.QLabel("Demo dataset")
        self._source_lbl.setWordWrap(True)
        self._experiment_lbl = QtWidgets.QLabel("Experiment type: AF")
        self._step_count_lbl = QtWidgets.QLabel("Steps: 0")
        self._fit_summary = QtWidgets.QLabel("Fit summary will appear here once a specimen is loaded.")
        self._fit_summary.setWordWrap(True)
        for widget in (self._source_lbl, self._experiment_lbl, self._step_count_lbl, self._fit_summary):
            widget.setStyleSheet("color: #5b544f;")
            source_layout.addWidget(widget)
        left_layout.addWidget(source_card)

        specimen_card = QtWidgets.QFrame()
        specimen_card.setObjectName("viewerCard")
        specimen_layout = QtWidgets.QVBoxLayout(specimen_card)
        specimen_layout.setContentsMargins(12, 12, 12, 12)
        specimen_layout.setSpacing(6)
        specimen_layout.addWidget(_section_label("Specimens"))
        self._specimen_list = QtWidgets.QListWidget()
        self._specimen_list.currentRowChanged.connect(self._change_specimen)
        specimen_layout.addWidget(self._specimen_list, 1)
        helper = QtWidgets.QLabel("Use the fit-range controls above to mimic quick lower/upper-bound PCA selection without opening the heavier PmagPy GUIs.")
        helper.setWordWrap(True)
        helper.setStyleSheet("color: #7a736d; font-size: 11px;")
        specimen_layout.addWidget(helper)
        left_layout.addWidget(specimen_card, 1)

        suggestion_card = QtWidgets.QFrame()
        suggestion_card.setObjectName("viewerCard")
        suggestion_layout = QtWidgets.QVBoxLayout(suggestion_card)
        suggestion_layout.setContentsMargins(12, 12, 12, 12)
        suggestion_layout.setSpacing(8)
        suggestion_layout.addWidget(_section_label("Next-Step Hint"))
        self._suggestion_title = QtWidgets.QLabel("No suggestion yet")
        self._suggestion_title.setStyleSheet("font-size: 13px; font-weight: 700; color: #7A0219;")
        self._suggestion_text = QtWidgets.QTextBrowser()
        self._suggestion_text.setOpenExternalLinks(True)
        self._suggestion_text.setStyleSheet("QTextBrowser { background: #ffffff; border: 1px solid #d9dde3; color: #374151; }")
        suggestion_layout.addWidget(self._suggestion_title)
        suggestion_layout.addWidget(self._suggestion_text, 1)
        left_layout.addWidget(suggestion_card, 1)

        main_split.addWidget(left_panel)

        right_split = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        right_split.setChildrenCollapsible(False)
        right_split.setHandleWidth(6)

        self._directional_stack = QtWidgets.QStackedWidget()
        self._zij = ZijderveldWidget()
        self._vector3d = Vector3DWidget()
        self._directional_stack.addWidget(self._zij)
        self._directional_stack.addWidget(self._vector3d)

        stereo_wrap = QtWidgets.QWidget()
        sl = QtWidgets.QHBoxLayout(stereo_wrap)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.setAlignment(QtCore.Qt.AlignCenter)
        self._stereo = StereonetWidget()
        self._stereo.setMinimumSize(350, 350)
        sl.addWidget(self._stereo)

        self._int_widget = IntensityWidget()
        self._pint_widget = PaleointensityWidget()

        plots_split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        plots_split.setChildrenCollapsible(False)
        plots_split.setHandleWidth(6)
        plots_split.addWidget(_panel_card("Directional", self._directional_stack, "PmagPy-style orthogonal view defaults to east-right and north-up."))

        secondary_split = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        secondary_split.setChildrenCollapsible(False)
        secondary_split.setHandleWidth(6)
        secondary_split.addWidget(_panel_card("Equal-Area Stereonet", stereo_wrap))

        self._analysis_split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self._analysis_split.setChildrenCollapsible(False)
        self._analysis_split.setHandleWidth(6)
        self._analysis_split.addWidget(_panel_card("Intensity Decay", self._int_widget))
        self._pint_card = _panel_card("Paleointensity (Arai)", self._pint_widget)
        self._analysis_split.addWidget(self._pint_card)
        secondary_split.addWidget(self._analysis_split)
        secondary_split.setSizes([320, 260])

        plots_split.addWidget(secondary_split)
        plots_split.setSizes([640, 430])
        right_split.addWidget(plots_split)

        self._table = DataTableWidget()
        right_split.addWidget(_panel_card("Measurement Table", self._table, "Vector columns follow the selected coordinate system."))
        right_split.setSizes([580, 220])
        main_split.addWidget(right_split)
        main_split.setSizes([290, 980])

        vl.addWidget(main_split, 1)

        # Status bar
        self.statusBar().showMessage("Ready — open a CIT .sam file, a MagIC directory, or use demo data.")
        self.setStyleSheet(
            "QFrame#viewerCard { background: #ffffff; border: 1px solid #d9dde3; border-radius: 12px; }"
            "QListWidget, QTableWidget, QComboBox, QTextBrowser { background: #ffffff; }"
        )

    # ── Menu ───────────────────────────────────────────────────────────────
    def _build_menu(self) -> None:
        mb = self.menuBar()
        fm = mb.addMenu("&File")
        fm.addAction("&Open File…", self._open_file)
        fm.addAction("Open &MagIC Directory…", self._open_directory)
        fm.addAction("Load &Demo Dataset", self._load_demo)
        fm.addSeparator()
        fm.addAction("E&xit", QtWidgets.QApplication.quit)

        hm = mb.addMenu("&Help")
        hm.addAction("CSV &Format Help", self._show_format_help)

    # ── Actions ────────────────────────────────────────────────────────────
    def _load_demo(self) -> None:
        dataset = _synthetic_dataset()
        self._apply_dataset(dataset, dataset_name="Demo — synthetic AF demagnetization", load_request=None)
        self.statusBar().showMessage(f"Demo dataset loaded: {len(dataset.specimens[0].steps)} steps.")

    def _open_file(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open CIT SAM or demag CSV",
            "",
            "Supported files (*.sam *.csv);;SAM files (*.sam);;CSV files (*.csv);;All files (*)",
        )
        if not path:
            return
        try:
            source_path = Path(path)
            self._apply_dataset(load_input(source_path), dataset_name=source_path.name, load_request=("input", source_path))
            self.statusBar().showMessage(f"Loaded dataset from {Path(path).name}")
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Load Error", str(exc))

    def _open_directory(self) -> None:
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Open MagIC directory", "")
        if not path:
            return
        try:
            source_path = Path(path)
            self._apply_dataset(load_magic_directory(source_path), dataset_name=source_path.name, load_request=("magic_dir", source_path))
            self.statusBar().showMessage(f"Loaded MagIC directory from {Path(path).name}")
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Load Error", str(exc))

    def _apply_dataset(
        self,
        dataset: ViewerDataset,
        dataset_name: str,
        load_request: tuple[str, Path] | None,
        preferred_specimen: str | None = None,
        preferred_fit_labels: tuple[str | None, str | None] | None = None,
    ) -> None:
        self._dataset = dataset
        self._load_request = load_request
        self._loaded_dataset_name = dataset_name
        self._dataset_lbl.setText(dataset_name)
        self._source_lbl.setText(f"Source: {dataset_name}\nKind: {dataset.source_kind.replace('_', ' ')}")
        self._specimen_list.blockSignals(True)
        self._specimen_list.clear()
        for specimen in dataset.specimens:
            self._specimen_list.addItem(specimen.name)
        self._specimen_list.blockSignals(False)
        if dataset.specimens:
            row = 0
            if preferred_specimen:
                for index, specimen in enumerate(dataset.specimens):
                    if specimen.name == preferred_specimen:
                        row = index
                        break
            self._specimen_list.setCurrentRow(row)
            self._change_specimen(row)
            if preferred_fit_labels is not None:
                self._restore_fit_range(*preferred_fit_labels)
        self._refresh_watch_paths()

    def _change_specimen(self, row: int) -> None:
        if not self._dataset or row < 0 or row >= len(self._dataset.specimens):
            self._current_specimen = None
            self._update_all()
            return
        self._current_specimen = self._dataset.specimens[row]
        self._populate_fit_ranges()
        self._update_all()

    def _populate_fit_ranges(self) -> None:
        self._fit_start.blockSignals(True)
        self._fit_end.blockSignals(True)
        self._fit_start.clear()
        self._fit_end.clear()
        if self._current_specimen:
            for index, step in enumerate(self._current_specimen.steps):
                label = f"{index + 1}. {step.demag_label}"
                self._fit_start.addItem(label, index)
                self._fit_end.addItem(label, index)
            self._fit_start.setCurrentIndex(0)
            self._fit_end.setCurrentIndex(len(self._current_specimen.steps) - 1)
        self._fit_start.blockSignals(False)
        self._fit_end.blockSignals(False)

    def _handle_fit_range_change(self) -> None:
        if self._fit_start.currentIndex() > self._fit_end.currentIndex():
            sender = self.sender()
            if sender is self._fit_start:
                self._fit_end.setCurrentIndex(self._fit_start.currentIndex())
            else:
                self._fit_start.setCurrentIndex(self._fit_end.currentIndex())
        self._update_all()

    def _sync_directional_stack(self) -> None:
        self._directional_stack.setCurrentIndex(self._view_combo.currentIndex())
        self._update_all()

    def _current_coordinate_system(self) -> str:
        value = self._coord_combo.currentData()
        return value if isinstance(value, str) else self._coord_combo.currentText().lower()

    def _current_fit_steps(self) -> list[MeasurementStep]:
        if not self._current_specimen or not self._current_specimen.steps:
            return []
        start = self._fit_start.currentData()
        end = self._fit_end.currentData()
        if start is None or end is None:
            return self._current_specimen.steps
        lo, hi = sorted((int(start), int(end)))
        return self._current_specimen.steps[lo : hi + 1]

    def _capture_view_state(self) -> tuple[str | None, tuple[str | None, str | None]]:
        specimen_name = self._current_specimen.name if self._current_specimen else None
        return specimen_name, (self._current_fit_label(self._fit_start), self._current_fit_label(self._fit_end))

    def _current_fit_label(self, combo: QtWidgets.QComboBox) -> str | None:
        text = combo.currentText().strip()
        if not text:
            return None
        return text.split(". ", 1)[-1]

    def _restore_fit_range(self, start_label: str | None, end_label: str | None) -> None:
        if start_label is not None:
            self._set_fit_combo_to_label(self._fit_start, start_label)
        if end_label is not None:
            self._set_fit_combo_to_label(self._fit_end, end_label)

    def _set_fit_combo_to_label(self, combo: QtWidgets.QComboBox, target_label: str) -> None:
        for index in range(combo.count()):
            if combo.itemText(index).split(". ", 1)[-1] == target_label:
                combo.setCurrentIndex(index)
                return

    def _refresh_watch_paths(self) -> None:
        existing = self._watcher.files() + self._watcher.directories()
        if existing:
            self._watcher.removePaths(existing)

        if not self._auto_watch.isChecked() or self._dataset is None or self._load_request is None:
            self._watch_lbl.setText("Watch off")
            return

        paths = [path for path in watch_paths_for_dataset(self._dataset) if path.exists()]
        files = [str(path) for path in paths if path.is_file()]
        dirs = [str(path) for path in paths if path.is_dir()]
        if files:
            self._watcher.addPaths(files)
        if dirs:
            self._watcher.addPaths(dirs)
        count = len(files) + len(dirs)
        self._watch_lbl.setText(f"Watching {count} path{'s' if count != 1 else ''}")

    def _schedule_watched_reload(self, changed_path: str) -> None:
        if not self._auto_watch.isChecked() or self._load_request is None:
            return
        self._watch_lbl.setText(f"Change detected: {Path(changed_path).name}")
        self._reload_timer.start()

    def _reload_watched_source(self) -> None:
        if self._load_request is None:
            return
        mode, source_path = self._load_request
        specimen_name, fit_labels = self._capture_view_state()
        try:
            dataset = load_magic_directory(source_path) if mode == "magic_dir" else load_input(source_path)
            self._apply_dataset(
                dataset,
                dataset_name=self._loaded_dataset_name or source_path.name,
                load_request=(mode, source_path),
                preferred_specimen=specimen_name,
                preferred_fit_labels=fit_labels,
            )
            stamp = QtCore.QTime.currentTime().toString("HH:mm:ss")
            self._watch_lbl.setText(f"Reloaded {stamp}")
            self.statusBar().showMessage(f"Auto-reloaded {source_path.name} at {stamp}")
        except Exception as exc:
            self._watch_lbl.setText("Auto-reload failed")
            self.statusBar().showMessage(f"Auto-reload failed: {exc}")

    def _update_all(self) -> None:
        coordinate_system = self._current_coordinate_system()
        fit_steps = self._current_fit_steps()
        scale_by_csd = self._scale_csd.isChecked()
        include_origin = self._include_origin.isChecked()
        self._zij.set_specimen(self._current_specimen, coordinate_system, scale_by_csd, fit_steps, include_origin)
        self._vector3d.set_specimen(self._current_specimen, coordinate_system, scale_by_csd, fit_steps, include_origin)
        self._stereo.set_specimen(self._current_specimen, coordinate_system, scale_by_csd)
        self._int_widget.set_specimen(self._current_specimen)
        self._pint_widget.set_specimen(self._current_specimen)
        self._table.set_specimen(self._current_specimen, coordinate_system)
        self._sync_analysis_panels()

        if not self._current_specimen:
            self._experiment_lbl.setText("Experiment type: n/a")
            self._step_count_lbl.setText("Steps: 0")
            self._fit_summary.setText("Fit summary will appear here once a specimen is loaded.")
            self._suggestion_title.setText("No suggestion yet")
            self._suggestion_text.setHtml("<p>Load a specimen to see interpretation guidance.</p>")
            return

        suggestion = next_step_suggestion(self._current_specimen)
        fit = principal_component_fit(fit_steps, coordinate_system=coordinate_system, include_origin=self._include_origin.isChecked()) if fit_steps else None
        self._experiment_lbl.setText(f"Experiment type: {self._current_specimen.experiment_type}")
        self._step_count_lbl.setText(f"Steps: {len(self._current_specimen.steps)} | Fit window: {len(fit_steps)} steps")
        self._fit_summary.setText(_format_fit_summary(fit))
        self._suggestion_title.setText(f"{suggestion.title} ({suggestion.confidence} confidence)")
        self._suggestion_text.setHtml(
            "<p><b>Suggested next step:</b> "
            + suggestion.suggested_step
            + "</p><ul>"
            + "".join(f"<li>{reason}</li>" for reason in suggestion.reasons)
            + "</ul>"
        )

    def _show_format_help(self) -> None:
        QtWidgets.QMessageBox.information(
            self,
            "Supported Inputs",
            "Open one of the following:\n\n"
            "  1. CIT-format .sam file\n"
            "     The viewer reads the listed specimen files next to the .sam file.\n\n"
            "  2. MagIC directory\n"
            "     The directory must contain measurements.txt. Specimens are grouped automatically.\n\n"
            "  3. Legacy CSV\n"
            "     Expected columns: label, n, e, u\n\n"
            "The experiment type is auto-detected from the step labels or MagIC treatment fields.\n"
            "Coordinate views follow the legacy RAPID convention: Specimen, Geographic, and Tilt-corrected.\n"
            "Turn on Auto-watch to refresh the plots automatically while files are being updated.\n"
            "Use the fit-range controls to inspect quick PCA lines without opening the heavier PmagPy GUIs.",
        )

    def _sync_analysis_panels(self) -> None:
        has_paleointensity = _has_paleointensity_data(self._current_specimen)
        self._pint_card.setVisible(has_paleointensity)
        if has_paleointensity:
            self._analysis_split.setSizes([1, 1])
        else:
            self._analysis_split.setSizes([1, 0])


# ── Helpers ───────────────────────────────────────────────────────────────────
def _vline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.VLine)
    f.setStyleSheet("color: rgba(122,2,25,0.20); margin: 8px 2px;")
    return f


def _section_label(text: str) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(text)
    label.setStyleSheet("font-size: 12px; font-weight: 700; color: #7A0219;")
    return label


def _panel_card(title: str, content: QtWidgets.QWidget, note: str | None = None) -> QtWidgets.QFrame:
    card = QtWidgets.QFrame()
    card.setObjectName("viewerCard")
    layout = QtWidgets.QVBoxLayout(card)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(8)
    layout.addWidget(_section_label(title))
    if note:
        note_label = QtWidgets.QLabel(note)
        note_label.setWordWrap(True)
        note_label.setStyleSheet("color: #6b7280; font-size: 11px;")
        layout.addWidget(note_label)
    layout.addWidget(content, 1)
    return card


def _dot_radius(error_angle: float, scale_by_csd: bool) -> float:
    if not scale_by_csd:
        return 5.0
    return max(4.0, min(12.0, 4.0 + float(error_angle) * 0.55))


def _format_fit_summary(fit) -> str:
    if fit is None:
        return "Fit summary unavailable. Select at least two steps in the fit window."
    return (
        f"Fit: Dec {fit.dec:.1f} deg, Inc {fit.inc:.1f} deg, "
        f"MAD {fit.mad:.1f} deg, DANG {fit.dang:.1f} deg."
    )


# ── Entry point ───────────────────────────────────────────────────────────────
def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    assets_dir = None
    try:
        from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon
        apply_liquid_glass_theme(app)
        assets = Path(__file__).resolve().parent.parent / "assets"
        set_app_icon(app, "data_viewer_icon.png", assets)
        assets_dir = assets
    except ImportError:
        pass  # works standalone without rapidpy_common

    win = ZijderveldWindow()
    if assets_dir is not None:  # type: ignore[name-defined]
        set_app_icon(win, "data_viewer_icon.png", assets_dir)
    win.show()
    return app.exec()
