"""Zijderveld Viewer — main window.

Standalone PySide6 app for viewing palaeomagnetic demagnetisation data.
Demonstrates the plot widgets that are also used inside rapid_main Phase 2.

Data model
----------
Each loaded dataset is a list of DemagStep namedtuples:
    step_label (str)   — "NRM", "5 mT", "10 mT", …
    n  (float)         — North component  (A/m or mA/m)
    e  (float)         — East  component
    u  (float)         — Up    component  (positive = up)
"""
from __future__ import annotations

import csv
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

import numpy as np
from PySide6 import QtCore, QtGui, QtWidgets

try:
    import pyqtgraph as pg
    _HAS_PG = True
except ImportError:
    _HAS_PG = False


# ── Data model ────────────────────────────────────────────────────────────────
@dataclass
class DemagStep:
    label: str
    n: float   # North (A/m)
    e: float   # East
    u: float   # Up


def _synthetic_dataset() -> list[DemagStep]:
    """Generate a synthetic NRM→AF demagnetisation sequence for demo use."""
    rng = np.random.default_rng(7)
    nrm = np.array([0.850, 0.450, -0.120])
    steps_mT = [0, 5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100]
    decay = 0.74
    data: list[DemagStep] = []
    vec = nrm.copy()
    for mT in steps_mT:
        noise = rng.normal(0, 0.004, 3)
        data.append(DemagStep(
            label="NRM" if mT == 0 else f"{mT} mT",
            n=float(vec[0] + noise[0]),
            e=float(vec[1] + noise[1]),
            u=float(vec[2] + noise[2]),
        ))
        vec = vec * decay
    return data


# ── Geometry helpers ──────────────────────────────────────────────────────────
def _cart_to_inc_dec(n: float, e: float, u: float) -> tuple[float, float]:
    h = math.hypot(n, e)
    inc = math.degrees(math.atan2(-u, h))
    dec = math.degrees(math.atan2(e, n)) % 360.0
    return inc, dec


def _equal_area(inc_deg: float, dec_deg: float) -> tuple[float, float]:
    inc = math.radians(abs(inc_deg))
    dec = math.radians(dec_deg)
    r = math.sqrt(2.0) * math.sin((math.pi / 2.0 - inc) / 2.0)
    return r * math.sin(dec), r * math.cos(dec)


# ── Stereonet widget ──────────────────────────────────────────────────────────
class StereonetWidget(QtWidgets.QWidget):
    """Lambert equal-area stereonet drawn with QPainter."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(300, 300)
        self._dataset: list[DemagStep] = []

    def set_dataset(self, data: list[DemagStep]) -> None:
        self._dataset = data
        self.update()

    def paintEvent(self, _e: QtGui.QPaintEvent) -> None:
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        w, h = self.width(), self.height()
        r = min(w, h) / 2.0 - 16
        cx, cy = w / 2.0, h / 2.0

        # Background circle
        p.setPen(QtGui.QPen(QtGui.QColor(180, 160, 155), 1.5))
        p.setBrush(QtGui.QBrush(QtGui.QColor(252, 249, 244)))
        p.drawEllipse(QtCore.QPointF(cx, cy), r, r)

        # Grid lines and cardinal ticks
        p.setPen(QtGui.QPen(QtGui.QColor(210, 195, 190), 1))
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
        p.setPen(QtGui.QPen(QtGui.QColor(155, 135, 130)))
        p.drawText(QtCore.QPointF(cx - 4, cy - r - 6), "N")
        p.drawText(QtCore.QPointF(cx + r + 6, cy + 4), "E")
        p.drawText(QtCore.QPointF(cx - 4, cy + r + 14), "S")
        p.drawText(QtCore.QPointF(cx - r - 14, cy + 4), "W")

        if not self._dataset:
            p.end()
            return

        # Plot path line
        pts_lower: list[QtCore.QPointF] = []
        pts_upper: list[QtCore.QPointF] = []
        dot_r = 6.0

        for i, step in enumerate(self._dataset):
            inc, dec = _cart_to_inc_dec(step.n, step.e, step.u)
            x, y = _equal_area(inc, dec)
            px = cx + x * r
            py = cy - y * r
            pt = QtCore.QPointF(px, py)

            if inc >= 0:
                pts_lower.append(pt)
            else:
                pts_upper.append(pt)

            # Label first and last
            is_first = i == 0
            is_last = i == len(self._dataset) - 1
            if is_first or is_last:
                font2 = p.font()
                font2.setPointSize(7)
                font2.setBold(False)
                p.setFont(font2)
                p.setPen(QtGui.QPen(QtGui.QColor(100, 80, 75)))
                p.drawText(pt + QtCore.QPointF(8, -4), step.label)

        # Draw connecting path
        pen_path = QtGui.QPen(QtGui.QColor(122, 2, 25, 100), 1, QtCore.Qt.DashLine)
        p.setPen(pen_path)
        all_pts = pts_lower + pts_upper
        for i in range(1, len(all_pts)):
            p.drawLine(all_pts[i - 1], all_pts[i])

        # Lower hemisphere: filled circles
        p.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25), 1.5))
        p.setBrush(QtGui.QBrush(QtGui.QColor(122, 2, 25)))
        for pt in pts_lower:
            p.drawEllipse(pt, dot_r, dot_r)

        # Upper hemisphere: open circles
        p.setBrush(QtCore.Qt.NoBrush)
        for pt in pts_upper:
            p.drawEllipse(pt, dot_r, dot_r)

        # Legend
        font3 = p.font()
        font3.setPointSize(8)
        p.setFont(font3)
        p.setPen(QtGui.QPen(QtGui.QColor(120, 100, 95)))
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
        self._dataset: list[DemagStep] = []

        if _HAS_PG:
            pg.setConfigOptions(antialias=True, background="#faf8f4", foreground="#2f2827")
            self._plot = pg.PlotWidget()
            self._plot.setAspectLocked(True)
            self._plot.showGrid(x=True, y=True, alpha=0.18)
            self._plot.getAxis("bottom").setLabel("North →")
            self._plot.getAxis("left").setLabel("← East     Down →")
            self._plot.addLegend(offset=(10, 10))

            self._horiz_line = self._plot.plot(
                [], [], pen=pg.mkPen("#7A0219", width=1.5),
                name="Horizontal (N, E)",
            )
            self._horiz_pts = self._plot.plot(
                [], [], pen=None,
                symbol="o", symbolSize=9,
                symbolBrush=pg.mkBrush(None),
                symbolPen=pg.mkPen("#7A0219", width=1.8),
            )
            self._vert_line = self._plot.plot(
                [], [], pen=pg.mkPen("#7A0219", width=1.5, style=QtCore.Qt.DashLine),
                name="Vertical (N, Down)",
            )
            self._vert_pts = self._plot.plot(
                [], [], pen=None,
                symbol="o", symbolSize=9,
                symbolBrush=pg.mkBrush("#7A0219"),
                symbolPen=pg.mkPen("#7A0219", width=1.8),
            )
            vl.addWidget(self._plot)
        else:
            lbl = QtWidgets.QLabel("pyqtgraph not installed\npip install pyqtgraph")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setStyleSheet("color: #9a8885; font-size: 13px;")
            vl.addWidget(lbl)

    def set_dataset(self, data: list[DemagStep]) -> None:
        self._dataset = data
        if not _HAS_PG:
            return
        n  = [s.n for s in data]
        e  = [s.e for s in data]
        d  = [-s.u for s in data]  # Down = -Up
        self._horiz_line.setData(n, e)
        self._horiz_pts.setData(n, e)
        self._vert_line.setData(n, d)
        self._vert_pts.setData(n, d)


# ── Intensity widget ──────────────────────────────────────────────────────────
class IntensityWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(0, 0, 0, 0)

        if _HAS_PG:
            pg.setConfigOptions(antialias=True, background="#faf8f4", foreground="#2f2827")
            self._plot = pg.PlotWidget()
            self._plot.showGrid(x=True, y=True, alpha=0.18)
            self._plot.getAxis("bottom").setLabel("Step #")
            self._plot.getAxis("left").setLabel("Intensity (A/m)")
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

    def set_dataset(self, data: list[DemagStep]) -> None:
        if not _HAS_PG:
            return
        xs = list(range(len(data)))
        ys = [math.sqrt(s.n**2 + s.e**2 + s.u**2) for s in data]
        self._curve.setData(xs, ys)


# ── Data table widget ─────────────────────────────────────────────────────────
class DataTableWidget(QtWidgets.QTableWidget):
    _COLS = ("Step", "N (A/m)", "E (A/m)", "U (A/m)", "Inc (°)", "Dec (°)", "Intensity")

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(0, len(self._COLS))
        self.setHorizontalHeaderLabels(list(self._COLS))
        self.horizontalHeader().setStretchLastSection(True)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.verticalHeader().setVisible(False)

    def set_dataset(self, data: list[DemagStep]) -> None:
        self.setRowCount(0)
        for step in data:
            row = self.rowCount()
            self.insertRow(row)
            inc, dec = _cart_to_inc_dec(step.n, step.e, step.u)
            intensity = math.sqrt(step.n**2 + step.e**2 + step.u**2)
            for col, val in enumerate([
                step.label,
                f"{step.n:.5f}",
                f"{step.e:.5f}",
                f"{step.u:.5f}",
                f"{inc:.2f}",
                f"{dec:.2f}",
                f"{intensity:.5f}",
            ]):
                item = QtWidgets.QTableWidgetItem(val)
                item.setTextAlignment(
                    QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter
                    if col > 0 else QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
                )
                self.setItem(row, col, item)
        self.resizeColumnsToContents()


# ── Main window ───────────────────────────────────────────────────────────────
class ZijderveldWindow(QtWidgets.QMainWindow):
    """Main application window for the Zijderveld Viewer."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Zijderveld Viewer — RAPID")
        self.resize(1120, 720)
        self.setMinimumSize(800, 520)
        self._dataset: list[DemagStep] = []
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
        tb_frame.setFixedHeight(46)
        tb_frame.setStyleSheet(
            "QFrame { background: #fef9f4; border-bottom: 1px solid rgba(122,2,25,0.15); }"
        )
        tbl = QtWidgets.QHBoxLayout(tb_frame)
        tbl.setContentsMargins(14, 0, 14, 0)
        tbl.setSpacing(8)

        title = QtWidgets.QLabel("⚗  Zijderveld Viewer")
        title.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        tbl.addWidget(title)

        tbl.addWidget(_vline())

        self._dataset_lbl = QtWidgets.QLabel("Demo dataset")
        self._dataset_lbl.setStyleSheet("color: #6b7280; font-size: 12px;")
        tbl.addWidget(self._dataset_lbl)

        tbl.addStretch()

        load_btn = QtWidgets.QPushButton("📂  Load CSV")
        load_btn.clicked.connect(self._load_csv)
        demo_btn = QtWidgets.QPushButton("🔄  Load Demo")
        demo_btn.clicked.connect(self._load_demo)
        tbl.addWidget(load_btn)
        tbl.addWidget(demo_btn)

        vl.addWidget(tb_frame)

        # Splitter: plots (top) / data table (bottom)
        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        # Plot tab widget
        self._tabs = QtWidgets.QTabWidget()

        self._zij = ZijderveldWidget()
        self._tabs.addTab(self._zij, "Zijderveld")

        stereo_wrap = QtWidgets.QWidget()
        sl = QtWidgets.QHBoxLayout(stereo_wrap)
        sl.setAlignment(QtCore.Qt.AlignCenter)
        self._stereo = StereonetWidget()
        self._stereo.setMinimumSize(350, 350)
        sl.addWidget(self._stereo)
        self._tabs.addTab(stereo_wrap, "Equal-Area Stereonet")

        self._int_widget = IntensityWidget()
        self._tabs.addTab(self._int_widget, "Intensity Decay")

        splitter.addWidget(self._tabs)

        # Data table
        self._table = DataTableWidget()
        splitter.addWidget(self._table)

        splitter.setSizes([480, 200])
        vl.addWidget(splitter, 1)

        # Status bar
        self.statusBar().showMessage("Ready — load a CSV or use demo data.")

    # ── Menu ───────────────────────────────────────────────────────────────
    def _build_menu(self) -> None:
        mb = self.menuBar()
        fm = mb.addMenu("&File")
        fm.addAction("&Load CSV…", self._load_csv)
        fm.addAction("Load &Demo Dataset", self._load_demo)
        fm.addSeparator()
        fm.addAction("E&xit", QtWidgets.QApplication.quit)

        hm = mb.addMenu("&Help")
        hm.addAction("CSV &Format Help", self._show_format_help)

    # ── Actions ────────────────────────────────────────────────────────────
    def _load_demo(self) -> None:
        self._dataset = _synthetic_dataset()
        self._dataset_lbl.setText("Demo — synthetic AF demagnetisation")
        self._update_all()
        self.statusBar().showMessage(
            f"Demo dataset loaded: {len(self._dataset)} steps."
        )

    def _load_csv(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Demagnetisation CSV", "", "CSV files (*.csv);;All files (*)"
        )
        if not path:
            return
        try:
            data = _parse_csv(path)
            self._dataset = data
            self._dataset_lbl.setText(Path(path).name)
            self._update_all()
            self.statusBar().showMessage(
                f"Loaded {len(data)} steps from {Path(path).name}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Load Error", str(exc))

    def _update_all(self) -> None:
        self._zij.set_dataset(self._dataset)
        self._stereo.set_dataset(self._dataset)
        self._int_widget.set_dataset(self._dataset)
        self._table.set_dataset(self._dataset)

    def _show_format_help(self) -> None:
        QtWidgets.QMessageBox.information(
            self,
            "CSV Format",
            "Expected CSV columns (header row required):\n\n"
            "  label, n, e, u\n\n"
            "Where n/e/u are the North, East, Up components in A/m.\n"
            "Example:\n"
            "  label,n,e,u\n"
            "  NRM,0.85,0.45,-0.12\n"
            "  5 mT,0.64,0.34,-0.09\n"
            "  ...",
        )


# ── CSV parser ────────────────────────────────────────────────────────────────
def _parse_csv(path: str) -> list[DemagStep]:
    data: list[DemagStep] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(DemagStep(
                label=row.get("label", row.get("Label", str(len(data)))),
                n=float(row.get("n", row.get("N", 0))),
                e=float(row.get("e", row.get("E", 0))),
                u=float(row.get("u", row.get("U", 0))),
            ))
    if not data:
        raise ValueError("CSV file appears empty or has unrecognised column names.")
    return data


# ── Helpers ───────────────────────────────────────────────────────────────────
def _vline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.VLine)
    f.setStyleSheet("color: rgba(122,2,25,0.20); margin: 8px 2px;")
    return f


# ── Entry point ───────────────────────────────────────────────────────────────
def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    try:
        from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon
        apply_liquid_glass_theme(app)
        assets = Path(__file__).resolve().parent.parent / "assets"
        set_app_icon(app, "zijderveld_icon.png", assets)
    except ImportError:
        pass  # works standalone without rapidpy_common

    win = ZijderveldWindow()
    win.show()
    return app.exec()
