from __future__ import annotations

import sys
from pathlib import Path

from PySide6 import QtCore, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import (  # noqa: E402
    MotorAxisConfig,
    MotorSerialClient,
    convert_position_to_hole,
)
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy DC Motor Control")
        self.resize(1160, 700)
        self.client = MotorSerialClient()
        self.axes = {
            "Changer (X)": MotorAxisConfig("ChangerX", 1, 1),
            "Turning": MotorAxisConfig("Turning", 2, 2),
            "Up/Down": MotorAxisConfig("UpDown", 3, 3),
            "Changer (Y)": MotorAxisConfig("ChangerY", 4, 4),
        }
        self._build_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        left = QtWidgets.QFrame()
        left.setObjectName("card")
        l = QtWidgets.QVBoxLayout(left)
        l.setContentsMargins(18, 18, 18, 18)

        title = QtWidgets.QLabel("Motor Control")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Direct move, spin, and hole-position workflows")
        subtitle.setObjectName("subtitle")
        l.addWidget(title)
        l.addWidget(subtitle)

        conn_row = QtWidgets.QHBoxLayout()
        self.port_edit = QtWidgets.QLineEdit("COM4")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")
        conn_row.addWidget(self.port_edit)
        conn_row.addWidget(self.connect_btn)
        conn_row.addWidget(self.disconnect_btn)
        l.addLayout(conn_row)

        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(list(self.axes.keys()))
        self.pos_spin = QtWidgets.QSpinBox()
        self.pos_spin.setRange(-5_000_000, 5_000_000)
        self.speed_spin = QtWidgets.QSpinBox()
        self.speed_spin.setRange(50, 5000)
        self.speed_spin.setValue(1200)
        self.move_btn = QtWidgets.QPushButton("Move Axis")

        form = QtWidgets.QFormLayout()
        form.addRow("Active Axis", self.axis_combo)
        form.addRow("Target Position", self.pos_spin)
        form.addRow("Speed", self.speed_spin)
        form.addRow("", self.move_btn)
        l.addLayout(form)

        self.spin_rps = QtWidgets.QDoubleSpinBox()
        self.spin_rps.setRange(0.01, 20.0)
        self.spin_rps.setValue(1.0)
        self.spin_btn = QtWidgets.QPushButton("Spin Turning Motor")
        spin_form = QtWidgets.QFormLayout()
        spin_form.addRow("Spin (rps)", self.spin_rps)
        spin_form.addRow("", self.spin_btn)
        l.addLayout(spin_form)

        self.hole_spin = QtWidgets.QDoubleSpinBox()
        self.hole_spin.setRange(1.0, 101.0)
        self.hole_spin.setDecimals(1)
        self.goto_hole_btn = QtWidgets.QPushButton("Changer Motor To Hole")
        self.read_hole_btn = QtWidgets.QPushButton("Read Changer Hole")
        hole_form = QtWidgets.QFormLayout()
        hole_form.addRow("Target Hole", self.hole_spin)
        hole_form.addRow("", self.goto_hole_btn)
        hole_form.addRow("", self.read_hole_btn)
        l.addLayout(hole_form)

        ops = QtWidgets.QGridLayout()
        self.home_top_btn = QtWidgets.QPushButton("Home Up/Down To Top")
        self.home_center_btn = QtWidgets.QPushButton("Home XY To Center")
        self.corner_btn = QtWidgets.QPushButton("Move XY To Corner")
        self.pickup_btn = QtWidgets.QPushButton("Sample Pickup")
        self.dropoff_btn = QtWidgets.QPushButton("Sample Dropoff")
        ops.addWidget(self.home_top_btn, 0, 0)
        ops.addWidget(self.home_center_btn, 0, 1)
        ops.addWidget(self.corner_btn, 1, 0)
        ops.addWidget(self.pickup_btn, 1, 1)
        ops.addWidget(self.dropoff_btn, 2, 0, 1, 2)
        l.addLayout(ops)

        self.status = QtWidgets.QLabel("Disconnected")
        self.status.setObjectName("valuePill")
        l.addWidget(self.status)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        l.addWidget(self.console, stretch=1)

        right = QtWidgets.QFrame()
        right.setObjectName("card")
        r = QtWidgets.QVBoxLayout(right)
        r.setContentsMargins(18, 18, 18, 18)
        legend = QtWidgets.QLabel("Active Controls")
        legend.setObjectName("title")
        legend.setStyleSheet("font-size:20px;")
        r.addWidget(legend)
        for axis_name in self.axes:
            pill = QtWidgets.QLabel(axis_name)
            pill.setObjectName("valuePill")
            r.addWidget(pill)
        r.addStretch(1)

        layout.addWidget(left, stretch=3)
        layout.addWidget(right, stretch=1)
        apply_card_shadow(left)
        apply_card_shadow(right)

        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.move_btn.clicked.connect(self._move)
        self.spin_btn.clicked.connect(self._spin)
        self.goto_hole_btn.clicked.connect(self._goto_hole)
        self.read_hole_btn.clicked.connect(self._read_hole)
        self.home_top_btn.clicked.connect(self._home_to_top)
        self.home_center_btn.clicked.connect(self._home_xy_center)
        self.corner_btn.clicked.connect(self._move_xy_corner)
        self.pickup_btn.clicked.connect(self._sample_pickup)
        self.dropoff_btn.clicked.connect(self._sample_dropoff)

    def _log(self, text: str) -> None:
        self.console.appendPlainText(text)
        self.status.setText(text)

    def _connect(self) -> None:
        try:
            self.client.connect(self.port_edit.text().strip(), baudrate=57600)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Connection Error", str(exc))
            return
        self._log(f"Connected to motor serial on {self.port_edit.text().strip()}")

    def _disconnect(self) -> None:
        self.client.disconnect()
        self._log("Disconnected")

    def _selected_axis(self) -> MotorAxisConfig:
        return self.axes[self.axis_combo.currentText()]

    def _move(self) -> None:
        axis = self._selected_axis()
        target = self.pos_spin.value()
        speed = self.speed_spin.value()
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before issuing movement commands.")
            return
        try:
            result = self.client.move_motor(axis, target, speed, wait_for_stop=True)
            self._log(
                f"{axis.name} move: target={result.target}, final={result.final_position}, "
                f"success={result.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Move Error", str(exc))

    def _spin(self) -> None:
        rps = self.spin_rps.value()
        turning = self.axes["Turning"]
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before spin commands.")
            return
        try:
            result = self.client.turning_motor_spin(turning, speed_rps=rps, duration_s=60.0)
            self._log(
                f"Turning spin armed: target={result.target}, final={result.final_position}, "
                f"success={result.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Spin Error", str(exc))

    def _goto_hole(self) -> None:
        hole = self.hole_spin.value()
        axis = self.axes["Changer (X)"]
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before changer-hole moves.")
            return
        try:
            result = self.client.changer_motor_to_hole(axis, hole, wait_for_stop=True)
            self.axis_combo.setCurrentText("Changer (X)")
            self.pos_spin.setValue(result.final_position)
            self._log(
                f"Changer to hole {hole:.2f}: target={result.target}, final={result.final_position}, "
                f"success={result.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Changer Move Error", str(exc))

    def _read_hole(self) -> None:
        hole = convert_position_to_hole(self.pos_spin.value(), slot_min=1, slot_max=101, one_step=-1000)
        self._log(f"Computed changer hole from current position: {hole:.2f}")

    def _home_to_top(self) -> None:
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before homing commands.")
            return
        try:
            result = self.client.home_to_top(self.axes["Up/Down"])
            self._log(f"HomeToTop complete: final={result.final_position}, success={result.success}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "HomeToTop Error", str(exc))

    def _home_xy_center(self) -> None:
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before homing commands.")
            return
        try:
            x_res, y_res = self.client.home_xy_to_center(
                self.axes["Changer (X)"],
                self.axes["Changer (Y)"],
                self.axes["Up/Down"],
            )
            self._log(
                "HomeToCenter complete: "
                f"x={x_res.final_position}, y={y_res.final_position}, success={x_res.success and y_res.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "HomeToCenter Error", str(exc))

    def _move_xy_corner(self) -> None:
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before corner moves.")
            return
        try:
            x_res, y_res = self.client.move_xy_to_corner(
                self.axes["Changer (X)"],
                self.axes["Changer (Y)"],
                self.axes["Up/Down"],
            )
            self._log(f"MoveToCorner complete: x={x_res.final_position}, y={y_res.final_position}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "MoveToCorner Error", str(exc))

    def _sample_pickup(self) -> None:
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before pickup/dropoff.")
            return
        try:
            result = self.client.sample_pickup(self.axes["Up/Down"])
            self._log(f"SamplePickup complete: final={result.final_position}, success={result.success}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "SamplePickup Error", str(exc))

    def _sample_dropoff(self) -> None:
        if not self.client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before pickup/dropoff.")
            return
        try:
            result = self.client.sample_dropoff(self.axes["Up/Down"], use_xy_table=True)
            self._log(f"SampleDropoff complete: final={result.final_position}, success={result.success}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "SampleDropoff Error", str(exc))


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()
