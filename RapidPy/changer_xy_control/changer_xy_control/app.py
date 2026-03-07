from __future__ import annotations

import csv
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6 import QtCore, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import MotorAxisConfig, MotorSerialClient  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


@dataclass(slots=True)
class QueueOptions:
    ascending: bool = True
    repeat_holder: bool = True
    return_to_start: bool = True


class MainWindow(QtWidgets.QMainWindow):
    SLOT_MIN = 1
    SLOT_MAX = 100

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Changer XY Control")
        self.resize(1180, 760)
        self.motor = MotorSerialClient()
        self.changer_axis = MotorAxisConfig("ChangerX", 1, 1)
        self.changer_y_axis = MotorAxisConfig("ChangerY", 4, 4)
        self.updown_axis = MotorAxisConfig("UpDown", 3, 3)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        left = QtWidgets.QFrame()
        left.setObjectName("card")
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(18, 18, 18, 18)

        title = QtWidgets.QLabel("Hole Sample List")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("VB6-like sample order + queue preparation")
        subtitle.setObjectName("subtitle")
        left_layout.addWidget(title)
        left_layout.addWidget(subtitle)

        self.table = QtWidgets.QTableWidget(self.SLOT_MAX, 2)
        self.table.setHorizontalHeaderLabels(["Hole", "Sample"])
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        for idx in range(self.SLOT_MAX):
            hole_item = QtWidgets.QTableWidgetItem(str(self.SLOT_MIN + idx))
            hole_item.setFlags(hole_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.table.setItem(idx, 0, hole_item)
            self.table.setItem(idx, 1, QtWidgets.QTableWidgetItem(""))
        left_layout.addWidget(self.table, stretch=1)

        actions = QtWidgets.QHBoxLayout()
        self.clear_btn = QtWidgets.QPushButton("Clear")
        self.load_csv_btn = QtWidgets.QPushButton("Load CSV")
        self.save_csv_btn = QtWidgets.QPushButton("Save CSV")
        self.process_btn = QtWidgets.QPushButton("Process to Queue")
        self.process_btn.setObjectName("accent")
        actions.addWidget(self.clear_btn)
        actions.addWidget(self.load_csv_btn)
        actions.addWidget(self.save_csv_btn)
        actions.addWidget(self.process_btn)
        left_layout.addLayout(actions)

        right = QtWidgets.QFrame()
        right.setObjectName("card")
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(18, 18, 18, 18)

        opts_title = QtWidgets.QLabel("Queue Options")
        opts_title.setObjectName("title")
        opts_title.setStyleSheet("font-size:20px;")
        right_layout.addWidget(opts_title)

        conn = QtWidgets.QHBoxLayout()
        self.port_edit = QtWidgets.QLineEdit("COM4")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")
        conn.addWidget(self.port_edit)
        conn.addWidget(self.connect_btn)
        conn.addWidget(self.disconnect_btn)
        right_layout.addLayout(conn)

        self.order_group = QtWidgets.QButtonGroup(self)
        self.asc = QtWidgets.QRadioButton("Ascending")
        self.desc = QtWidgets.QRadioButton("Descending")
        self.asc.setChecked(True)
        self.order_group.addButton(self.asc)
        self.order_group.addButton(self.desc)
        right_layout.addWidget(self.asc)
        right_layout.addWidget(self.desc)

        self.repeat_holder = QtWidgets.QCheckBox("Repeat holder measurements")
        self.repeat_holder.setChecked(True)
        self.return_to_start = QtWidgets.QCheckBox("Return to start after queue")
        self.return_to_start.setChecked(True)
        right_layout.addWidget(self.repeat_holder)
        right_layout.addWidget(self.return_to_start)

        self.current_hole = QtWidgets.QSpinBox()
        self.current_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        self.goto_hole = QtWidgets.QSpinBox()
        self.goto_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        self.goto_btn = QtWidgets.QPushButton("Goto Hole")
        self.home_center_btn = QtWidgets.QPushButton("Home XY To Center")
        self.corner_btn = QtWidgets.QPushButton("Move XY To Corner")
        form = QtWidgets.QFormLayout()
        form.addRow("Current Hole", self.current_hole)
        form.addRow("Target Hole", self.goto_hole)
        form.addRow("", self.goto_btn)
        form.addRow("", self.home_center_btn)
        form.addRow("", self.corner_btn)
        right_layout.addLayout(form)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        right_layout.addWidget(self.console, stretch=1)

        layout.addWidget(left, stretch=3)
        layout.addWidget(right, stretch=2)
        apply_card_shadow(left)
        apply_card_shadow(right)

        self.clear_btn.clicked.connect(self._clear)
        self.load_csv_btn.clicked.connect(self._load_csv)
        self.save_csv_btn.clicked.connect(self._save_csv)
        self.process_btn.clicked.connect(self._process)
        self.goto_btn.clicked.connect(self._goto_hole)
        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.home_center_btn.clicked.connect(self._home_xy_center)
        self.corner_btn.clicked.connect(self._move_xy_corner)

    def _connect(self) -> None:
        try:
            self.motor.connect(self.port_edit.text().strip(), baudrate=57600)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Connection Error", str(exc))
            return
        self._append(f"Connected changer motor on {self.port_edit.text().strip()}.")

    def _disconnect(self) -> None:
        self.motor.disconnect()
        self._append("Disconnected changer motor.")

    def _append(self, msg: str) -> None:
        self.console.appendPlainText(msg)

    def _clear(self) -> None:
        for row in range(self.SLOT_MAX):
            self.table.item(row, 1).setText("")
        self._append("Cleared changer sample grid.")

    def _load_csv(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load Sample CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        with open(path, "r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            self._clear()
            for row in reader:
                hole = int(row.get("hole", "0"))
                sample = row.get("sample", "")
                if self.SLOT_MIN <= hole <= self.SLOT_MAX:
                    self.table.item(hole - self.SLOT_MIN, 1).setText(sample)
        self._append(f"Loaded sample grid from {path}.")

    def _save_csv(self) -> None:
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Sample CSV", "changer_samples.csv", "CSV Files (*.csv)")
        if not path:
            return
        with open(path, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["hole", "sample"])
            writer.writeheader()
            for row in range(self.SLOT_MAX):
                sample = self.table.item(row, 1).text().strip()
                if sample:
                    writer.writerow({"hole": row + self.SLOT_MIN, "sample": sample})
        self._append(f"Saved sample grid to {path}.")

    def _process(self) -> None:
        options = QueueOptions(
            ascending=self.asc.isChecked(),
            repeat_holder=self.repeat_holder.isChecked(),
            return_to_start=self.return_to_start.isChecked(),
        )
        sample_rows = []
        for row in range(self.SLOT_MAX):
            sample = self.table.item(row, 1).text().strip()
            if sample:
                sample_rows.append((row + self.SLOT_MIN, sample))
        sample_rows.sort(key=lambda item: item[0], reverse=not options.ascending)
        self._append(
            f"Queue prepared: {len(sample_rows)} sample entries, ascending={options.ascending}, "
            f"repeat_holder={options.repeat_holder}, return_to_start={options.return_to_start}."
        )

    def _goto_hole(self) -> None:
        target_hole = float(self.goto_hole.value())
        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect changer motor serial before goto-hole operations.")
            return
        try:
            result = self.motor.changer_motor_to_hole(self.changer_axis, target_hole, wait_for_stop=True)
            hole = self.current_hole.value()
            self.current_hole.setValue(int(round(target_hole)))
            self._append(
                f"Moved changer from hole {hole} to {target_hole:.2f}: "
                f"target_pos={result.target}, final_pos={result.final_position}, success={result.success}."
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Changer Move Error", str(exc))

    def _home_xy_center(self) -> None:
        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before homing commands.")
            return
        try:
            x_res, y_res = self.motor.home_xy_to_center(
                self.changer_axis,
                self.changer_y_axis,
                self.updown_axis,
            )
            self._append(
                "HomeToCenter complete: "
                f"x={x_res.final_position}, y={y_res.final_position}, success={x_res.success and y_res.success}."
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "HomeToCenter Error", str(exc))

    def _move_xy_corner(self) -> None:
        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before corner moves.")
            return
        try:
            x_res, y_res = self.motor.move_xy_to_corner(
                self.changer_axis,
                self.changer_y_axis,
                self.updown_axis,
            )
            self._append(f"MoveToCorner complete: x={x_res.final_position}, y={y_res.final_position}.")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "MoveToCorner Error", str(exc))


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()
