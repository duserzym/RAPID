from __future__ import annotations

import json
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path

import serial
from serial.tools import list_ports
from PySide6 import QtCore, QtGui, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorControllerConfig, MotorSerialClient, MoveResult  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


CONFIG_PATH = Path.home() / ".rapidpy_changer_xy_control.json"
TRAY_GRID_COLS = 10
TRAY_GRID_ROWS = 10
TRAY_CENTER_DROP_HOLE = 46
TRAY_ROW_OFFSET = 0.5
TRAY_ROW_STEP = 0.88
TRAY_CUP_DIAMETER_IN = 1.0
TRAY_CUP_PITCH_IN = 1.32
TRAY_CENTER_DROP_DIAMETER_IN = 1.12
MOTION_DEFAULTS = MotorControllerConfig()


def tray_logical_hole_positions(slot_min: int, slot_max: int) -> dict[int, QtCore.QPointF]:
    positions: dict[int, QtCore.QPointF] = {}
    for hole in range(slot_min, slot_max + 1):
        row = (hole - 1) // TRAY_GRID_COLS
        col = (hole - 1) % TRAY_GRID_COLS
        x = col + 0.5 + (TRAY_ROW_OFFSET if row % 2 else 0.0)
        y = row * TRAY_ROW_STEP + 0.5
        positions[hole] = QtCore.QPointF(x, y)
    return positions


def tray_logical_bounds(slot_min: int, slot_max: int) -> tuple[float, float, float, float]:
    positions = tray_logical_hole_positions(slot_min, slot_max)
    xs = [point.x() for point in positions.values()]
    ys = [point.y() for point in positions.values()]
    return min(xs), max(xs), min(ys), max(ys)


@dataclass(slots=True)
class HoleCalibration:
    hole: int
    x: int | None = None
    y: int | None = None


@dataclass(slots=True)
class AppConfig:
    x_port: str = "COM4"
    y_port: str = "COM5"
    updown_port: str = "COM6"
    sensor_port: str = ""
    selected_hole: int = 1
    drag_moves_enabled: bool = False
    calibrations: list[HoleCalibration] = field(default_factory=list)


class SerialProbe:
    def __init__(self) -> None:
        self._serial: serial.Serial | None = None

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    def connect(self, port: str, baudrate: int = 9600, timeout: float = 0.35) -> None:
        self.disconnect()
        self._serial = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
            write_timeout=timeout,
        )

    def disconnect(self) -> None:
        if self._serial is not None:
            try:
                self._serial.close()
            finally:
                self._serial = None


class ChangerStageController:
    def __init__(self) -> None:
        self.x_axis = MotorAxisConfig("ChangerX", 1, 1)
        self.y_axis = MotorAxisConfig("ChangerY", 4, 4)
        self.updown_axis = MotorAxisConfig("UpDown", 3, 3)
        self._motor_clients: dict[str, MotorSerialClient] = {}
        self._axis_ports: dict[str, str] = {"x": "", "y": "", "updown": ""}
        self._sensor_probe = SerialProbe()
        self._sensor_port = ""

    def shutdown(self) -> None:
        for client in list(self._motor_clients.values()):
            client.disconnect()
        self._motor_clients.clear()
        self._axis_ports = {"x": "", "y": "", "updown": ""}
        self.disconnect_sensor()

    def axis_port(self, key: str) -> str:
        return self._axis_ports.get(key, "")

    def connect_axis(self, key: str, port: str) -> None:
        port = port.strip()
        if not port:
            raise HardwareError("A COM port is required.")
        existing_port = self._axis_ports.get(key, "")
        if existing_port == port and port in self._motor_clients and self._motor_clients[port].is_connected:
            return
        if existing_port and existing_port != port:
            self.disconnect_axis(key)
        client = self._motor_clients.get(port)
        if client is None or not client.is_connected:
            client = MotorSerialClient()
            client.connect(port, baudrate=57600)
            self._motor_clients[port] = client
        self._axis_ports[key] = port

    def disconnect_axis(self, key: str) -> None:
        port = self._axis_ports.get(key, "")
        if not port:
            return
        self._axis_ports[key] = ""
        if port and port not in self._axis_ports.values():
            client = self._motor_clients.pop(port, None)
            if client is not None:
                client.disconnect()

    def connect_sensor(self, port: str) -> None:
        port = port.strip()
        if not port:
            raise HardwareError("A COM port is required.")
        self._sensor_probe.connect(port)
        self._sensor_port = port

    def disconnect_sensor(self) -> None:
        self._sensor_probe.disconnect()
        self._sensor_port = ""

    @property
    def sensor_port(self) -> str:
        return self._sensor_port

    @property
    def sensor_connected(self) -> bool:
        return self._sensor_probe.is_connected

    def axis_connected(self, key: str) -> bool:
        port = self._axis_ports.get(key, "")
        return bool(port and port in self._motor_clients and self._motor_clients[port].is_connected)

    def _client_for_axis(self, key: str) -> MotorSerialClient:
        port = self._axis_ports.get(key, "")
        if not port:
            raise HardwareError(f"Axis {key.upper()} is not assigned to a COM port.")
        client = self._motor_clients.get(port)
        if client is None or not client.is_connected:
            raise HardwareError(f"Axis {key.upper()} is not connected.")
        return client

    def _axis(self, key: str) -> MotorAxisConfig:
        if key == "x":
            return self.x_axis
        if key == "y":
            return self.y_axis
        return self.updown_axis

    def read_axis_position(self, key: str) -> int:
        client = self._client_for_axis(key)
        return client.read_position(self._axis(key))

    def read_xy_position(self) -> tuple[int | None, int | None]:
        x_pos = self.read_axis_position("x") if self.axis_connected("x") else None
        y_pos = self.read_axis_position("y") if self.axis_connected("y") else None
        return x_pos, y_pos

    def check_status_bit(self, key: str, bit: int) -> int:
        client = self._client_for_axis(key)
        return client.check_internal_status(self._axis(key), bit)

    def reference_switch_states(self) -> dict[str, bool | None]:
        states: dict[str, bool | None] = {"z_top": None, "x_neg": None, "x_pos": None, "y_neg": None, "y_pos": None}
        if self.axis_connected("updown"):
            states["z_top"] = self.check_status_bit("updown", 4) == 1
        if self.axis_connected("x"):
            states["x_neg"] = self.check_status_bit("x", 4) == 0
            states["x_pos"] = self.check_status_bit("x", 5) == 0
        if self.axis_connected("y"):
            states["y_neg"] = self.check_status_bit("y", 5) == 0
            states["y_pos"] = self.check_status_bit("y", 6) == 0
        return states

    def home_to_top(self) -> MoveResult:
        client = self._client_for_axis("updown")
        return client.home_to_top(self.updown_axis)

    def _require_axes(self, *keys: str) -> None:
        missing = [key.upper() for key in keys if not self.axis_connected(key)]
        if missing:
            raise HardwareError(f"Missing motor connections: {', '.join(missing)}")

    def home_xy_to_center(self) -> tuple[MoveResult, MoveResult]:
        self._require_axes("x", "y", "updown")
        if self.check_status_bit("updown", 4) == 0:
            self.home_to_top()
        if self.check_status_bit("updown", 4) == 0:
            raise HardwareError("Cannot home XY to center: up/down axis is not homed to top.")

        start = time.monotonic()
        stop_x = False
        stop_y = False

        while (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out during negative-limit XY homing pass.")
            if self.check_status_bit("x", 4) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=-30_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 5) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=-30_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        if (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            raise HardwareError("Homed XY center pass 1 failed: did not hit negative limit switches.")

        start = time.monotonic()
        stop_x = False
        stop_y = False
        while (self.check_status_bit("x", 5) != 0) or (self.check_status_bit("y", 6) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out during positive-limit XY homing pass.")
            if self.check_status_bit("x", 5) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=30_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 6) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=30_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-3,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        if (self.check_status_bit("x", 5) != 0) or (self.check_status_bit("y", 6) != 0):
            raise HardwareError("Homed XY center pass 2 failed: did not hit positive limit switches.")

        x_client = self._client_for_axis("x")
        y_client = self._client_for_axis("y")
        x_client.zero_target_pos(self.x_axis)
        y_client.zero_target_pos(self.y_axis)
        x_pos = x_client.read_position(self.x_axis)
        y_pos = y_client.read_position(self.y_axis)
        return (
            MoveResult(target=0, final_position=x_pos, success=True),
            MoveResult(target=0, final_position=y_pos, success=True),
        )

    def move_xy_to_corner(self) -> tuple[MoveResult, MoveResult]:
        self._require_axes("x", "y", "updown")
        self.home_to_top()
        if self.check_status_bit("updown", 4) == 0:
            raise HardwareError("Cannot move XY to corner: up/down axis is not homed to top.")

        start = time.monotonic()
        stop_x = False
        stop_y = False
        while (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out moving XY stage to corner.")
            if self.check_status_bit("x", 4) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=-3_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 5) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=-3_000_000,
                    velocity=8_000_000,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        x_pos = self.read_axis_position("x")
        y_pos = self.read_axis_position("y")
        return (
            MoveResult(target=x_pos, final_position=x_pos, success=True),
            MoveResult(target=y_pos, final_position=y_pos, success=True),
        )

    def move_xy_absolute(self, x_target: int, y_target: int) -> tuple[MoveResult, MoveResult]:
        self._require_axes("x", "y", "updown")
        if self.check_status_bit("updown", 4) == 0:
            self.home_to_top()
        x_result = self._client_for_axis("x").move_motor(
            self.x_axis,
            target=int(x_target),
            velocity=8_000_000,
            wait_for_stop=False,
            acceleration=483184,
            relative_mode=False,
        )
        y_result = self._client_for_axis("y").move_motor(
            self.y_axis,
            target=int(y_target),
            velocity=8_000_000,
            wait_for_stop=True,
            acceleration=483184,
            relative_mode=False,
        )
        final_x = self.read_axis_position("x")
        final_y = self.read_axis_position("y")
        return (
            MoveResult(target=x_result.target, final_position=final_x, success=True),
            MoveResult(target=y_result.target, final_position=final_y, success=True),
        )

    def move_axis_relative(self, key: str, delta: int, velocity: int | None = None) -> MoveResult:
        self._require_axes(key)
        current = self.read_axis_position(key)
        axis = self._axis(key)
        move_velocity = int(velocity) if velocity is not None else 8_000_000
        return self._client_for_axis(key).move_motor(
            axis,
            target=int(current + delta),
            velocity=move_velocity,
            wait_for_stop=True,
            acceleration=483184,
            relative_mode=False,
        )


class StageScene(QtWidgets.QWidget):
    holeSelected = QtCore.Signal(int)
    holeDragged = QtCore.Signal(int)
    holeActivated = QtCore.Signal(int)
    CENTER_DROP_HOLE = TRAY_CENTER_DROP_HOLE

    def __init__(self, slot_min: int, slot_max: int, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._slot_min = slot_min
        self._slot_max = slot_max
        self._target_hole = slot_min
        self._current_hole: int | None = None
        self._current_xy: tuple[int | None, int | None] = (None, None)
        self._calibrations: dict[int, tuple[int, int]] = {}
        self._hole_points: dict[int, QtCore.QPointF] = {}
        self._dragging = False
        self._logical_positions = tray_logical_hole_positions(slot_min, slot_max)
        self._min_x, self._max_x, self._min_y, self._max_y = tray_logical_bounds(slot_min, slot_max)
        self.setMinimumSize(520, 520)
        self.setMouseTracking(True)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return width

    def sizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(760, 760)

    def set_target_hole(self, hole: int) -> None:
        self._target_hole = hole
        self.update()

    def set_current_hole(self, hole: int | None) -> None:
        self._current_hole = hole
        self.update()

    def set_calibrations(self, calibrations: dict[int, tuple[int, int]]) -> None:
        self._calibrations = dict(calibrations)
        self.update()

    def set_current_xy(self, x_pos: int | None, y_pos: int | None) -> None:
        self._current_xy = (x_pos, y_pos)
        self.update()

    def _stage_rect(self) -> QtCore.QRectF:
        margin = 8.0
        side = max(120.0, min(self.width(), self.height()) - 2 * margin)
        return QtCore.QRectF((self.width() - side) * 0.5, (self.height() - side) * 0.5, side, side)

    def _layout_rect(self) -> tuple[QtCore.QRectF, float]:
        stage = self._stage_rect()
        usable = QtCore.QRectF(
            stage.left() + stage.width() * 0.06,
            stage.top() + stage.height() * 0.10,
            stage.width() * 0.88,
            stage.height() * 0.80,
        )
        width_units = self._max_x - self._min_x
        height_units = self._max_y - self._min_y
        scale = min(usable.width() / width_units, usable.height() / height_units)
        actual = QtCore.QRectF(0, 0, width_units * scale, height_units * scale)
        actual.moveCenter(usable.center())
        return actual, scale

    def _deck_rect(self) -> QtCore.QRectF:
        rect, scale = self._layout_rect()
        return rect.adjusted(-0.92 * scale, -0.92 * scale, 0.92 * scale, 0.92 * scale)

    def _cup_radius(self) -> float:
        _, scale = self._layout_rect()
        return scale * (TRAY_CUP_DIAMETER_IN / TRAY_CUP_PITCH_IN) * 0.5

    def _drop_radius(self) -> float:
        _, scale = self._layout_rect()
        return scale * (TRAY_CENTER_DROP_DIAMETER_IN / TRAY_CUP_PITCH_IN) * 0.5

    def _logical_to_scene(self, point: QtCore.QPointF) -> QtCore.QPointF:
        rect, scale = self._layout_rect()
        x = rect.left() + (point.x() - self._min_x) * scale
        y = rect.top() + (point.y() - self._min_y) * scale
        return QtCore.QPointF(x, y)

    def _hole_centers(self) -> dict[int, QtCore.QPointF]:
        if self._hole_points:
            return self._hole_points
        self._hole_points = {hole: self._logical_to_scene(point) for hole, point in self._logical_positions.items()}
        return self._hole_points

    def _drop_center(self) -> QtCore.QPointF:
        return self._hole_centers()[self.CENTER_DROP_HOLE]

    def _crosshair_color(self) -> QtGui.QColor:
        return QtGui.QColor("#f7f3d7")

    def _cup_label_rect(self, point: QtCore.QPointF, radius: float) -> QtCore.QRectF:
        return QtCore.QRectF(point.x() - radius, point.y() - radius * 0.72, radius * 2, radius * 1.44)

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        self._hole_points.clear()
        super().resizeEvent(event)

    def _nearest_hole(self, pos: QtCore.QPointF) -> int:
        centers = self._hole_centers()
        best_hole = self._slot_min
        best_dist = float("inf")
        for hole, point in centers.items():
            dist = (point.x() - pos.x()) ** 2 + (point.y() - pos.y()) ** 2
            if dist < best_dist:
                best_dist = dist
                best_hole = hole
        return best_hole

    def _current_marker_point(self) -> QtCore.QPointF | None:
        x_pos, y_pos = self._current_xy
        if x_pos is None or y_pos is None:
            return None
        if self._current_hole is not None and self._current_hole in self._hole_centers():
            return self._hole_centers()[self._current_hole]
        if len(self._calibrations) < 2:
            return None
        points = self._hole_centers()
        calibrated_points = [(points[hole], xy) for hole, xy in self._calibrations.items() if hole in points]
        if len(calibrated_points) < 2:
            return None
        scene_xs = [point.x() for point, _ in calibrated_points]
        scene_ys = [point.y() for point, _ in calibrated_points]
        motor_xs = [xy[0] for _, xy in calibrated_points]
        motor_ys = [xy[1] for _, xy in calibrated_points]
        if max(motor_xs) == min(motor_xs) or max(motor_ys) == min(motor_ys):
            return None
        x_ratio = (x_pos - min(motor_xs)) / float(max(motor_xs) - min(motor_xs))
        y_ratio = (y_pos - min(motor_ys)) / float(max(motor_ys) - min(motor_ys))
        x_ratio = max(0.0, min(1.0, x_ratio))
        y_ratio = max(0.0, min(1.0, y_ratio))
        return QtCore.QPointF(
            min(scene_xs) + x_ratio * (max(scene_xs) - min(scene_xs)),
            min(scene_ys) + y_ratio * (max(scene_ys) - min(scene_ys)),
        )

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() != QtCore.Qt.LeftButton:
            return
        self._dragging = True
        self.holeSelected.emit(self._nearest_hole(event.position()))

    def mouseMoveEvent(self, event: QtGui.QMouseEvent) -> None:
        if not self._dragging:
            return
        self.holeDragged.emit(self._nearest_hole(event.position()))

    def mouseDoubleClickEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() != QtCore.Qt.LeftButton:
            return
        self._dragging = False
        hole = self._nearest_hole(event.position())
        self.holeSelected.emit(hole)
        self.holeActivated.emit(hole)
        event.accept()

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        self._dragging = False
        super().mouseReleaseEvent(event)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        del event
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        centers = self._hole_centers()
        cup_radius = self._cup_radius()
        drop_center = self._drop_center()
        drop_radius = self._drop_radius()
        font = painter.font()
        font.setPointSize(9)
        painter.setFont(font)
        for hole, point in centers.items():
            tick_rect = QtCore.QRectF(point.x() - cup_radius * 0.10, point.y() - cup_radius * 1.58, cup_radius * 0.20, cup_radius * 0.34)
            painter.setPen(QtCore.Qt.NoPen)
            painter.setBrush(QtGui.QColor("#1f1f1f"))
            painter.drawRoundedRect(tick_rect, 1.5, 1.5)

            if hole == self.CENTER_DROP_HOLE:
                continue
            engraved = QtGui.QRadialGradient(point, cup_radius * 1.17)
            engraved.setColorAt(0.0, QtGui.QColor(244, 236, 229, 72))
            engraved.setColorAt(0.75, QtGui.QColor(210, 195, 181, 118))
            engraved.setColorAt(1.0, QtGui.QColor(154, 138, 126, 126))
            painter.setPen(QtGui.QPen(QtGui.QColor(250, 246, 242, 148), 1.08))
            painter.setBrush(engraved)
            painter.drawEllipse(point, cup_radius, cup_radius)

            inner = QtGui.QRadialGradient(point, cup_radius * 0.80)
            inner.setColorAt(0.0, QtGui.QColor(255, 255, 255, 16))
            inner.setColorAt(1.0, QtGui.QColor(112, 96, 86, 30))
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 251, 247, 82), 0.8))
            painter.setBrush(inner)
            painter.drawEllipse(point, cup_radius * 0.82, cup_radius * 0.82)

            if hole in self._calibrations:
                painter.setPen(QtGui.QPen(QtGui.QColor("#2ca58d"), 2.0))
                painter.setBrush(QtCore.Qt.NoBrush)
                painter.drawEllipse(point, cup_radius + 3, cup_radius + 3)
            if hole == self._target_hole:
                painter.setPen(QtGui.QPen(QtGui.QColor("#ff9f43"), 2.8))
                painter.drawEllipse(point, cup_radius + 7, cup_radius + 7)
            if hole == self._current_hole:
                painter.setPen(QtGui.QPen(self._crosshair_color(), 2.4))
                painter.drawEllipse(point, cup_radius + 11, cup_radius + 11)

            painter.setPen(QtGui.QColor(82, 74, 66, 180))
            painter.drawText(self._cup_label_rect(point, cup_radius * 0.82), QtCore.Qt.AlignCenter, str(hole))

        drop_shadow = QtGui.QRadialGradient(drop_center, drop_radius * 1.22)
        drop_shadow.setColorAt(0.0, QtGui.QColor(74, 54, 34, 118))
        drop_shadow.setColorAt(0.65, QtGui.QColor(88, 66, 48, 78))
        drop_shadow.setColorAt(1.0, QtGui.QColor(90, 68, 50, 0))
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(drop_shadow)
        painter.drawEllipse(drop_center, drop_radius * 1.22, drop_radius * 1.22)

        painter.setPen(QtGui.QPen(QtGui.QColor("#fff0d8"), 1.9))
        painter.setBrush(QtGui.QColor("#6b5138"))
        painter.drawEllipse(drop_center, drop_radius, drop_radius)
        painter.setPen(QtGui.QPen(QtGui.QColor("#f4dec2"), 1.05))
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.drawEllipse(drop_center, drop_radius + 4, drop_radius + 4)
        painter.setPen(QtGui.QColor(110, 84, 60, 170))
        painter.drawText(self._cup_label_rect(drop_center, cup_radius * 0.82), QtCore.Qt.AlignCenter, str(self.CENTER_DROP_HOLE))

        if self._target_hole == self.CENTER_DROP_HOLE or self._current_hole == self.CENTER_DROP_HOLE:
            ring_color = QtGui.QColor("#ffb25c") if self._target_hole == self.CENTER_DROP_HOLE else self._crosshair_color()
            painter.setPen(QtGui.QPen(ring_color, 3))
            painter.drawEllipse(drop_center, drop_radius + 13, drop_radius + 13)

        marker = self._current_marker_point()
        if marker is not None:
            painter.setPen(QtGui.QPen(self._crosshair_color(), 2))
            painter.setBrush(self._crosshair_color())
            painter.drawEllipse(marker, 5, 5)
            painter.drawLine(marker.x() - 16, marker.y(), marker.x() + 16, marker.y())
            painter.drawLine(marker.x(), marker.y() - 16, marker.x(), marker.y() + 16)
        painter.end()


class MainWindow(QtWidgets.QMainWindow):
    SLOT_MIN = 1
    SLOT_MAX = 100
    POSITION_TOLERANCE = 1000

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Changer XY Control")
        self.controller = ChangerStageController()
        self.motion_defaults = MOTION_DEFAULTS
        self.config = self._load_config()
        self._port_boxes: dict[str, QtWidgets.QComboBox] = {}
        self._port_status: dict[str, QtWidgets.QLabel] = {}
        self._port_toggle_buttons: dict[str, tuple[QtWidgets.QPushButton, QtWidgets.QPushButton]] = {}
        self._last_poll_error = ""
        self._switch_labels: dict[str, QtWidgets.QLabel] = {}
        self._current_stage_hole: int | None = None
        self._current_z_pos: int | None = None
        self._build_ui()
        self._apply_compact_font()
        self._apply_module_style_overrides()
        self.setMinimumSize(1420, 860)
        self.resize(1620, 940)
        self._load_config_into_widgets()
        self._poll_timer = QtCore.QTimer(self)
        self._poll_timer.setInterval(900)
        self._poll_timer.timeout.connect(self._poll_stage_state)
        self._poll_timer.start()
        QtCore.QTimer.singleShot(0, self._poll_stage_state)

    def _load_config(self) -> AppConfig:
        if not CONFIG_PATH.exists():
            return AppConfig()
        try:
            payload = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return AppConfig()
        calibrations = [HoleCalibration(**item) for item in payload.get("calibrations", []) if isinstance(item, dict)]
        return AppConfig(
            x_port=str(payload.get("x_port", "COM4")),
            y_port=str(payload.get("y_port", "COM5")),
            updown_port=str(payload.get("updown_port", "COM6")),
            sensor_port=str(payload.get("sensor_port", "")),
            selected_hole=int(payload.get("selected_hole", 1)),
            drag_moves_enabled=bool(payload.get("drag_moves_enabled", False)),
            calibrations=calibrations,
        )

    def _save_config(self) -> None:
        self.config.x_port = self._port_boxes["x"].currentText().strip()
        self.config.y_port = self._port_boxes["y"].currentText().strip()
        self.config.updown_port = self._port_boxes["updown"].currentText().strip()
        self.config.sensor_port = self._port_boxes["sensor"].currentText().strip()
        self.config.selected_hole = self.target_hole.value()
        self.config.drag_moves_enabled = self.drag_move_check.isChecked()
        self.config.calibrations = [
            HoleCalibration(hole=hole, x=x_pos, y=y_pos)
            for hole, (x_pos, y_pos) in sorted(self._calibration_map().items())
        ]
        CONFIG_PATH.write_text(json.dumps(asdict(self.config), indent=2), encoding="utf-8")

    def _build_card(self, title_text: str, subtitle_text: str | None = None) -> tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        title = QtWidgets.QLabel(title_text)
        title.setObjectName("title")
        layout.addWidget(title)
        if subtitle_text:
            subtitle = QtWidgets.QLabel(subtitle_text)
            subtitle.setObjectName("subtitle")
            subtitle.setWordWrap(True)
            layout.addWidget(subtitle)
        apply_card_shadow(card)
        return card, layout

    def _apply_compact_font(self) -> None:
        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(9.0, compact_size - 0.4))
            self.setFont(compact_font)

    def _apply_module_style_overrides(self) -> None:
        self.setStyleSheet(
            """
            QScrollArea#panelScroll {
                background: transparent;
                border: none;
            }
            QScrollArea#panelScroll > QWidget > QWidget {
                background: transparent;
            }
            QLabel#mapStatus,
            QLabel#renderStatus {
                background: rgba(255, 255, 255, 0.86);
                border: 1px solid rgba(122, 2, 25, 0.18);
                border-radius: 14px;
                padding: 8px 10px;
                color: #4d3a39;
            }
            QLabel#tableHint {
                color: #6a5b54;
                background: rgba(255, 255, 255, 0.78);
                border: 1px solid rgba(122, 2, 25, 0.10);
                border-radius: 12px;
                padding: 7px 9px;
            }
            QWidget#portToggleGroup {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(122, 2, 25, 0.28);
                border-radius: 12px;
            }
            QPushButton#portToggle {
                min-width: 32px;
                padding: 4px 6px;
                border: 1px solid transparent;
                border-radius: 9px;
                background: transparent;
                color: #5d4d49;
            }
            QPushButton#portToggle:checked {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7A0219, stop:1 #5a0013);
                color: #ffffff;
                border: 1px solid rgba(122, 2, 25, 0.82);
            }
            QWidget#renderViewport {
                background: #e8e0d3;
                border-radius: 18px;
            }
            QTableWidget#cupTable {
                font-size: 11px;
            }
            """
        )

    def _stage_limit_instructions(self) -> str:
        return (
            "VB6 limits: XY jog velocity 100000-"
            f"{self.motion_defaults.changer_speed} counts/s; XY homing sweeps "
            f"{self.motion_defaults.xy_neg_homing_distance} to {self.motion_defaults.xy_pos_homing_distance} and requires Z top first. "
            "Z top uses bit 4, X uses bits 4/5, Y uses bits 5/6. Z jog velocity 100000-"
            f"{self.motion_defaults.lift_speed_fast} counts/s; VB6 references are slow={self.motion_defaults.lift_speed_slow}, "
            f"normal={self.motion_defaults.lift_speed_normal}, fast={self.motion_defaults.lift_speed_fast}."
        )

    def _set_port_toggle_state(self, key: str, connected: bool) -> None:
        buttons = self._port_toggle_buttons.get(key)
        if buttons is None:
            return
        on_button, off_button = buttons
        on_button.setChecked(connected)
        off_button.setChecked(not connected)

    def _build_port_row(self, layout: QtWidgets.QGridLayout, row: int, key: str, label_text: str) -> None:
        label = QtWidgets.QLabel(label_text)
        combo = QtWidgets.QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        combo.setMinimumWidth(76)
        combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        status = QtWidgets.QLabel("Off")
        status.setAlignment(QtCore.Qt.AlignCenter)
        status.setMinimumWidth(50)

        toggle_group = QtWidgets.QWidget()
        toggle_group.setObjectName("portToggleGroup")
        toggle_layout = QtWidgets.QHBoxLayout(toggle_group)
        toggle_layout.setContentsMargins(3, 3, 3, 3)
        toggle_layout.setSpacing(2)

        on_button = QtWidgets.QPushButton("On")
        on_button.setObjectName("portToggle")
        on_button.setCheckable(True)
        on_button.setAutoExclusive(True)
        off_button = QtWidgets.QPushButton("Off")
        off_button.setObjectName("portToggle")
        off_button.setCheckable(True)
        off_button.setAutoExclusive(True)
        off_button.setChecked(True)
        on_button.clicked.connect(lambda: self._connect_port(key))
        off_button.clicked.connect(lambda: self._disconnect_port(key))
        toggle_layout.addWidget(on_button)
        toggle_layout.addWidget(off_button)

        layout.addWidget(label, row, 0)
        layout.addWidget(combo, row, 1)
        layout.addWidget(toggle_group, row, 2)
        layout.addWidget(status, row, 3)
        self._port_boxes[key] = combo
        self._port_status[key] = status
        self._port_toggle_buttons[key] = (on_button, off_button)

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        shell = QtWidgets.QHBoxLayout(root)
        shell.setContentsMargins(12, 12, 12, 12)
        shell.setSpacing(10)

        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setObjectName("panelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        left_scroll.setMinimumWidth(320)
        left_scroll.setMaximumWidth(370)
        left_host = QtWidgets.QWidget()
        left_scroll.setWidget(left_host)
        left = QtWidgets.QVBoxLayout(left_host)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(10)

        center_host = QtWidgets.QWidget()
        center_host.setMinimumWidth(650)
        center_layout = QtWidgets.QVBoxLayout(center_host)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(12)

        right_scroll = QtWidgets.QScrollArea()
        right_scroll.setObjectName("panelScroll")
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        right_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        right_scroll.setMinimumWidth(360)
        right_scroll.setMaximumWidth(420)
        right_host = QtWidgets.QWidget()
        right_scroll.setWidget(right_host)
        right = QtWidgets.QVBoxLayout(right_host)
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(10)

        connections_card, connections_layout = self._build_card(
            "Connections",
            "Expose X, Y, Z, and reference-monitor serial links separately. Matching COM ports can still be reused.",
        )
        toolbar = QtWidgets.QHBoxLayout()
        self.refresh_ports_btn = QtWidgets.QPushButton("Refresh Ports")
        self.read_status_btn = QtWidgets.QPushButton("Read Switches")
        toolbar.addWidget(self.refresh_ports_btn)
        toolbar.addWidget(self.read_status_btn)
        connections_layout.addLayout(toolbar)
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(4)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        self._build_port_row(grid, 0, "x", "X Motor")
        self._build_port_row(grid, 1, "y", "Y Motor")
        self._build_port_row(grid, 2, "updown", "Z Axis")
        self._build_port_row(grid, 3, "sensor", "Reference")
        connections_layout.addLayout(grid)
        switches = QtWidgets.QGridLayout()
        switch_rows = (
            ("z_top", "Z TOP"),
            ("x_neg", "X NEG"),
            ("x_pos", "X POS"),
            ("y_neg", "Y NEG"),
            ("y_pos", "Y POS"),
        )
        for row, (key, title) in enumerate(switch_rows):
            label = QtWidgets.QLabel("Unknown")
            label.setMinimumWidth(72)
            self._switch_labels[key] = label
            switches.addWidget(QtWidgets.QLabel(title), row, 0)
            switches.addWidget(label, row, 1)
        connections_layout.addLayout(switches)

        motion_card, motion_layout = self._build_card(
            "Stage Controls",
            "Home Z first if the quartz rod is not clear, then move XY.",
        )
        metrics = QtWidgets.QGridLayout()
        self.x_pos_label = QtWidgets.QLabel("--")
        self.y_pos_label = QtWidgets.QLabel("--")
        self.z_pos_label = QtWidgets.QLabel("--")
        self.current_hole_label = QtWidgets.QLabel("Unknown")
        metrics.addWidget(QtWidgets.QLabel("X Position"), 0, 0)
        metrics.addWidget(self.x_pos_label, 0, 1)
        metrics.addWidget(QtWidgets.QLabel("Y Position"), 1, 0)
        metrics.addWidget(self.y_pos_label, 1, 1)
        metrics.addWidget(QtWidgets.QLabel("Z Position"), 2, 0)
        metrics.addWidget(self.z_pos_label, 2, 1)
        metrics.addWidget(QtWidgets.QLabel("Nearest Cup"), 3, 0)
        metrics.addWidget(self.current_hole_label, 3, 1)
        motion_layout.addLayout(metrics)

        selection = QtWidgets.QFormLayout()
        self.target_hole = QtWidgets.QSpinBox()
        self.target_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        self.calibration_hole = QtWidgets.QSpinBox()
        self.calibration_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        selection.addRow("Target Cup", self.target_hole)
        selection.addRow("Calibration Cup", self.calibration_hole)
        motion_layout.addLayout(selection)

        commands = QtWidgets.QGridLayout()
        self.goto_btn = QtWidgets.QPushButton("Move To Selected Cup")
        self.goto_btn.setObjectName("accent")
        self.goto_btn.setText("Move To Cup")
        self.home_top_btn = QtWidgets.QPushButton("Home Z")
        self.home_center_btn = QtWidgets.QPushButton("Home XY")
        self.corner_btn = QtWidgets.QPushButton("Corner")
        self.refresh_live_btn = QtWidgets.QPushButton("Refresh")
        commands.addWidget(self.goto_btn, 0, 0, 1, 2)
        commands.addWidget(self.home_top_btn, 1, 0)
        commands.addWidget(self.home_center_btn, 1, 1)
        commands.addWidget(self.corner_btn, 2, 0)
        commands.addWidget(self.refresh_live_btn, 2, 1)
        motion_layout.addLayout(commands)

        left.addWidget(motion_card)

        cup_card, cup_layout = self._build_card(
            "Cup Position Sheet",
            "Compact editable X/Y count table inspired by the VB6 settings grid. Edit counts directly; blank both X and Y cells to clear a cup.",
        )
        self.cup_table_summary = QtWidgets.QLabel("0 cups stored. Capture XY or edit counts directly.")
        self.cup_table_summary.setObjectName("tableHint")
        self.cup_table_summary.setWordWrap(True)
        cup_layout.addWidget(self.cup_table_summary)
        self.cup_table = QtWidgets.QTableWidget(self.SLOT_MAX - self.SLOT_MIN + 1, 3)
        self.cup_table.setObjectName("cupTable")
        self.cup_table.setHorizontalHeaderLabels(["Cup", "X counts", "Y counts"])
        self.cup_table.verticalHeader().setVisible(False)
        self.cup_table.setAlternatingRowColors(True)
        self.cup_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.cup_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.cup_table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        self.cup_table.setMinimumHeight(248)
        self.cup_table.setMaximumHeight(312)
        header = self.cup_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        cup_layout.addWidget(self.cup_table)
        left.addWidget(cup_card)
        left.addStretch(1)

        stage_card, stage_layout = self._build_card(
            "Stage Model",
            "Stage-first tray view matched to the photographed acrylic plate. Inset grip slots and center drop-off hole are drawn to match the physical holder.",
        )
        stage_help_text = (
            "Click a cup to choose the next target. Double-click a cup to run the same move as the Move To Cup button. "
            "Drag across cups only when you want to scrub through selections. Use the side panels for connections, motion tuning, jogs, and calibration. "
            + self._stage_limit_instructions()
        )

        self.stage_scene = StageScene(self.SLOT_MIN, self.SLOT_MAX)
        self.stage_scene.setToolTip(stage_help_text)
        self.stage_status = QtWidgets.QLabel("Click a cup to select it. Double-click a cup to move there.")
        self.stage_status.setObjectName("mapStatus")
        self.stage_status.setWordWrap(True)
        self.stage_status.setMinimumHeight(50)
        self.stage_status.setToolTip(stage_help_text)
        stage_layout.addWidget(self.stage_scene, stretch=1)
        stage_layout.addWidget(self.stage_status)
        center_layout.addWidget(stage_card, stretch=1)

        console_card, console_layout = self._build_card("Console")
        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        self.console.setMinimumHeight(96)
        self.console.setMaximumHeight(132)
        console_layout.addWidget(self.console)
        center_layout.addWidget(console_card)

        tuning_card, tuning_layout = self._build_card(
            "Velocity And Jog",
            "Keep motion tuning near the stage while leaving the tray itself visually dominant.",
        )
        self.model_help_btn = QtWidgets.QToolButton()
        self.model_help_btn.setText("?")
        self.model_help_btn.setAutoRaise(True)
        self.model_help_btn.setToolTip(stage_help_text)
        self.model_xy_speed = QtWidgets.QSpinBox()
        self.model_xy_speed.setRange(100_000, self.motion_defaults.changer_speed)
        self.model_xy_speed.setSingleStep(100_000)
        self.model_xy_speed.setValue(self.motion_defaults.changer_speed)
        self.model_xy_speed.setToolTip(
            "XY velocity in counts per second. VB6/Python ChangerSpeed default is "
            f"{self.motion_defaults.changer_speed}. Valid range: 100000-{self.motion_defaults.changer_speed}."
        )
        self.model_xy_speed.setWhatsThis(self.model_xy_speed.toolTip())
        self.model_z_speed = QtWidgets.QSpinBox()
        self.model_z_speed.setRange(100_000, self.motion_defaults.lift_speed_fast)
        self.model_z_speed.setSingleStep(100_000)
        self.model_z_speed.setValue(self.motion_defaults.lift_speed_slow)
        self.model_z_speed.setToolTip(
            "Z velocity in counts per second. VB6 lift references are slow="
            f"{self.motion_defaults.lift_speed_slow}, normal={self.motion_defaults.lift_speed_normal}, "
            f"fast={self.motion_defaults.lift_speed_fast}. Valid range: 100000-{self.motion_defaults.lift_speed_fast}."
        )
        self.model_z_speed.setWhatsThis(self.model_z_speed.toolTip())
        self.model_step = QtWidgets.QSpinBox()
        self.model_step.setRange(100, 2_000_000)
        self.model_step.setSingleStep(100)
        self.model_step.setValue(25_000)
        self.model_step.setToolTip(
            "Relative jog distance in counts. Use smaller values near cups and switches."
        )
        self.model_x_minus_btn = QtWidgets.QPushButton("X-")
        self.model_x_plus_btn = QtWidgets.QPushButton("X+")
        self.model_y_minus_btn = QtWidgets.QPushButton("Y-")
        self.model_y_plus_btn = QtWidgets.QPushButton("Y+")
        self.model_z_minus_btn = QtWidgets.QPushButton("Z-")
        self.model_z_plus_btn = QtWidgets.QPushButton("Z+")
        self.model_home_xy_btn = QtWidgets.QPushButton("Home XY")
        self.model_home_z_btn = QtWidgets.QPushButton("Home Z")
        tuning_grid = QtWidgets.QGridLayout()
        tuning_grid.addWidget(QtWidgets.QLabel("XY velocity"), 0, 0)
        tuning_grid.addWidget(self.model_xy_speed, 0, 1)
        tuning_grid.addWidget(self.model_help_btn, 0, 2)
        tuning_grid.addWidget(QtWidgets.QLabel("Z velocity"), 1, 0)
        tuning_grid.addWidget(self.model_z_speed, 1, 1)
        tuning_grid.addWidget(QtWidgets.QLabel("Jog step"), 2, 0)
        tuning_grid.addWidget(self.model_step, 2, 1)
        tuning_grid.addWidget(self.model_x_minus_btn, 3, 0)
        tuning_grid.addWidget(self.model_x_plus_btn, 3, 1)
        tuning_grid.addWidget(self.model_y_minus_btn, 4, 0)
        tuning_grid.addWidget(self.model_y_plus_btn, 4, 1)
        tuning_grid.addWidget(self.model_z_minus_btn, 5, 0)
        tuning_grid.addWidget(self.model_z_plus_btn, 5, 1)
        tuning_grid.addWidget(self.model_home_xy_btn, 6, 0)
        tuning_grid.addWidget(self.model_home_z_btn, 6, 1)
        tuning_grid.setColumnStretch(0, 1)
        tuning_grid.setColumnStretch(1, 1)
        tuning_layout.addLayout(tuning_grid)
        right.addWidget(tuning_card)

        calibration_card, calibration_layout = self._build_card(
            "Calibration",
            "Capture the current X/Y counts for each cup to build the tray lookup table.",
        )
        self.capture_btn = QtWidgets.QPushButton("Capture XY")
        self.capture_btn.setObjectName("accent")
        self.clear_calibration_btn = QtWidgets.QPushButton("Clear Cup")
        self.drag_move_check = QtWidgets.QCheckBox("Enable drag selection to also move stage")
        self.drag_move_check.setChecked(False)
        self.calibration_summary = QtWidgets.QLabel("0 calibrated cups")
        self.calibration_summary.setWordWrap(True)
        calibration_layout.addWidget(self.capture_btn)
        calibration_layout.addWidget(self.clear_calibration_btn)
        calibration_layout.addWidget(self.drag_move_check)
        calibration_layout.addWidget(self.calibration_summary)
        right.addWidget(calibration_card)

        right.addWidget(connections_card)
        right.addStretch(1)

        shell.addWidget(left_scroll)
        shell.addWidget(center_host, stretch=1)
        shell.addWidget(right_scroll)

        self.refresh_ports_btn.clicked.connect(self._refresh_port_boxes)
        self.read_status_btn.clicked.connect(self._poll_stage_state)
        self.goto_btn.clicked.connect(self._move_to_selected_hole)
        self.home_top_btn.clicked.connect(self._home_top)
        self.home_center_btn.clicked.connect(self._home_xy_center)
        self.corner_btn.clicked.connect(self._move_xy_corner)
        self.refresh_live_btn.clicked.connect(self._poll_stage_state)
        self.capture_btn.clicked.connect(self._capture_calibration)
        self.clear_calibration_btn.clicked.connect(self._clear_calibration)
        self.target_hole.valueChanged.connect(self._sync_selected_hole)
        self.calibration_hole.valueChanged.connect(self._sync_calibration_hole)
        self.drag_move_check.toggled.connect(self._save_config)
        self.cup_table.itemChanged.connect(self._on_cup_table_item_changed)
        self.cup_table.currentCellChanged.connect(self._on_cup_table_current_cell_changed)
        self.stage_scene.holeSelected.connect(self._on_scene_hole_selected)
        self.stage_scene.holeDragged.connect(self._on_scene_hole_dragged)
        self.stage_scene.holeActivated.connect(self._on_scene_hole_activated)
        self.model_x_minus_btn.clicked.connect(lambda: self._jog_axis("x", -self.model_step.value(), self._model_speed("x")))
        self.model_x_plus_btn.clicked.connect(lambda: self._jog_axis("x", self.model_step.value(), self._model_speed("x")))
        self.model_y_minus_btn.clicked.connect(lambda: self._jog_axis("y", -self.model_step.value(), self._model_speed("y")))
        self.model_y_plus_btn.clicked.connect(lambda: self._jog_axis("y", self.model_step.value(), self._model_speed("y")))
        self.model_z_minus_btn.clicked.connect(lambda: self._jog_axis("updown", -self.model_step.value(), self._model_speed("updown")))
        self.model_z_plus_btn.clicked.connect(lambda: self._jog_axis("updown", self.model_step.value(), self._model_speed("updown")))
        self.model_home_xy_btn.clicked.connect(self._home_xy_center)
        self.model_home_z_btn.clicked.connect(self._home_top)

    def _available_ports(self) -> list[str]:
        ports = sorted(port.device for port in list_ports.comports())
        return ports or [""]

    def _set_combo_value(self, combo: QtWidgets.QComboBox, value: str) -> None:
        current = combo.currentText()
        items = [combo.itemText(index) for index in range(combo.count())]
        if value and value not in items:
            combo.addItem(value)
        combo.setCurrentText(value or current)

    def _refresh_port_boxes(self) -> None:
        ports = self._available_ports()
        for combo in self._port_boxes.values():
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(ports)
            combo.setCurrentText(current)
            combo.blockSignals(False)

    def _load_config_into_widgets(self) -> None:
        self._refresh_port_boxes()
        self._set_combo_value(self._port_boxes["x"], self.config.x_port)
        self._set_combo_value(self._port_boxes["y"], self.config.y_port)
        self._set_combo_value(self._port_boxes["updown"], self.config.updown_port)
        self._set_combo_value(self._port_boxes["sensor"], self.config.sensor_port)
        self.target_hole.setValue(self.config.selected_hole)
        self.calibration_hole.setValue(self.config.selected_hole)
        self.drag_move_check.setChecked(self.config.drag_moves_enabled)
        self._update_stage_calibrations()
        self.stage_scene.set_target_hole(self.config.selected_hole)
        self._sync_cup_table_selection(self.config.selected_hole)

    def _model_speed(self, axis: str) -> int:
        if axis == "updown":
            return int(self.model_z_speed.value())
        return int(self.model_xy_speed.value())

    def _append(self, message: str) -> None:
        self.console.appendPlainText(message)

    def _connect_port(self, key: str) -> None:
        port = self._port_boxes[key].currentText().strip()
        try:
            if key == "sensor":
                self.controller.connect_sensor(port)
            else:
                self.controller.connect_axis(key, port)
        except Exception as exc:
            self._set_port_toggle_state(key, False)
            QtWidgets.QMessageBox.warning(self, "Connection Error", str(exc))
            return
        self._port_status[key].setText(f"On {port}")
        self._set_port_toggle_state(key, True)
        self._append(f"Connected {key.upper()} on {port}.")
        self._save_config()
        self._poll_stage_state()

    def _disconnect_port(self, key: str) -> None:
        if key == "sensor":
            self.controller.disconnect_sensor()
        else:
            self.controller.disconnect_axis(key)
        self._port_status[key].setText("Off")
        self._set_port_toggle_state(key, False)
        self._append(f"Disconnected {key.upper()}.")
        self._save_config()
        self._poll_stage_state()

    def _sync_selected_hole(self, value: int) -> None:
        if self.calibration_hole.value() != value:
            self.calibration_hole.blockSignals(True)
            self.calibration_hole.setValue(value)
            self.calibration_hole.blockSignals(False)
        self.stage_scene.set_target_hole(value)
        self._sync_cup_table_selection(value)
        self._save_config()

    def _sync_calibration_hole(self, value: int) -> None:
        if self.target_hole.value() != value:
            self.target_hole.blockSignals(True)
            self.target_hole.setValue(value)
            self.target_hole.blockSignals(False)
        self.stage_scene.set_target_hole(value)
        self._sync_cup_table_selection(value)
        self._save_config()

    def _calibration_map(self) -> dict[int, tuple[int, int]]:
        mapping: dict[int, tuple[int, int]] = {}
        for item in self.config.calibrations:
            if item.x is not None and item.y is not None:
                mapping[item.hole] = (int(item.x), int(item.y))
        return mapping

    def _set_calibration(self, hole: int, x_pos: int, y_pos: int) -> None:
        updated = False
        for item in self.config.calibrations:
            if item.hole == hole:
                item.x = x_pos
                item.y = y_pos
                updated = True
                break
        if not updated:
            self.config.calibrations.append(HoleCalibration(hole=hole, x=x_pos, y=y_pos))
        self._update_stage_calibrations()
        self._save_config()

    def _clear_calibration(self) -> None:
        hole = self.calibration_hole.value()
        self.config.calibrations = [item for item in self.config.calibrations if item.hole != hole]
        self._append(f"Cleared cup {hole} calibration.")
        self._update_stage_calibrations()
        self._save_config()

    def _update_stage_calibrations(self) -> None:
        mapping = self._calibration_map()
        self.stage_scene.set_calibrations(mapping)
        self.calibration_summary.setText(f"{len(mapping)} calibrated cups")
        self.cup_table_summary.setText(
            f"{len(mapping)} cups stored. Double-click X or Y to edit counts; blank both cells to clear a cup."
        )
        self._refresh_cup_table()

    def _refresh_cup_table(self) -> None:
        blocker = QtCore.QSignalBlocker(self.cup_table)
        del blocker
        mapping = self._calibration_map()
        self.cup_table.blockSignals(True)
        try:
            for row, hole in enumerate(range(self.SLOT_MIN, self.SLOT_MAX + 1)):
                cup_item = self.cup_table.item(row, 0)
                if cup_item is None:
                    cup_item = QtWidgets.QTableWidgetItem()
                    cup_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    cup_item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.cup_table.setItem(row, 0, cup_item)
                cup_item.setText(str(hole))

                x_item = self.cup_table.item(row, 1)
                if x_item is None:
                    x_item = QtWidgets.QTableWidgetItem()
                    x_item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                    self.cup_table.setItem(row, 1, x_item)

                y_item = self.cup_table.item(row, 2)
                if y_item is None:
                    y_item = QtWidgets.QTableWidgetItem()
                    y_item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                    self.cup_table.setItem(row, 2, y_item)

                if hole in mapping:
                    x_val, y_val = mapping[hole]
                    x_item.setText(str(x_val))
                    y_item.setText(str(y_val))
                else:
                    x_item.setText("")
                    y_item.setText("")
            self._sync_cup_table_selection(self.target_hole.value())
        finally:
            self.cup_table.blockSignals(False)

    def _sync_cup_table_selection(self, hole: int) -> None:
        if not hasattr(self, "cup_table"):
            return
        row = max(0, min(self.SLOT_MAX - self.SLOT_MIN, hole - self.SLOT_MIN))
        self.cup_table.blockSignals(True)
        try:
            self.cup_table.setCurrentCell(row, 0)
            self.cup_table.selectRow(row)
        finally:
            self.cup_table.blockSignals(False)

    def _on_cup_table_current_cell_changed(self, current_row: int, current_column: int, previous_row: int, previous_column: int) -> None:
        del current_column, previous_row, previous_column
        if current_row < 0:
            return
        hole = self.SLOT_MIN + current_row
        if self.target_hole.value() != hole:
            self.target_hole.setValue(hole)
        elif self.calibration_hole.value() != hole:
            self.calibration_hole.setValue(hole)

    def _on_cup_table_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        if item.column() not in (1, 2):
            return
        row = item.row()
        hole = self.SLOT_MIN + row
        x_item = self.cup_table.item(row, 1)
        y_item = self.cup_table.item(row, 2)
        x_text = "" if x_item is None else x_item.text().strip()
        y_text = "" if y_item is None else y_item.text().strip()

        if not x_text and not y_text:
            self.config.calibrations = [entry for entry in self.config.calibrations if entry.hole != hole]
            self.stage_status.setText(f"Cleared cup {hole} from the cup-position sheet.")
            self._update_stage_calibrations()
            self._save_config()
            return

        if not x_text or not y_text:
            self.stage_status.setText(f"Cup {hole} needs both X and Y counts before it can be saved.")
            return

        try:
            x_pos = int(x_text)
            y_pos = int(y_text)
        except ValueError:
            self.stage_status.setText(f"Cup {hole} counts must be whole integers.")
            self._refresh_cup_table()
            return

        self._set_calibration(hole, x_pos, y_pos)
        self.stage_status.setText(f"Saved cup {hole} counts from the cup-position sheet.")

    def _nearest_hole_from_xy(self, x_pos: int | None, y_pos: int | None) -> int | None:
        if x_pos is None or y_pos is None:
            return None
        for hole, (hole_x, hole_y) in self._calibration_map().items():
            if abs(x_pos - hole_x) < self.POSITION_TOLERANCE and abs(y_pos - hole_y) < self.POSITION_TOLERANCE:
                return hole
        return None

    def _poll_stage_state(self) -> None:
        for key in ("x", "y", "updown"):
            connected = self.controller.axis_connected(key)
            self._set_port_toggle_state(key, connected)
            self._port_status[key].setText(f"On {self.controller.axis_port(key)}" if connected else "Off")
        self._set_port_toggle_state("sensor", self.controller.sensor_connected)
        self._port_status["sensor"].setText(f"On {self.controller.sensor_port}" if self.controller.sensor_connected else "Off")
        try:
            x_pos, y_pos = self.controller.read_xy_position()
            switches = self.controller.reference_switch_states()
            z_pos = self.controller.read_axis_position("updown") if self.controller.axis_connected("updown") else None
        except Exception as exc:
            x_pos, y_pos = None, None
            z_pos = None
            switches = {"z_top": None, "x_neg": None, "x_pos": None, "y_neg": None, "y_pos": None}
            message = str(exc)
            if message and message != self._last_poll_error:
                self._last_poll_error = message
                self._append(f"Polling warning: {message}")
        else:
            self._last_poll_error = ""

        self.x_pos_label.setText("--" if x_pos is None else str(x_pos))
        self.y_pos_label.setText("--" if y_pos is None else str(y_pos))
        self.z_pos_label.setText("--" if z_pos is None else str(z_pos))
        current_hole = self._nearest_hole_from_xy(x_pos, y_pos)
        self.current_hole_label.setText("Unknown" if current_hole is None else str(current_hole))
        self.stage_scene.set_current_hole(current_hole)
        self.stage_scene.set_current_xy(x_pos, y_pos)
        self._current_stage_hole = current_hole
        self._current_z_pos = z_pos

        for key, label in self._switch_labels.items():
            state = switches.get(key)
            if state is None:
                label.setText("Unknown")
            else:
                label.setText("TRIPPED" if state else "Released")

    def _home_top(self) -> None:
        try:
            result = self.controller.home_to_top()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Home Z Error", str(exc))
            return
        self._append(f"HomeToTop complete: z={result.final_position}, success={result.success}.")
        self._poll_stage_state()

    def _home_xy_center(self) -> None:
        try:
            x_res, y_res = self.controller.home_xy_to_center()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Home XY Error", str(exc))
            return
        self._append(
            "HomeToCenter complete: "
            f"x={x_res.final_position}, y={y_res.final_position}, success={x_res.success and y_res.success}."
        )
        self._poll_stage_state()

    def _move_xy_corner(self) -> None:
        try:
            x_res, y_res = self.controller.move_xy_to_corner()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Corner Move Error", str(exc))
            return
        self._append(f"MoveToCorner complete: x={x_res.final_position}, y={y_res.final_position}.")
        self._poll_stage_state()

    def _move_to_selected_hole(self) -> None:
        hole = self.target_hole.value()
        calibration = self._calibration_map().get(hole)
        if calibration is None:
            QtWidgets.QMessageBox.warning(
                self,
                "Uncalibrated Cup",
                f"Cup {hole} does not have a stored X/Y calibration yet. Capture the current position first.",
            )
            return
        try:
            x_res, y_res = self.controller.move_xy_absolute(calibration[0], calibration[1])
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Move Error", str(exc))
            return
        self._append(
            f"Moved to cup {hole}: x_target={x_res.target}, x_final={x_res.final_position}, "
            f"y_target={y_res.target}, y_final={y_res.final_position}."
        )
        self.stage_status.setText(f"Cup {hole} is the active target.")
        self._poll_stage_state()

    def _capture_calibration(self) -> None:
        try:
            x_pos, y_pos = self.controller.read_xy_position()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Read Position Error", str(exc))
            return
        if x_pos is None or y_pos is None:
            QtWidgets.QMessageBox.warning(self, "Read Position Error", "Connect X and Y before capturing calibration.")
            return
        hole = self.calibration_hole.value()
        self._set_calibration(hole, x_pos, y_pos)
        self._append(f"Captured cup {hole} calibration: x={x_pos}, y={y_pos}.")
        self.stage_status.setText(f"Cup {hole} calibration updated from the live stage position.")
        self._poll_stage_state()

    def _select_scene_hole(self, hole: int) -> str:
        self.target_hole.setValue(hole)
        self.calibration_hole.setValue(hole)
        return "measurement drop-off" if hole == StageScene.CENTER_DROP_HOLE else f"cup {hole}"

    def _on_scene_hole_selected(self, hole: int) -> None:
        label = self._select_scene_hole(hole)
        self.stage_status.setText(f"Selected {label}. Double-click to move there.")

    def _on_scene_hole_dragged(self, hole: int) -> None:
        label = self._select_scene_hole(hole)
        if self.drag_move_check.isChecked():
            self.stage_status.setText(f"Dragging across {label}. Moving because drag-move is enabled.")
            self._move_to_selected_hole()
        else:
            self.stage_status.setText(f"Scrubbing target over {label}. Double-click to move there.")

    def _on_scene_hole_activated(self, hole: int) -> None:
        label = self._select_scene_hole(hole)
        self.stage_status.setText(f"Moving to {label}.")
        self._move_to_selected_hole()

    def _jog_axis(self, key: str, delta: int, velocity: int | None = None) -> None:
        try:
            result = self.controller.move_axis_relative(key, delta, velocity=velocity)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Jog Error", str(exc))
            return
        self._append(f"Jogged {key.upper()} by {delta}: final={result.final_position}.")
        self._poll_stage_state()

    def _render_overlay_payload(self) -> str:
        target = self.target_hole.value()
        current = self._current_stage_hole
        target_text = "Cup 46 / measurement drop-off" if target == StageScene.CENTER_DROP_HOLE else f"Cup {target}"
        if current is None:
            current_text = "Unknown"
        elif current == StageScene.CENTER_DROP_HOLE:
            current_text = "Cup 46 / measurement drop-off"
        else:
            current_text = f"Cup {current}"
        return (
            "Stage overlay: "
            f"target={target_text}; current={current_text}; calibrated={len(self._calibration_map())} cups"
        )

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._save_config()
        self.controller.shutdown()
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "changer_xy_control_icon.png", assets_dir)
    window = MainWindow()
    set_app_icon(window, "changer_xy_control_icon.png", assets_dir)
    window.show()
    return app.exec()
