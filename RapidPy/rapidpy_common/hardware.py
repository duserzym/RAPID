from __future__ import annotations

from dataclasses import dataclass
import math
import re
import threading
import time
from typing import Callable, Optional

import serial


class HardwareError(RuntimeError):
    """Raised when a serial command fails."""


@dataclass(slots=True)
class MotorAxisConfig:
    name: str
    motor_id: int
    address: int
    port: str = ""


@dataclass(slots=True)
class MotorControllerConfig:
    slot_min: int = 1
    slot_max: int = 101
    one_step: float = -1000
    sample_hole_alignment_offset: int = 0
    changer_speed: int = 8_000_000
    turner_speed: int = 120_000_000
    turning_motor_full_rotation: int = -172_800
    turning_motor_1rps: int = 157_950_000
    lift_speed_slow: int = 2_500_000
    lift_speed_normal: int = 8_500_000
    lift_speed_fast: int = 20_000_000
    lift_acceleration: int = 96_637
    meas_pos: int = 950_000
    sample_bottom: int = 425_000
    sample_height: int = 175_000
    updown_torque_factor: int = 15
    updown_max_torque: int = 32000
    pickup_torque_throttle: float = 0.6
    xy_neg_homing_distance: int = -30_000_000
    xy_pos_homing_distance: int = 30_000_000
    xy_corner_distance: int = -3_000_000


@dataclass(slots=True)
class MoveResult:
    target: int
    final_position: int
    success: bool


@dataclass(slots=True, frozen=True)
class MotorTelemetry:
    """One controller telemetry sample in native Quicksilver units."""

    timestamp: float
    axis_name: str
    target_position: int
    actual_position: int
    position_error: int | None = None
    velocity_1: int | None = None
    velocity_2: int | None = None
    actual_torque: int | None = None


_HEX4_RE = re.compile(r"[0-9A-Fa-f]{4}")


class MotorSerialClient:
    """VB6-aligned Quicksilver serial protocol wrapper for DC motors."""

    def __init__(
        self,
        config: Optional[MotorControllerConfig] = None,
        *,
        trace: Callable[[str, str, str], None] | None = None,
    ) -> None:
        self._serial: Optional[serial.Serial] = None
        self._port: str = ""
        self._trace = trace
        self._last_command = ""
        self._io_lock = threading.RLock()
        self._last_position: dict[int, int] = {1: 0, 2: 0, 3: 0, 4: 0}
        self._xy_last_pos: tuple[int, int] = (0, 0)
        self._axis_torque_capability: dict[int, bool] = {}
        self.config = config or MotorControllerConfig()

    def _emit_trace(self, direction: str, payload: str, detail: str) -> None:
        if self._trace is None:
            return
        try:
            self._trace(str(direction), str(payload), str(detail))
        except Exception:
            # Evidence hooks must never alter controller behavior.
            return

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    def connect(self, port: str, baudrate: int = 57600, timeout: float = 0.35) -> None:
        with self._io_lock:
            self.disconnect()
            self._port = port
            try:
                self._serial = serial.Serial(
                    port=port,
                    baudrate=baudrate,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_TWO,
                    timeout=timeout,
                    write_timeout=timeout,
                )
            except Exception as exc:
                self._emit_trace(
                    "ERROR", port, f"motor connect {port}:{baudrate} failed: {exc}"
                )
                raise
            try:
                self._serial.dtr = True
                self._serial.rts = True
                self._serial.reset_input_buffer()
                self._serial.reset_output_buffer()
                # VB6 issues this broadcast command at connect to set ACK delay.
                self.send_ascii("@255 173 416")
                time.sleep(0.03)
                self._serial.reset_input_buffer()
            except Exception as exc:
                try:
                    self._serial.close()
                except Exception as cleanup:
                    raise HardwareError(f'Motor setup failed: {exc}; original handle close failed: {cleanup}') from exc
                else:
                    self._serial = None
                self._emit_trace(
                    "ERROR", port, f"motor setup {port}:{baudrate} failed: {exc}"
                )
                raise
            self._emit_trace("INFO", "", f"motor connected {port}:{baudrate}")

    def disconnect(self) -> None:
        with self._io_lock:
            if self._serial is not None:
                self._serial.close()  # Retain the original handle for retry if close fails.
                self._serial = None
                self._emit_trace("INFO", "", f"motor disconnected {self._port}")

    def send_ascii(self, command: str) -> None:
        if not self.is_connected or self._serial is None:
            raise HardwareError("Motor serial connection is not open.")
        self._serial.dtr = True
        self._serial.rts = True
        payload = f"{command}\r\n".encode("ascii", errors="ignore")
        self._last_command = str(command)
        try:
            self._serial.write(payload)
            self._serial.flush()
        except Exception as exc:
            self._emit_trace(
                "ERROR", payload.decode("ascii", errors="replace"),
                f"motor write failed command={command!r}: {exc}",
            )
            raise
        self._emit_trace(
            "TX", payload.decode("ascii", errors="replace"), "motor command"
        )

    def read_ascii(self) -> str:
        if not self.is_connected or self._serial is None:
            raise HardwareError("Motor serial connection is not open.")
        raw = self._serial.read_until(b"\r")
        if not raw:
            error = HardwareError("No response from motor controller.")
            self._emit_trace(
                "ERROR", "", f"reply to {self._last_command!r} failed: {error}"
            )
            raise error
        if not raw.endswith(b"\r"):
            error = HardwareError("Partial motor response without CR terminator.")
            self._emit_trace(
                "ERROR",
                raw.decode("ascii", errors="replace"),
                f"reply to {self._last_command!r} failed: {error}",
            )
            raise error
        try:
            response = raw.decode("ascii", errors="strict").strip()
        except UnicodeDecodeError as exc:
            error = HardwareError(f"Motor response is not valid ASCII: {exc}")
            self._emit_trace(
                "ERROR",
                raw.decode("ascii", errors="replace"),
                f"reply to {self._last_command!r} failed: {error}",
            )
            raise error from exc
        if not response:
            error = HardwareError("Empty response from motor controller.")
            self._emit_trace(
                "ERROR", raw.decode("ascii", errors="replace"),
                f"reply to {self._last_command!r} failed: {error}",
            )
            raise error
        self._emit_trace(
            "RX", raw.decode("ascii", errors="replace"),
            f"reply to {self._last_command!r}",
        )
        return response

    def query_ascii(self, command: str) -> str:
        # A Quicksilver request and its response are one transaction. Telemetry
        # and motion workers may share a multi-drop serial connection, so they
        # must never consume each other's replies.
        with self._io_lock:
            self.send_ascii(command)
            return self.read_ascii()

    @staticmethod
    def _address(axis: MotorAxisConfig) -> str:
        return f"@{axis.address} "

    @staticmethod
    def _parse_status_word(response: str) -> int:
        if len(response) >= 15:
            token = response[10:14]
            try:
                return int(token, 16)
            except ValueError:
                pass
        groups = _HEX4_RE.findall(response)
        if groups:
            return int(groups[-1], 16)
        raise HardwareError(f"Unable to parse status word from response: {response!r}")

    @staticmethod
    def _parse_position(response: str) -> int:
        if len(response) >= 20:
            token = response[10:14] + response[15:19]
            try:
                raw = int(token, 16)
                if raw & 0x80000000:
                    raw -= 0x100000000
                return raw
            except ValueError:
                pass
        groups = _HEX4_RE.findall(response)
        if len(groups) >= 2:
            raw = int(groups[-2] + groups[-1], 16)
            if raw & 0x80000000:
                raw -= 0x100000000
            return raw
        raise HardwareError(f"Unable to parse position from response: {response!r}")

    @staticmethod
    def _parse_register_values(response: str, count: int) -> tuple[int, ...]:
        """Parse one or more signed 32-bit values returned by RRG (command 12)."""
        groups = _HEX4_RE.findall(response)
        required = int(count) * 2
        if count < 1 or len(groups) < required:
            raise HardwareError(
                f"Expected {count} register value(s), got response: {response!r}"
            )
        payload = groups[-required:]
        values: list[int] = []
        for index in range(0, required, 2):
            raw = int(payload[index] + payload[index + 1], 16)
            if raw & 0x80000000:
                raw -= 0x100000000
            values.append(raw)
        return tuple(values)

    @staticmethod
    def _signed_word(value: int, *, high: bool = False) -> int:
        raw = ((value >> 16) if high else value) & 0xFFFF
        return raw - 0x10000 if raw & 0x8000 else raw

    def read_registers(
        self,
        axis: MotorAxisConfig,
        registers: tuple[int, ...],
    ) -> tuple[int, ...]:
        """Read up to four Quicksilver registers in one serial transaction."""
        if not 1 <= len(registers) <= 4:
            raise ValueError("Quicksilver RRG accepts between one and four registers.")
        register_text = " ".join(str(int(register)) for register in registers)
        response = self.query_ascii(f"{self._address(axis)}12 {register_text}")
        try:
            return self._parse_register_values(response, len(registers))
        except Exception as exc:
            self._emit_trace("ERROR", response, f"motor register parse failed: {exc}")
            raise

    def read_telemetry(self, axis: MotorAxisConfig) -> MotorTelemetry:
        """Read target, feedback, error, velocity, and torque efficiently.

        Registers 0, 1, 6, and 7 are batched in one RRG request. Register 9 is
        read separately because the protocol permits at most four registers per
        request. Word-oriented fields are sign-extended according to the
        SilverLode dedicated-register map.
        """
        target, actual, error_register, velocity_register = self.read_registers(
            axis, (0, 1, 6, 7)
        )
        axis_id = axis.motor_id
        supports_torque = self._axis_torque_capability.get(axis_id)
        torque_reading: int | None
        if supports_torque is False:
            torque_reading = None
        else:
            try:
                (torque_register,) = self.read_registers(axis, (9,))
            except Exception:
                if supports_torque is None:
                    self._axis_torque_capability[axis_id] = False
                torque_reading = None
            else:
                self._axis_torque_capability[axis_id] = True
                torque_reading = self._signed_word(torque_register)
        self._last_position[axis.motor_id] = actual
        return MotorTelemetry(
            timestamp=time.monotonic(),
            axis_name=axis.name,
            target_position=target,
            actual_position=actual,
            position_error=self._signed_word(error_register),
            velocity_1=self._signed_word(velocity_register, high=True),
            velocity_2=self._signed_word(velocity_register),
            actual_torque=torque_reading,
        )

    def poll_motor(self, axis: MotorAxisConfig) -> str:
        return self.query_ascii(f"{self._address(axis)}0")

    def clear_poll_status(self, axis: MotorAxisConfig) -> str:
        return self.query_ascii(f"{self._address(axis)}1 65535")

    def read_position(self, axis: MotorAxisConfig) -> int:
        response = self.query_ascii(f"{self._address(axis)}12 1")
        try:
            pos = self._parse_position(response)
        except Exception as exc:
            self._emit_trace("ERROR", response, f"motor position parse failed: {exc}")
            raise
        self._last_position[axis.motor_id] = pos
        return pos

    def check_internal_status(self, axis: MotorAxisConfig, bit: int) -> int:
        response = self.query_ascii(f"{self._address(axis)}20")
        try:
            status_word = self._parse_status_word(response)
        except Exception as exc:
            self._emit_trace("ERROR", response, f"motor status parse failed: {exc}")
            raise
        return (status_word // (2 ** bit)) % 2

    def halt(self, axis: MotorAxisConfig) -> str:
        return self.query_ascii(f"{self._address(axis)}2")

    def stop(self, axis: MotorAxisConfig) -> str:
        return self.query_ascii(f"{self._address(axis)}3 0")

    def reset(self, axis: MotorAxisConfig) -> str:
        return self.query_ascii(f"{self._address(axis)}4")

    def zero_target_pos(self, axis: MotorAxisConfig) -> None:
        self.query_ascii(f"{self._address(axis)}145")
        self.read_position(axis)

    def set_torques(
        self,
        axis: MotorAxisConfig,
        closed_hold: int,
        closed_move: int,
        open_hold: int,
        open_move: int,
    ) -> str:
        # VB6 SetTorques accepts percentages and converts each to controller
        # units using the station's configured 100% torque value.
        percentages = (closed_hold, closed_move, open_hold, open_move)
        if any(not 0 <= value <= 100 for value in percentages) or not 0 < self.config.updown_max_torque <= 32767:
            raise HardwareError("Torque percentages or native maximum torque are invalid.")
        per_torque = round(.01 * self.config.updown_max_torque)
        closed_hold, closed_move, open_hold, open_move = (round(value) * per_torque for value in percentages)
        return self.query_ascii(
            f"{self._address(axis)}149 {closed_hold} {closed_move} {open_hold} {open_move}"
        )

    def move_motor(
        self,
        axis: MotorAxisConfig,
        target: int,
        velocity: int,
        wait_for_stop: bool = True,
        stop_enable: int = 0,
        stop_condition: int = 0,
        acceleration: int = 96637,
        relative_mode: bool = False,
    ) -> MoveResult:
        for name, value in (("target", target), ("velocity", velocity), ("acceleration", acceleration)):
            if not math.isfinite(value) or int(value) != value or not -(2**31) <= value < 2**31:
                raise HardwareError(f"Motor {name} must be a signed 32-bit integer.")
        if velocity <= 0 or acceleration <= 0:
            raise HardwareError("Motor velocity and acceleration must be positive.")
        self._check_motion_cancel()
        self.poll_motor(axis)
        self.clear_poll_status(axis)
        opcode = 135 if relative_mode else 134
        use_accel = int(acceleration)
        if axis.motor_id == 3:
            use_accel = int(self.config.lift_acceleration)
        elif axis.motor_id in (1, 4):
            use_accel = 483184
        elif axis.motor_id == 2:
            use_accel = 4831
            if abs(int(velocity)) > int(self.config.turning_motor_1rps):
                scale = abs(float(velocity)) / float(self.config.turning_motor_1rps)
                use_accel = int(use_accel * scale * 1.1)

        self._check_motion_cancel()
        self.query_ascii(
            f"{self._address(axis)}{opcode} {target} {use_accel} {velocity} {stop_enable} {stop_condition}"
        )
        if wait_for_stop:
            self.wait_for_motor_stop(axis)
        final_pos = self.read_position(axis)
        return MoveResult(target=target, final_position=final_pos, success=True)

    def _check_motion_cancel(self, axis=None):
        check = getattr(self, '_motion_cancel_check', None)
        if check is not None and check():
            if axis is not None:
                self.halt(axis)
            raise HardwareError('Motor motion cancelled or its operation deadline expired.')

    def set_position_observer(self, observer: "Callable[[MotorAxisConfig, float, int], None] | None") -> None:
        """Receive ``(axis, perf_counter_s, position)`` for each in-motion position poll.

        Used to timestamp lift height / turn angle while a SQUID stream records
        the specimen in motion.  The observer runs on the motion thread and any
        exception it raises is swallowed so it can never disturb motion.
        """
        self._position_observer = observer

    def _notify_position(self, axis: MotorAxisConfig, pos: int) -> None:
        observer = getattr(self, "_position_observer", None)
        if observer is None:
            return
        try:
            observer(axis, time.perf_counter(), pos)
        except Exception:
            pass

    def wait_for_motor_stop(self, axis: MotorAxisConfig, timeout_s: float = 120.0) -> None:
        start = time.monotonic()
        old1 = 2**7
        old0 = -(2**7)
        while True:
            self._check_motion_cancel(axis)
            time.sleep(0.05)
            pos = self.read_position(axis)
            self._notify_position(axis, pos)
            if old1 == old0 == pos:
                break
            if abs(old1 - old0) < 5 and abs(old0 - pos) < 5:
                break
            old1, old0 = old0, pos
            if time.monotonic() - start > timeout_s:
                raise HardwareError(f"Timed out waiting for motor {axis.name} to stop.")
        self.stop(axis)

    def convert_hole_to_pos(self, target_hole: float, current_pos: int) -> int:
        full_loop = abs((self.config.slot_max - self.config.slot_min + 1) * self.config.one_step)
        target_hole_pos_raw = int(self.config.one_step * target_hole)
        steps_to_go = (target_hole_pos_raw - current_pos) % full_loop
        if abs(steps_to_go) > (full_loop / 2):
            if steps_to_go > 0:
                steps_to_go -= full_loop
            else:
                steps_to_go += full_loop
        if not changer_is_hole(target_hole, self.config.slot_min, self.config.slot_max):
            steps_to_go += self.config.sample_hole_alignment_offset * self.config.one_step
        return int(steps_to_go + current_pos)

    def changer_motor_to_hole(
        self,
        axis: MotorAxisConfig,
        hole: float,
        wait_for_stop: bool = True,
    ) -> MoveResult:
        starting_pos = self.read_position(axis)
        target = self.convert_hole_to_pos(hole, starting_pos)
        result = self.move_motor(
            axis,
            target,
            self.config.changer_speed,
            wait_for_stop=wait_for_stop,
            acceleration=483184,
            relative_mode=False,
        )
        if not wait_for_stop:
            return result
        curpos = self.read_position(axis)
        tol_holes = abs(curpos - target) / abs(self.config.one_step)
        if tol_holes > 0.02:
            self.move_motor(
                axis,
                target,
                int(0.5 * self.config.changer_speed),
                wait_for_stop=True,
                acceleration=483184,
                relative_mode=False,
            )
            curpos = self.read_position(axis)
            tol_holes = abs(curpos - target) / abs(self.config.one_step)
        return MoveResult(target=target, final_position=curpos, success=tol_holes <= 0.02)

    def turning_motor_rotate(
        self,
        axis: MotorAxisConfig,
        angle: float,
        wait_for_stop: bool = True,
    ) -> MoveResult:
        target = int(-self.config.turning_motor_full_rotation * angle / 360.0)
        result = self.move_motor(
            axis,
            target,
            self.config.turner_speed,
            wait_for_stop=wait_for_stop,
            acceleration=4831,
            relative_mode=False,
        )
        if not wait_for_stop:
            return result
        final_angle = convert_pos_to_angle(result.final_position, self.config.turning_motor_full_rotation)
        angle_delta = abs((final_angle - angle) % 360)
        if angle_delta > 180:
            angle_delta = 360 - angle_delta
        return MoveResult(target=result.target, final_position=result.final_position, success=angle_delta <= 5.0)

    def turning_motor_spin(
        self,
        axis: MotorAxisConfig,
        speed_rps: float,
        duration_s: float = 60.0,
    ) -> MoveResult:
        if not math.isfinite(speed_rps) or not math.isfinite(duration_s) or duration_s <= 0 or abs(speed_rps) > 40:
            raise HardwareError("Turning speed must be finite within +/-40 rps and duration positive.")
        if speed_rps == 0:
            self.stop(axis)
            self.wait_for_motor_stop(axis)
            pos = self.read_position(axis)
            return MoveResult(target=pos, final_position=pos, success=True)
        start = self.read_position(axis)
        target = int(start - self.config.turning_motor_full_rotation * speed_rps * duration_s)
        velocity = int(abs(self.config.turning_motor_1rps * speed_rps))
        if not -(2**31) <= target < 2**31 or not 0 < velocity < 2**31:
            raise HardwareError("Turning spin target/velocity exceeds native signed 32-bit limits.")
        return self.move_motor(axis, target, velocity, wait_for_stop=False, acceleration=4831)

    def updown_move(
        self,
        axis: MotorAxisConfig,
        target: int,
        speed_index: int,
        wait_for_stop: bool = True,
    ) -> MoveResult:
        speed_index = max(0, min(2, int(speed_index)))
        speeds = [self.config.lift_speed_slow, self.config.lift_speed_normal, self.config.lift_speed_fast]
        result = self.move_motor(
            axis,
            target,
            speeds[speed_index],
            wait_for_stop=wait_for_stop,
            acceleration=96637,
            relative_mode=False,
        )
        if not wait_for_stop:
            return result
        curpos = self.read_position(axis)
        if abs(curpos - target) > 100 and target != 0:
            mid = int((curpos + self._last_position.get(axis.motor_id, curpos)) / 2)
            self.move_motor(axis, mid, self.config.lift_speed_slow, wait_for_stop=True, acceleration=96637)
            self.move_motor(axis, target, speeds[speed_index], wait_for_stop=True, acceleration=96637)
            curpos = self.read_position(axis)
        return MoveResult(target=target, final_position=curpos, success=(abs(curpos - target) <= 150 or target == 0))

    def relabel_pos(self, axis: MotorAxisConfig, pos: int, tolerance: int = 10, max_cycles: int = 20) -> None:
        for _ in range(max_cycles):
            self._check_motion_cancel()
            current = self.read_position(axis)
            if abs(current - pos) < tolerance:
                return
            self._check_motion_cancel()
            self.zero_target_pos(axis)
            self.query_ascii(f"{self._address(axis)}11 10 {-int(pos)}")
            self.query_ascii(f"{self._address(axis)}165 1802")
        raise HardwareError(f"Unable to relabel motor {axis.name} to position {pos}.")

    def home_to_top(self, updown_axis: MotorAxisConfig) -> MoveResult:
        self._check_motion_cancel()
        if self.check_internal_status(updown_axis, 4) == 1:
            self._check_motion_cancel()
            self.zero_target_pos(updown_axis)
            pos = self.read_position(updown_axis)
            return MoveResult(target=0, final_position=pos, success=True)

        up_pos = self.read_position(updown_axis)
        if abs(up_pos) > abs(self.config.sample_bottom):
            speed = self.config.lift_speed_normal
        else:
            speed = int(0.25 * (self.config.lift_speed_normal + 3 * self.config.lift_speed_slow))

        self.move_motor(
            updown_axis,
            target=-2 * int(self.config.meas_pos),
            velocity=int(speed),
            wait_for_stop=True,
            stop_enable=-1,
            stop_condition=1,
        )

        if self.check_internal_status(updown_axis, 4) != 1:
            self.move_motor(
                updown_axis,
                target=-2 * int(self.config.meas_pos),
                velocity=int(self.config.lift_speed_slow),
                wait_for_stop=True,
            )

        if self.check_internal_status(updown_axis, 4) != 1:
            raise HardwareError("Homed to top but did not hit switch (internal status bit 4).")

        final_pos = self.read_position(updown_axis)
        self._check_motion_cancel()
        self.zero_target_pos(updown_axis)
        return MoveResult(target=0, final_position=final_pos, success=True)

    def home_xy_to_center(
        self,
        x_axis: MotorAxisConfig,
        y_axis: MotorAxisConfig,
        updown_axis: MotorAxisConfig,
    ) -> tuple[MoveResult, MoveResult]:
        self._check_motion_cancel()
        if self.check_internal_status(updown_axis, 4) == 0:
            raise HardwareError("Cannot home XY to center: up/down axis is not homed to top.")

        stop_x = False
        stop_y = False

        while (self.check_internal_status(x_axis, 4) != 0) or (self.check_internal_status(y_axis, 5) != 0):
            self._check_motion_cancel()
            time.sleep(0.05)
            if self.check_internal_status(x_axis, 4) != 0 and not stop_x:
                self.move_motor(
                    x_axis,
                    target=int(self.config.xy_neg_homing_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_internal_status(y_axis, 5) != 0 and not stop_y:
                self.move_motor(
                    y_axis,
                    target=int(self.config.xy_neg_homing_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True

        if (self.check_internal_status(x_axis, 4) != 0) or (self.check_internal_status(y_axis, 5) != 0):
            raise HardwareError("Homed XY center pass 1 failed: did not hit negative limit switches.")

        stop_x = False
        stop_y = False

        while (self.check_internal_status(x_axis, 5) != 0) or (self.check_internal_status(y_axis, 6) != 0):
            self._check_motion_cancel()
            time.sleep(0.05)
            if self.check_internal_status(x_axis, 5) != 0 and not stop_x:
                self.move_motor(
                    x_axis,
                    target=int(self.config.xy_pos_homing_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_internal_status(y_axis, 6) != 0 and not stop_y:
                self.move_motor(
                    y_axis,
                    target=int(self.config.xy_pos_homing_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-3,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True

        if (self.check_internal_status(x_axis, 5) != 0) or (self.check_internal_status(y_axis, 6) != 0):
            raise HardwareError("Homed XY center pass 2 failed: did not hit positive limit switches.")

        self._check_motion_cancel()
        self.zero_target_pos(x_axis)
        self._check_motion_cancel()
        self.zero_target_pos(y_axis)
        x_pos = self.read_position(x_axis)
        y_pos = self.read_position(y_axis)
        self._xy_last_pos = (x_pos, y_pos)
        return (
            MoveResult(target=0, final_position=x_pos, success=True),
            MoveResult(target=0, final_position=y_pos, success=True),
        )

    def move_xy_to_corner(
        self,
        x_axis: MotorAxisConfig,
        y_axis: MotorAxisConfig,
        updown_axis: MotorAxisConfig,
    ) -> tuple[MoveResult, MoveResult]:
        self._check_motion_cancel()
        self.home_to_top(updown_axis)
        if self.check_internal_status(updown_axis, 4) == 0:
            raise HardwareError("Cannot move XY to corner: up/down axis is not homed to top.")

        stop_x = False
        stop_y = False
        while (self.check_internal_status(x_axis, 4) != 0) or (self.check_internal_status(y_axis, 5) != 0):
            self._check_motion_cancel()
            time.sleep(0.05)
            if self.check_internal_status(x_axis, 4) != 0 and not stop_x:
                self.move_motor(
                    x_axis,
                    target=int(self.config.xy_corner_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_internal_status(y_axis, 5) != 0 and not stop_y:
                self.move_motor(
                    y_axis,
                    target=int(self.config.xy_corner_distance),
                    velocity=int(self.config.changer_speed),
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True

        x_pos = self.read_position(x_axis)
        y_pos = self.read_position(y_axis)
        self._xy_last_pos = (x_pos, y_pos)
        return (
            MoveResult(target=x_pos, final_position=x_pos, success=True),
            MoveResult(target=y_pos, final_position=y_pos, success=True),
        )

    def sample_dropoff(self, updown_axis: MotorAxisConfig, use_xy_table: bool = True) -> MoveResult:
        self._check_motion_cancel()
        self.poll_motor(updown_axis)
        self.clear_poll_status(updown_axis)

        if use_xy_table:
            target = int(self.config.sample_bottom + (self.config.sample_height - 0.1 * self.config.sample_height))
        else:
            target = int(self.config.sample_bottom + 1.1 * self.config.sample_height)

        result = self.move_motor(
            updown_axis,
            target=target,
            velocity=int(self.config.lift_speed_slow),
            wait_for_stop=True,
        )
        if self.check_internal_status(updown_axis, 4) == 1:
            raise HardwareError("Dropped off sample but homing switch is still set (bit 4).")
        return result

    def sample_pickup(self, updown_axis: MotorAxisConfig) -> MoveResult:
        self._check_motion_cancel()
        pickup_torque = round(self.config.pickup_torque_throttle * self.config.updown_torque_factor)
        self.set_torques(updown_axis, pickup_torque, pickup_torque, pickup_torque, pickup_torque)
        result = self.move_motor(
            updown_axis,
            target=int(self.config.sample_bottom),
            velocity=int(self.config.lift_speed_slow),
            wait_for_stop=True,
        )
        time.sleep(1.0)
        current_pos = self.read_position(updown_axis)
        self._check_motion_cancel()
        self.zero_target_pos(updown_axis)
        self._check_motion_cancel()
        self.relabel_pos(updown_axis, current_pos)
        if self.check_internal_status(updown_axis, 4) == 1:
            raise HardwareError("Quartz tube at sample top but homing switch is still set (bit 4).")
        return MoveResult(target=result.target, final_position=current_pos, success=True)

    def move_to_position(self, axis: MotorAxisConfig, position: int, speed: int = 1200) -> str:
        # Backward-compatible shim for existing app code paths.
        result = self.move_motor(axis, position, speed, wait_for_stop=True)
        return f"target={result.target} final={result.final_position} success={result.success}"


def convert_hole_to_chain_position(hole: float, one_step: int) -> int:
    return int(one_step * hole)


def convert_position_to_hole(pos: int, slot_min: int, slot_max: int, one_step: int) -> float:
    full_loop = slot_max - slot_min + 1
    hole = (pos / one_step) % full_loop
    if hole <= 0:
        hole += full_loop
    return float(hole)


def changer_is_hole(num: float, slot_min: int, slot_max: int) -> bool:
    hole_slot_num = slot_max - slot_min + 1
    if hole_slot_num <= 0:
        return False
    return (int(num) % hole_slot_num) == 0


def convert_angle_to_pos(angle: float, turning_motor_full_rotation: int) -> int:
    return int(-turning_motor_full_rotation * angle / 360.0)


def convert_pos_to_angle(pos: int, turning_motor_full_rotation: int) -> float:
    if turning_motor_full_rotation == 0:
        raise HardwareError("Turning motor full rotation cannot be zero.")
    return (pos / -float(turning_motor_full_rotation)) * 360.0
