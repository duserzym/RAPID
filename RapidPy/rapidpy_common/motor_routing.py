"""Route logical axes to their own Quicksilver serial connections.

The controller address is only unique within a serial port. Composite motion
operations retain the common client's algorithms and route each nested axis
operation, under one reentrant lock, to its configured physical connection.
"""
from __future__ import annotations

from functools import wraps
import inspect
import uuid

from .hardware import HardwareError, MotorAxisConfig, MotorSerialClient


class RoutedMotorSerialClient(MotorSerialClient):
    def __init__(self, config, axes, *, trace=None):
        super().__init__(config, trace=trace)
        self._axes_by_id = {axis.motor_id: axis for axis in axes}
        self._bindings = {axis.motor_id: (axis.name, axis.address, axis.port) for axis in axes}
        self._connections = {}
        self._connection_id = None
        self._routing_depth = 0
        seen = set()
        for axis in axes:
            if not axis.port.strip() or not 1 <= axis.address <= 254:
                raise HardwareError(f"Invalid motor port/address for {axis.name}.")
            binding = (axis.port.upper(), axis.address)
            if binding in seen:
                raise HardwareError(f"Duplicate motor port/address for {axis.name}.")
            seen.add(binding)
        if len(self._axes_by_id) != len(axes):
            raise HardwareError("Logical motor identifiers must be distinct.")

    @property
    def is_connected(self):
        if self._routing_depth:
            return super().is_connected
        expected = {axis.port.upper() for axis in self._axes_by_id.values()}
        return bool(expected) and set(self._connections) == expected and all(client.is_connected for client in self._connections.values())

    def connect(self, port=None, baudrate=57600, timeout=.35):
        with self._io_lock:
            self.disconnect()
            try:
                for axis in self._axes_by_id.values():
                    key = axis.port.upper()
                    if key in self._connections:
                        continue
                    client = MotorSerialClient(self.config, trace=lambda direction, payload, detail, port=axis.port:
                                               self._trace(direction, payload, f"port={port} {detail}") if self._trace else None)
                    self._connections[key] = client
                    client.connect(axis.port, baudrate, timeout)
                self._connection_id = uuid.uuid4().hex
            except Exception:
                self.disconnect()
                raise

    def disconnect(self):
        with self._io_lock:
            self._connection_id = None
            failures = []
            for port, client in list(self._connections.items()):
                try:
                    client.disconnect()
                except Exception as exc:
                    failures.append(str(exc))
                else:
                    del self._connections[port]
            self._serial = None
            if failures:
                raise HardwareError("Motor disconnect failed: " + "; ".join(failures))

    def _emit_trace(self, direction, payload, detail):
        super()._emit_trace(direction, payload, f"port={self._port} {detail}")

    def _with_axis(self, method, axis, args, kwargs):
        with self._io_lock:
            configured = self._axes_by_id.get(axis.motor_id)
            if configured != axis or self._bindings.get(axis.motor_id) != (axis.name, axis.address, axis.port):
                raise HardwareError(f"Unregistered motor axis {axis.name}.")
            client = self._connections.get(axis.port.upper())
            if client is None or not client.is_connected:
                raise HardwareError(f"Motor port {axis.port} is not open.")
            previous = self._serial, self._port, self._last_command
            self._serial, self._port = client._serial, client._port
            self._routing_depth += 1
            try:
                return method(self, axis, *args, **kwargs)
            finally:
                self._routing_depth -= 1
                self._serial, self._port, self._last_command = previous


def _route_axis_method(method, axis_parameter):
    @wraps(method)
    def routed(self, *args, **kwargs):
        if args:
            return self._with_axis(method, args[0], args[1:], kwargs)
        remaining = dict(kwargs)
        axis = remaining.pop(axis_parameter)
        return self._with_axis(method, axis, (), remaining)
    return routed


# Only public instance methods with a first axis parameter participate. Pure
# conversion helpers and raw serial commands cannot choose a port implicitly.
for _name, _method in inspect.getmembers(MotorSerialClient, inspect.isfunction):
    _params = list(inspect.signature(_method).parameters)
    if not _name.startswith("_") and len(_params) > 1 and _params[0] == "self" and (
        _params[1] == "axis" or _params[1].endswith("_axis")
    ):
        setattr(RoutedMotorSerialClient, _name, _route_axis_method(_method, _params[1]))
