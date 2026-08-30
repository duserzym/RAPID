"""DAC command planning mapped from VB6 ``frmDAC_Comm``.

This module is intentionally hardware-adapter neutral. It provides the stable
command surface RapidPy needs before binding MCC/DAQ hardware: channel
configuration, explicit voltage validation, safe-zero commands, and a tiny
execution shim for adapters exposing ``write_voltage(channel, voltage)``.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class DacChannelConfig:
    """Configuration for one analog output channel."""

    channel: int
    name: str = "DAC"
    minimum_v: float = -10.0
    maximum_v: float = 10.0
    safe_v: float = 0.0

    def __post_init__(self) -> None:
        if self.channel < 0:
            raise ValueError("DAC channel must be non-negative")
        if self.minimum_v >= self.maximum_v:
            raise ValueError("DAC channel minimum must be below maximum")
        if not self.minimum_v <= self.safe_v <= self.maximum_v:
            raise ValueError("DAC safe voltage must be within channel limits")
        object.__setattr__(self, "name", str(self.name).strip() or f"DAC{self.channel}")


@dataclass(frozen=True)
class DacVoltageCommand:
    """Validated analog output request."""

    channel: DacChannelConfig
    requested_v: float
    allowed: bool
    reason: str = ""
    purpose: str = "write"

    @property
    def channel_number(self) -> int:
        return self.channel.channel

    @property
    def voltage_v(self) -> float:
        return float(self.requested_v)

    @property
    def command_summary(self) -> str:
        if not self.allowed:
            return f"BLOCKED {self.channel.name} ch{self.channel_number}: {self.reason}"
        return f"{self.purpose.upper()} {self.channel.name} ch{self.channel_number}: {self.voltage_v:.6g} V"


@dataclass(frozen=True)
class DacStartupReport:
    """Startup readiness report for retained DAC/MCC interfaces."""

    adapter_present: bool
    adapter_ready: bool
    safe_zero_commands: tuple[DacVoltageCommand, ...]
    blockers: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()

    @property
    def ok(self) -> bool:
        return not self.blockers

    @property
    def safe_zero_summary(self) -> tuple[str, ...]:
        return tuple(command.command_summary for command in self.safe_zero_commands)


class DacCommandPlanner:
    """Create validated DAC output commands for an adapter."""

    def __init__(self, channels: list[DacChannelConfig]) -> None:
        if not channels:
            raise ValueError("at least one DAC channel is required")
        self.channels = {channel.channel: channel for channel in channels}
        if len(self.channels) != len(channels):
            raise ValueError("DAC channel numbers must be unique")

    def plan_write(self, channel: int, voltage_v: float, *, purpose: str = "write") -> DacVoltageCommand:
        config = self._channel(channel)
        voltage = float(voltage_v)
        if not config.minimum_v <= voltage <= config.maximum_v:
            return DacVoltageCommand(
                channel=config,
                requested_v=voltage,
                allowed=False,
                reason=(
                    f"requested {voltage:.6g} V outside "
                    f"{config.minimum_v:.6g}..{config.maximum_v:.6g} V"
                ),
                purpose=purpose,
            )
        return DacVoltageCommand(channel=config, requested_v=voltage, allowed=True, purpose=purpose)

    def plan_zero(self, channel: int) -> DacVoltageCommand:
        config = self._channel(channel)
        return self.plan_write(config.channel, config.safe_v, purpose="safe-zero")

    def _channel(self, channel: int) -> DacChannelConfig:
        try:
            return self.channels[int(channel)]
        except KeyError as exc:
            raise ValueError(f"unknown DAC channel {channel}") from exc


def execute_dac_command(adapter: Any, command: DacVoltageCommand) -> Any:
    """Execute a validated command against an adapter.

    The adapter must expose ``write_voltage(channel, voltage)``. Blocked
    commands are rejected before any adapter call.
    """

    if not command.allowed:
        raise RuntimeError(command.command_summary)
    writer = getattr(adapter, "write_voltage", None)
    if not callable(writer):
        raise TypeError("DAC adapter must expose write_voltage(channel, voltage)")
    return writer(command.channel_number, command.voltage_v)


def inspect_dac_startup(planner: DacCommandPlanner, adapter: Any | None = None) -> DacStartupReport:
    """Return DAC/MCC startup readiness without writing to hardware."""
    safe_zero_commands = tuple(
        planner.plan_zero(channel_number)
        for channel_number in sorted(planner.channels)
    )
    warnings: list[str] = []
    blockers: list[str] = []

    if adapter is None:
        warnings.append("No DAC/MCC adapter configured; retained interface is planning-only.")
        return DacStartupReport(
            adapter_present=False,
            adapter_ready=False,
            safe_zero_commands=safe_zero_commands,
            blockers=tuple(blockers),
            warnings=tuple(warnings),
        )

    writer = getattr(adapter, "write_voltage", None)
    if not callable(writer):
        blockers.append("DAC/MCC adapter must expose write_voltage(channel, voltage).")
    return DacStartupReport(
        adapter_present=True,
        adapter_ready=not blockers,
        safe_zero_commands=safe_zero_commands,
        blockers=tuple(blockers),
        warnings=tuple(warnings),
    )
