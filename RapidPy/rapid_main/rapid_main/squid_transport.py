"""Live transport and motion adapters for bracketed SQUID acquisition.

This module binds the narrow protocols in :mod:`rapid_main.acquisition` to the
real instruments:

* :class:`RawSquidTransport` drives the 2G 581 through the shared
  ``updown_control`` serial client (latch, counter read, DVM read, range,
  counter reset), keeping every raw reply for the audit record.
* :class:`MotorVerticalController` and :class:`MotorTurningController` drive the
  QuickSilver lift and turning axes through ``rapidpy_common.hardware`` and
  report whether the commanded motion was actually reached.
* :class:`BracketedSquidBackend` is the measurement-backend facade the worker
  consumes: ``read_squid`` returns a whole :class:`BracketedMeasurementBlock`
  and ``recover_flux_count_discontinuity`` performs the real 2G recovery.

Nothing here falls back to simulation.  A missing dependency, port, or
calibration raises so hardware mode fails preflight instead of quietly
producing synthetic numbers.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from rapid_main.acquisition import (
    AcquisitionConfig,
    AcquisitionError,
    AxisReply,
    BlockContext,
    BracketedAcquisition,
    BracketedAcquisitionService,
    LatchResult,
    MotionOutcome,
    RecoveryRecord,
)
from rapid_main.config import AppConfig
from rapid_main.magnetometer import (
    BracketedMeasurementBlock,
    Position4,
    Vector3,
    ZeroPairValidation,
)
from rapidpy_common.hardware import (
    HardwareError,
    MotorAxisConfig,
    MotorSerialClient,
    convert_angle_to_pos,
    convert_pos_to_angle,
)


class SquidTransportError(RuntimeError):
    """Raised when the 2G transport is unusable."""


@dataclass(frozen=True)
class SquidTransportConfig:
    """Serial and latch settings for the 2G transport."""

    port: str
    baud: int = 1200
    settle_delay_s: float = 1.0
    range_value: float = 1.0


class RawSquidTransport:
    """:class:`~rapid_main.acquisition.SquidTransport` over the 2G serial client.

    ``latch_id`` is minted per latch cycle and stamped onto every axis reply so
    the acquisition service can prove that X, Y, and Z came from one
    ``ALC``/``ALD`` pair.
    """

    def __init__(self, client: object, *, config: SquidTransportConfig) -> None:
        self._client = client
        self._config = config
        self._latch_serial = 0
        self._latch_id = ""

    @property
    def client(self) -> object:
        return self._client

    @property
    def latch_id(self) -> str:
        return self._latch_id

    def is_connected(self) -> bool:
        return bool(getattr(self._client, "is_connected", False))

    def clear_and_reset_counters(self, axis: str = "A") -> None:
        self._require_connected()
        self._client.clear_and_reset(axis)  # type: ignore[attr-defined]

    def set_range(self, axis: str, range_label: str) -> None:
        self._require_connected()
        self._client.set_range(axis, range_label)  # type: ignore[attr-defined]

    def latch(self, axis: str = "A", *, settle: bool = False) -> LatchResult:
        self._require_connected()
        settle_s = float(self._config.settle_delay_s) if settle else 0.0
        commands = self._client.latch(axis, settle_s=settle_s)  # type: ignore[attr-defined]
        self._latch_serial += 1
        self._latch_id = f"{axis}-{self._latch_serial:06d}"
        return LatchResult(latch_id=self._latch_id, commands=tuple(commands))

    def read_axis(self, axis: str) -> AxisReply:
        self._require_connected()
        if not self._latch_id:
            raise SquidTransportError("read_axis called before a latch cycle")
        sample = self._client.read_axis(axis, range_value=self._config.range_value)  # type: ignore[attr-defined]
        return AxisReply(
            counts=float(sample.counts),
            dvm=float(sample.dvm),
            latch_id=self._latch_id,
            range_value=float(sample.range_value),
            count_command=str(sample.count_command),
            count_reply=str(sample.count_reply),
            data_command=str(sample.data_command),
            data_reply=str(sample.data_reply),
        )

    def _require_connected(self) -> None:
        if not self.is_connected():
            raise SquidTransportError(
                f"SQUID transport is not connected ({self._config.port}:{self._config.baud})."
            )


class MotorVerticalController:
    """Verified lift motion for the up/down axis."""

    def __init__(self, client: MotorSerialClient, axis: MotorAxisConfig) -> None:
        self._client = client
        self._axis = axis

    def move_to(self, position: int, *, speed_index: int = 0) -> MotionOutcome:
        try:
            result = self._client.updown_move(self._axis, int(position), int(speed_index), wait_for_stop=True)
        except HardwareError as exc:
            return MotionOutcome(target=float(position), actual=float("nan"), ok=False, detail=str(exc))
        return MotionOutcome(
            target=float(result.target),
            actual=float(result.final_position),
            ok=bool(result.success),
            detail="" if result.success else "lift did not settle within tolerance",
        )

    def position(self) -> int:
        return int(self._client.read_position(self._axis))


class MotorTurningController:
    """Verified rotation for the turning axis."""

    def __init__(self, client: MotorSerialClient, axis: MotorAxisConfig) -> None:
        self._client = client
        self._axis = axis

    @property
    def _full_rotation(self) -> int:
        return int(self._client.config.turning_motor_full_rotation)

    def rotate_to(self, angle_deg: float) -> MotionOutcome:
        try:
            result = self._client.turning_motor_rotate(self._axis, float(angle_deg), wait_for_stop=True)
        except HardwareError as exc:
            return MotionOutcome(target=float(angle_deg), actual=float("nan"), ok=False, detail=str(exc))
        actual = convert_pos_to_angle(result.final_position, self._full_rotation)
        return MotionOutcome(
            target=float(angle_deg),
            actual=float(actual),
            ok=bool(result.success),
            detail="" if result.success else "turn exceeded the 5 degree tolerance",
        )

    def angle(self) -> float:
        return float(convert_pos_to_angle(self._client.read_position(self._axis), self._full_rotation))

    def set_reference_angle(self, angle_deg: float) -> None:
        """VB6 ``SetTurningMotorAngle``: wrap into one rotation, then relabel."""

        position = int(self._client.read_position(self._axis))
        wrapped = wrap_turning_position(position, self._full_rotation)
        target = wrapped + convert_angle_to_pos(float(angle_deg), self._full_rotation)
        self._client.relabel_pos(self._axis, int(target))


def wrap_turning_position(position: int, full_rotation: int, *, max_cycles: int = 64) -> int:
    """Port of the VB6 ``SetTurningMotorAngle`` position-wrapping loops."""

    if full_rotation == 0:
        raise HardwareError("Turning motor full rotation cannot be zero.")
    pos = int(position)
    sign = 1 if full_rotation > 0 else -1
    cycles = 0
    if pos < 0 and sign == 1:
        while not pos > -full_rotation * 0.95 and cycles < max_cycles:
            pos += full_rotation
            cycles += 1
    elif pos > 0 and sign == -1:
        while not pos < -full_rotation * 0.05 and cycles < max_cycles:
            pos += full_rotation
            cycles += 1
    elif pos < 0 and sign == -1:
        while not pos > full_rotation * 0.95 and cycles < max_cycles:
            pos -= full_rotation
            cycles += 1
    elif abs(pos) < abs(full_rotation * 0.05):
        return pos
    else:
        while not pos < abs(full_rotation * 0.05) and cycles < max_cycles:
            pos -= abs(full_rotation)
            cycles += 1
    if cycles >= max_cycles:
        raise HardwareError(
            f"Unable to wrap turning position {position} into one rotation of {full_rotation}."
        )
    return pos


class BracketedSquidBackend:
    """Measurement backend that returns coherent bracketed blocks.

    The worker treats a rejected block as fatal unless this backend can safely
    re-zero, so ``recover_flux_count_discontinuity`` is only present when a real
    recovery path exists.  It never returns a value of its own: the whole block
    is discarded and re-acquired from the zero-before step.
    """

    read_timeout: float | None = None
    return_timeout: float | None = None

    def __init__(
        self,
        service: BracketedAcquisitionService,
        *,
        holder_provider: Callable[[], Position4 | None] | None = None,
        direction_provider: Callable[[], bool] | None = None,
        context_provider: Callable[[], BlockContext] | None = None,
        flux_discontinuity_retries: int = 2,
        simulated: bool = False,
    ) -> None:
        self._service = service
        self._holder_provider = holder_provider
        self._direction_provider = direction_provider
        self._context_provider = context_provider
        self.flux_discontinuity_retries = max(0, int(flux_discontinuity_retries))
        self._simulated = bool(simulated)
        self._last_acquisition: BracketedAcquisition | None = None
        self._recoveries: list[RecoveryRecord] = []

    @property
    def simulated(self) -> bool:
        return self._simulated

    @property
    def last_acquisition(self) -> BracketedAcquisition | None:
        return self._last_acquisition

    @property
    def recovery_records(self) -> tuple[RecoveryRecord, ...]:
        return tuple(self._recoveries)

    def read_squid(self) -> BracketedMeasurementBlock:
        context = self._context_provider() if self._context_provider else BlockContext()
        if self._simulated and not context.simulated:
            context = BlockContext(
                sample_name=context.sample_name,
                treatment_label=context.treatment_label,
                run_id=context.run_id,
                operator=context.operator,
                software_version=context.software_version,
                config_hash=context.config_hash,
                holder_record_id=context.holder_record_id,
                holder_recorded_iso=context.holder_recorded_iso,
                is_holder_block=context.is_holder_block,
                simulated=True,
            )
        is_up = bool(self._direction_provider()) if self._direction_provider else True
        holder = self._holder_provider() if self._holder_provider else None
        acquisition = self._service.acquire(
            is_up=is_up,
            holder_positions=holder,
            context=context,
        )
        self._last_acquisition = acquisition
        return acquisition.block

    def recover_flux_count_discontinuity(
        self, validation: ZeroPairValidation | None = None
    ) -> RecoveryRecord:
        record = self._service.recover_flux_count_discontinuity(
            validation, attempt=len(self._recoveries) + 1
        )
        self._recoveries.append(record)
        return record

    def return_to_safe_state(self) -> None:
        try:
            self._service.return_to_safe_state()
        except AcquisitionError:
            raise
        except Exception as exc:  # adapters raise their own transport errors
            raise AcquisitionError(f"failed to return to a safe state: {exc}") from exc


def acquisition_config_from_app_config(
    config: AppConfig,
    *,
    zero_position: int,
    measurement_position: int,
) -> AcquisitionConfig:
    """Build acquisition settings from calibration/SQUID configuration."""

    calibration = config.calibration
    axis_calibration: Vector3 = (
        float(calibration.cal_x),
        float(calibration.cal_y),
        float(calibration.cal_z),
    )
    return AcquisitionConfig(
        zero_position=int(zero_position),
        measurement_position=int(measurement_position),
        axis_calibration=axis_calibration,
        range_factor=float(calibration.range_factor),
        range_label=_range_label(config.squid.range_label),
        holder_range_label="1",
        settle_delay_s=float(config.squid.settle_time),
    )


def _range_label(label: str) -> str:
    """Map the UI range wording onto the 2G control-rate letters."""

    text = (label or "").strip().lower()
    if "flux" in text or text in {"f", "extended"}:
        return "F"
    for prefix, code in (("1000", "E"), ("100", "H"), ("10", "T"), ("1", "1")):
        if text.startswith(prefix):
            return code
    return "1"
