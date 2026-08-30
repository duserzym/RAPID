"""Replay recorded 2G blocks without hardware.

A replay fixture is a recorded sequence of six SQUID observations (zero-before,
four orientations, zero-after) stored as raw counter and DVM values, exactly as
the instrument reports them. Replaying one exercises the whole reduction,
rejection, holder, and output path with no magnetometer attached.

This is production code, not test scaffolding: step D1 of the hardware
acceptance procedure replays the archived failing block through RapidPy and
requires proof that the holder correction and the output directory are
unchanged afterwards.

Fixture schema (``rapidpy.replay.bracketed_block.v1``)::

    {
      "schema": "rapidpy.replay.bracketed_block.v1",
      "name": "holder-archived-x-step",
      "description": "...",
      "source": "where the recording came from",
      "axis_calibration": [x, y, z],
      "range_value": 1.0,
      "range_factor": 1.0e-5,
      "is_up": true,
      "is_holder_block": true,
      "expected": {"accepted": false, "discontinuous_axes": ["X"]},
      "observations": [
        {"role": "zero-before", "axes": {"X": {"counts": 0, "dvm": 1.0}, ...}},
        ...
      ]
    }

``counts`` and ``dvm`` are the verbatim ``<axis>SC`` and ``<axis>SD`` values.
The calibrated value is ``(-dvm - counts * range_value) * axis_calibration``,
matching VB6 ``getVal`` followed by ``getData``.
"""
from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from rapid_main.magnetometer import (
    AXIS_NAMES,
    AxisObservation,
    BlockAudit,
    BracketedMeasurementBlock,
    Position4,
    SquidObservation,
    Vector3,
    ZeroPairValidation,
)

REPLAY_SCHEMA = "rapidpy.replay.bracketed_block.v1"

BLOCK_ROLES: tuple[str, ...] = (
    "zero-before",
    "position-1",
    "position-2",
    "position-3",
    "position-4",
    "zero-after",
)
POSITION_ANGLES: tuple[float, ...] = (0.0, 0.0, 90.0, 180.0, 270.0, 0.0)


class ReplayFixtureError(ValueError):
    """Raised when a replay fixture cannot be turned into a coherent block."""


@dataclass(frozen=True)
class ReplayExpectation:
    """What the fixture asserts should happen when the block is reduced."""

    accepted: bool = True
    discontinuous_axes: tuple[str, ...] = ()
    flux_step_axes: tuple[str, ...] = ()
    reason_contains: str = ""


@dataclass(frozen=True)
class ReplayFixture:
    """One recorded bracketed block plus its expected verdict."""

    name: str
    block: BracketedMeasurementBlock
    expectation: ReplayExpectation
    description: str = ""
    source: str = ""
    path: Path | None = None


def load_replay_fixture(path: str | Path) -> ReplayFixture:
    """Read a replay fixture from disk."""

    fixture_path = Path(path)
    try:
        payload = json.loads(fixture_path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise ReplayFixtureError(f"cannot read replay fixture {fixture_path}: {exc}") from exc
    except ValueError as exc:
        raise ReplayFixtureError(f"replay fixture {fixture_path} is not valid JSON: {exc}") from exc
    fixture = replay_fixture_from_dict(payload)
    return ReplayFixture(
        name=fixture.name,
        block=fixture.block,
        expectation=fixture.expectation,
        description=fixture.description,
        source=fixture.source,
        path=fixture_path,
    )


def replay_fixture_from_dict(payload: Mapping[str, Any]) -> ReplayFixture:
    """Build a fixture (and its block) from a decoded fixture document."""

    schema = str(payload.get("schema", ""))
    if schema != REPLAY_SCHEMA:
        raise ReplayFixtureError(f"unsupported replay schema {schema!r}, expected {REPLAY_SCHEMA!r}")

    name = str(payload.get("name", "")) or "unnamed"
    calibration = _vector(payload.get("axis_calibration", (1.0, 1.0, 1.0)), "axis_calibration")
    range_value = float(payload.get("range_value", 1.0) or 1.0)
    range_factor = float(payload.get("range_factor", 1.0e-5) or 1.0e-5)
    is_up = bool(payload.get("is_up", True))
    is_holder_block = bool(payload.get("is_holder_block", False))
    holder_positions = payload.get("holder_positions")

    raw_observations = payload.get("observations")
    if not isinstance(raw_observations, Sequence) or len(raw_observations) != len(BLOCK_ROLES):
        raise ReplayFixtureError(
            f"fixture {name!r} needs exactly {len(BLOCK_ROLES)} observations"
        )

    observations: list[SquidObservation] = []
    for index, (role, entry) in enumerate(zip(BLOCK_ROLES, raw_observations)):
        recorded_role = str(entry.get("role", role))
        if recorded_role != role:
            raise ReplayFixtureError(
                f"fixture {name!r} observation {index} is {recorded_role!r}, expected {role!r}"
            )
        axes_payload = entry.get("axes") or {}
        axes: list[AxisObservation] = []
        for axis_index, axis in enumerate(AXIS_NAMES):
            values = axes_payload.get(axis)
            if values is None:
                raise ReplayFixtureError(
                    f"fixture {name!r} observation {role!r} is missing axis {axis}"
                )
            counts = float(values.get("counts", 0.0))
            dvm = float(values.get("dvm", 0.0))
            axes.append(
                AxisObservation(
                    axis=axis,
                    counts=counts,
                    dvm=dvm,
                    range_value=range_value,
                    calibration=float(calibration[axis_index]),
                    count_command=f"{axis}SC",
                    count_reply=str(values.get("count_reply", _format(counts))),
                    data_command=f"{axis}SD",
                    data_reply=str(values.get("data_reply", _format(dvm))),
                )
            )
        observations.append(
            SquidObservation(
                role=role,
                sequence_index=index,
                axes=(axes[0], axes[1], axes[2]),
                latch_id=f"replay-{index:02d}",
                latch_commands=("ALC", "ALD"),
                range_label=str(payload.get("range_label", "1")),
                turn_angle_deg=POSITION_ANGLES[index],
            )
        )

    vectors = [observation.calibrated_vector for observation in observations]
    block = BracketedMeasurementBlock(
        zero_before=vectors[0],
        positions=(vectors[1], vectors[2], vectors[3], vectors[4]),
        zero_after=vectors[5],
        holder_positions=_position4(holder_positions),
        is_up=is_up,
        range_factor=range_factor,
        # Calibration is already applied per axis above, exactly once.
        axis_calibration=(1.0, 1.0, 1.0),
        observations=tuple(observations),
        audit=BlockAudit(
            block_id=f"replay:{name}",
            sample_name=str(payload.get("sample_name", "")),
            treatment_label=str(payload.get("treatment_label", "")),
            range_label=str(payload.get("range_label", "1")),
            range_factor=range_factor,
            axis_calibration_applied=calibration,
            is_holder_block=is_holder_block,
        ),
    )

    expected = payload.get("expected") or {}
    expectation = ReplayExpectation(
        accepted=bool(expected.get("accepted", True)),
        discontinuous_axes=tuple(str(axis) for axis in expected.get("discontinuous_axes", ())),
        flux_step_axes=tuple(str(axis) for axis in expected.get("flux_step_axes", ())),
        reason_contains=str(expected.get("reason_contains", "")),
    )
    return ReplayFixture(
        name=name,
        block=block,
        expectation=expectation,
        description=str(payload.get("description", "")),
        source=str(payload.get("source", "")),
    )


class ReplaySquidBackend:
    """Measurement backend that returns recorded blocks instead of reading hardware.

    It is marked ``simulated`` so its output is labelled and kept out of the
    production output path, and it deliberately has no recovery hook by
    default: a recorded discontinuity cannot be fixed by re-zeroing, so the
    default replay proves the "reject and keep previous state" behavior.
    """

    simulated = True
    flux_discontinuity_retries = 0

    def __init__(
        self,
        fixtures: Iterable[ReplayFixture | BracketedMeasurementBlock],
        *,
        repeat_last: bool = True,
    ) -> None:
        blocks: list[BracketedMeasurementBlock] = []
        for item in fixtures:
            blocks.append(item.block if isinstance(item, ReplayFixture) else item)
        if not blocks:
            raise ReplayFixtureError("ReplaySquidBackend needs at least one recorded block")
        self._blocks = blocks
        self._repeat_last = bool(repeat_last)
        self._index = 0
        self.reads = 0
        self.recovery_calls: list[ZeroPairValidation | None] = []

    def read_squid(self) -> BracketedMeasurementBlock:
        self.reads += 1
        if self._index >= len(self._blocks):
            if not self._repeat_last:
                raise ReplayFixtureError("replay sequence is exhausted")
            return self._blocks[-1]
        block = self._blocks[self._index]
        self._index += 1
        return block

    def read_susceptibility(self) -> float:
        return 0.0

    def set_demag_step(self, label: str) -> None:
        del label

    def is_available(self) -> bool:
        return True

    def preflight(self):
        from rapid_main.hardware_contracts import PreflightResult

        return PreflightResult(
            ok=True,
            blockers=(),
            warnings=("REPLAY MODE: recorded blocks, not hardware evidence.",),
        )

    def return_to_safe_state(self) -> None:
        return


def _vector(values: Iterable[float], name: str) -> Vector3:
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise ReplayFixtureError(f"{name} requires exactly three values")
    return (items[0], items[1], items[2])


def _position4(values: Any) -> Position4:
    if values is None:
        return ((0.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0))
    items = tuple(_vector(vector, "holder_positions") for vector in values)
    if len(items) != 4:
        raise ReplayFixtureError("holder_positions requires exactly four vectors")
    return (items[0], items[1], items[2], items[3])


def _format(value: float) -> str:
    return f"{value:g}"
