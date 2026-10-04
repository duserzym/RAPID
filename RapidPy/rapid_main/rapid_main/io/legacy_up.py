"""Raw VB6 .UP interchange, as written by Sample.WriteUpMeasurements.

Each acquisition block contains two Z, four S and four H rows. The file does
not record calibration, treatment identity or acquisition audit evidence; those
must come from the enclosing per-step source context. Reading it never creates
live observations or a hardware audit.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import math
import re
from typing import Sequence

from rapid_main.magnetometer import BracketedMeasurementBlock, Position4, Vector3

HEADER = "Sample|Direction|Blocks|MsmtType|Block|MsmtNum|X,Y,Z"
ROW_ORDER = (("Z", 1), ("Z", 2), *(('S', i) for i in range(1, 5)),
             *(('H', i) for i in range(1, 5)))


def _vector(values: Sequence[float]) -> Vector3:
    if len(values) != 3 or any(isinstance(value, bool) for value in values):
        raise ValueError(".UP vectors require three finite numbers")
    result = tuple(float(value) for value in values)
    if not all(math.isfinite(value) for value in result):
        raise ValueError(".UP vectors require three finite numbers")
    return result


def _sample(value: str) -> str:
    if not isinstance(value, str) or not value.strip() or any(
        char == '|' or ord(char) < 32 for char in value
    ):
        raise ValueError("invalid .UP specimen identity")
    return value


def _timestamp(value: str | None) -> str | None:
    if value is None:
        return None  # Older seven-column records have no timestamp.
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", value):
        raise ValueError("invalid .UP timestamp")
    datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    return value


@dataclass(frozen=True)
class LegacyUpBlock:
    zero_before: Vector3
    positions: Position4
    zero_after: Vector3
    holder_positions: Position4
    is_up: bool
    timestamps: tuple[str | None, ...]

    @classmethod
    def from_measurement_block(cls, block: BracketedMeasurementBlock, *,
                               timestamps: Sequence[str | None]) -> LegacyUpBlock:
        """Export raw vectors; callers retain the full audit in a separate artifact."""
        result = cls(block.zero_before, block.positions, block.zero_after,
                     block.holder_positions, block.is_up, tuple(timestamps))
        vectors = _block_vectors(result)
        return cls(vectors[0], tuple(vectors[2:6]), vectors[1], tuple(vectors[6:10]),
                   result.is_up, result.timestamps)

    def to_measurement_block(self, *, range_factor: float,
                             axis_calibration: Vector3) -> BracketedMeasurementBlock:
        """Restore raw data with explicit external calibration, without an audit."""
        if isinstance(range_factor, bool) or not math.isfinite(range_factor) or range_factor <= 0:
            raise ValueError(".UP range factor must be finite and positive")
        calibration = _vector(axis_calibration)
        vectors = _block_vectors(self)
        return BracketedMeasurementBlock(
            zero_before=vectors[0], positions=tuple(vectors[2:6]),
            zero_after=vectors[1], holder_positions=tuple(vectors[6:10]),
            is_up=self.is_up, range_factor=float(range_factor), axis_calibration=calibration,
        )


@dataclass(frozen=True)
class LegacyUpRun:
    sample: str
    blocks: tuple[LegacyUpBlock, ...]


def _block_vectors(block: LegacyUpBlock) -> tuple[Vector3, ...]:
    if type(block.is_up) is not bool or len(block.positions) != 4 or len(block.holder_positions) != 4:
        raise ValueError("invalid .UP block shape or direction")
    if len(block.timestamps) != 10:
        raise ValueError(".UP block needs ten row timestamps")
    for timestamp in block.timestamps:
        _timestamp(timestamp)
    return tuple(_vector(vector) for vector in (
        block.zero_before, block.zero_after, *block.positions, *block.holder_positions))


def parse_up_file(text: str) -> tuple[LegacyUpRun, ...]:
    """Read every complete run; malformed/truncated records fail closed.

    Repeated specimens are distinct runs. Unlike VB6's permissive Val parsing,
    corrupted counters, missing vectors and partial trailing runs are rejected.
    """
    lines = text.splitlines()
    while lines and not lines[-1]:
        lines.pop()
    if not lines or lines[0] != HEADER:
        raise ValueError("invalid .UP header")
    runs = []
    cursor = 1
    while cursor < len(lines):
        initial = lines[cursor].split('|')
        if len(initial) not in (7, 8):
            raise ValueError("invalid .UP row column count")
        sample, direction, count_text = initial[:3]
        _sample(sample)
        if direction not in ('U', 'D') or not re.fullmatch(r'[1-9]\d*', count_text):
            raise ValueError("invalid .UP direction or block count")
        count = int(count_text)
        if count > (len(lines) - cursor) // 10:
            raise ValueError("truncated .UP run")
        blocks = []
        for number in range(1, count + 1):
            vectors, timestamps = [], []
            for role, position in ROW_ORDER:
                fields = lines[cursor].split('|')
                cursor += 1
                expected = [sample, direction, count_text, role, str(number), str(position)]
                if len(fields) not in (7, 8) or fields[:6] != expected:
                    raise ValueError("inconsistent .UP run identity, count or row order")
                vectors.append(_vector(fields[6].split(',')))
                timestamps.append(_timestamp(fields[7] if len(fields) == 8 else None))
            blocks.append(LegacyUpBlock(vectors[0], tuple(vectors[2:6]), vectors[1],
                                       tuple(vectors[6:10]), direction == 'U', tuple(timestamps)))
        runs.append(LegacyUpRun(sample, tuple(blocks)))
    return tuple(runs)


def read_up_measurements(text: str, sample: str) -> LegacyUpRun:
    """Select the latest complete Up run for the exact specimen, as in VB6."""
    _sample(sample)
    matches = [run for run in parse_up_file(text)
               if run.sample == sample and run.blocks[0].is_up]
    if not matches:
        raise ValueError("specimen has no complete Up measurements in .UP file")
    return matches[-1]


def encode_up_file(runs: Sequence[LegacyUpRun]) -> str:
    """Emit VB6's header and ten ordered pipe-delimited rows per raw block."""
    lines = [HEADER]
    for run in runs:
        _sample(run.sample)
        if not run.blocks:
            raise ValueError(".UP run must contain at least one block")
        for number, block in enumerate(run.blocks, start=1):
            vectors = _block_vectors(block)
            if block.is_up != run.blocks[0].is_up:
                raise ValueError("mixed directions in .UP run")
            direction = 'U' if block.is_up else 'D'
            for (role, position), vector, timestamp in zip(ROW_ORDER, vectors, block.timestamps):
                fields = [run.sample, direction, str(len(run.blocks)), role, str(number), str(position),
                          ','.join(format(value, '.7E') for value in vector)]
                if timestamp is not None:
                    fields.append(timestamp)
                lines.append('|'.join(fields))
    return '\r\n'.join(lines) + '\r\n'
