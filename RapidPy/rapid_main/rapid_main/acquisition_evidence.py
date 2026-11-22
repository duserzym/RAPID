"""Lossless canonical encoding of complete bracketed acquisitions."""
from dataclasses import fields
import json
import math

from rapid_main.magnetometer import (AxisObservation, BlockAudit, BracketedMeasurementBlock,
    CommandEvent, SquidObservation)


def canonical(payload):
    return json.dumps(payload, sort_keys=True, separators=(',', ':'), allow_nan=False)


def _construct(kind, payload):
    if not isinstance(payload, dict) or set(payload) != {field.name for field in fields(kind)}:
        raise ValueError('acquisition has incomplete or unknown acquisition fields')
    return kind(**payload)


def _vector(value):
    if not isinstance(value, (list, tuple)) or len(value) != 3 or any(
            type(number) not in (float, int) or not math.isfinite(number) for number in value):
        raise ValueError('acquisition vectors must contain three finite numbers')
    return tuple(float(number) for number in value)


def _positions(value):
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        raise ValueError('acquisition requires four position vectors')
    return tuple(_vector(vector) for vector in value)


def restore_block(payload):
    """Restore the entire original acquisition, including counter/DVM evidence."""
    payload = dict(payload)
    for name in ('zero_before', 'zero_after', 'axis_calibration'):
        payload[name] = _vector(payload[name])
    for name in ('positions', 'holder_positions'):
        payload[name] = _positions(payload[name])
    observations = []
    for raw in payload['observations']:
        raw = dict(raw)
        raw['axes'] = tuple(_construct(AxisObservation, axis) for axis in raw['axes'])
        raw['latch_commands'] = tuple(raw['latch_commands'])
        observations.append(_construct(SquidObservation, raw))
    payload['observations'] = tuple(observations)
    if payload['audit'] is not None:
        audit = dict(payload['audit'])
        audit['axis_calibration_applied'] = _vector(audit['axis_calibration_applied'])
        audit['commands'] = tuple(_construct(CommandEvent, command) for command in audit['commands'])
        payload['audit'] = _construct(BlockAudit, audit)
    return _construct(BracketedMeasurementBlock, payload)
