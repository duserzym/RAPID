"""Reproducible, immutable evidence for an accepted holder collection."""
from dataclasses import asdict, fields
from functools import lru_cache
import hashlib
import json
import math

from rapid_main.block_statistics import block_collection_statistics
from rapid_main.magnetometer import (AxisObservation, BlockAudit, BracketedMeasurementBlock,
    CommandEvent, SquidObservation, reduce_bracketed_measurement)

SCHEMA = 'rapidpy.holder_collection.v1'


def canonical(payload):
    return json.dumps(payload, sort_keys=True, separators=(',', ':'), allow_nan=False)


def _construct(kind, payload):
    if not isinstance(payload, dict) or set(payload) != {field.name for field in fields(kind)}:
        raise ValueError('holder collection has incomplete or unknown acquisition fields')
    return kind(**payload)


def _vector(value):
    if not isinstance(value, (list, tuple)) or len(value) != 3 or any(
            type(number) not in (float, int) or not math.isfinite(number) for number in value):
        raise ValueError('holder collection vectors must contain three finite numbers')
    return tuple(float(number) for number in value)


def _positions(value):
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        raise ValueError('holder collection requires four position vectors')
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
        observations.append(_construct(SquidObservation, raw))
    payload['observations'] = tuple(observations)
    if payload['audit'] is not None:
        audit = dict(payload['audit'])
        audit['axis_calibration_applied'] = _vector(audit['axis_calibration_applied'])
        audit['commands'] = tuple(_construct(CommandEvent, command) for command in audit['commands'])
        payload['audit'] = _construct(BlockAudit, audit)
    return _construct(BracketedMeasurementBlock, payload)


def derive(blocks):
    blocks = tuple(blocks)
    statistics = block_collection_statistics(blocks)
    if len({block.is_up for block in blocks}) != 1:
        raise ValueError('holder collection cannot mix directions')
    if any(any(any(number != 0 for number in vector) for vector in block.holder_positions) for block in blocks):
        raise ValueError('holder collection must use a blank holder correction')
    audits = [block.audit for block in blocks if block.audit is not None]
    if any(audit.is_holder_block is not True for audit in audits):
        raise ValueError('holder collection contains a specimen acquisition')
    identifiers = [audit.block_id for audit in audits if audit.block_id]
    if len(identifiers) != len(set(identifiers)):
        raise ValueError('holder collection repeats an acquisition identity')
    results = tuple(reduce_bracketed_measurement(block) for block in blocks)
    count = len(results)
    positions = tuple(tuple(math.fsum(result.baseline_adjusted_raw[position][axis] / count
                                     for result in results) for axis in range(3)) for position in range(4))
    magnitude = math.hypot(*statistics.mean_raw)
    induced = tuple(math.fsum(result.induced_raw[axis] / count for result in results) for axis in range(3))
    induced_magnitude = math.hypot(*induced)
    metrics = dict(magnitude_raw=magnitude, magnitude_emu=math.hypot(*statistics.moment_emu),
                   induced_magnitude_raw=induced_magnitude,
                   asymmetry_ratio=induced_magnitude / magnitude if magnitude else 0.,
                   drift_magnitude_raw=math.fsum(math.hypot(*result.drift_raw) / count for result in results),
                   fischer_sd_deg=statistics.fischer_sd_deg)
    canonical(metrics)  # Reject non-finite aggregate values before installation.
    return positions, metrics, statistics, results


def make_collection(blocks, *, holder_id, hole, measured_at_iso):
    """Freeze all raw/audit/observation data and bind it to the holder identity."""
    # Round-trip first to detach any caller-owned mutable containers.
    raw_blocks = json.loads(canonical([asdict(block) for block in blocks]))
    frozen_blocks = tuple(restore_block(raw) for raw in raw_blocks)
    positions, metrics, statistics, results = derive(frozen_blocks)
    payload = dict(schema=SCHEMA, holder_id=holder_id, hole=hole, measured_at_iso=measured_at_iso,
                   blocks=raw_blocks, positions=positions, metrics=metrics, statistics=statistics.payload())
    text = canonical(payload)
    return text, hashlib.sha256(text.encode('utf-8')).hexdigest(), positions, metrics, results


def verify_collection(text, digest):
    """Check the digest and rederive every aggregate from the retained raw data."""
    if not isinstance(text, str) or not isinstance(digest, str) or hashlib.sha256(text.encode('utf-8')).hexdigest() != digest:
        raise ValueError('holder collection evidence digest changed')
    payload = json.loads(text)
    if set(payload) != {'schema', 'holder_id', 'hole', 'measured_at_iso', 'blocks', 'positions', 'metrics', 'statistics'} or payload['schema'] != SCHEMA:
        raise ValueError('invalid holder collection schema')
    blocks = tuple(restore_block(raw) for raw in payload['blocks'])
    positions, metrics, statistics, results = derive(blocks)
    if canonical(payload['positions']) != canonical(positions) or payload['metrics'] != metrics or payload['statistics'] != json.loads(canonical(statistics.payload())):
        raise ValueError('holder collection aggregates disagree with raw acquisitions')
    return payload, blocks, results


@lru_cache(maxsize=8)
def verified_top_fields(text, digest):
    """Cache only an immutable proof summary for immutable evidence strings."""
    payload, blocks, results = verify_collection(text, digest)
    block, result = blocks[-1], results[-1]
    audit = block.audit
    expected = dict(holder_id=payload['holder_id'], hole=payload['hole'],
        measured_at_iso=payload['measured_at_iso'], averaging_cycles=len(blocks),
        positions=payload['positions'], metrics=payload['metrics'], is_up=block.is_up,
        range_factor=block.range_factor, range_label=getattr(audit, 'range_label', ''),
        axis_calibration=getattr(audit, 'axis_calibration_applied', block.axis_calibration),
        raw_positions=block.positions, holder_frame_positions=result.holder_frame_raw,
        zero_before=block.zero_before, zero_after=block.zero_after,
        block_id=getattr(audit, 'block_id', ''), run_id=getattr(audit, 'run_id', ''),
        config_hash=getattr(audit, 'config_hash', ''), operator=getattr(audit, 'operator', ''),
        software_version=getattr(audit, 'software_version', ''), simulated=bool(getattr(audit, 'simulated', False)),
        validation_reason=result.validation.reason, validation_deltas=result.validation.deltas)
    return canonical(expected)
