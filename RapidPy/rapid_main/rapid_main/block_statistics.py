"""VB6 MeasurementBlocks statistics over complete bracketed acquisitions."""
from dataclasses import dataclass, asdict
import math
from typing import Sequence

from rapid_main.magnetometer import BracketedMeasurementBlock, Vector3, reduce_bracketed_measurement


@dataclass(frozen=True)
class BlockCollectionStatistics:
    block_count: int
    position_count: int
    mean_raw: Vector3
    moment_emu: Vector3
    axis_sd_raw: Vector3
    axis_sd_emu: Vector3
    fischer_sd_deg: float
    sig_drift: float
    sig_holder: float
    sig_induced: float
    sig_noise: float
    up_to_down: float
    error_horizontal_deg: float

    def payload(self) -> dict:
        return asdict(self)


def _mean(vectors: Sequence[Vector3]) -> Vector3:
    return tuple(math.fsum(vector[axis] / len(vectors) for vector in vectors) for axis in range(3))


def _norm(vector: Vector3) -> float:
    return math.hypot(*vector)


def _ratio(signal: float, denominator: float) -> float:
    return signal / (denominator if denominator else 1e-9)


def _angle(vector: Vector3) -> float:
    # VB6 modVector3d.Atan2(x,y) returns [0,2pi), including pi/2 for (0,0).
    if vector[0] == 0 and vector[1] == 0:
        return math.pi / 2
    return math.atan2(vector[1], vector[0]) % (2 * math.pi)


def block_collection_statistics(blocks: Sequence[BracketedMeasurementBlock]) -> BlockCollectionStatistics:
    """Reduce every block before combining; one rejected block rejects the set.

    Raw statistics follow MeasurementBlocks.cls: positions have equal weight,
    component SD uses N-1, drift/holder use mean block magnitudes and induced
    uses the mean induced vector. Up/Down subsets retain their own directions.
    Calibration must be identical; a legacy .UP import supplies it explicitly.
    """
    blocks = tuple(blocks)
    if not blocks:
        raise ValueError("block collection cannot be empty")
    calibration = tuple(blocks[0].axis_calibration)
    factor = blocks[0].range_factor
    if not math.isfinite(factor) or factor <= 0 or isinstance(factor, bool):
        raise ValueError("block collection range factor must be finite and positive")
    if len(calibration) != 3 or any(isinstance(value, bool) or not math.isfinite(value) for value in calibration):
        raise ValueError("block collection calibration must contain three finite gains")
    for block in blocks:
        if type(block.is_up) is not bool:
            raise ValueError("block collection direction must be boolean")
        if isinstance(block.range_factor, bool) or any(isinstance(value, bool) for value in block.axis_calibration):
            raise ValueError("block collection calibration cannot contain boolean gains")
        if block.range_factor != factor or tuple(block.axis_calibration) != calibration:
            raise ValueError("block collection calibration changed between acquisitions")
    audits = [block.audit for block in blocks]
    if any(audit is not None for audit in audits):
        if any(audit is None for audit in audits):
            raise ValueError("block collection mixed audited and unaudited acquisitions")
        contexts = {(audit.sample_name, audit.treatment_label, audit.run_id, audit.config_hash,
                     audit.simulated, audit.holder_record_id) for audit in audits}
        if len(contexts) != 1:
            raise ValueError("block collection acquisition context changed")
    results = [reduce_bracketed_measurement(block) for block in blocks]
    positions = tuple(vector for result in results for vector in result.holder_frame_raw)
    count = len(positions)
    mean = _mean(positions)
    sd = tuple(math.sqrt(math.fsum((vector[axis] - mean[axis]) ** 2 for vector in positions)
                         / (count - 1)) for axis in range(3))
    units = tuple(tuple(value / _norm(vector) for value in vector) if _norm(vector) else (0., 0., 0.)
                  for vector in positions)
    resultant = _norm(tuple(math.fsum(vector[axis] for vector in units) for axis in range(3)))
    # Preserve the existing VB6 exact-alignment convention; physical acceptance
    # must separately review that historical convention rather than silently alter it.
    resultant = min(float(count), resultant)  # Bound floating-point roundoff.
    kappa = 1e-9 if resultant == count else (count - 1) / (count - resultant)
    fischer_sd = 81 / math.sqrt(kappa)
    signal = _norm(mean)
    drift = math.fsum(_norm(result.drift_raw) / len(results) for result in results)
    holder = math.fsum(_norm(result.holder_average_raw) / len(results) for result in results)
    induced = _norm(_mean(tuple(result.induced_raw for result in results)))
    up = tuple(result.mean_raw for block, result in zip(blocks, results) if block.is_up)
    down = tuple(result.mean_raw for block, result in zip(blocks, results) if not block.is_up)
    up_mean = _mean(up) if up else (0., 0., 0.)
    down_mean = _mean(down) if down else (0., 0., 0.)
    down_magnitude = _norm(down_mean)
    statistics = BlockCollectionStatistics(
        len(blocks), count, mean,
        tuple(value * gain * factor for value, gain in zip(mean, calibration)),
        sd, tuple(value * abs(gain * factor) for value, gain in zip(sd, calibration)),
        fischer_sd, _ratio(signal, drift), _ratio(signal, holder), _ratio(signal, induced),
        _ratio(signal, _norm(sd)), _norm(up_mean) / down_magnitude if down_magnitude else 0.,
        math.degrees(_angle(up_mean) - _angle(down_mean)) if up and down else 0.,
    )
    for value in statistics.payload().values():
        values = value if isinstance(value, tuple) else (value,)
        if any(not math.isfinite(number) for number in values):
            raise ValueError("block collection statistics are non-finite")
    return statistics
