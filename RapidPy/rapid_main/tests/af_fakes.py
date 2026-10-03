"""Reviewed AF fixture; never opens hardware."""
from rapid_main.config import AfDemagConfig


def configured_af(**overrides):
    values = dict(enabled=True, calibration_source="test-fixture", coil_position=-20000,
                  peak=50, axial_calibration=[[1, 100], [2, 200]],
                  transverse_calibration=[[1, 50], [2, 100]],
                  axial_calibrated=True, transverse_calibrated=True,
                  axial_min_mT=1, axial_max_mT=200, transverse_min_mT=1, transverse_max_mT=100,
                  axial_frequency_hz=100, transverse_frequency_hz=200,
                  axial_ramp_max_v=10, transverse_ramp_max_v=10,
                  axial_monitor_max_v=10, transverse_monitor_max_v=4,
                  axial_ramp_up_vps=5, transverse_ramp_up_vps=5,
                  ramp_up_min_ms=500, ramp_up_max_ms=1000,
                  ramp_down_min_periods=100, ramp_down_max_periods=5000,
                  ramp_down_periods_per_v=200, hold_peak_periods=100,
                  io_rate_hz=50000, axial_relay_bit=1, transverse_relay_bit=2)
    values.update(overrides)
    return AfDemagConfig(**values)
