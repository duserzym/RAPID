"""Independent calibrated ARM DC bias, including active-low gate cleanup."""
import math
import time
from .communication_log import CommunicationLogger
from rapidpy_common.mcc_daq import MccDaq


def plan_arm_bias(bias_mT, cfg):
    bias = float(bias_mT)
    if not cfg.arm_enabled or not cfg.arm_calibration_source.strip():
        raise ValueError("Enabled and calibrated ARM bias circuitry is required.")
    if any(not math.isfinite(float(value)) or value <= 0 for value in
           (cfg.arm_bias_max_mT, cfg.arm_voltage_per_mT, cfg.arm_voltage_max)):
        raise ValueError("ARM bias field/voltage limits and conversion must be positive and finite.")
    if not math.isfinite(bias) or not 0 <= bias <= cfg.arm_bias_max_mT:
        raise ValueError("ARM bias is outside its calibrated field limits.")
    voltage = bias * cfg.arm_voltage_per_mT
    if not 0 <= voltage <= min(10, cfg.arm_voltage_max):
        raise ValueError("ARM bias exceeds its calibrated voltage limit.")
    if any(not isinstance(value, int) or value < 0 for value in
           (cfg.arm_board, cfg.arm_dac_channel, cfg.arm_gate_bit, cfg.arm_digital_port, cfg.arm_voltage_range)):
        raise ValueError("ARM MCC board, DAC channel and gate must be configured.")
    return voltage


class ArmBiasBackend:
    simulated = False

    def __init__(self, cfg, *, controller=None):
        plan_arm_bias(0, cfg)
        self.cfg = cfg
        self.controller = controller if controller is not None else MccDaq(cfg.arm_board)
        if getattr(self.controller, "simulated", False) is True:
            raise ValueError("A simulated DAQ cannot supply live ARM bias.")
        self.logger = CommunicationLogger("MCC_ARM", port=f"board:{cfg.arm_board}")
        self.should_cancel = lambda: False
        self.sleep = time.sleep

    def is_connected(self):
        return self.controller is not None

    def status(self):
        return "Independent MCC ARM bias circuit connected"

    def communication_events(self):
        return tuple(self.logger.transcript.events)

    def _write(self, action, operation):
        self.logger.sent(action, detail="ARM bias circuit output")
        try:
            result = operation()
        except Exception as exc:
            self.logger.error(str(exc), payload=action)
            raise
        self.logger.received(str(result), detail="MCC output acknowledged; physical field readback pending")

    def _voltage(self, voltage):
        self._write(f"DAC channel={self.cfg.arm_dac_channel} voltage={voltage:g}",
                    lambda: self.controller.analog_output(self.cfg.arm_dac_channel, voltage, self.cfg.arm_voltage_range))

    def _gate(self, high):
        self._write(f"gate bit={self.cfg.arm_gate_bit} high={int(high)}",
                    lambda: self.controller.digital_output(self.cfg.arm_digital_port, self.cfg.arm_gate_bit, high))

    def _wait(self, seconds, *, cancellable=True):
        remaining = seconds
        while remaining > 0:
            if cancellable and self.should_cancel():
                raise InterruptedError("ARM bias setup cancelled.")
            interval = min(.1, remaining)
            self.sleep(interval)
            remaining -= interval

    def set_bias_mT(self, bias_mT):
        voltage = plan_arm_bias(bias_mT, self.cfg)
        try:
            self._set_bias_voltage(voltage)
        except Exception as exc:
            try:
                self.clear_bias()
            except Exception as cleanup:
                raise RuntimeError(f"{exc}; {cleanup}") from exc
            raise

    def _set_bias_voltage(self, voltage):
        self._voltage(0)
        self._wait(.5)
        if voltage > 0:
            self._gate(False)
            self._wait(.5)
            self._voltage(voltage)
            self._wait(1)
        else:
            self.clear_bias()

    def clear_bias(self):
        errors = []
        for operation in (lambda: self._voltage(0), lambda: self._wait(.5, cancellable=False), lambda: self._gate(True)):
            try:
                operation()
            except Exception as exc:
                errors.append(str(exc))
        if errors:
            raise RuntimeError("ARM bias cleanup failed: " + "; ".join(errors))
