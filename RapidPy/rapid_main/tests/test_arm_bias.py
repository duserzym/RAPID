"""ARM independent bias circuitry and checked MCC driver calls; no hardware."""
import ctypes
from dataclasses import asdict
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from rapid_main.arm_bias import ArmBiasBackend, plan_arm_bias
from rapid_main.config import IrmArmConfig
from rapid_main.af_treatment import AfTreatmentService, AfTreatmentError, plan_af_treatment
from rapidpy_common.mcc_daq import MccDaq, MccError
from tests.acquisition_fakes import FakeClock
from tests.af_fakes import configured_af


def bias_config(**overrides):
    values = dict(arm_enabled=True, arm_calibration_source="test-fixture",
                  arm_bias_max_mT=1, arm_voltage_per_mT=2, arm_voltage_max=2,
                  arm_board=0, arm_dac_channel=0, arm_gate_bit=0)
    values.update(overrides)
    return IrmArmConfig(**values)


class ArmBiasTests(unittest.TestCase):
    def adapter(self):
        self.journal = []
        self.controller = Mock()
        self.controller.analog_output.side_effect = lambda channel,voltage,vrange: self.journal.append(("voltage",voltage))
        self.controller.digital_output.side_effect = lambda port,bit,high: self.journal.append(("gate",high))
        adapter = ArmBiasBackend(bias_config(), controller=self.controller)
        self.clock = FakeClock()
        adapter.sleep = self.clock.sleep
        return adapter

    def test_setup_and_clear_follow_legacy_zero_and_active_low_gate_sequence(self):
        adapter = self.adapter()
        adapter.set_bias_mT(.5)
        self.assertEqual(self.journal, [("voltage",0),("gate",False),("voltage",1)])
        self.assertAlmostEqual(sum(self.clock.slept), 2)
        adapter.clear_bias()
        self.assertEqual(self.journal[-2:], [("voltage",0),("gate",True)])
        self.assertAlmostEqual(sum(self.clock.slept), 2.5)
        self.assertEqual(len(adapter.communication_events()), 10)

    def test_bias_failure_still_attempts_gate_disconnection(self):
        adapter = self.adapter()
        self.controller.analog_output.side_effect = RuntimeError("DAC fault")
        with self.assertRaisesRegex(RuntimeError, "DAC fault"):
            adapter.clear_bias()
        self.controller.digital_output.assert_called_once_with(1,0,True)

    def test_af_ramp_remains_unchanged_and_bias_is_cleared_before_return(self):
        adapter = self.adapter()
        af = Mock(simulated=False)
        af.is_connected.return_value = True
        af.apply_calibrated_af.side_effect = lambda ramp: self.journal.append(("ramp",ramp.field_mT))
        vertical, turning = Mock(), Mock()
        outcome = SimpleNamespace(ok=True, actual=0, target=0, detail="")
        vertical.home_to_top.return_value = vertical.move_to.return_value = outcome
        turning.rotate_to.return_value = outcome
        service = AfTreatmentService(af,vertical,turning,bias_adapter=adapter,bias_mT=.5,
                                     sleep=self.clock.sleep,clock=self.clock.now)
        record = service.execute(plan_af_treatment("AFZ50",configured_af(),3000),sample_id="S1",run_id="R1")
        self.assertIn(("ramp",50), self.journal)
        self.assertEqual(record.bias_mT, .5)
        self.assertTrue(record.safe_state_confirmed)
        self.assertIn("bias_clear", [p.name for p in record.phases])

    def test_cancel_during_bias_setup_clears_outputs_without_any_af_ramp(self):
        adapter = self.adapter()
        af = Mock(simulated=False)
        af.is_connected.return_value = True
        vertical, turning = Mock(), Mock()
        vertical.home_to_top.return_value = SimpleNamespace(ok=True, actual=0)
        service = AfTreatmentService(af,vertical,turning,bias_adapter=adapter,bias_mT=.5,
                    sleep=self.clock.sleep,clock=self.clock.now,should_cancel=lambda: bool(self.clock.slept))
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(plan_af_treatment("AFZ50",configured_af(),3000),sample_id="S1",run_id="R1")
        self.assertTrue(result.exception.record.safe_state_confirmed)
        self.assertEqual(self.journal[-2:], [("voltage",0),("gate",True)])
        af.apply_calibrated_af.assert_not_called()
        vertical.move_to.assert_not_called()

    def test_invalid_bias_or_configuration_never_changes_config_or_outputs(self):
        for cfg,bias in ((bias_config(),-1),(bias_config(),1.01),(bias_config(),float("nan")),
                         (bias_config(arm_voltage_max=.2),.5),(bias_config(arm_enabled=False),.5),
                         (bias_config(arm_gate_bit=-1),.5)):
            before = asdict(cfg)
            with self.subTest(bias=bias), self.assertRaises(ValueError):
                plan_arm_bias(bias,cfg)
            self.assertEqual(before,asdict(cfg))


class MccDriverTests(unittest.TestCase):
    def driver(self):
        dll = Mock()
        def board_name(board, buffer):
            buffer.value = b"PCI-DAS6030"
            return 0
        def conversion(board,vrange,voltage,pointer):
            ctypes.cast(pointer,ctypes.POINTER(ctypes.c_ushort)).contents.value = 1234
            return 0
        dll.cbGetBoardName.side_effect = board_name
        dll.cbFromEngUnits.side_effect = conversion
        for name in ("cbAOut","cbDConfigBit","cbDBitOut"):
            getattr(dll,name).return_value = 0
        self.dll = dll
        return MccDaq(0,dll=dll)

    def test_probe_is_read_only_and_dac_conversion_is_used(self):
        driver = self.driver()
        self.assertEqual(driver.board_name,"PCI-DAS6030")
        self.dll.cbAOut.assert_not_called()
        self.dll.cbDConfigBit.assert_not_called()
        self.assertEqual(driver.analog_output(0,1,100),1234)
        self.dll.cbAOut.assert_called_once_with(0,0,100,1234)

    def test_conversion_failure_cannot_write_analog_output(self):
        driver = self.driver()
        self.dll.cbFromEngUnits.side_effect = None
        self.dll.cbFromEngUnits.return_value = 7
        with self.assertRaisesRegex(MccError,"cbFromEngUnits.*7"):
            driver.analog_output(0,1,100)
        self.dll.cbAOut.assert_not_called()

    def test_bit_configuration_failure_cannot_write_digital_output(self):
        driver = self.driver()
        self.dll.cbDConfigBit.return_value = 5
        with self.assertRaisesRegex(MccError,"cbDConfigBit.*5"):
            driver.digital_output(1,0,False)
        self.dll.cbDBitOut.assert_not_called()

    def test_bit_configuration_cached_only_after_acknowledgment(self):
        driver = self.driver()
        driver.digital_output(1,0,False)
        driver.digital_output(1,0,True)
        self.dll.cbDConfigBit.assert_called_once_with(0,1,0,1)
        self.assertEqual(self.dll.cbDBitOut.call_count,2)

    def test_adc_read_uses_vendor_conversion_and_checks_each_status(self):
        driver = self.driver()
        def read(board,channel,vrange,pointer):
            ctypes.cast(pointer,ctypes.POINTER(ctypes.c_ushort)).contents.value = 1974
            return 0
        def convert(board,vrange,raw,pointer):
            self.assertEqual(raw,1974)
            ctypes.cast(pointer,ctypes.POINTER(ctypes.c_float)).contents.value = 1.2345
            return 0
        self.dll.cbAIn.side_effect = read
        self.dll.cbToEngUnits.side_effect = convert
        self.assertAlmostEqual(driver.analog_input(0,100),1.2345,places=5)
        self.dll.cbAIn.side_effect = None
        self.dll.cbAIn.return_value = 9
        self.dll.cbToEngUnits.reset_mock()
        with self.assertRaisesRegex(MccError,"cbAIn.*9"):
            driver.analog_input(0,100)
        self.dll.cbToEngUnits.assert_not_called()
