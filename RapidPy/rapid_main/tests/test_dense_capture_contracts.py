from dataclasses import replace
import unittest
from unittest.mock import Mock, patch

from rapidpy_common.adwin_af import (AdwinAFController, AdwinBoardConfig, AdwinCoilLimits,
                                     AdwinDenseCaptureRequest, AdwinRampRequest, AdwinError, validate_dense_capture_request)


class DenseCaptureContractsTests(unittest.TestCase):
    def request(self, **changes):
        return replace(AdwinDenseCaptureRequest(800., .5, 10000., .1, active_coil="axial"), **changes)

    def controller(self):
        controller = object.__new__(AdwinAFController)
        controller.board = AdwinBoardConfig()
        controller.limits = AdwinCoilLimits()
        controller._dll = Mock()
        self.word = 0
        self.journal = []
        def boot():
            self.word = 0
            self.journal.append("boot")
        def digital(word):
            self.word = word
            self.journal.append(("relay", word))
        def start(*_args):
            self.journal.append(("start", self.word))
            return 0
        controller.boot_board = Mock(side_effect=boot)
        controller.set_digout = Mock(side_effect=digital)
        controller.get_digout = Mock(side_effect=lambda: self.word)
        controller.clear_all_processes, controller.load_process = Mock(), Mock()
        controller.set_par, controller.set_fpar = Mock(), Mock()
        controller._dll.ADB_Start.side_effect = start
        controller.get_par = Mock(side_effect=lambda index: {4: 7, 5: 12, 6: 12, 7: 0, 8: 1}[index])
        controller.get_fpar = Mock(side_effect=lambda index: {6: .0001, 7: 12.5}[index])
        controller.get_data_long = Mock(return_value=[32768, 33000])
        return controller

    def test_coil_reselected_after_board_boot_and_present_at_process_start(self):
        controller = self.controller()
        with patch("rapidpy_common.adwin_af.time.sleep"):
            capture = controller.run_dense_loopback(self.request())
        self.assertEqual(self.journal[-2:], [("relay", 1), ("start", 1)])
        self.assertEqual(len(capture.adc_v), 2)

    def test_loopback_default_keeps_coil_relays_off(self):
        controller = self.controller()
        with patch("rapidpy_common.adwin_af.time.sleep"):
            controller.run_dense_loopback(self.request(active_coil="off"))
        self.assertEqual(self.journal[-1], ("start", 0))

    def test_unsupported_requests_fail_before_boot_relay_or_process_start(self):
        for changes in ({"amplitude_v": 11.}, {"amplitude_v": float("nan")}, {"io_rate_hz": 100000.},
                        {"duration_s": 200.}, {"ramp_up_slope_vps": 0.}, {"adc_chan": 17},
                        {"dac_chan": 3}, {"sine_freq_hz": 4000.}, {"ramp_mode": 2}):
            with self.subTest(changes=changes):
                controller = self.controller()
                with self.assertRaises(AdwinError):
                    controller.run_dense_loopback(self.request(**changes))
                controller.boot_board.assert_not_called()
                controller.set_digout.assert_not_called()
                controller._dll.ADB_Start.assert_not_called()

    def test_timeout_and_pre_cancel_are_validated_before_any_output(self):
        for options in ({"timeout_s": .001}, {"should_stop": lambda: True}):
            controller = self.controller()
            with self.assertRaises(AdwinError):
                controller.run_dense_loopback(self.request(), **options)
            controller.boot_board.assert_not_called()

    def test_coil_limit_violation_is_rejected_without_clipping(self):
        controller = self.controller()
        controller.limits = AdwinCoilLimits(axial_ramp_max=.25)
        with self.assertRaisesRegex(AdwinError, "coil"):
            controller.run_dense_loopback(self.request())
        controller.boot_board.assert_not_called()

    def test_relay_readback_failure_prevents_process_start(self):
        controller = self.controller()
        controller.get_digout.return_value = 0
        controller.get_digout.side_effect = None
        with patch("rapidpy_common.adwin_af.time.sleep"):
            with self.assertRaisesRegex(AdwinError, "relay readback"):
                controller.run_dense_loopback(self.request())
        controller._dll.ADB_Start.assert_not_called()

    def test_invalid_counts_or_timing_are_rejected_before_bulk_read(self):
        for counts, timestep in (({5: 2000010, 6: 12}, .0001), ({5: 10, 6: 12}, .0001), ({5: 12, 6: 12}, float("nan"))):
            controller = self.controller()
            controller.get_par.side_effect = lambda index: {4: 7, 7: 0, 8: 1, **counts}[index]
            controller.get_fpar.side_effect = lambda index: timestep if index == 6 else 12.5
            with patch("rapidpy_common.adwin_af.time.sleep"):
                with self.assertRaisesRegex(AdwinError, "invalid timing"):
                    controller.run_dense_loopback(self.request())
            controller.get_data_long.assert_not_called()

    def test_short_bulk_arrays_cannot_be_published_as_complete_capture(self):
        controller = self.controller()
        controller.get_data_long.return_value = [32768]
        with patch("rapidpy_common.adwin_af.time.sleep"):
            with self.assertRaisesRegex(AdwinError, "incomplete"):
                controller.run_dense_loopback(self.request())

    def test_long_steady_capture_budget_includes_ramp_and_periods(self):
        request = self.request(duration_s=40., io_rate_hz=10000.)
        estimated = validate_dense_capture_request(request)
        self.assertGreater(estimated, 40.)

    def test_calibrated_af_rejects_process_rate_that_would_be_silently_clamped(self):
        from tests.af_fakes import configured_af
        from rapid_main.af_treatment import plan_calibrated_af_ramp
        cfg = configured_af()
        cfg.io_rate_hz = 100000.
        with self.assertRaisesRegex(ValueError, "50 kHz"):
            plan_calibrated_af_ramp("AFZ50", 50, "axial", cfg)

    def test_field_recovery_attempts_all_processes_and_both_outputs_after_failure(self):
        controller = self.controller()
        controller.get_par.side_effect = lambda index: 1 if index == -99 else 0
        controller._dll.ADB_Stop.return_value = 5
        controller._raise_if_error = Mock()
        controller.set_dac = Mock()
        with self.assertRaisesRegex(AdwinError, "failed with code 5"):
            controller.recover_safe_field()
        self.assertIn(-90, [call.args[0] for call in controller.get_par.call_args_list])
        self.assertEqual([call.args for call in controller.set_dac.call_args_list], [(1, 0.0), (2, 0.0)])
        controller.set_digout.assert_not_called()

    def test_ramp_invalid_timing_or_voltage_is_rejected_before_boot(self):
        request = AdwinRampRequest(1., 1., .5, 800., .5, "axial", io_rate_hz=10000.)
        for changes in ({"io_rate_hz": 100000.}, {"slope_up": 0.}, {"peak_monitor_voltage": 11.},
                        {"ramp_peak_voltage": float("nan")}, {"hold_ms": -1}, {"sine_freq_hz": 50000.}):
            controller = self.controller()
            with self.assertRaises(AdwinError):
                controller.run_ramp(replace(request, **changes))
            controller.boot_board.assert_not_called()
            controller.set_digout.assert_not_called()

    def test_ramp_respects_coil_voltage_limit_without_clipping(self):
        controller = self.controller()
        controller.limits = AdwinCoilLimits(axial_ramp_max=.25)
        with self.assertRaisesRegex(AdwinError, "voltage limit"):
            controller.run_ramp(AdwinRampRequest(1., 1., .5, 800., .5, "axial", io_rate_hz=10000.))
        controller.boot_board.assert_not_called()
