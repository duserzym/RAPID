from dataclasses import asdict
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rapid_main.config import AppConfig
from rapid_main.hardware_contracts import QueueHardwareBackend
from rapid_main.queue_specimen_geometry import QueueSpecimenGeometry
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests.af_fakes import configured_af
from tests.test_pulse_circuit import circuit_config


def safe_record(schema='rapidpy.af.treatment.v1'):
    return SimpleNamespace(safe_state_confirmed=True, simulated=False, schema=schema,
        to_dict=lambda: {'safe_state_confirmed': True, 'simulated': False, 'schema': schema})


class QueueSpecimenGeometryTests(QueueLiftFixture, unittest.TestCase):
    def additional_stage_profiles(self):
        config = AppConfig()
        config.general.nocomm = False
        config.motion.zero_pos, config.motion.meas_pos = -50000, -97500
        config.motion.sample_bottom, config.motion.sample_top = -5000, -2000
        station = config.motor_station
        station.calibration_source, station.hole_slot, station.use_xy_table = 'station.ini', 46, True
        station.xy_home = [-3, -2]
        station.xy_positions = {str(slot): [x, y] for slot, x, y in self.geometry.xy_positions}
        station.controller = asdict(self.motor.config)
        station.ports = {key: axis.port for key, axis in self.axes.items()}
        station.addresses = {key: axis.address for key, axis in self.axes.items()}
        config.af_demag = configured_af()
        config.pulse_irm = circuit_config()
        config.susceptibility.coil_position = -22700
        config.susceptibility.moment_factor_cgs = .0000097914
        backend = object.__new__(QueueHardwareBackend)
        backend._config, backend._client, backend._axes = config, self.motor, self.axes
        backend._safety_store = self.store
        backend._sample_name, backend._run_id = 'S1', 'run-1'
        backend._queue_specimen_geometry = None
        backend._geometry_stage_token = None
        backend._bracketed_geometry_signature = None
        backend._bracketed, backend._acquisition_clock, backend._halt_check = None, None, None
        backend._measurement = Mock(simulated=False, raw_client=Mock())
        backend._af_demag, backend._pulse_irm = Mock(simulated=False), Mock(simulated=False)
        backend._arm_bias = None
        backend._af_treatment_records, backend._pulse_treatment_records = [], []
        backend._connected = True
        backend._measuring_holder = False
        backend._susceptibility = Mock(simulated=False)
        self.backend = backend
        return dict({family: backend._safety_profile() for family in ('af', 'arm', 'pulse', 'rrm')},
                    acquisition=backend._acquisition_safety_profile())

    def setUp(self):
        super().setUp()
        self.load()
        self.backend._safety_store = self.session.child_store
        with self.session.claim(): self.backend.bind_specimen_geometry(self.session)

    def test_verified_live_height_does_not_modify_station_calibration(self):
        with self.session.claim():
            self.assertEqual(self.backend._sample_height(), 2800)
            self.assertEqual(self.backend._queue_specimen_geometry.positions('S1'), (-48600, -96100))
        self.assertEqual(self.backend._config.motion.sample_height, 3000)
        self.assertEqual(self.motor.config.sample_height, 3000)

    def test_unclaimed_wrong_sample_and_changed_station_fail_closed(self):
        with self.assertRaises(HardwareSafetyError): self.backend._sample_height()
        with self.session.claim():
            self.backend._sample_name = 'S2'
            with self.assertRaises(HardwareSafetyError): self.backend._sample_height()
            self.backend._sample_name = 'S1'
            self.backend._config.motor_station.xy_home[0] += 1
            with self.assertRaises(HardwareSafetyError): self.backend._sample_height()

    def test_af_plan_and_durable_stage_use_measured_height(self):
        def execute(plan, **kwargs):
            self.assertEqual(plan.target_position, -18600)
            self.assertEqual(self.backend._sample_height(), 2800)
            self.assertEqual(self.store.pending()['stage']['plan']['specimen_geometry']['sample_height'], 2800)
            from rapid_main.hardware_contracts import HardwareError
            with self.assertRaises(HardwareError): self.backend._ensure_bracketed()
            return safe_record()
        with patch('rapid_main.af_treatment.AfTreatmentService') as service, self.session.claim():
            service.return_value.execute.side_effect = execute
            self.backend._execute_af_treatment('AF50')
        self.assertIsNone(self.backend._geometry_stage_token)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')

    def test_arm_uses_same_measured_specimen_centre(self):
        with patch('rapid_main.af_treatment.AfTreatmentService') as service, self.session.claim():
            service.return_value.execute.return_value = safe_record()
            self.backend._execute_af_treatment('ARM50', peak_af_mT=50, bias_mT=0)
            self.assertEqual(service.return_value.execute.call_args.args[0].target_position, -18600)

    def test_pulse_execution_receives_measured_height_before_stage_begins(self):
        with patch('rapid_main.pulse_treatment.PulseTreatmentService') as service, self.session.claim():
            service.return_value.execute.return_value = safe_record('rapidpy.irm.treatment.v1')
            self.backend._execute_pulse_treatment(50, axis='Z')
            self.assertEqual(service.return_value.execute.call_args.kwargs['sample_height'], 2800)

    def test_rrm_plan_uses_measured_height_and_preserves_native_speed_calibration(self):
        from rapid_main.rrm_treatment import plan_rrm_treatment
        with self.session.claim():
            plan = plan_rrm_treatment('RRMZ50/1', self.backend._config, sample_height=self.backend._sample_height())
        self.assertEqual(plan.target_position, -18600)
        self.assertEqual(plan.native_1rps, self.motor.config.turning_motor_1rps)

    def test_bracketed_zero_and_measurement_use_measured_specimen_centre(self):
        with self.session.claim():
            self.backend._ensure_bracketed()
            config = self.backend._bracketed._service.config
            self.assertEqual((config.zero_position, config.measurement_position), (-48600, -96100))
            original = self.backend._bracketed
            self.backend._ensure_bracketed()
            self.assertIs(self.backend._bracketed, original)

    def test_pending_unrelated_stage_cannot_borrow_cached_geometry(self):
        with self.session.claim() as child:
            child.begin('acquisition', {}, self.backend._acquisition_safety_profile(), sample_id='S1')
            with self.assertRaises(HardwareSafetyError): self.backend._sample_height()

    def test_returned_specimen_cannot_keep_using_cached_height(self):
        self.call('lower_for_dropoff')
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.backend._sample_height()

    def test_susceptibility_config_receives_measured_height(self):
        record = SimpleNamespace(sample_id='S1', is_holder=False)
        with patch('rapid_main.susceptibility_acquisition.SusceptibilityAcquisitionService') as service, self.session.claim():
            service.return_value.acquire.return_value = record
            self.backend._susceptibility_records = []
            self.backend._acquire_susceptibility(sample_id='S1', is_holder=False)
            config = service.call_args.kwargs['config']
            self.assertEqual((config.sample_height, config.target_position), (2800, -21300))

    def test_geometry_can_only_be_cleared_after_verified_original_slot_return(self):
        from rapid_main.hardware_contracts import HardwareError
        with self.session.claim(), self.assertRaises(HardwareError): self.backend.clear_specimen_geometry()
        self.call('lower_for_dropoff')
        with self.session.claim(): self.outputs(False)
        self.call('clear_after_release')
        with self.session.claim(): self.backend.clear_specimen_geometry()
        self.assertIsNone(self.backend._queue_specimen_geometry)

    def test_odd_measured_height_uses_vb6_floor_for_negative_targets(self):
        self.call('lower_for_dropoff')
        with self.session.claim(): self.outputs(False)
        self.call('clear_after_release')
        self.lift_serial().home_offset = 199
        self.load()
        with self.session.claim():
            self.backend.bind_specimen_geometry(self.session)
            self.assertEqual(self.backend._sample_height(), 2801)
            self.backend._ensure_bracketed()
            config = self.backend._bracketed._service.config
            self.assertEqual((config.zero_position, config.measurement_position), (-48600, -96100))
        from rapid_main.susceptibility_acquisition import SusceptibilityAcquisitionConfig
        self.assertEqual(SusceptibilityAcquisitionConfig(-22700, 2801, 1).target_position, -21300)

    def test_changed_scientific_positions_or_wiring_cannot_reuse_loaded_geometry(self):
        for section, name, value in (('motion', 'zero_pos', -51000),
                                     ('squid', 'settle_time', 3),
                                     ('susceptibility', 'coil_position', -22000)):
            target = getattr(self.backend._config, section)
            original = getattr(target, name)
            setattr(target, name, value)
            with self.subTest(section=section), self.session.claim(), self.assertRaises(HardwareSafetyError):
                self.backend._sample_height()
            setattr(target, name, original)

    def test_foreign_stage_store_cannot_execute_a_measured_queue_treatment(self):
        from rapid_main.hardware_contracts import HardwareError
        from rapid_main.af_treatment import plan_af_treatment
        self.backend._safety_store = self.store
        plan = plan_af_treatment('AF50', self.backend._config.af_demag, 2800)
        before = self.store.pending()['stage_count']
        with self.session.claim(), self.assertRaises(HardwareError):
            with self.backend._safety_operation('af', plan):
                self.fail('A foreign store borrowed the geometry.')
        self.assertEqual(self.store.pending()['stage_count'], before)
