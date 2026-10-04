"""Native blank rod poses, zero-height acquisition and immutable queue ownership."""
from dataclasses import replace
import unittest
from unittest.mock import patch

from rapid_main.acquisition import BlockContext
from rapid_main.holder_state import HolderStateStore
from rapid_main.queue_holder_geometry import QueueBlankHolderMotion, QueueHolderState
from rapid_main.squid_transport import BracketedSquidBackend, MotorVerticalController, MotorTurningController
from rapid_main.susceptibility_acquisition import SusceptibilityAcquisitionConfig, SusceptibilityAcquisitionService, SusceptibilityAcquisitionError
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests import test_queue_specimen_geometry as geometry_fixture
from tests import test_susceptibility_acquisition as bridge_fixture
from tests.test_acquisition import _build_service


class QueueHolderGeometryTests(QueueLiftFixture, unittest.TestCase):
    def additional_stage_profiles(self):
        profiles = geometry_fixture.QueueSpecimenGeometryTests.additional_stage_profiles(self)
        self.backend._config.susceptibility.enabled = True
        profiles['acquisition'] = self.backend._acquisition_safety_profile()
        return profiles

    def setUp(self):
        super().setUp()
        self.backend._safety_store = self.session.child_store
        self.backend._holder_store = HolderStateStore(self.path.parent / 'holder.json')
        self.backend._direction_up = True
        self.backend._susceptibility_records = []
        self.backend._susceptibility = bridge_fixture._Bridge(12.5)
        self.backend._susceptibility.test_connection = lambda: True
        self.blank = QueueBlankHolderMotion(self.table, self.vacuum)
        with self.session.claim():
            self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=True)

    def prepare(self):
        with self.session.claim():
            return self.blank.prepare(self.session, 46, reference_verified=True, field_outputs_off_verified=True)

    def bind(self):
        self.prepare()
        with self.session.claim(): self.backend.bind_holder_geometry(self.session, self.vacuum)

    def injected_squid(self):
        service, transport, vertical, turning, clock = _build_service()
        service._config = replace(service.config, zero_position=-50000, measurement_position=-97500)
        service._vertical = MotorVerticalController(self.motor, self.axes['updown'])
        service._turning = MotorTurningController(self.motor, self.axes['turning'])
        squid = BracketedSquidBackend(service, transport_retries=0,
            context_provider=lambda: BlockContext(sample_name=self.backend._sample_name,
                run_id='run-1', is_holder_block=True))
        def ensure():
            self.assertEqual(self.store.pending()['stage']['family'], 'acquisition')
            self.assertEqual(self.backend._sample_height(), 0)
            self.backend._bracketed = squid
        return patch.object(self.backend, '_ensure_bracketed', ensure)

    def test_pose_is_distinct_verified_zero_height_context(self):
        record = self.prepare()
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.holder_context['sample_height'], 0)
        self.assertIsNone(self.store.latest_transfer_context(self.session.token))
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertEqual(len(record.cleanup_observations), 4)

    def test_empty_holder_squid_uses_unshifted_positions_and_no_correction(self):
        self.bind()
        with self.injected_squid(), self.session.claim(): block = self.backend.read_squid()
        self.assertTrue(block.audit.is_holder_block)
        self.assertEqual((block.audit.zero_position, block.audit.measurement_position), (-50000, -97500))
        self.assertEqual(block.holder_positions, ((0., 0., 0.),) * 4)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertEqual(self.store.pending()['stage']['record']['evidence']['block']['audit']['sample_name'], 'holder-046')

    def test_blank_susceptibility_zero_height_requires_no_prior_correction(self):
        self.bind()
        with self.session.claim(): value = self.backend.read_susceptibility()
        self.assertEqual(value, 12.5)
        record = self.backend._susceptibility_records[-1]
        self.assertEqual((record.sample_height, record.target_position), (0, -22700))
        self.assertTrue(record.is_holder)
        self.assertIsNone(record.holder_scaled_value)
        self.assertFalse(self.vacuum.is_valve_connected())

    def test_complete_holder_measurement_installs_magnetic_and_bridge_evidence(self):
        self.bind()
        with self.injected_squid(), self.session.claim(): outcome = self.backend.measure_bound_holder()
        self.assertTrue(outcome.installed)
        self.assertEqual(outcome.correction.holder_id, 'holder-046')
        self.assertEqual(outcome.correction.susceptibility_raw, 12.5)
        self.assertTrue((self.path.parent / 'holder_susceptibility' /
                         (outcome.correction.susceptibility_evidence_id + '.json')).exists())
        self.assertEqual(self.lift_serial().position, -50000)
        with self.session.claim():
            self.blank.return_to_clearance(self.session, field_outputs_off_verified=True)
            self.backend.clear_holder_geometry()
        self.assertEqual(self.lift_serial().position, 0)
        self.assertEqual(self.backend._sample_name, 'S1')
        self.assertIsNone(self.backend._queue_specimen_geometry)
        self.assertTrue(self.vacuum.is_pump_on())

    def test_binding_without_pose_fails_before_io(self):
        count = self.store.pending()['stage_count']
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.backend.bind_holder_geometry(self.session, self.vacuum)
        self.assertEqual(count, self.store.pending()['stage_count'])

    def test_pose_at_specimen_slot_is_rejected(self):
        with self.session.claim():
            self.table.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=True)
        with self.assertRaisesRegex(HardwareSafetyError, 'empty hole'): self.prepare()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_active_specimen_cannot_be_reinterpreted_as_blank(self):
        with self.session.claim():
            self.table.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=True)
        self.load()
        before = self.store.pending()['stage_count']
        with self.assertRaisesRegex(HardwareSafetyError, 'Return the original specimen'): self.prepare()
        self.assertEqual(before, self.store.pending()['stage_count'])
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_unknown_valve_cannot_authorize_blank_pose(self):
        self.vacuum.output_state_known = False
        with self.assertRaisesRegex(HardwareSafetyError, 'acknowledged'): self.prepare()

    def test_failed_stop_ack_does_not_verify_blank_pose(self):
        stop = self.motor.stop
        def failed(axis):
            stop(axis)
            if axis.name == 'changer_x': raise RuntimeError('bad ACK')
        with patch.object(self.motor, 'stop', side_effect=failed):
            with self.assertRaisesRegex(HardwareSafetyError, 'bad ACK'): self.prepare()
        self.assertEqual(self.store.latest_holder_context(self.session.token)['phase'], 'unverified')

    def test_pose_publication_failure_keeps_pending(self):
        with patch.object(self.store, '_publish_event', side_effect=OSError('disk')):
            with self.assertRaises(OSError): self.prepare()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.backend.bind_holder_geometry(self.session, self.vacuum)

    def test_holder_context_cannot_accept_positive_height_or_nonempty_slot(self):
        value = QueueHolderState(self.session.token, 'holder-046', 46).to_dict()
        for changed in (dict(value, sample_height=3000), dict(value, hole=1), dict(value, sample_height=False)):
            with self.assertRaises(HardwareSafetyError): QueueHolderState.read(changed, self.session, self.geometry)

    def test_zero_height_bridge_config_is_explicit_and_holder_only(self):
        with self.assertRaises(ValueError): SusceptibilityAcquisitionConfig(-22700, 0, 1).validate()
        with self.assertRaises(ValueError): SusceptibilityAcquisitionConfig(-22700, 3000, 1, blank_holder=True).validate()
        config = SusceptibilityAcquisitionConfig(-22700, 0, 1, blank_holder=True)
        config.validate()
        bridge, motion = bridge_fixture._Bridge(), bridge_fixture._Motion()
        service = SusceptibilityAcquisitionService(bridge, motion, config=config)
        with self.assertRaises(SusceptibilityAcquisitionError): service.acquire(sample_id='S1', is_holder=False)
        self.assertEqual(bridge.calls, [])
        self.assertEqual(motion.calls, [])

    def test_holder_binding_cannot_authorize_specimen_treatment(self):
        self.bind()
        from rapid_main.hardware_contracts import HardwareError
        from rapid_main.af_treatment import plan_af_treatment
        plan = plan_af_treatment('AF50', self.backend._config.af_demag, 2800)
        with self.session.claim(), self.assertRaisesRegex(HardwareError, 'specimen field treatment'):
            with self.backend._safety_operation('af', plan): self.fail('Started specimen treatment on blank geometry.')

    def test_holder_binding_cannot_clear_before_verified_rod_return(self):
        self.bind()
        from rapid_main.hardware_contracts import HardwareError
        with self.session.claim(), self.assertRaises(HardwareError): self.backend.clear_holder_geometry()

    def test_missing_live_reference_or_field_off_proof_prevents_pose_stage(self):
        count = self.store.pending()['stage_count']
        with self.session.claim(), self.assertRaises(HardwareSafetyError):
            self.blank.prepare(self.session, 46, field_outputs_off_verified=True)
        with self.session.claim(), self.assertRaises(HardwareSafetyError):
            self.blank.prepare(self.session, 46, reference_verified=True)
        self.assertEqual(count, self.store.pending()['stage_count'])

    def test_changed_xy_pose_prevents_instrument_read_and_lift_lowering(self):
        self.bind()
        self.motor._connections['COM3']._serial.position = 9590
        self.motor._connections['COM4']._serial.position = -11916
        before = len(self.lift_serial().writes)
        with self.injected_squid(), self.session.claim(), self.assertRaisesRegex(HardwareSafetyError, 'empty hole'):
            self.backend.read_squid()
        new_commands = self.lift_serial().writes[before:]
        self.assertFalse(any(b'134 ' in command for command in new_commands))
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertEqual(len(self.store.pending()['stage']['record']['initial_observations']), 4)

    def test_lost_valve_evidence_invalidates_holder_binding_before_stage(self):
        self.bind()
        self.vacuum.output_state_known = False
        count = self.store.pending()['stage_count']
        with self.session.claim(), self.assertRaisesRegex(HardwareSafetyError, 'acknowledged'):
            self.backend.read_susceptibility()
        self.assertEqual(count, self.store.pending()['stage_count'])
        self.assertEqual(self.backend._susceptibility.calls, [])

    def test_changed_empty_hole_during_acquisition_prevents_handoff(self):
        self.bind()
        # Alter coordinates before cleanup readbacks for X and Y.
        def during_measure():
            self.motor._connections['COM3']._serial.position = 9590
            self.motor._connections['COM4']._serial.position = -11916
            return 12.5
        with patch.object(self.backend._susceptibility, 'measure', side_effect=during_measure), self.session.claim():
            with self.assertRaises(HardwareSafetyError): self.backend.read_susceptibility()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_lost_original_motor_connection_cannot_reconnect_inside_stage(self):
        self.bind()
        self.backend._connected = False
        with self.session.claim(), self.assertRaisesRegex(HardwareSafetyError, 'reconnecting'):
            self.backend.read_susceptibility()
        self.assertEqual(self.backend._susceptibility.calls, [])
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_unowned_connection_guard_still_blocks_pending_queue(self):
        from rapid_main.hardware_contracts import HardwareError
        self.bind()
        with self.session.claim() as child:
            child.begin('acquisition', {}, self.backend._acquisition_safety_profile(), sample_id='holder-046')
            with self.assertRaises(HardwareError): self.backend._ensure_connected()

    def test_blank_binding_cannot_be_overwritten_by_specimen_binding(self):
        from rapid_main.hardware_contracts import HardwareError
        self.bind()
        original = self.backend._queue_specimen_geometry
        with self.session.claim(), self.assertRaisesRegex(HardwareError, 'clear the original blank-holder'):
            self.backend.bind_specimen_geometry(self.session)
        self.assertIs(original, self.backend._queue_specimen_geometry)
