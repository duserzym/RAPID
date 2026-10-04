"""Real queue journal and routed stop telemetry with injected instruments."""
from dataclasses import replace
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from rapid_main.acquisition import BlockContext
from rapid_main.squid_transport import BracketedSquidBackend
from rapid_main.susceptibility_acquisition import SusceptibilityAcquisitionRecord, SusceptibilityPhase
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests import test_queue_specimen_geometry as geometry_fixture
from tests.test_acquisition import _build_service


class QueueAcquisitionTests(QueueLiftFixture, unittest.TestCase):
    def additional_stage_profiles(self):
        return geometry_fixture.QueueSpecimenGeometryTests.additional_stage_profiles(self)

    def setUp(self):
        super().setUp()
        self.load()
        self.backend._safety_store = self.session.child_store
        with self.session.claim(): self.backend.bind_specimen_geometry(self.session)
        self.backend._holder_store = SimpleNamespace(require_valid=lambda **kw: SimpleNamespace(
            require_susceptibility=lambda: 0, susceptibility_evidence_id='holder-1'))
        self.backend._direction_up = True
        self.backend._susceptibility_records = []
        service, transport, vertical, turning, clock = _build_service()
        service._config = replace(service.config, zero_position=-48600, measurement_position=-96100)
        self.squid = BracketedSquidBackend(service, transport_retries=0,
            context_provider=lambda: BlockContext(sample_name='S1', run_id='run-1'))
        self.transport = transport
        self.patch_bracket = patch.object(self.backend, '_ensure_bracketed', self.ensure_bracket)
        self.patch_bracket.start()
        self.addCleanup(self.patch_bracket.stop)
        self.susc_record = SusceptibilityAcquisitionRecord('susc-1', 'S1', False,
            '2026-10-03T00:00:00+00:00', '2026-10-03T00:00:01+00:00', 'completed',
            -22700, 2800, -21300, 0, -21300, 0, 0,
            bridge_zero_reply='OK', bridge_scaled_value=1, holder_scaled_value=0,
            holder_evidence_id='holder-1', susceptibility=1, safe_state_confirmed=True,
            phases=(SusceptibilityPhase('measure', '', '', True),))

    def ensure_bracket(self):
        stage = self.store.pending()['stage']
        self.assertEqual((stage['family'], stage['status']), ('acquisition', 'pending'))
        self.assertEqual(self.backend._sample_height(), 2800)
        self.backend._bracketed = self.squid

    def read(self):
        with self.session.claim(): return self.backend.read_squid()

    def susceptibility(self, record=None):
        record = record or self.susc_record
        def acquire(**kwargs):
            self.assertEqual(self.store.pending()['stage']['family'], 'acquisition')
            self.assertEqual(self.backend._sample_height(), 2800)
            self.backend._susceptibility_records.append(record)
            return record
        with patch.object(self.backend, '_susceptibility_blockers', return_value=[]), \
             patch.object(self.backend, '_ensure_connected'), \
             patch.object(self.backend, '_acquire_susceptibility', side_effect=acquire), self.session.claim():
            return self.backend.read_susceptibility()

    def test_squid_publishes_six_observations_and_all_axis_stop_checks(self):
        result = self.read()
        stage = self.store.pending()['stage']
        self.assertEqual(stage['status'], 'verified')
        self.assertEqual(len(stage['record']['evidence']['block']['observations']), 6)
        self.assertEqual(len(stage['record']['cleanup_observations']), 4)
        self.assertEqual(stage['plan']['specimen_geometry']['sample_height'], 2800)
        self.assertIs(result, self.squid.last_acquisition.block)
        self.assertIsNone(self.backend._geometry_stage_token)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'lifted')

    def test_susceptibility_settles_original_geometry_and_keeps_grip(self):
        self.assertEqual(self.susceptibility(), 1)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_instrument_failure_keeps_pending_and_stops_all_axes(self):
        with patch.object(self.squid, 'read_squid', side_effect=RuntimeError('read failed')):
            with self.assertRaisesRegex(HardwareSafetyError, 'read failed'): self.read()
        record = self.store.pending()['stage']['record']
        self.assertEqual(len(record['cleanup_observations']), 4)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertTrue(self.vacuum.is_valve_connected())
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.backend._sample_height()

    def test_failed_stop_ack_stays_pending_even_with_zero_velocity(self):
        stop = self.motor.stop
        def failed(axis):
            stop(axis)
            if axis.name == 'changer_x': raise RuntimeError('missing ACK')
        with patch.object(self.motor, 'stop', side_effect=failed):
            with self.assertRaisesRegex(HardwareSafetyError, 'missing ACK'): self.read()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertEqual(len(self.store.pending()['stage']['record']['cleanup_observations']), 4)

    def test_publication_failure_cannot_handoff_a_completed_block(self):
        with patch.object(self.store, '_publish_event', side_effect=OSError('disk failure')):
            with self.assertRaisesRegex(OSError, 'disk failure'): self.read()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertIsNone(self.backend._geometry_stage_token)

    def test_late_cancel_preserves_pending_and_grip(self):
        original = self.squid.read_squid
        def read():
            result = original()
            self.backend._halt_check = lambda: True
            return result
        with patch.object(self.squid, 'read_squid', side_effect=read):
            with self.assertRaisesRegex(HardwareSafetyError, 'cancelled'): self.read()
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_pre_cancel_does_not_begin_stage_or_read(self):
        old = self.store.pending()['stage_count']
        self.backend._halt_check = lambda: True
        with self.assertRaises(InterruptedError): self.read()
        self.assertEqual(self.store.pending()['stage_count'], old)
        self.assertIsNone(self.squid.last_acquisition)

    def test_simulated_or_partial_squid_evidence_cannot_settle(self):
        self.squid._simulated = True
        with self.assertRaisesRegex(HardwareSafetyError, 'unverified'): self.read()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_unverified_susceptibility_safe_return_cannot_settle(self):
        with self.assertRaisesRegex(HardwareSafetyError, 'unverified'):
            self.susceptibility(replace(self.susc_record, safe_state_confirmed=False))
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_foreign_store_is_rejected_before_stage_or_read(self):
        self.backend._safety_store = self.store
        with self.assertRaisesRegex(HardwareSafetyError, 'original claimed'): self.read()
        self.assertIsNone(self.squid.last_acquisition)

    def test_flux_recovery_is_a_new_owned_stage_before_motion(self):
        self.read()
        count = self.store.pending()['stage_count']
        recover = self.squid.recover_flux_count_discontinuity
        def verified(validation):
            self.assertEqual(self.store.pending()['stage']['status'], 'pending')
            self.assertEqual(self.store.pending()['stage']['plan']['action'], 'flux_recovery')
            self.assertEqual(self.backend._sample_height(), 2800)
            return recover(validation)
        with patch.object(self.squid, 'recover_flux_count_discontinuity', side_effect=verified), self.session.claim():
            self.backend.recover_flux_count_discontinuity()
        self.assertEqual(self.store.pending()['stage_count'], count + 1)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')

    def test_failed_acquisition_cannot_replay_flux_recovery(self):
        with patch.object(self.squid, 'read_squid', side_effect=RuntimeError('failed')):
            with self.assertRaises(HardwareSafetyError): self.read()
        before = len(self.transport.commands)
        with self.session.claim(), self.assertRaises(HardwareSafetyError):
            self.backend.recover_flux_count_discontinuity()
        self.assertEqual(len(self.transport.commands), before)

    def test_failed_susceptibility_retains_new_failed_bridge_record(self):
        def failed(**kwargs):
            self.backend._susceptibility_records.append(replace(self.susc_record,
                outcome='failed', error='bridge timeout', susceptibility=None))
            raise RuntimeError('bridge timeout')
        with patch.object(self.backend, '_susceptibility_blockers', return_value=[]), \
             patch.object(self.backend, '_ensure_connected'), \
             patch.object(self.backend, '_acquire_susceptibility', side_effect=failed), self.session.claim():
            with self.assertRaisesRegex(HardwareSafetyError, 'bridge timeout'):
                self.backend.read_susceptibility()
        stage = self.store.pending()['stage']
        self.assertEqual(stage['status'], 'pending')
        self.assertEqual(stage['record']['evidence']['error'], 'bridge timeout')
        self.assertEqual(len(stage['record']['cleanup_observations']), 4)

    def test_scientific_setting_change_during_cleanup_prevents_settlement(self):
        stop = self.motor.stop
        def changed(axis):
            stop(axis)
            if axis.name == 'turning': self.backend._config.motion.zero_pos -= 1
        with patch.object(self.motor, 'stop', side_effect=changed):
            with self.assertRaisesRegex(HardwareSafetyError, 'original live queue'): self.read()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_partial_observation_payload_is_rejected_before_settlement(self):
        read = self.squid.read_squid
        def partial():
            read()
            acquisition = self.squid.last_acquisition
            block = replace(acquisition.block, observations=acquisition.block.observations[:-1])
            self.squid._last_acquisition = replace(acquisition, block=block)
            return block
        with patch.object(self.squid, 'read_squid', side_effect=partial):
            with self.assertRaisesRegex(HardwareSafetyError, '6 observations'): self.read()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
