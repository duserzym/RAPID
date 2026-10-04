from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from rapid_main.config import VacuumConfig
from rapid_main.diagnostic_services import VacuumBackendAdapter
from rapid_main.queue_lift_transfer import QueueLiftTransfer
from rapid_main.queue_station import QueueStationGeometry
from rapid_main.queue_table_motion import QueueXYTableMotion
from rapidpy_common.hardware import MotorAxisConfig, MotorControllerConfig
from rapidpy_common.hardware_safety import HardwareSafetyError, HardwareSafetyStore
from rapidpy_common.motor_routing import RoutedMotorSerialClient
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
from updown_control.app import VacuumController
from tests.test_queue_table_motion import PositionSerial
from tests.test_queue_vacuum import SerialInterface


class LiftSerial(PositionSerial):
    home_offset = 200
    def write(self, payload):
        super().write(payload)
        if self.port != 'COM5':
            return
        if self.command.startswith('134 '):
            target = int(self.command.split()[1])
            if target == -5000:
                self.position, self.status = -2000, 0
            elif target == 195000:
                self.position, self.status = self.home_offset, 16
            else:
                self.status = 16 if target == 0 else 0
        elif self.command == '145':
            self.position = 0
        elif self.command.startswith('11 10 '):
            self.relabel_target = -int(self.command.split()[-1])
        elif self.command == '165 1802':
            self.position = self.relabel_target


class Clock:
    def __init__(self): self.time = 0
    def monotonic(self): return self.time
    def sleep(self, seconds): self.time += seconds


class QueueLiftFixture:
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.path = Path(temporary.name) / 'safety.json'
        self.store = QueueSafetyStore(self.path)
        self.vacuum_serial = SerialInterface({'10V00': 'CLOSED', '10MFF': 'ON',
                                             '10VFF': 'OPEN', '10M00': 'OFF'})
        def serial_factory(**kwargs):
            return self.vacuum_serial if kwargs['port'] == 'COM11' else LiftSerial(**kwargs)
        for item in (patch('rapidpy_common.hardware.serial.Serial', side_effect=serial_factory),
                     patch('rapidpy_common.hardware.time.sleep', return_value=None)):
            item.start()
            self.addCleanup(item.stop)
        self.axes = {key: MotorAxisConfig(key, ident, 16, port) for key, ident, port in
                     (('changer_x', 1, 'COM3'), ('turning', 2, 'COM6'),
                      ('updown', 3, 'COM5'), ('changer_y', 4, 'COM4'))}
        config = MotorControllerConfig(slot_max=100, one_step=-1010.1010101,
                                       sample_bottom=-5000, sample_height=3000, meas_pos=-97500)
        self.motor = RoutedMotorSerialClient(config, list(self.axes.values()))
        self.motor.connect()
        self.addCleanup(self.motor.disconnect)
        self.geometry = QueueStationGeometry('station.ini', True, 1, 100, 46, config.one_step,
                                             ((1, 9590, -11916), (46, 0, 39)), (-3, -2))
        self.clock = Clock()
        self.table = QueueXYTableMotion(self.motor, self.axes, self.geometry, sleep=self.clock.sleep,
            transfer_settings={'grip_settle_s': .3, 'dropoff_delay_s': 1.2})
        self.vacuum = VacuumBackendAdapter(VacuumConfig(port='COM11', warn_threshold=0))
        self.vacuum._safety_store = HardwareSafetyStore(self.path)
        self.profile = dict(helper='rapid_main_queue', resources={'vacuum': self.vacuum._binding()},
            stage_profiles={'motion': self.table.profile, 'vacuum': self.vacuum.queue_station_binding(), 'acquisition': {'test': 1}})
        self.profile['stage_profiles'].update(self.additional_stage_profiles())
        self.session = QueueWorkflowSession.start(self.store, self.queue_plan(), self.profile, run_id='run-1')
        self.addCleanup(self.session.release)
        self.lift = QueueLiftTransfer(self.table, self.vacuum, monotonic=self.clock.monotonic)
        with self.session.claim():
            self.vacuum.queue_connect(self.session)
            self.table.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=True)
            self.outputs(False)
        self.addCleanup(self.vacuum._pump_only_controller.disconnect)

    def outputs(self, valve):
        return self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=valve,
            motors_stopped_verified=True, specimen_secured=True, specimen_at_pickup_verified=True,
            field_outputs_off_verified=True)

    def additional_stage_profiles(self):
        return {}

    def queue_plan(self):
        return {'sample': 'S1'}

    def call(self, method, *args, **kwargs):
        with self.session.claim():
            return getattr(self.lift, method)(self.session, *args, field_outputs_off_verified=True, **kwargs)

    def load(self):
        self.call('pickup', 1, 'S1', file_id='F1')
        with self.session.claim(): self.outputs(True)
        return self.call('home_loaded')

    def context(self):
        with self.session.claim(): return self.lift.context(self.session)

    def lift_serial(self): return self.motor._connections['COM5']._serial

class QueueLiftTransferTests(QueueLiftFixture, unittest.TestCase):
    def test_pickup_and_top_reference_persist_dynamic_height_and_original_identity(self):
        record = self.load()
        context = self.context()
        self.assertEqual((context.original_slot, context.sample_id, context.file_id), (1, 'S1', 'F1'))
        self.assertEqual(context.pickup_position_raw, -2000)
        self.assertEqual(context.sample_height, 2800)
        self.assertEqual(context.phase, 'lifted')
        self.assertEqual(self.motor.config.sample_height, 3000)
        self.assertTrue(record.safe_state_confirmed)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertTrue(self.store.pending())
        self.assertEqual(self.store.pending()['stage']['run_id'], 'run-1')

    def test_supported_return_uses_measured_height_and_release_delay_before_clear(self):
        self.load()
        self.call('lower_for_dropoff')
        self.assertEqual(self.lift_serial().position, -2480)
        self.assertEqual(self.context().phase, 'supported')
        before_release = self.clock.time
        with self.session.claim(): self.outputs(False)
        self.call('clear_after_release')
        self.assertGreaterEqual(self.clock.time - before_release, 1.2)
        self.assertEqual(self.context().phase, 'clear')
        self.assertEqual(self.lift_serial().position, 0)
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertIsNotNone(self.store.pending())

    def test_lift_home_needs_recorded_pickup_and_acknowledged_grip(self):
        with self.assertRaises(HardwareSafetyError): self.call('home_loaded')
        self.call('pickup', 1, 'S1')
        with self.assertRaisesRegex(HardwareSafetyError, 'grip'): self.call('home_loaded')
        self.assertEqual(self.context().phase, 'picked')

    def test_another_specimen_cannot_replace_an_active_transfer(self):
        self.call('pickup', 1, 'S1')
        with self.assertRaisesRegex(HardwareSafetyError, 'original specimen'):
            self.call('pickup', 1, 'S2')
        self.assertEqual(self.context().sample_id, 'S1')

    def test_wrong_xy_slot_cannot_authorize_supported_dropoff(self):
        self.load()
        with self.session.claim():
            self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=True)
        before = len(self.lift_serial().writes)
        with self.assertRaisesRegex(HardwareSafetyError, 'original registered'):
            self.call('lower_for_dropoff')
        self.assertFalse(any(b'134 ' in value for value in self.lift_serial().writes[before:]))
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'unverified')

    def test_raise_for_return_requires_empty_location_and_keeps_grip(self):
        self.load()
        with self.session.claim():
            self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=True)
        self.lift_serial().position, self.lift_serial().status = -96000, 0
        self.call('raise_loaded_to_clearance')
        self.assertEqual(self.lift_serial().position, 0)
        self.assertEqual(self.context().phase, 'lifted')
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_grip_must_be_released_before_lift_clearance(self):
        self.load()
        self.call('lower_for_dropoff')
        with self.assertRaisesRegex(HardwareSafetyError, 'valve release'):
            self.call('clear_after_release')
        self.assertEqual(self.context().phase, 'supported')

    def test_failed_home_preserves_grip_and_original_context_for_recovery(self):
        self.call('pickup', 1, 'S1')
        with self.session.claim(): self.outputs(True)
        self.lift_serial().fail_move = True
        with self.assertRaises(HardwareSafetyError): self.call('home_loaded')
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'unverified')
        self.assertEqual(self.context().original_slot, 1)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_context_survives_intervening_acquisition_evidence(self):
        self.load()
        from types import SimpleNamespace
        record = SimpleNamespace(safe_state_confirmed=True, simulated=False,
            to_dict=lambda: {'safe_state_confirmed': True, 'simulated': False})
        with self.session.claim() as child:
            token = child.begin('acquisition', {}, {'test': 1})
            child.finish(token, {'test': 1}, record)
        self.assertEqual(self.context().phase, 'lifted')
        self.assertEqual(self.context().sample_height, 2800)

    def test_failed_publication_never_claims_a_completed_pickup(self):
        with patch.object(self.session.child_store, 'finish', side_effect=OSError('disk full')):
            with self.assertRaises(OSError): self.call('pickup', 1, 'S1')
        self.assertEqual(self.context().phase, 'unverified')
        self.assertIsNone(self.context().pickup_position_raw)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_restart_reads_original_identity_but_never_replays_pickup(self):
        self.load()
        self.session.release()
        self.session = QueueWorkflowSession.recover(self.store, self.profile)
        self.addCleanup(self.session.release)
        self.assertEqual(self.context().sample_id, 'S1')
        with self.assertRaisesRegex(HardwareSafetyError, 'never replay'):
            self.call('pickup', 1, 'S2')

    def test_historical_transfer_evidence_corruption_fails_closed(self):
        self.load()
        head = self.store.pending()['history_head']
        self.store._event_path(self.session.token, head['id']).write_bytes(b'corrupt')
        with self.assertRaises(HardwareSafetyError): self.context()

    def test_cancellation_during_grip_settling_keeps_grip_and_never_homes(self):
        self.call('pickup', 1, 'S1')
        with self.session.claim(): self.outputs(True)
        before = len(self.lift_serial().writes)
        deadline = self.clock.time + .35
        self.table.should_cancel = lambda: self.clock.time >= deadline
        with self.assertRaises(HardwareSafetyError): self.call('home_loaded')
        self.assertFalse(any(b'134 ' in value for value in self.lift_serial().writes[before:]))
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'unverified')

    def test_failed_clearance_after_valve_release_keeps_pump_ownership(self):
        self.load()
        self.call('lower_for_dropoff')
        with self.session.claim(): self.outputs(False)
        self.lift_serial().fail_move = True
        with self.assertRaises(HardwareSafetyError): self.call('clear_after_release')
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertTrue(self.vacuum.outputs_held)
        self.assertEqual(self.context().phase, 'unverified')
        self.assertEqual(self.context().original_slot, 1)

    def test_bad_top_reference_cannot_publish_a_negative_specimen_height(self):
        self.call('pickup', 1, 'S1')
        with self.session.claim(): self.outputs(True)
        self.lift_serial().home_offset = 5001
        with self.assertRaises(HardwareSafetyError): self.call('home_loaded')
        self.assertIsNone(self.context().sample_height)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_typed_specimen_identity_rejects_wrong_owner_and_fractional_counts(self):
        from rapid_main.queue_lift_transfer import QueueSpecimenState
        self.load()
        valid = self.context().to_dict()
        for change in ({'queue_token': 'f' * 32}, {'original_slot': 46}, {'phase': 'ready'},
                       {'pickup_position_raw': -2000.0}, {'sample_height': True}, {'sample_id': ''}):
            with self.subTest(change=change), self.assertRaises(HardwareSafetyError):
                QueueSpecimenState.read(dict(valid, **change), self.session, self.geometry)

    def test_complete_empty_hole_measurement_return_preserves_original_specimen(self):
        self.load()
        with self.session.claim():
            self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=True)
        # Inject the physical acquisition pose; the acquisition service is tested
        # separately and still needs the measured-height integration.
        self.lift_serial().position, self.lift_serial().status = -96000, 0
        self.call('raise_loaded_to_clearance')
        with self.session.claim():
            self.table.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=True)
        self.call('lower_for_dropoff')
        with self.session.claim(): self.outputs(False)
        self.call('clear_after_release')
        context = self.context()
        self.assertEqual((context.phase, context.original_slot, context.sample_id), ('clear', 1, 'S1'))
        self.assertEqual(self.geometry.slot_from_xy_counts(
            self.motor._connections['COM3']._serial.position, self.motor._connections['COM4']._serial.position), 1)
        self.assertEqual(context.sample_height, 2800)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertGreater(self.store.verify_history(self.store.pending()), 8)

    def test_pending_acquisition_marks_transfer_uncertain_without_losing_identity(self):
        self.load()
        with self.session.claim() as child:
            child.begin('acquisition', {}, {'test': 1})
        self.assertEqual(self.context().phase, 'unverified')
        self.assertEqual(self.context().sample_id, 'S1')
        self.assertEqual(self.context().sample_height, 2800)

    def test_original_slot_is_persisted_before_every_native_pickup_command(self):
        observed = []
        def observe(command):
            stage = self.store.pending()['stage']
            self.assertEqual(stage['status'], 'pending')
            self.assertEqual(stage['plan']['transfer_context']['original_slot'], 1)
            self.assertEqual(stage['sample_id'], 'S1')
            observed.append(command)
        self.lift_serial().on_write = observe
        self.call('pickup', 1, 'S1')
        self.assertTrue(any(command.startswith('149 ') for command in observed))
        self.assertTrue(any(command.startswith('134 ') for command in observed))

    def test_changed_or_invalid_dropoff_delay_cannot_start_transfer(self):
        self.lift.dropoff_delay_s = 0
        before = len(self.lift_serial().writes)
        with self.assertRaises(HardwareSafetyError): self.call('pickup', 1, 'S1')
        self.assertEqual(len(self.lift_serial().writes), before)
        for delay in (None, True, -1, float('nan'), float('inf')):
            self.table.transfer_settings['dropoff_delay_s'] = delay
            with self.subTest(delay=delay), self.assertRaises(HardwareSafetyError):
                QueueLiftTransfer(self.table, self.vacuum)

    def test_settle_deadline_never_issues_a_negative_sleep(self):
        readings = iter((0, .99, 1.01))
        sleeps = []
        self.lift.monotonic = lambda: next(readings)
        self.table.sleep = sleeps.append
        self.lift._delay(1)
        self.assertEqual(len(sleeps), 1)
        self.assertGreater(sleeps[0], 0)
