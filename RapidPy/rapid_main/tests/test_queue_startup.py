"""Original native startup with injected DLLs and serial transports only."""
import copy
import unittest
from unittest.mock import patch

from rapid_main.pulse_circuit import PulseCircuit
from rapid_main.queue_command_worker import QueueCommandWorker
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests import test_queue_backend_coordinator as backend_fixture
from tests.test_queue_vacuum import SerialInterface


class QueueStartupFixture(backend_fixture.QueueBackendCoordinatorFixture):
    def setUp(self):
        super().setUp()
        worker = self.command(self.backend.finish_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        self.session.release()
        self.vacuum_serial = SerialInterface({'10V00': 'CLOSED', '10MFF': 'ON',
                                             '10VFF': 'OPEN', '10M00': 'OFF'})
        self.addCleanup(lambda: self.vacuum._pump_only_controller.disconnect()
            if self.vacuum._pump_only_controller is not None else None)
        self.backend._config.vacuum = copy.deepcopy(self.vacuum._cfg)
        self.backend._config.irm_arm = copy.deepcopy(self.arm.cfg)
        self.backend._pulse_irm.make_circuit = lambda: PulseCircuit(self.backend._config.pulse_irm,
            self.daq, self.af, sleep=self.clock.sleep, monotonic=self.clock.monotonic)
        self.backend._queue_startup = None

    def prepare(self, **kwargs):
        arguments = dict(run_id='new-run', operator='Operator', empty_rod_confirmed=True)
        arguments.update(kwargs)
        startup = self.backend.prepare_queue_lifetime(self.vacuum, {'samples': ['S1']}, **arguments)
        self.addCleanup(startup.session.release)
        self.session = startup.session
        # Use deterministic clocks for native verification; commands remain real
        # adapter calls against the injected serial/DLL interfaces.
        startup.table.sleep = self.clock.sleep
        return startup

    def arm_xy_edges(self):
        native = self.motor.connect
        def connect(*args, **kwargs):
            native(*args, **kwargs)
            for key, negative, positive in (('changer_x', 4, 5), ('changer_y', 5, 6)):
                serial = self.motor._connections[self.axes[key].port]._serial
                serial.status = (1 << negative) | (1 << positive)
                def observe(command, serial=serial, neg=negative, pos=positive):
                    if command.startswith('135 '):
                        distance = int(command.split()[1])
                        serial.position = -1000 if distance < 0 else 1000
                        serial.status = (1 << pos) if distance < 0 else (1 << neg)
                serial.on_write = observe
        return patch.object(self.motor, 'connect', side_effect=connect)

class QueueStartupTests(QueueStartupFixture, unittest.TestCase):
    def test_retained_squid_handle_blocks_preparation_before_new_root_or_io(self):
        raw = self.backend._measurement.prepare_raw_client()
        raw._serial = object()
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        try:
            with self.assertRaisesRegex(HardwareSafetyError, 'retained SQUID'):
                self.prepare()
            self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)
        finally:
            raw._serial = None

    def test_squid_motor_port_alias_blocks_before_new_root_or_io(self):
        self.backend._config.squid.port = chr(92) * 2 + '.' + chr(92) + 'com3'
        before = self.store.read(), len(self.commands)
        with self.assertRaisesRegex(HardwareSafetyError, 'distinct serial circuits'):
            self.prepare()
        self.assertEqual((self.store.read(), len(self.commands)), before)

    def test_vacuum_motor_port_collision_blocks_before_new_root_or_io(self):
        self.backend._config.vacuum.port = self.vacuum._cfg.port = 'COM3'
        before = self.store.read(), len(self.commands)
        with self.assertRaisesRegex(HardwareSafetyError, 'distinct serial circuits'):
            self.prepare()
        self.assertEqual((self.store.read(), len(self.commands)), before)

    def test_prepare_records_empty_rod_and_all_profiles_without_any_io(self):
        before = len(self.commands), len(self.vacuum_serial.writes)
        startup = self.prepare()
        state = startup.session.store.read()
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertEqual(state['stage_count'], 0)
        self.assertEqual(state['plan']['startup']['operator'], 'Operator')
        self.assertTrue(state['plan']['startup']['empty_rod_confirmed'])
        self.assertEqual(set(state['profile']['stage_profiles']),
            {'motion', 'vacuum', 'field_outputs', 'acquisition', 'af', 'arm', 'pulse', 'rrm'})
        self.assertIs(self.backend._safety_store, startup.session.child_store)

    def test_native_startup_then_terminal_settlement_uses_original_worker_claim(self):
        startup = self.prepare()
        with self.arm_xy_edges():
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        self.assertTrue(startup.completed)
        self.assertIs(self.backend._queue_coordinator.session, startup.session)
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertFalse(self.vacuum.is_valve_connected())
        with startup.session.claim():
            self.assertEqual(self.backend._queue_coordinator.reference.require(startup.session).phase, 'referenced')
        worker = self.command(self.backend.finish_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        self.assertIsNone(startup.session.store.pending())
        self.assertIsNone(self.backend._queue_startup)
        self.assertIsNotNone(startup.session._lease)

    def test_missing_confirmation_or_named_operator_never_begins_or_actuates(self):
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        for kwargs in ({'empty_rod_confirmed': False}, {'empty_rod_confirmed': 1}, {'operator': ''}):
            with self.assertRaises(HardwareSafetyError):
                self.prepare(**kwargs)
        self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_field_cutoff_failure_never_opens_motor_or_vacuum_ports(self):
        startup = self.prepare()
        self.cap_voltage = .5
        with patch.object(self.motor, 'connect') as motor_connect, patch.object(self.vacuum, 'queue_connect') as vacuum_connect:
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        motor_connect.assert_not_called()
        vacuum_connect.assert_not_called()
        self.assertEqual(startup.session.store.pending()['stage']['family'], 'field_outputs')
        retry = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(retry.ok)
        self.assertIn('cannot replay', retry.error)

    def test_motion_stage_precedes_motor_connection_and_broadcast(self):
        startup = self.prepare()
        native = self.motor.connect
        def observe(*args, **kwargs):
            state = startup.session.store.pending()
            self.assertEqual(state['stage']['status'], 'pending')
            self.assertEqual(state['stage']['plan']['action'], 'connect_empty_station_and_verify_stop')
            return native(*args, **kwargs)
        with patch.object(self.motor, 'connect', side_effect=observe), patch.object(self.vacuum, 'queue_connect', side_effect=RuntimeError('boundary reached')):
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('boundary reached', worker.error)
        self.assertEqual(startup.session.store.pending()['stage']['record']['schema'], 'rapidpy.queue_startup_stop.v1')

    def test_failed_stop_never_connects_vacuum_or_homes(self):
        startup = self.prepare()
        with patch.object(startup.table, '_stopped', return_value=({}, 'moving readback')), patch.object(self.vacuum, 'queue_connect') as connect:
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('moving readback', worker.error)
        connect.assert_not_called()
        self.assertEqual(startup.session.store.pending()['stage']['status'], 'pending')

    def test_changed_prepared_configuration_fails_before_field_or_connection_io(self):
        self.prepare()
        self.backend._config.squid.samples_per_pos += 1
        before = len(self.commands)
        worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)
        self.assertIn('bindings changed', worker.error)

    def test_unclaimed_startup_cannot_issue_commands(self):
        startup = self.prepare()
        before = len(self.commands)
        with self.assertRaises(HardwareSafetyError): startup.run()
        self.assertEqual(len(self.commands), before)

    def test_changed_original_motor_calibration_cannot_prepare_a_new_owner(self):
        self.backend._config.motor_station.controller['changer_speed'] += 1
        before = self.store.read(), len(self.commands)
        with self.assertRaisesRegex(HardwareSafetyError, 'accepted native'):
            self.prepare()
        self.assertEqual((self.store.read(), len(self.commands)), before)

    def test_retained_motor_handle_cannot_be_replaced_by_startup(self):
        self.motor.connect()
        handles = dict(self.motor._connections)
        with self.assertRaisesRegex(HardwareSafetyError, 'retained motor'):
            self.prepare()
        self.assertEqual(self.motor._connections, handles)

    def test_precancelled_startup_worker_leaves_root_pending_without_io(self):
        startup = self.prepare()
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = QueueCommandWorker(self.backend, self.backend.start_queue_lifetime, recover_on_error=False)
        worker.stop()
        worker.run()
        self.assertFalse(worker.ok)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertEqual(startup.session.store.pending()['stage_count'], 0)

    def test_failed_empty_lift_top_never_zeroes_or_references_xy(self):
        startup = self.prepare()
        native = self.motor.connect
        def connect(*args, **kwargs):
            native(*args, **kwargs)
            serial = self.motor._connections['COM5']._serial
            serial.status = 0
            serial.on_write = lambda command: setattr(serial, 'status', 0)
        with patch.object(self.motor, 'connect', side_effect=connect), patch.object(self.motor, 'zero_target_pos') as zero:
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('top switch', worker.error)
        zero.assert_not_called()
        self.assertEqual(startup.session.store.pending()['stage']['status'], 'pending')
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertFalse(self.vacuum.is_valve_connected())

    def test_empty_lift_from_below_top_uses_one_switch_guarded_sweep_before_zero(self):
        startup = self.prepare()
        native_connect, native_move = self.motor.connect, self.motor.move_motor
        moves = []
        def connect(*args, **kwargs):
            native_connect(*args, **kwargs)
            serial = self.motor._connections['COM5']._serial
            serial.status, serial.position = 0, -2000
        def move(axis, target, speed, **kwargs):
            moves.append((axis.name, target, kwargs))
            state = startup.session.store.pending()
            self.assertEqual(state['stage']['plan']['action'], 'reference_confirmed_empty_lift')
            self.assertEqual(state['stage']['status'], 'pending')
            return native_move(axis, target, speed, **kwargs)
        with patch.object(self.motor, 'connect', side_effect=connect), patch.object(self.motor, 'move_motor', side_effect=move), patch('rapid_main.queue_startup.QueueXYReference.home', side_effect=RuntimeError('XY boundary')):
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('XY boundary', worker.error)
        self.assertEqual(len(moves), 1)
        self.assertEqual(moves[0][1], 195000)
        self.assertEqual((moves[0][2]['stop_enable'], moves[0][2]['stop_condition']), (-1, 1))
        self.assertEqual(self.motor._connections['COM5']._serial.position, 0)
        self.assertEqual(startup.session.store.pending()['stage']['record']['schema'], 'rapidpy.queue_empty_lift_reference.v1')

    def test_cancel_after_empty_lift_sweep_stops_without_zero_or_xy_motion(self):
        startup = self.prepare()
        native_connect, native_move = self.motor.connect, self.motor.move_motor
        cancelled = False
        startup.table.should_cancel = lambda: cancelled
        def connect(*args, **kwargs):
            native_connect(*args, **kwargs)
            self.motor._connections['COM5']._serial.status = 0
        def move(*args, **kwargs):
            nonlocal cancelled
            result = native_move(*args, **kwargs)
            cancelled = True
            return result
        with patch.object(self.motor, 'connect', side_effect=connect), patch.object(self.motor, 'move_motor', side_effect=move), patch.object(self.motor, 'zero_target_pos') as zero:
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('cancelled', worker.error)
        zero.assert_not_called()
        self.assertEqual(startup.session.store.pending()['stage']['status'], 'pending')
        self.assertFalse(self.vacuum.is_valve_connected())
