"""Native MCC/ADwin calls with injected DLLs; no physical field or motion."""
import ctypes
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rapid_main.arm_bias import ArmBiasBackend
from rapid_main.pulse_circuit import PulseCircuit
from rapid_main.queue_field_outputs import QueueFieldOutputs
from rapidpy_common.adwin_af import AdwinAFController, AdwinBoardConfig, AdwinCoilLimits
from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.mcc_daq import MccDaq, MccError
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests.test_arm_bias import bias_config
from tests.test_pulse_circuit import circuit_config


class QueueFieldOutputsTests(QueueLiftFixture, unittest.TestCase):
    def additional_stage_profiles(self):
        self.commands, self.bits, self.cap_voltage = [], {}, 0.0
        self.processes = {number: 1 for number in range(1, 11)}
        self.relay_word = 7
        self.dll = Mock()
        self.dll.ADGetErrorCode.return_value = 0
        self.dll.Get_ADBPar.side_effect = lambda index, dev: self.processes[index + 100]
        def stop(number, dev):
            self.commands.append(('stop', number))
            self.processes[number] = 0
            return 0
        self.dll.ADB_Stop.side_effect = stop
        def dac(channel, count, dev):
            self.commands.append(('adwin_dac', channel, count))
            return 0
        self.dll.Set_DAC.side_effect = dac
        def relays(word, dev):
            self.commands.append(('relays', word))
            self.relay_word = word
            return 0
        self.dll.Set_Digout.side_effect = relays
        self.dll.Get_Digout.side_effect = lambda dev: self.relay_word
        self.af = object.__new__(AdwinAFController)
        self.af.board, self.af.limits, self.af._dll = AdwinBoardConfig(), AdwinCoilLimits(), self.dll
        self.af._last_digout_bit = -1
        mcc = Mock()
        mcc.cbGetBoardName.side_effect = lambda board, buffer: (setattr(buffer, 'value', b'PCI-DAS6030') or 0)
        def conversion(board, voltage_range, voltage, pointer):
            ctypes.cast(pointer, ctypes.POINTER(ctypes.c_ushort)).contents.value = round(voltage * 1000)
            return 0
        mcc.cbFromEngUnits.side_effect = conversion
        mcc.cbAOut.side_effect = lambda board, channel, vrange, raw: (self.commands.append(('mcc_dac', channel, raw)) or 0)
        mcc.cbDConfigBit.return_value = 0
        def bitout(board, port, bit, high):
            self.commands.append(('bit', bit, high))
            self.bits[bit] = high
            return 0
        def bitin(board, port, bit, pointer):
            ctypes.cast(pointer, ctypes.POINTER(ctypes.c_ushort)).contents.value = self.bits.get(bit, 0)
            self.commands.append(('bit_read', bit))
            return 0
        mcc.cbDBitOut.side_effect, mcc.cbDBitIn.side_effect = bitout, bitin
        mcc.cbAIn.side_effect = lambda *args: 0
        def tovoltage(board, vrange, raw, pointer):
            ctypes.cast(pointer, ctypes.POINTER(ctypes.c_float)).contents.value = self.cap_voltage
            self.commands.append(('capacitor_read', self.cap_voltage))
            return 0
        mcc.cbToEngUnits.side_effect = tovoltage
        self.mcc = mcc
        self.daq = MccDaq(0, dll=mcc)
        self.arm = ArmBiasBackend(bias_config(), controller=self.daq)
        self.arm.sleep = self.clock.sleep
        self.pulse = PulseCircuit(circuit_config(), self.daq, self.af,
                                 sleep=self.clock.sleep, monotonic=self.clock.monotonic)
        self.fields = QueueFieldOutputs(self.af, self.pulse, self.arm)
        return dict(field_outputs=self.fields.profile, af={'test': 1})

    def off(self, **kwargs):
        with self.session.claim(): return self.fields.verify_off(self.session, **kwargs)

    def test_journal_precedes_native_off_commands_and_all_evidence_is_linked(self):
        native = self.mcc.cbAOut.side_effect
        def observed(*args):
            self.assertEqual(self.store.pending()['stage']['family'], 'field_outputs')
            self.assertEqual(self.store.pending()['stage']['status'], 'pending')
            return native(*args)
        self.mcc.cbAOut.side_effect = observed
        context = self.off()
        self.assertEqual(context['phase'], 'off')
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        record = self.store.pending()['stage']['record']
        self.assertTrue(record['arm']['safe_state_confirmed'])
        self.assertEqual(record['arm']['gate_readback'], 1)
        self.assertTrue(record['pulse']['discharged'])
        self.assertTrue(record['af']['safe_state_confirmed'])
        self.assertEqual(len([item for item in record['af']['observations'] if item['action'] == 'process_status_after']), 10)
        first_relay = next(index for index, command in enumerate(self.commands) if command[0] == 'relays')
        before_relay = self.commands[:first_relay]
        self.assertEqual(len([command for command in before_relay if command[0] == 'stop']), 10)
        self.assertEqual(len([command for command in before_relay if command[0] == 'adwin_dac']), 2)
        self.assertTrue(record['af']['initial_output_cutoff']['outputs_zero_confirmed'])
        self.assertFalse(any(command[0] == 'bit' and command[1] == 1 and command[2] == 0 for command in self.commands))
        self.assertFalse(any(command[0] == 'mcc_dac' and command[2] != 0 for command in self.commands))
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertEqual(self.lift_serial().position, 0)

    def test_capacitor_unknown_withholds_all_relays_but_attempts_arm_and_af_zero(self):
        self.cap_voltage = .5
        with self.assertRaisesRegex(HardwareSafetyError, 'discharge'): self.off()
        self.assertFalse(any(command[0] == 'relays' for command in self.commands))
        self.assertIn(('bit', 0, 1), self.commands)
        self.assertEqual(len([command for command in self.commands if command[0] == 'adwin_dac']), 2)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_failed_arm_gate_readback_keeps_stage_pending_but_cleans_other_circuits(self):
        self.mcc.cbDBitIn.side_effect = None
        self.mcc.cbDBitIn.return_value = 7
        with self.assertRaisesRegex(HardwareSafetyError, 'cbDBitIn'): self.off()
        self.assertTrue(self.fields.last_record.pulse['safe_state_confirmed'])
        self.assertTrue(self.fields.last_record.af['safe_state_confirmed'])
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_failed_af_process_stop_attempts_both_dacs_and_preserves_unsafe_evidence(self):
        self.dll.ADB_Stop.return_value = 8
        self.dll.ADB_Stop.side_effect = None
        with self.assertRaisesRegex(HardwareSafetyError, 'Stop_Process'): self.off()
        self.assertEqual(self.dll.ADB_Stop.call_count, 10)
        self.assertEqual(self.dll.Set_DAC.call_count, 2)
        self.assertFalse(self.fields.last_record.af['safe_state_confirmed'])
        self.assertTrue(self.fields.last_record.pulse['discharged'])
        self.assertFalse(any(command[0] == 'relays' for command in self.commands))

    def test_relay_high_after_clear_does_not_publish_off_proof(self):
        self.dll.Get_Digout.side_effect = lambda dev: 7
        with self.assertRaises(HardwareSafetyError): self.off()
        self.assertEqual(self.store.latest_field_context(self.session.token)['phase'], 'unverified')

    def test_failed_publication_keeps_pending_and_cannot_return_proof(self):
        with patch.object(self.store, '_publish_event', side_effect=OSError('disk')):
            with self.assertRaises(OSError): self.off()
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.fields.require(self.session)

    def test_changed_station_binding_is_rejected_before_any_output(self):
        self.af.board.board_num = 2
        with self.assertRaisesRegex(HardwareSafetyError, 'binding'): self.off()
        self.assertEqual(self.commands, [])

    def test_original_recovery_reuses_failed_stage_and_never_fires(self):
        self.cap_voltage = .5
        with self.assertRaises(HardwareSafetyError): self.off()
        token, count = self.store.pending()['stage']['token'], self.store.pending()['stage_count']
        self.cap_voltage = 0
        self.session.is_recovery = True
        proof = self.off(recovery=True)
        self.assertEqual(proof['phase'], 'off')
        self.assertEqual(self.store.pending()['stage']['token'], token)
        self.assertEqual(self.store.pending()['stage_count'], count)
        self.assertFalse(any(command[0] == 'bit' and command[1] == 1 and command[2] == 0 for command in self.commands))

    def test_live_retry_cannot_replay_a_pending_cutoff(self):
        self.cap_voltage = .5
        with self.assertRaises(HardwareSafetyError): self.off()
        count = len(self.commands)
        with self.assertRaisesRegex(HardwareSafetyError, 'original unfinished'): self.off(recovery=True)
        self.assertEqual(count, len(self.commands))

    def test_another_service_instance_cannot_reuse_original_off_proof(self):
        self.off()
        other = QueueFieldOutputs(self.af, self.pulse, self.arm)
        with self.session.claim(), self.assertRaises(HardwareSafetyError): other.require(self.session)

    def test_digital_input_checks_binary_and_status_without_configuring_output(self):
        before = self.mcc.cbDConfigBit.call_count
        self.bits[0] = 1
        self.assertEqual(self.daq.digital_input(1, 0), 1)
        self.assertEqual(before, self.mcc.cbDConfigBit.call_count)
        self.bits[0] = 2
        with self.assertRaises(MccError): self.daq.digital_input(1, 0)

    def test_aliased_arm_and_pulse_outputs_are_rejected_before_io(self):
        self.arm.cfg.arm_gate_bit = self.pulse.cfg.fire_bit
        with self.assertRaisesRegex(HardwareSafetyError, 'alias'):
            QueueFieldOutputs(self.af, self.pulse, self.arm)
        self.assertEqual(self.commands, [])

    def test_native_transfer_and_grip_phases_persist_original_field_proof(self):
        proof = self.off()
        with self.session.claim():
            pickup = self.lift.pickup(self.session, 1, 'S1', file_id='F1', field_outputs_off_verified=proof)
            self.assertEqual(pickup.operation['field_outputs_proof'], dict(proof))
            self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=True,
                motors_stopped_verified=True, specimen_at_pickup_verified=True, field_outputs_off_verified=proof)
            self.lift.home_loaded(self.session, field_outputs_off_verified=proof)
            self.lift.lower_for_dropoff(self.session, field_outputs_off_verified=proof)
            self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=False,
                motors_stopped_verified=True, specimen_secured=True, field_outputs_off_verified=proof)
            returned = self.lift.clear_after_release(self.session, field_outputs_off_verified=proof)
        self.assertEqual(returned.transfer_context['phase'], 'clear')
        self.assertEqual(returned.operation['field_outputs_proof'], dict(proof))
        self.assertTrue(self.vacuum.is_pump_on())

    def test_table_transfer_records_live_field_proof_and_rejects_copied_dict(self):
        proof = self.off()
        with self.session.claim():
            moved = self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=proof)
            self.assertEqual(moved.operation['field_outputs_proof'], dict(proof))
            with self.assertRaises(HardwareSafetyError):
                self.table.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=dict(proof))

    def test_a_later_verified_field_treatment_consumes_off_proof(self):
        proof = self.off()
        with self.session.claim() as child:
            token = child.begin('af', {}, {'test': 1})
            record = SimpleNamespace(safe_state_confirmed=True, simulated=False,
                to_dict=lambda: dict(safe_state_confirmed=True, simulated=False))
            child.finish(token, {'test': 1}, record)
            with self.assertRaisesRegex(HardwareSafetyError, 'again after'): self.fields.require(self.session)
            with self.assertRaises(HardwareSafetyError):
                self.table.move_to_slot(self.session, 46, reference_verified=True, field_outputs_off_verified=proof)

    def test_arm_enabled_gate_readback_never_authorizes_off(self):
        self.mcc.cbDBitIn.side_effect = None
        self.mcc.cbDBitIn.return_value = 0
        with self.assertRaisesRegex(HardwareSafetyError, 'remains enabled'): self.off()
        self.assertEqual(self.fields.last_record.arm['gate_readback'], 0)

    def test_failed_arm_dac_still_disconnects_gate_and_checks_other_circuits(self):
        self.mcc.cbAOut.side_effect = lambda board, channel, vrange, raw: 7 if channel == 0 else 0
        with self.assertRaisesRegex(HardwareSafetyError, 'cbAOut'): self.off()
        self.assertEqual(self.bits[0], 1)
        self.assertTrue(self.fields.last_record.pulse['safe_state_confirmed'])
        self.assertTrue(self.fields.last_record.af['safe_state_confirmed'])
