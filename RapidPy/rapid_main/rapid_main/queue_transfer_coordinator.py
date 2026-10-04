"""Compose native specimen transfer phases on the original queue worker."""
from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_field_outputs import QueueFieldOutputs
from .queue_lift_transfer import QueueLiftTransfer
from .queue_xy_reference import QueueXYReference


class QueueTransferCoordinator:
    def __init__(self, session, lift, reference, fields):
        if (not isinstance(lift, QueueLiftTransfer) or not isinstance(reference, QueueXYReference)
                or not isinstance(fields, QueueFieldOutputs) or reference.table is not lift.table
                or reference.vacuum is not lift.vacuum):
            raise HardwareSafetyError('Transfer coordination requires the original native station services.')
        self.session, self.lift, self.reference, self.fields = session, lift, reference, fields
        self.table, self.vacuum = lift.table, lift.vacuum

    def _validate(self):
        self.table._validate(self.session)
        binding = self.vacuum._queue_binding
        if (binding is None or binding.session is not self.session
                or self.vacuum.output_state_known is not True or self.vacuum.is_pump_on() is not True):
            raise HardwareSafetyError('Original acknowledged queue pump ownership is required.')
        binding._validate_owner()
        self.fields._validate(self.session)
        from .queue_holder_geometry import QueueHolderState
        holder = self.session.store.latest_holder_context(self.session.token)
        if holder is not None and QueueHolderState.read(holder, self.session, self.table.geometry).phase != 'clear':
            raise HardwareSafetyError('Clear the original blank-holder rod before specimen transfer.')
        return self.reference.require(self.session)

    def _outputs(self, valve, field_proof):
        pose = self.lift.verify_vacuum_pose(self.session, field_outputs_off_verified=field_proof)
        return self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=valve,
            transfer_pose=pose, field_outputs_off_verified=field_proof)

    def load(self, original_slot, sample_id, *, file_id=''):
        """Load once, reference measured height and park above the empty hole."""
        reference = self._validate()
        previous = self.lift.context(self.session)
        if previous is not None and previous.phase != 'clear':
            raise HardwareSafetyError('Return the original specimen before loading another.')
        self.table.geometry.specimen_slot(original_slot)
        if not isinstance(sample_id, str) or not sample_id.strip() or not isinstance(file_id, str):
            raise HardwareSafetyError('Original specimen and file identity are required.')
        if self.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('Loading requires the original acknowledged gripper valve OFF.')
        field_proof = self.fields.verify_off(self.session)
        self.table.move_to_slot(self.session, original_slot, reference_verified=reference,
            field_outputs_off_verified=field_proof, sample_id=sample_id)
        self.lift.pickup(self.session, original_slot, sample_id, file_id=file_id,
            field_outputs_off_verified=field_proof)
        self._outputs(True, field_proof)
        self.lift.home_loaded(self.session, field_outputs_off_verified=field_proof)
        self.table.move_to_slot(self.session, self.table.geometry.nearest_empty(original_slot),
            reference_verified=reference, field_outputs_off_verified=field_proof, sample_id=sample_id)
        return self.lift.context(self.session)

    def return_specimen(self):
        """Return only the original measured specimen; release after support proof."""
        reference = self._validate()
        state = self.lift.context(self.session)
        if state is None or state.phase != 'lifted':
            raise HardwareSafetyError('An original verified loaded specimen is required for return.')
        field_proof = self.fields.verify_off(self.session)
        self.lift.raise_loaded_to_clearance(self.session, field_outputs_off_verified=field_proof)
        self.table.move_to_slot(self.session, state.original_slot, reference_verified=reference,
            field_outputs_off_verified=field_proof, sample_id=state.sample_id)
        self.lift.lower_for_dropoff(self.session, field_outputs_off_verified=field_proof)
        self._outputs(False, field_proof)
        self.lift.clear_after_release(self.session, field_outputs_off_verified=field_proof)
        return self.lift.context(self.session)
