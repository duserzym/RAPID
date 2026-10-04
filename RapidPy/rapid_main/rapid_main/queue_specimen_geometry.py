"""Immutable measured geometry borrowed by a claimed production queue worker."""
import copy
from dataclasses import asdict
import json
import math

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_routing import RoutedMotorSerialClient
from .queue_lift_transfer import QueueSpecimenState
from .queue_station import QueueStationGeometry


def _snapshot(value):
    return json.loads(json.dumps(value, allow_nan=False))


class QueueSpecimenGeometry:
    def __init__(self, session, config, motor, axes, store):
        if not isinstance(motor, RoutedMotorSerialClient):
            raise HardwareSafetyError('Measured queue geometry requires the original routed native motor client.')
        self.session, self.config, self.motor, self.axes = session, config, motor, axes
        original = getattr(store, 'session', None)
        path = original.store.path if original is not None else store.path
        if path.resolve() != session.store.path.resolve() or (original is not None and original is not session):
            raise HardwareSafetyError('Measured geometry and actuators must share the original queue journal.')
        self.station = QueueStationGeometry.from_config(config, use_xy_table=config.motor_station.use_xy_table)
        self.profile = copy.deepcopy(session.store._queue(session.token)['profile']['stage_profiles']['motion'])
        self._validate_station()
        self.context = self._read_context()

    def _validate_station(self):
        from .hardware_contracts import _build_motor_controller_config
        self.session.child_store._owned()
        root = self.session.store._queue(self.session.token)
        station = QueueStationGeometry.from_config(self.config, use_xy_table=self.config.motor_station.use_xy_table)
        acquisition_profile = dict(helper='rapid_main_queue_acquisition',
            **{name: asdict(getattr(self.config, name)) for name in ('motion', 'squid', 'calibration', 'susceptibility')})
        if (self.session.is_recovery or self.config.general.nocomm is not False
                or root['profile'].get('helper') != 'rapid_main_queue'
                or root['profile']['stage_profiles'].get('motion') != self.profile
                or root['profile']['stage_profiles'].get('acquisition') != _snapshot(acquisition_profile)
                or self.profile.get('helper') != 'rapid_main_queue_xy'
                or self.profile.get('geometry') != _snapshot(asdict(station))
                or self.profile.get('controller') != _snapshot(asdict(self.motor.config))
                or self.profile.get('controller') != _snapshot(asdict(_build_motor_controller_config(self.config)))
                or self.config.motor_station.ports != {key: axis.port for key, axis in self.axes.items()}
                or self.config.motor_station.addresses != {key: axis.address for key, axis in self.axes.items()}
                or self.profile.get('axes') != _snapshot({key: asdict(axis) for key, axis in self.axes.items()})):
            raise HardwareSafetyError('Measured specimen geometry requires the original live queue station binding.')
        for axis in self.axes.values():
            if self.motor._bindings.get(axis.motor_id) != (axis.name, axis.address, axis.port):
                raise HardwareSafetyError('Actual motor routing differs from the measured specimen station.')
        return root

    def _read_context(self):
        value = self.session.store.latest_transfer_context(self.session.token)
        state = QueueSpecimenState.read(value, self.session, self.station)
        if (state.phase != 'lifted' or type(state.sample_height) is not int
                or not 0 < state.sample_height <= abs(self.motor.config.sample_bottom)):
            raise HardwareSafetyError('Verified loaded specimen height is required before treatment or acquisition.')
        return state

    def height(self, sample_id, *, own_stage_token=None):
        root = self._validate_station()
        if sample_id != self.context.sample_id:
            raise HardwareSafetyError('Measured specimen height belongs to a different sample.')
        stage = root['stage']
        if stage and stage['status'] == 'pending':
            if (own_stage_token is None or stage['token'] != own_stage_token
                    or stage['sample_id'] != sample_id
                    or stage['plan'].get('specimen_geometry') != self.context.to_dict()):
                raise HardwareSafetyError('Recover the unfinished original stage before using specimen geometry.')
        elif self._read_context() != self.context:
            raise HardwareSafetyError('Original loaded specimen geometry changed; bind the new verified transfer.')
        return self.context.sample_height

    def positions(self, sample_id, *, own_stage_token=None):
        height = self.height(sample_id, own_stage_token=own_stage_token)
        return (math.floor(self.config.motion.zero_pos + height / 2),
                math.floor(self.config.motion.meas_pos + height / 2))
