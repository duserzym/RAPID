"""Calibrate a tray standard under the original production transfer journal."""
from dataclasses import asdict, dataclass
import time
from uuid import uuid4

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_diagnostic_safety import verify_stopped_in_place

from .queue_acquisition import run_queue_acquisition
from .queue_station import QueueStationGeometry
from .rod_calibration import RodFit, RodFitError, RodReading, RodScanPlan, plan_rod_scan, scan_rod
from .squid_rod_calibration import SquidRodFit, SquidRodPlan, coil_factors, plan_squid_rod, scan_squid_rod


def validate_tray_standard(config, plan):
    """Check the same typed settings used by live queue load and the proposed scan."""
    request = plan.request
    station = QueueStationGeometry.from_config(config, use_xy_table=config.motor_station.use_xy_table)
    if type(request.tray_slot) is not int:
        raise ValueError('Choose the registered slot containing the calibration standard.')
    slot = station.specimen_slot(request.tray_slot)
    if station.use_xy_table:
        station.xy_target(slot)
        station.xy_target(station.nearest_empty(slot))
    scale = config.motor_station.lift_counts_per_cm
    if isinstance(plan, SquidRodPlan):
        regular = plan_squid_rod(request, zero_pos=config.motion.zero_pos, meas_pos=config.motion.meas_pos,
            counts_per_cm=scale, axis_factors_emu=coil_factors(config.calibration), settle_s=config.squid.settle_time)
        if config.squid.range_label != '1×':
            raise ValueError('Select the calibrated 1× SQUID range before scanning the standard.')
    elif isinstance(plan, RodScanPlan):
        bridge = config.susceptibility
        regular = plan_rod_scan(request, coil=bridge.coil_position, travel_limit=config.motion.meas_pos,
                                counts_per_cm=scale, moment_factor_cgs=bridge.moment_factor_cgs)
        if bridge.enabled is not True:
            raise ValueError('Enable the accepted Bartington bridge before scanning the standard.')
    else:
        raise ValueError('Select SQUID or susceptibility rod calibration for automatic loading.')
    if plan != regular or not 0 < request.height_cm * scale <= abs(config.motion.sample_bottom):
        raise ValueError('The certified height and scan plan must fit the accepted specimen pickup geometry.')
    return station


@dataclass(frozen=True)
class _CalibrationEvidence:
    fit: RodFit | SquidRodFit | None
    scientific_error: str
    readings: tuple


class TrayRodScan:
    def __init__(self, backend, vacuum, plan, *, operator, chain_slot=None):
        self.backend = backend
        self.vacuum = vacuum
        self.plan = plan
        self.operator = operator
        self.chain_slot = chain_slot
        self.startup = None
        self.run_id = 'rod-calibration-' + uuid4().hex

    def __call__(self, plan, readings, evidence, *, cancelled):
        if self.startup is not None or plan != self.plan:
            raise HardwareSafetyError('Run the original tray calibration plan once under a new queue owner.')
        backend = self.backend
        validate_tray_standard(backend._config, plan)
        if cancelled():
            raise InterruptedError('Tray calibration was canceled before startup.')
        original_sample = backend._sample_name
        original_run = backend._run_id
        original_cancel = backend._halt_check
        self.startup = backend.prepare_queue_lifetime(self.vacuum,
            dict(kind='tray_rod_calibration', calibration_plan=asdict(plan)), run_id=self.run_id,
            empty_rod_confirmed=True, operator=self.operator, chain_slot=self.chain_slot)
        session = self.startup.session
        transfer = evidence['tray_transfer'] = dict(queue_token=session.token, run_id=self.run_id)
        evidence['mount_confirmed_by'] = self.operator
        try:
            backend.set_halt_check(cancelled)
            with session.claim():
                coordinator = backend.start_queue_lifetime()
                backend.set_measurement_context(run_id=self.run_id)
                request = plan.request
                loaded = backend.load_queue_specimen(request.tray_slot, request.standard, file_id=self.run_id)
                transfer['loaded'] = loaded.to_dict()
                certified_counts = request.height_cm * plan.counts_per_cm
                transfer['certified_height_counts'] = certified_counts
                tolerance = max(1., certified_counts * request.threshold_percent / 200)
                if abs(loaded.sample_height - certified_counts) > tolerance:
                    outcome = _CalibrationEvidence(None,
                        'The measured tray standard height disagrees with its certified height.', ())
                else:
                    kind = 'squid' if isinstance(plan, SquidRodPlan) else 'susceptibility'
                    detector_label = 'squid' if kind == 'squid' else 'bridge'
                    coordinator._validate()
                    coordinator.fields.verify_off(session)
                    outcome = run_queue_acquisition(backend, kind,
                        lambda: self._scan(coordinator, kind, readings, cancelled), rod_plan=plan)
                    transfer['acquisition_record_id'] = backend._last_queue_acquisition.record_id
                    evidence[detector_label] = self.startup.instruments.profile[kind]
                returned = backend.return_queue_specimen()
                transfer['returned'] = returned.to_dict()
                terminal = backend._queue_terminal_cleanup
                record = backend.finish_queue_lifetime()
                transfer['shutdown'] = record.to_dict()
                evidence['motor_closed'] = True
                evidence['squid_closed'] = self.startup.instruments.closed['squid']
                evidence['bridge_closed'] = self.startup.instruments.closed['susceptibility']
                evidence['final_stopped_positions'] = {
                    axis['axis']: axis['samples'][-1]['position_raw']
                    for axis in terminal._close_plan['pose']['observations']
                }
            backend._run_id = original_run
            backend._sample_name = original_sample
        except Exception as exc:
            evidence['pending_queue_token'] = session.token
            raise HardwareSafetyError(f'{exc} Use the original interrupted queue recovery before further motion.') from exc
        finally:
            backend.set_halt_check(original_cancel)
        if outcome.scientific_error:
            raise RodFitError(outcome.scientific_error)
        return outcome.fit

    def _scan(self, coordinator, kind, readings, cancelled):
        plan, backend = self.plan, self.backend
        table, session = coordinator.table, coordinator.session
        motor, axis = table.motor, table.axes['updown']
        instruments = self.startup.instruments
        is_squid = kind == 'squid'
        sensor = instruments.squid if is_squid else instruments.bridge
        low = plan.deepest_target if is_squid else min(plan.positions)
        high = plan.baseline_position if is_squid else max(plan.positions)
        point_budget = len(plan.positions) + (2 * plan.fine_half_count + 6 if is_squid else 2)
        deadline = time.monotonic() + 360 + (30 + (plan.settle_s if is_squid else 0)) * point_budget
        before_cancel = motor._motion_cancel_check

        def verify_grip():
            motor._check_motion_cancel()
            binding = coordinator.vacuum._queue_binding
            if (binding is None or binding.session is not session or not binding.held
                    or coordinator.vacuum.output_state_known is not True
                    or coordinator.vacuum.is_pump_on() is not True
                    or coordinator.vacuum.is_valve_connected() is not True):
                raise HardwareSafetyError('The original acknowledged standard grip is required during calibration.')
            binding._validate_owner()

        def home():
            verify_grip()
            if motor.home_to_top(axis).success is not True:
                raise HardwareSafetyError('The loaded calibration standard did not reach lift top.')
            samples, error = verify_stopped_in_place(motor, axis, sleep=table.sleep)
            if error:
                raise HardwareSafetyError(error)
            table._clearance({'updown': samples})

        def move(target):
            verify_grip()
            if not low <= target <= high:
                raise HardwareSafetyError('The calibration target lies outside its accepted scan envelope.')
            outcome = motor.updown_move(axis, target, 0 if is_squid else 1)
            samples, error = verify_stopped_in_place(motor, axis, sleep=table.sleep)
            if error:
                raise HardwareSafetyError(error)
            position = samples[-1]['position_raw']
            if (outcome.success is not True or abs(position - target) > plan.readback_tolerance
                    or motor.check_internal_status(axis, 4) != 0):
                raise HardwareSafetyError('The loaded lift did not verify the calibration position.')
            verify_grip()
            return position

        def sample_bridge(target):
            home()
            zero = sensor.zero()
            return RodReading(target, move(target), sensor.measure(), zero)

        def sample_squid(target):
            position = move(target)
            ready = time.monotonic() + plan.settle_s
            while (remaining := ready - time.monotonic()) > 0:
                verify_grip()
                table.sleep(min(.05, remaining))
            verify_grip()
            return position, sensor.read_raw_xyz()

        motor._motion_cancel_check = lambda: cancelled() or time.monotonic() >= deadline or (
            before_cancel is not None and before_cancel())
        try:
            table._validate(session)
            verify_grip()
            home()
            if is_squid:
                sensor.connect_for_acquisition()
            else:
                sensor.test_connection()
            instruments.validate()
            try:
                fit = (scan_squid_rod(plan, sample_squid, readings, cancelled=cancelled) if is_squid else
                       scan_rod(plan, sample_bridge, readings, cancelled=cancelled))
                scientific_error = ''
            except RodFitError as exc:
                fit, scientific_error = None, str(exc)
            home()
            return _CalibrationEvidence(fit, scientific_error, tuple(readings))
        finally:
            motor._motion_cancel_check = before_cancel


class TrayRodReservation:
    """Hand the original application reservation back only after worker exit."""
    def __init__(self, scan, lease, settle):
        self.scan, self.lease, self.settle = scan, lease, settle

    def release(self):
        lease, self.lease = self.lease, None
        if lease is not None:
            self.settle(self.scan, lease)
