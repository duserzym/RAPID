"""VB6 Sample.WriteRockmagData rows, with Gauss for field levels."""

from __future__ import annotations

import csv
from dataclasses import dataclass, replace
from pathlib import Path
import re

from rapid_main.data_model import MeasurementStep, RmgRecord, SpecimenMeta
from rapid_main.rrm_treatment import parse_rrm_request
from rapid_main.treatment_labels import parse_field_treatment


# Keep the historical intensity value under VB6's original M (emu) heading.
RMG_COLUMNS = ('', 'Level', 'Bias Field (G)', 'Spin Speed (rps)', 'Hold Time (s)',
    'Mz (emu)', 'Std. Dev. Z', 'Mz/Vol', 'Moment Susceptibility (emu/Oe)',
    'Mx (emu)', 'Std. Dev. X', 'My (emu)', 'Std. Dev. Y', 'Remarks', 'Core Dec',
    'Core Inc', 'M (emu)', 'CSD', 'Sample Height (cm)', 'Date/Time')
_STEP_TYPES = {'AFMAX': 'AFmax', 'AFZ': 'AFz', 'RRMZ': 'RRMz'}
_SIGNED_LEVEL = re.compile(r'[+-]?\d+(?:\.\d+)?')


@dataclass(frozen=True)
class RmgTreatment:
    label: str
    step_type: str
    level: float | None = None
    bias_field_g: float | None = None
    spin_speed_rps: float | None = None
    hold_time_s: float | None = None
    evidence_id: str = ''


def describe_rmg_treatment(label: str, *, af=None, pulse=None) -> RmgTreatment:
    """Use accepted treatment evidence when present and explicit label values otherwise."""
    text = label.strip().upper()
    if text.startswith('RRM'):
        try:
            family, field, speed, bias = parse_rrm_request(label)
        except ValueError:
            # Early simulation programs also accepted a speed-only RRM label.
            return RmgTreatment(label, 'RRM')
        treatment = RmgTreatment(label, _STEP_TYPES.get(family, family), field * 10.,
            bias * 10. if bias is not None else None, speed)
        if af is not None:
            return replace(treatment, level=af.plan.ramp.field_mT * 10., spin_speed_rps=af.plan.speed_rps,
                bias_field_g=af.bias_mT * 10. if af.bias_mT is not None else None,
                hold_time_s=af.plan.ramp.hold_ms / 1000., evidence_id=af.treatment_id)
        return treatment

    try:
        request = parse_field_treatment(label)
    except ValueError:
        prefix = text.split(maxsplit=1)[0] if text else ''
        thermal = re.fullmatch(r'(TT|TH|TEMP)\s*([+-]?\d+(?:\.\d+)?)', text)
        if thermal:
            return RmgTreatment(label, thermal[1], float(thermal[2]))
        # Historical AF variants contain an axis numeral in the prefix.
        legacy_af = re.fullmatch(r'(ATAF[XYZ]|UAF[XYZ][12])\s*([+-]?\d+(?:\.\d+)?)', text)
        if legacy_af:
            family = legacy_af[1].replace('ATAF', 'aTAF')
            return RmgTreatment(label, family, float(legacy_af[2]))
        return RmgTreatment(label, prefix, 0. if text in {'NRM', 'SUSC'} else None)

    field, bias = request.field_mT, request.bias_mT
    step_type = _STEP_TYPES.get(request.family, request.family)
    if request.family == 'IRM':
        if pulse is not None:
            plan = pulse.circuit.plan
            step_type = 'IRMz' if plan.coil == 'axial' else 'IRMx'
            return RmgTreatment(label, step_type, plan.field_mT * 10., evidence_id=pulse.treatment_id)
        if request.axis:
            step_type = 'IRM' + request.axis.lower()
    hold = None
    if af is not None:
        ramps = [entry.ramp for entry in af.plan.passes]
        field = min(ramp.field_mT for ramp in ramps)
        bias = af.bias_mT
        holds = {ramp.hold_ms for ramp in ramps}
        if len(holds) == 1:
            hold = holds.pop() / 1000.
    return RmgTreatment(label, step_type, field * 10. if field is not None else None,
        bias * 10. if bias is not None else None, hold_time_s=hold,
        evidence_id=af.treatment_id if af is not None else '')


def _row(record: RmgRecord) -> list:
    if record.raw_fields:
        return record.raw_fields
    return [record.step_type, record.step_value, record.bias_field_g, record.spin_speed_rps, record.hold_time_s,
        record.moment_z_emu, record.sd_z_emu, record.intensity_z_emu_per_cm3, record.susceptibility,
        record.moment_x_emu, record.sd_x_emu, record.moment_y_emu, record.sd_y_emu, record.remarks,
        record.core_dec, record.core_inc, record.intensity_emu_per_cm3, record.csd, record.sample_height_cm,
        record.timestamp]


def append_rmg_record(path: str | Path, step: MeasurementStep, susceptibility: float = 0.0, *,
                      meta: SpecimenMeta | None = None, treatment: RmgTreatment | None = None,
                      axis_sd_emu: tuple[float, float, float] | None = None,
                      sample_height_cm: float | None = None) -> None:
    """Append the scientific row. Earlier direct callers retain their label's units."""
    path = Path(path)
    if treatment is None:
        match = _SIGNED_LEVEL.search(step.demag_label)
        prefix = step.demag_label[:match.start()].strip() if match else step.demag_label
        treatment = RmgTreatment(step.demag_label, prefix, float(match.group()) if match else 0.)
    if treatment.label != step.demag_label:
        raise ValueError('RMG treatment evidence belongs to another measurement step.')
    volume = meta.volume if meta is not None and meta.volume > 0 else 1.
    sd_x, sd_y, sd_z = axis_sd_emu or (None, None, None)
    intensity = step.intensity_emu_per_cm3 if step.intensity_emu_per_cm3 is not None else step.moment / volume
    record = RmgRecord(treatment.step_type, treatment.level, susceptibility,
        bias_field_g=treatment.bias_field_g, spin_speed_rps=treatment.spin_speed_rps, hold_time_s=treatment.hold_time_s,
        moment_z_emu=step.sdz, sd_z_emu=sd_z, intensity_z_emu_per_cm3=step.sdz / volume,
        moment_x_emu=step.sdx, sd_x_emu=sd_x, moment_y_emu=step.sdy, sd_y_emu=sd_y,
        core_dec=step.crdec, core_inc=step.crinc, intensity_emu_per_cm3=intensity, csd=step.error_angle,
        sample_height_cm=sample_height_cm, timestamp=step.timestamp.strftime('%Y-%m-%d %H:%M:%S'))
    new_file = not path.is_file() or path.stat().st_size == 0
    with path.open('a', encoding='latin-1', newline='') as handle:
        writer = csv.writer(handle)
        if new_file and meta is not None:
            writer.writerow((meta.name, meta.comment, f'Vol: {volume:g}', ''))
            writer.writerow(RMG_COLUMNS)
        writer.writerow(_row(record))


def write_rmg_records(path: str | Path, records: list[RmgRecord]) -> None:
    """Write the original columns where available, including unknown extensions."""
    with Path(path).open('w', encoding='latin-1', newline='') as handle:
        csv.writer(handle).writerows(_row(record) for record in records)
