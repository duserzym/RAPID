"""Read both VB6's complete RMG rows and the earlier sparse sidecars."""

from pathlib import Path
import csv
import math

from rapid_main.data_model import RmgRecord


def read_rmg(path: str | Path) -> list[RmgRecord]:
    records = []
    with Path(path).open(encoding='latin-1', newline='') as handle:
        for fields in csv.reader(handle):
            if len(fields) < 2 or not fields[0].strip():
                continue
            padded = fields + [''] * (20 - len(fields))

            def number(index):
                text = padded[index].strip()
                if not text:
                    return None
                value = float(text)
                if not math.isfinite(value):
                    raise ValueError('Nonfinite RMG field')
                return value

            try:
                level = number(1)
                if level is None and len(fields) < 9:
                    continue
                records.append(RmgRecord(padded[0].strip(), level, number(8) or 0., fields,
                    bias_field_g=number(2), spin_speed_rps=number(3), hold_time_s=number(4),
                    moment_z_emu=number(5), sd_z_emu=number(6), intensity_z_emu_per_cm3=number(7),
                    moment_x_emu=number(9), sd_x_emu=number(10), moment_y_emu=number(11), sd_y_emu=number(12),
                    remarks=padded[13].strip(), core_dec=number(14), core_inc=number(15),
                    intensity_emu_per_cm3=number(16), csd=number(17), sample_height_cm=number(18),
                    timestamp=padded[19].strip()))
            except ValueError:
                # Specimen metadata, headings and Rockmag info lines can
                # precede or interrupt the numeric rows.
                continue
    return records
