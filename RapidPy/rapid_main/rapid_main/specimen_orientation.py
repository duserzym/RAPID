"""VB6 specimen, geographic and tilt/fold directions in the down-positive convention."""

from dataclasses import dataclass
import math

from .data_model import SpecimenMeta
from .geometry import Angular3D, angular3d_to_cartesian3d, atan2_custom, rad_to_deg


@dataclass(frozen=True)
class SpecimenDirections:
    core: Angular3D
    geographic: Angular3D
    structural: Angular3D


def _declination(north, east):
    return rad_to_deg(atan2_custom(north, east), wrap_0_360=True)


def _untilt(direction, dip_direction_deg, dip_deg):
    """Rotate in the vertical plane of greatest dip, as in Measure_Rotate."""
    if dip_deg == 0.:
        return direction
    vector = angular3d_to_cartesian3d(direction)
    azimuth, dip = map(math.radians, (dip_direction_deg, dip_deg))
    north, east = math.cos(azimuth), math.sin(azimuth)
    horizontal = vector.x * north + vector.y * east
    change = horizontal * (math.cos(dip) - 1) + vector.z * math.sin(dip)
    x, y = vector.x + north * change, vector.y + east * change
    z = vector.z * math.cos(dip) - horizontal * math.sin(dip)
    return Angular3D(_declination(x, y), math.degrees(math.atan2(z, math.hypot(x, y))), direction.mag)


def specimen_directions(vector, meta: SpecimenMeta, *, label='', raw_squid=False) -> SpecimenDirections:
    """Apply Measure_Unfold and Measure_Bedding to an accepted moment direction."""
    x, y, z = vector
    if label[:5].upper() == 'ATAFX':
        x, y, z = -z, x, -y
    elif label[:5].upper() == 'ATAFY':
        x, y, z = x, z, -y
    magnitude = math.hypot(x, y, z)
    horizontal = math.hypot(x, y) + (1e-5 if raw_squid else 0.)
    core = Angular3D(_declination(x, y), math.degrees(math.atan2(z, horizontal)), magnitude)
    if not meta.has_core_orientation or magnitude == 0:
        return SpecimenDirections(core, core, core)

    magnetic_declination = meta.magnetic_declination_e or 0.
    dip = math.radians(90. - meta.core_plate_dip)
    x, y, z = x / magnitude, y / magnitude, z / magnitude
    north = x * math.sin(dip) + z * math.cos(dip)
    down = z * math.sin(dip) - x * math.cos(dip)
    geographic = Angular3D(
        (_declination(north, y) + meta.core_plate_strike + magnetic_declination - 90.) % 360.,
        math.degrees(math.atan2(down, math.hypot(north, y) or 1e-5)), magnitude)
    strike = meta.bedding_strike + magnetic_declination
    if meta.fold_axis is None and meta.fold_plunge is None:
        structural = _untilt(geographic, strike + 90., meta.bedding_dip)
    else:
        fold_direction = (meta.fold_axis or 0.) + magnetic_declination
        fold_plunge = meta.fold_plunge or 0.
        unfolded = _untilt(geographic, fold_direction, fold_plunge)
        normal = _untilt(Angular3D(strike - 90., 90. - meta.bedding_dip, 1.), fold_direction, fold_plunge)
        structural = _untilt(unfolded, normal.dec + 180., 90. - normal.inc)
    return SpecimenDirections(core, geographic, structural)
