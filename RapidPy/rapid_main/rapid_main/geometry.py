"""Geometry utilities migrated from VB6 vector helpers."""

from __future__ import annotations

from dataclasses import dataclass
import math

PI: float = 3.141592653589
RAD: float = PI / 180.0
DEG: float = 180.0 / PI


def deg_to_rad(angle_deg: float) -> float:
    """Convert degrees to radians."""

    return angle_deg * RAD


def rad_to_deg(angle_rad: float, wrap_0_360: bool = False) -> float:
    """Convert radians to degrees.

    Parameters
    ----------
    wrap_0_360:
        When true, wrap into [0, 360).
    """

    deg = angle_rad * DEG
    if not wrap_0_360:
        return deg

    while deg >= 360.0:
        deg -= 360.0
    while deg < 0.0:
        deg += 360.0
    return deg


def acos(theta: float) -> float:
    """VB6-compatible arccos-style helper."""

    if theta >= 1:
        return 0.0
    if theta <= -1:
        return PI
    return math.atan2(math.sqrt(1 - theta * theta), theta)


def arcos(angle: float) -> float:
    """Alias kept for compatibility with VB6 ``arcos``."""

    return acos(angle)


def atn(x: float, y: float) -> float:
    """VB6-compatible ``atan`` style helper (keeps legacy quadrant branches)."""

    if x > 0:
        if y > 0:
            return math.atan(y / x)
        return 2 * PI + math.atan(y / x)
    if x < 0:
        return PI + math.atan(y / x)
    if y > 0:
        return PI
    if y < 0:
        return -PI
    return 0.0


def atan2_custom(x_c: float, y_c: float) -> float:
    """VB6-compatible ``Atan2`` used by legacy workflows."""

    if x_c == 0 and y_c >= 0:
        return PI / 2
    if x_c == 0 and y_c < 0:
        return 3 * PI / 2

    angle = math.atan(abs(y_c / x_c))
    if y_c >= 0 and x_c > 0:
        return angle
    if y_c >= 0 and x_c < 0:
        return PI - angle
    if y_c < 0 and x_c < 0:
        return PI + angle
    if y_c < 0 and x_c > 0:
        return 2 * PI - angle
    return angle


@dataclass
class Angular3D:
    """Direction vector represented by declination/inclination/magnitude."""

    dec: float
    inc: float
    mag: float

    @property
    def x(self) -> float:
        """X component in the geographic/north convention."""

        p = math.cos(deg_to_rad(self.inc))
        return self.mag * p * math.cos(deg_to_rad(self.dec))

    @property
    def y(self) -> float:
        """Y component in the geographic/east convention."""

        p = math.cos(deg_to_rad(self.inc))
        return self.mag * p * math.sin(deg_to_rad(self.dec))

    @property
    def z(self) -> float:
        """Z component (down-positive convention)."""

        return self.mag * math.sin(deg_to_rad(self.inc))


@dataclass
class Cartesian3D:
    """Cartesian vector in north/east/down convention."""

    x: float
    y: float
    z: float

    @property
    def dec(self) -> float:
        """Direction declination in degrees."""

        return rad_to_deg(atn(self.x, self.y))

    @property
    def inc(self) -> float:
        """Direction inclination in degrees."""

        horiz_sq = self.x * self.x + self.y * self.y
        return rad_to_deg(atn(math.sqrt(horiz_sq), self.z))

    @property
    def mag(self) -> float:
        """Vector magnitude."""

        return math.sqrt(self.x * self.x + self.y * self.y + self.z * self.z)

    @property
    def unit_x(self) -> float:
        magnitude = self.mag
        return 1.0 if magnitude == 0 else self.x / magnitude

    @property
    def unit_y(self) -> float:
        magnitude = self.mag
        return 0.0 if magnitude == 0 else self.y / magnitude

    @property
    def unit_z(self) -> float:
        magnitude = self.mag
        return 0.0 if magnitude == 0 else self.z / magnitude


@dataclass(frozen=True)
class InterpolationRange:
    """One linear interpolation segment.

    Used as the RapidPy replacement for VB6 ``InterpolationRange`` class
    behavior in fitting/transverse workflows.
    """

    start: float
    end: float
    start_value: float
    end_value: float

    def __post_init__(self) -> None:
        if self.end == self.start:
            raise ValueError("InterpolationRange requires distinct start/end values")

    @property
    def low(self) -> float:
        return min(self.start, self.end)

    @property
    def high(self) -> float:
        return max(self.start, self.end)

    def contains(self, x: float) -> bool:
        return self.low <= float(x) <= self.high

    def value_at(self, x: float, *, clamp: bool = False) -> float:
        x_value = float(x)
        if clamp:
            x_value = max(self.low, min(self.high, x_value))
        elif not self.contains(x_value):
            raise ValueError(f"{x_value} is outside interpolation range {self.low}..{self.high}")

        fraction = (x_value - self.start) / (self.end - self.start)
        return self.start_value + fraction * (self.end_value - self.start_value)


@dataclass(frozen=True)
class InterpolationRanges:
    """Ordered collection of interpolation segments."""

    ranges: tuple[InterpolationRange, ...]

    def __init__(self, ranges: list[InterpolationRange] | tuple[InterpolationRange, ...]) -> None:
        ordered = tuple(sorted(ranges, key=lambda item: item.low))
        if not ordered:
            raise ValueError("InterpolationRanges requires at least one range")
        object.__setattr__(self, "ranges", ordered)

    @property
    def low(self) -> float:
        return self.ranges[0].low

    @property
    def high(self) -> float:
        return self.ranges[-1].high

    def range_for(self, x: float) -> InterpolationRange:
        for item in self.ranges:
            if item.contains(x):
                return item
        raise ValueError(f"{float(x)} is outside interpolation ranges {self.low}..{self.high}")

    def value_at(self, x: float, *, clamp: bool = False) -> float:
        try:
            return self.range_for(x).value_at(x)
        except ValueError:
            if not clamp:
                raise
            if float(x) < self.low:
                return self.ranges[0].value_at(self.low)
            return self.ranges[-1].value_at(self.high)


def angular3d_to_cartesian3d(vector: Angular3D) -> Cartesian3D:
    """Convert angular direction into Cartesian coordinates."""

    inc_rad = deg_to_rad(vector.inc)
    dec_rad = deg_to_rad(vector.dec)
    cos_inc = math.cos(inc_rad)
    return Cartesian3D(
        x=vector.mag * cos_inc * math.cos(dec_rad),
        y=vector.mag * cos_inc * math.sin(dec_rad),
        z=vector.mag * math.sin(inc_rad),
    )


def angular3d_to_viewer_cartesian3d(vector: Angular3D, viewer_down_positive: bool = False) -> Cartesian3D:
    """Convert angular direction for legacy specimen-reader orientation conventions.

    Parameters
    ----------
    viewer_down_positive:
        When False (legacy rapid_main viewer behavior), this uses ``z = -mag*sin(inc)``.
    """

    cart = angular3d_to_cartesian3d(vector)
    if viewer_down_positive:
        return cart
    return Cartesian3D(cart.x, cart.y, -cart.z)


def cartesian3d_to_angular3d(vector: Cartesian3D) -> Angular3D:
    """Convert Cartesian coordinates into angular direction."""

    horizontal_sq = vector.x * vector.x + vector.y * vector.y
    mag = vector.mag
    inc = rad_to_deg(atn(math.sqrt(horizontal_sq), vector.z))
    if inc > 180:
        inc -= 360
    dec = rad_to_deg(atn(vector.x, vector.y))
    return Angular3D(dec=dec, inc=inc, mag=mag)


def cartesian3d_dot_product(v1: Cartesian3D, v2: Cartesian3D) -> float:
    return v1.x * v2.x + v1.y * v2.y + v1.z * v2.z


def cartesian3d_diff_angle(v1: Cartesian3D, v2: Cartesian3D) -> float:
    if v1.mag == 0 or v2.mag == 0:
        return 0.0
    return rad_to_deg(acos(cartesian3d_dot_product(v1, v2) / (v1.mag * v2.mag)))


def cartesian3d_sum(v1: Cartesian3D, v2: Cartesian3D) -> Cartesian3D:
    return Cartesian3D(v1.x + v2.x, v1.y + v2.y, v1.z + v2.z)


def cartesian3d_difference(v1: Cartesian3D, v2: Cartesian3D) -> Cartesian3D:
    return Cartesian3D(v1.x - v2.x, v1.y - v2.y, v1.z - v2.z)


def cartesian3d_average(v1: Cartesian3D, v2: Cartesian3D) -> Cartesian3D:
    return Cartesian3D((v1.x + v2.x) / 2.0, (v1.y + v2.y) / 2.0, (v1.z + v2.z) / 2.0)


def cartesian3d_scalar_div(v1: Cartesian3D, divisor: float) -> Cartesian3D:
    return Cartesian3D(v1.x / divisor, v1.y / divisor, v1.z / divisor)


def cartesian3d_scalar_mult(v1: Cartesian3D, multiplier: float) -> Cartesian3D:
    return Cartesian3D(v1.x * multiplier, v1.y * multiplier, v1.z * multiplier)


def cartesian3d_square(v1: Cartesian3D) -> Cartesian3D:
    return Cartesian3D(v1.x * v1.x, v1.y * v1.y, v1.z * v1.z)


def cartesian3d_square_root(v1: Cartesian3D) -> Cartesian3D:
    return Cartesian3D(math.sqrt(v1.x), math.sqrt(v1.y), math.sqrt(v1.z))


def cartesian3d_zero() -> Cartesian3D:
    return Cartesian3D(0.0, 0.0, 0.0)
