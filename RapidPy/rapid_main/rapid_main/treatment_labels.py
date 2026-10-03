"""Typed field labels with explicit unit conversion at the actuator boundary."""
from dataclasses import dataclass
from decimal import Decimal
import math
import re


_FIELD = re.compile(r"^(AFMAX|AFZ|AF|IRM[XYZ]?|ARM)\s*([+-]?\d+(?:\.\d+)?)?\s*(MT|G)?(?:_([0-9]+(?:\.[0-9]+)?)\s*(MT|G)?)?$")


def format_field_value(value: float) -> str:
    """Keep a field's decimal value without rounding it into a different step."""
    numeric = float(value)
    if not math.isfinite(numeric): raise ValueError("Treatment fields must be finite.")
    text = format(Decimal(str(numeric)), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


@dataclass(frozen=True)
class FieldTreatmentRequest:
    label: str
    family: str
    field_mT: float | None
    bias_mT: float | None
    axis: str | None


def parse_field_treatment(label: str) -> FieldTreatmentRequest:
    text = label.strip().upper()
    match = _FIELD.fullmatch(text)
    if match is None:
        raise ValueError(f"Invalid field treatment label: {label!r}")
    prefix, value, unit, bias, bias_unit = match.groups()
    if (unit and value is None) or (bias is not None and prefix != "ARM"):
        raise ValueError(f"Invalid field/bias units in treatment label: {label!r}")
    def convert(raw, units):
        if raw is None: return None
        numeric = float(raw) * (.1 if units == "G" else 1.)
        if not math.isfinite(numeric): raise ValueError("Treatment fields must be finite.")
        return numeric
    family = "IRM" if prefix.startswith("IRM") else prefix
    axis = prefix[-1] if prefix in {"IRMX", "IRMY", "IRMZ"} else None
    # Unsuffixed labels preserve RapidPy's existing mT execution contract.
    return FieldTreatmentRequest(label, family, convert(value, unit), convert(bias, bias_unit), axis)
