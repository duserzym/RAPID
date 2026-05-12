"""
magic_method_codes.py — Map RAPID step-label prefixes to MagIC LP-* method codes.

Reference: resources/MagIC Method Codes.json
MagIC v3.0 standard method codes used for measurements table.
"""
from __future__ import annotations


# Ordered list of (prefix, method_code) pairs — checked longest-first
_PREFIX_MAP: list[tuple[str, str]] = [
    ("AFMAX",  "LP-DIR-AF"),
    ("AFZ",    "LP-DIR-AF"),
    ("AF",     "LP-DIR-AF"),
    ("TT",     "LP-DIR-T"),
    ("TH",     "LP-DIR-T"),
    ("TEMP",   "LP-DIR-T"),
    ("IRM",    "LP-IRM"),
    ("ARM",    "LP-ARM"),
    ("NRM",    "LP-NRM"),
    ("PTRM",   "LP-PI-TRM"),
    ("ZF",     "LP-NONE"),   # zero-field step
    ("IF",     "LP-NONE"),   # in-field step
]

# Additional suffix codes appended for combined treatments
_COMBINED_CODE = "LP-DIR-AF:LP-IRM"  # e.g. IRM+AF demagnetisation


def label_to_method_codes(label: str) -> str:
    """
    Return the MagIC method code string for a given step label.

    Parameters
    ----------
    label:
        Step label from the specimen file, e.g. ``"NRM"``, ``"AF20"``,
        ``"TT400"``, ``"IRM1000"``.

    Returns
    -------
    str
        Colon-separated MagIC method codes, e.g. ``"LP-DIR-AF"``.
    """
    upper = label.strip().upper()
    for prefix, code in _PREFIX_MAP:
        if upper.startswith(prefix):
            return code
    # Unknown prefix — return a generic measurement code
    return "LP-NONE"


def method_codes_for_step_type(step_type: str) -> str:
    """
    Convenience wrapper for .rmg step_type strings (e.g. "AF", "TT").
    """
    return label_to_method_codes(step_type)
