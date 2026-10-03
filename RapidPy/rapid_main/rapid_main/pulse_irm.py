"""Calibrated capacitor pulse planning from frmIRMARM's field conversion.

Plans are immutable and perform no hardware I/O. Circuit execution, relay
binding and capacitor-readback qualification are separate requirements.
"""
from dataclasses import dataclass
import math


@dataclass(frozen=True)
class PulseIrmPlan:
    field_mT: float
    coil: str
    backfield: bool
    capacitor_v: float
    control_v: float
    charge_control_v: float
    capacitor_tolerance_v: float
    fire_hold_s: float
    calibration_source: str
    system: str
    operation: str = "charged_pulse"


def plan_pulse_irm(field_mT, coil, cfg):
    if coil not in {"axial", "transverse"}:
        raise ValueError("Pulse IRM coil must be axial or transverse.")
    if cfg.system.upper() not in {"OLD", "ASC", "MATSUSADA"}:
        raise ValueError("A supported calibrated pulse IRM system is required.")
    if not getattr(cfg, f"{coil}_enabled") or not cfg.calibration_source.strip():
        raise ValueError(f"Enabled and configured {coil} pulse IRM is required.")
    field = float(field_mT)
    if not math.isfinite(field):
        raise ValueError("Pulse IRM field must be finite.")
    if field == 0:
        # RockmagStep fires two residual discharges at LOAD, moves to the coil,
        # and explicitly skips FireIRMAtField. No magnetic field table is used.
        values = (cfg.feedback_v_per_capacitor_v, cfg.control_v_per_capacitor_v,
                  cfg.control_max_v, getattr(cfg, f"{coil}_capacitor_max_v"), cfg.discharged_max_v)
        if any(not math.isfinite(float(value)) or value <= 0 for value in values) or cfg.control_max_v > 10 or cfg.discharged_max_v > 10:
            raise ValueError("Zero-field IRM requires accepted capacitor feedback and safe voltage limits.")
        return PulseIrmPlan(0., coil, False, 0., 0., 0., cfg.discharged_max_v,
                            1 if coil == "axial" else 3, cfg.calibration_source, cfg.system, "zero_field")
    if not getattr(cfg, f"{coil}_calibrated"):
        raise ValueError(f"Accepted {coil} pulse IRM field calibration is required.")
    if field < 0 and not cfg.backfield_enabled:
        raise ValueError("Backfield IRM requires an enabled polarity relay route.")
    magnitude = abs(field)
    minimum, maximum = (float(getattr(cfg,f"{coil}_{bound}_mT")) for bound in ("min","max"))
    if not all(math.isfinite(value) for value in (minimum,maximum)) or not 0 <= minimum <= maximum or maximum <= 0:
        raise ValueError("Invalid pulse IRM field limits.")
    if not minimum <= magnitude <= maximum:
        raise ValueError("Pulse IRM field is outside its accepted calibration limits.")
    try:
        raw = getattr(cfg,f"{coil}_calibration")
        points = [(float(point[0]),float(point[1])) for point in raw if len(point)==2]
        if len(points) != len(raw) or len(points)<2:
            raise ValueError("Incomplete capacitor calibration.")
    except (TypeError,ValueError,IndexError) as exc:
        raise ValueError("Invalid pulse IRM capacitor calibration.") from exc
    if any(not math.isfinite(x) or not math.isfinite(y) or x < 0 or y <= 0 for x,y in points):
        raise ValueError("Invalid pulse IRM capacitor calibration.")
    if any(b[0] < a[0] or b[1] <= a[1] for a,b in zip(points,points[1:])):
        raise ValueError("Pulse IRM calibration fields must increase and capacitor voltage must not decrease.")
    points = [(0.,0.),*points]
    low, high = points[-2:]
    for a,b in zip(points,points[1:]):
        if magnitude <= b[1]:
            low,high = a,b
            break
    capacitor = low[0]+(magnitude-low[1])*(high[0]-low[0])/(high[1]-low[1])
    limits = (getattr(cfg,f"{coil}_capacitor_max_v"),cfg.control_v_per_capacitor_v,
              cfg.feedback_v_per_capacitor_v,cfg.control_max_v)
    if any(not math.isfinite(float(value)) or value <= 0 for value in limits):
        raise ValueError("Pulse IRM voltage limits and conversions must be positive and finite.")
    if capacitor > limits[0] or cfg.control_max_v > 10:
        raise ValueError("Pulse IRM capacitor/control voltage exceeds its calibrated limit.")
    control = capacitor*cfg.control_v_per_capacitor_v
    boost = 1
    if cfg.system.upper()=="ASC":
        if any(not math.isfinite(float(value)) or value < 1 for value in (cfg.asc_boost_at_min,cfg.asc_boost_at_max)):
            raise ValueError("Invalid ASC charge boost calibration.")
        # CalculateAscBoostMultiplier uses the axial capacitor maximum.
        if not math.isfinite(cfg.axial_capacitor_max_v) or cfg.axial_capacitor_max_v <= 0:
            raise ValueError("ASC axial capacitor maximum must be calibrated.")
        boost = round(cfg.asc_boost_at_min+(capacitor/cfg.axial_capacitor_max_v)*(cfg.asc_boost_at_max-cfg.asc_boost_at_min),2)
    charge = control*boost
    if cfg.system.upper()=="MATSUSADA":
        # ScaleUp modifies the capacitor target by reference in the legacy.
        boost = (1.5 if capacitor <= 20 else 1.33 if capacitor <= 50 else 1.2 if capacitor <= 100
                 else 1.1 if capacitor <= 200 else 1.02 if capacitor <= 350 else 1.01)
        charge = control*boost
    if not 0 <= control <= cfg.control_max_v or not 0 <= charge <= cfg.control_max_v:
        raise ValueError("Pulse IRM charge target exceeds its calibrated DAC limit; no silent clipping is permitted.")
    tolerance = max(.5,capacitor*.005) if cfg.system.upper()=="OLD" else capacitor*.005
    return PulseIrmPlan(field,coil,field<0,capacitor,control,charge,tolerance,
                        1 if coil=="axial" else 3,cfg.calibration_source,cfg.system)
