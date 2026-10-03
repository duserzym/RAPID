"""Manual ARM uses the queue lifecycle and retains success/failure evidence."""
import hashlib
import json
import math
import re
from pathlib import Path
import uuid

from .af_treatment import write_af_treatment_record


class ManualArmTreatment:
    simulated = False
    def __init__(self, backend, output_root):
        self.backend = backend
        self.output_root = Path(output_root)

    def is_connected(self):
        if self.backend.has_unresolved_rotation_fault or getattr(self.backend, "has_unresolved_hardware_fault", False) is True:
            source = getattr(self.backend, "_af_demag", None)
            if source is not None and source.is_connected():
                return True
        for name in ("_arm_bias","_pulse_irm"):
            source = getattr(self.backend,name,None)
            if source is not None and source.is_connected():
                return True
        return False

    def status(self):
        return "Calibrated ARM and capacitor pulse IRM workflows; disabled modules reject treatment requests."

    def apply(self, *, sample_id, peak_af_mT, bias_mT, should_cancel):
        if not sample_id.strip():
            raise ValueError("Enter the specimen identity before a manual treatment.")
        if not all(math.isfinite(float(value)) and float(value) >= 0 for value in (peak_af_mT, bias_mT)):
            raise ValueError("Manual ARM fields must be finite and non-negative.")
        label = f"ARM{float(peak_af_mT):g}MT_{float(bias_mT):g}MT"
        return self._execute(label=label,sample_id=sample_id,should_cancel=should_cancel,kind="ARM")

    def apply_irm(self,*,sample_id,max_field_mT,axis,should_cancel):
        if not math.isfinite(float(max_field_mT)) or float(max_field_mT)==0:
            raise ValueError("Manual pulse IRM field must be finite and nonzero.")
        if axis not in {"Z (up-axis)","X","Y"}:
            raise ValueError("Manual pulse IRM axis must be Z, X or Y.")
        axis_code = {"Z (up-axis)": "Z", "X": "X", "Y": "Y"}[axis]
        return self._execute(label=f"IRM{axis_code}{float(max_field_mT):g}MT",sample_id=sample_id,should_cancel=should_cancel,kind="IRM")

    def _execute(self,*,label,sample_id,should_cancel,kind):
        if not sample_id.strip():
            raise ValueError("Enter the specimen identity before a manual treatment.")
        ready = self.backend.validate_treatment_plan((label,))
        if not ready.ok:
            raise RuntimeError("; ".join(ready.blockers))
        run_id = "manual-" + uuid.uuid4().hex
        folder = self.output_root / run_id
        # Verify the evidence destination is writable before physical output.
        folder.mkdir(parents=True, exist_ok=False)
        index_path = folder / "artifact_index.json"
        index_path.write_text(json.dumps({"run_id": run_id, "state": "started", "artifacts": []}) + "\n", encoding="utf-8")
        record_attribute = "af_treatment_records" if kind=="ARM" else "pulse_treatment_records"
        start = len(getattr(self.backend,record_attribute))
        previous_halt = self.backend._halt_check
        previous_context = {name: getattr(self.backend, name) for name in ("_sample_name", "_run_id", "_treatment_label")}
        error = None
        artifacts = []
        try:
            self.backend.set_measurement_context(sample_name=sample_id.strip(), run_id=run_id, treatment_label=label)
            self.backend.set_halt_check(should_cancel)
            self.backend.set_demag_step(label)
        except Exception as exc:
            error = exc
            if kind=="IRM":
                try:
                    self.backend.return_to_safe_state()
                except Exception as recovery:
                    error = RuntimeError(f"{exc}; pulse safe-state recovery failed: {recovery}")
        finally:
            self.backend.set_halt_check(previous_halt)
            for name, value in previous_context.items():
                setattr(self.backend, name, value)
            publication_errors = []
            for record in getattr(self.backend,record_attribute)[start:]:
                try:
                    if not re.fullmatch(r"(?:af|irm)-[0-9a-f]{32}", record.treatment_id):
                        raise ValueError("Invalid manual treatment artifact identity.")
                    target = write_af_treatment_record(folder / f"{record.treatment_id}.json", record)
                    artifacts.append({"relative_path": target.name, "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
                                      "size_bytes": target.stat().st_size, "required": True})
                except Exception as exc:
                    publication_errors.append(str(exc))
            if not artifacts:
                publication_errors.append("No treatment evidence was produced.")
            payload = {"schema": "rapidpy.manual_treatment.artifact_index.v1", "run_id": run_id,
                       "simulated": False, "physical_acceptance_required": True,
                       "sample_id": sample_id.strip(), "label": label,
                       "state": "failed" if error or publication_errors else "completed",
                       "error": str(error) if error else "", "publication_errors": publication_errors,
                       "artifacts": artifacts}
            index_path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        if publication_errors:
            raise RuntimeError(f"{error or kind+' treatment'}; evidence publication failed: " + "; ".join(publication_errors)) from error
        if error:
            raise error
        return f"{kind} completed. Treatment evidence: {folder}"

    def reset(self):
        start = len(self.backend.pulse_treatment_records)
        af_start = len(self.backend.af_treatment_records)
        error = None
        try:
            self.backend.return_to_safe_state()
        except Exception as exc:
            error = exc
        records = (*self.backend.pulse_treatment_records[start:], *self.backend.af_treatment_records[af_start:])
        if records:
            folder = self.output_root/("recovery-"+uuid.uuid4().hex)
            try:
                folder.mkdir(parents=True,exist_ok=False)
                artifacts = []
                for record in records:
                    if not re.fullmatch(r"(?:irm|af)-[0-9a-f]{32}",record.treatment_id):
                        raise ValueError("Invalid field/rotation recovery identity.")
                    path = write_af_treatment_record(folder/f"{record.treatment_id}.json",record)
                    artifacts.append({"relative_path":path.name,"sha256":hashlib.sha256(path.read_bytes()).hexdigest(),"required":True})
                (folder/"artifact_index.json").write_text(json.dumps({"schema":"rapidpy.manual_recovery.artifact_index.v1",
                    "state":"failed" if error else "completed","error":str(error) if error else "", "artifacts":artifacts},indent=2)+"\n",encoding="utf-8")
            except Exception as publication:
                raise RuntimeError(f"{error or 'Field reset completed'}; recovery evidence publication failed: {publication}") from error
        if error:
            raise error
        return "Field reset and safe-state recovery completed."
