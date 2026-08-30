from __future__ import annotations

from collections import deque
from pathlib import Path
from typing import Callable

from PySide6 import QtCore, QtWidgets

from rapid_main.calibration import (
    CalibrationResult,
    IrmVoltageCalibrationPoint,
    fit_irm_voltage_calibration,
    format_calibration_result,
    run_automated_squid_baseline_calibration,
    run_manual_calibration,
    write_calibration_result_artifact,
    write_irm_voltage_calibration_artifact,
)
from rapid_main.hardware_contracts import MeasurementBackend
from rapid_main.thermal import (
    ThermalSafetyLimits,
    compile_thermal_routine,
    write_thermal_routine_artifact,
)


class CalibrationCenterPanel(QtWidgets.QWidget):
    """Calibration workspace with automated and manual pathways.

    This is intentionally lightweight and UI-first so VB6 migration can move
    routine procedures into a single place before full backend coverage is in
    production.
    """

    PROCEDURES = {
        "gaussmeter_baseline": {
            "name": "SQUID / Gaussmeter Baseline",
            "expected_default": 1.0,
            "tolerance_default": 0.05,
        },
        "irm_voltage": {
            "name": "IRM Voltage Calibration",
            "expected_default": 1.0,
            "tolerance_default": 0.05,
        },
        "thermal_routine": {
            "name": "Thermal Routine Planning",
            "expected_default": 1.0,
            "tolerance_default": 0.05,
        },
    }

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        *,
        backend_provider: Callable[[], MeasurementBackend] | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend_provider = backend_provider
        self._history: deque[str] = deque(maxlen=8)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 12, 16, 16)
        root.setSpacing(10)

        hdr = QtWidgets.QLabel("CALIBRATION CENTER")
        hdr.setObjectName("sectionHdr")
        root.addWidget(hdr)

        card = QtWidgets.QFrame()
        card.setFrameShape(QtWidgets.QFrame.StyledPanel)
        fl = QtWidgets.QFormLayout(card)
        fl.setContentsMargins(16, 12, 16, 12)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._procedure = QtWidgets.QComboBox()
        self._procedure.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self._procedure.setMinimumContentsLength(12)
        self._procedure.addItem("SQUID / Gaussmeter Baseline", "gaussmeter_baseline")
        self._procedure.addItem("IRM Voltage Calibration", "irm_voltage")
        self._procedure.addItem("Thermal Routine Planning", "thermal_routine")
        self._procedure.currentIndexChanged.connect(self._on_procedure_changed)
        fl.addRow("Procedure:", self._procedure)

        self._mode = QtWidgets.QComboBox()
        self._mode.addItem("Automated (backend)", "automated")
        self._mode.addItem("Manual (operator-led)", "manual")
        self._mode.currentIndexChanged.connect(self._on_mode_changed)
        fl.addRow("Run mode:", self._mode)

        self._expected = QtWidgets.QDoubleSpinBox()
        self._expected.setRange(-1.0, 1.0)
        self._expected.setDecimals(6)
        self._expected.setValue(1.0)
        self._expected.setSuffix(" (target)")
        fl.addRow("Expected:", self._expected)

        self._samples = QtWidgets.QSpinBox()
        self._samples.setRange(1, 100)
        self._samples.setValue(5)
        fl.addRow("Automated sample count:", self._samples)

        self._tolerance = QtWidgets.QDoubleSpinBox()
        self._tolerance.setRange(0.0, 10.0)
        self._tolerance.setDecimals(6)
        self._tolerance.setValue(0.05)
        self._tolerance.setSuffix(" (max residual)")
        fl.addRow("Tolerance:", self._tolerance)

        self._manual_value = QtWidgets.QDoubleSpinBox()
        self._manual_value.setRange(-1.0, 1.0)
        self._manual_value.setDecimals(6)
        self._manual_value.setValue(1.0)
        self._manual_value.setSuffix(" (manual measured)")
        fl.addRow("Manual measured:", self._manual_value)

        self._baseline_context = QtWidgets.QLineEdit("operator-entry")
        self._baseline_operator = QtWidgets.QLineEdit()
        self._baseline_artifact_dir = QtWidgets.QLineEdit(
            str(Path.home() / ".rapid" / "calibrations")
        )
        fl.addRow("Run context:", self._baseline_context)
        fl.addRow("Operator:", self._baseline_operator)
        fl.addRow("Artifact folder:", self._baseline_artifact_dir)

        self._irm_box = QtWidgets.QGroupBox("IRM voltage fit")
        irm_form = QtWidgets.QFormLayout(self._irm_box)
        irm_form.setContentsMargins(12, 10, 12, 10)
        irm_form.setSpacing(8)
        irm_form.setLabelAlignment(QtCore.Qt.AlignRight)
        self._irm_field_1 = self._spin(0.0, 5000.0, 0.0, " mT")
        self._irm_voltage_1 = self._spin(0.0, 10.0, 0.0, " V")
        self._irm_field_2 = self._spin(0.0, 5000.0, 1000.0, " mT")
        self._irm_voltage_2 = self._spin(0.0, 10.0, 10.0, " V")
        self._irm_max_voltage = self._spin(0.001, 20.0, 10.0, " V")
        self._irm_force_zero = QtWidgets.QCheckBox("Force zero intercept")
        self._irm_context = QtWidgets.QLineEdit("operator-entry")
        self._irm_operator = QtWidgets.QLineEdit()
        self._irm_artifact_dir = QtWidgets.QLineEdit(
            str(Path.home() / ".rapid" / "calibrations")
        )
        irm_form.addRow("Point 1 field:", self._irm_field_1)
        irm_form.addRow("Point 1 voltage:", self._irm_voltage_1)
        irm_form.addRow("Point 2 field:", self._irm_field_2)
        irm_form.addRow("Point 2 voltage:", self._irm_voltage_2)
        irm_form.addRow("Max voltage:", self._irm_max_voltage)
        irm_form.addRow("", self._irm_force_zero)
        irm_form.addRow("Run context:", self._irm_context)
        irm_form.addRow("Operator:", self._irm_operator)
        irm_form.addRow("Artifact folder:", self._irm_artifact_dir)
        self._run_irm_btn = QtWidgets.QPushButton("Record IRM Voltage Fit")
        self._run_irm_btn.clicked.connect(self._run_irm_voltage)
        irm_form.addRow("", self._run_irm_btn)
        fl.addRow("", self._irm_box)

        self._thermal_box = QtWidgets.QGroupBox("Thermal routine planning")
        thermal_form = QtWidgets.QFormLayout(self._thermal_box)
        thermal_form.setContentsMargins(12, 10, 12, 10)
        thermal_form.setSpacing(8)
        thermal_form.setLabelAlignment(QtCore.Qt.AlignRight)
        self._thermal_name = QtWidgets.QLineEdit("Thermal routine")
        self._thermal_targets = QtWidgets.QLineEdit("100, 150, 200")
        self._thermal_prefix = QtWidgets.QComboBox()
        self._thermal_prefix.addItems(["TT", "TH", "TEMP"])
        self._thermal_hold_seconds = QtWidgets.QSpinBox()
        self._thermal_hold_seconds.setRange(0, 24 * 60 * 60)
        self._thermal_hold_seconds.setValue(600)
        self._thermal_hold_seconds.setSuffix(" s")
        self._thermal_max_temp = self._spin(1.0, 1200.0, 700.0, " C")
        self._thermal_ambient = self._spin(-50.0, 200.0, 25.0, " C")
        self._thermal_cool_to = self._spin(-50.0, 400.0, 50.0, " C")
        self._thermal_ramp_rate = self._spin(0.001, 500.0, 20.0, " C/min")
        self._thermal_context = QtWidgets.QLineEdit("operator-entry")
        self._thermal_operator = QtWidgets.QLineEdit()
        self._thermal_artifact_dir = QtWidgets.QLineEdit(
            str(Path.home() / ".rapid" / "calibrations")
        )
        thermal_form.addRow("Routine name:", self._thermal_name)
        thermal_form.addRow("Targets:", self._thermal_targets)
        thermal_form.addRow("Label prefix:", self._thermal_prefix)
        thermal_form.addRow("Hold time:", self._thermal_hold_seconds)
        thermal_form.addRow("Max safe temp:", self._thermal_max_temp)
        thermal_form.addRow("Ambient:", self._thermal_ambient)
        thermal_form.addRow("Cool-to threshold:", self._thermal_cool_to)
        thermal_form.addRow("Ramp rate:", self._thermal_ramp_rate)
        thermal_form.addRow("Run context:", self._thermal_context)
        thermal_form.addRow("Operator:", self._thermal_operator)
        thermal_form.addRow("Artifact folder:", self._thermal_artifact_dir)
        self._run_thermal_btn = QtWidgets.QPushButton("Record Thermal Plan")
        self._run_thermal_btn.clicked.connect(self._run_thermal_routine)
        thermal_form.addRow("", self._run_thermal_btn)
        fl.addRow("", self._thermal_box)

        button_row = QtWidgets.QHBoxLayout()
        self._run_auto_btn = QtWidgets.QPushButton("Run Automated")
        self._run_auto_btn.clicked.connect(self._run_automated)
        self._run_manual_btn = QtWidgets.QPushButton("Record Manual")
        self._run_manual_btn.clicked.connect(self._run_manual)
        button_row.addWidget(self._run_auto_btn)
        button_row.addWidget(self._run_manual_btn)
        button_row.addStretch()
        fl.addRow("", button_row)

        self._status = QtWidgets.QLabel("Ready")
        self._status.setStyleSheet("color: #475569;")
        fl.addRow("Status:", self._status)
        root.addWidget(card)

        self._result = QtWidgets.QTextEdit()
        self._result.setReadOnly(True)
        self._result.setMinimumHeight(180)
        self._result.setPlaceholderText("Run a calibration to see a result...")
        root.addWidget(self._result)

        history_label = QtWidgets.QLabel("Recent results")
        history_label.setObjectName("sectionHdr")
        root.addWidget(history_label)

        self._history_view = QtWidgets.QTextEdit()
        self._history_view.setReadOnly(True)
        self._history_view.setMaximumHeight(160)
        self._history_view.setPlaceholderText("No recent results")
        root.addWidget(self._history_view)

        root.addStretch()
        self._on_procedure_changed()

    @staticmethod
    def _spin(
        low: float,
        high: float,
        value: float,
        suffix: str,
    ) -> QtWidgets.QDoubleSpinBox:
        spin = QtWidgets.QDoubleSpinBox()
        spin.setRange(low, high)
        spin.setDecimals(6)
        spin.setSingleStep(1.0)
        spin.setValue(value)
        spin.setSuffix(suffix)
        return spin

    def _on_procedure_changed(self) -> None:
        key = self._procedure.currentData()
        info = self.PROCEDURES.get(key if key else "gaussmeter_baseline")
        if info is None:
            return
        self._expected.setValue(float(info["expected_default"]))
        self._tolerance.setValue(float(info["tolerance_default"]))
        is_irm = key == "irm_voltage"
        is_thermal = key == "thermal_routine"
        self._irm_box.setVisible(is_irm)
        self._thermal_box.setVisible(is_thermal)
        for widget in (
            self._expected,
            self._samples,
            self._tolerance,
            self._manual_value,
            self._baseline_context,
            self._baseline_operator,
            self._baseline_artifact_dir,
            self._run_auto_btn,
            self._run_manual_btn,
        ):
            widget.setEnabled(not is_irm and not is_thermal)
        for widget in (
            self._irm_field_1,
            self._irm_voltage_1,
            self._irm_field_2,
            self._irm_voltage_2,
            self._irm_max_voltage,
            self._irm_force_zero,
            self._irm_context,
            self._irm_operator,
            self._irm_artifact_dir,
            self._run_irm_btn,
        ):
            widget.setEnabled(is_irm)
        for widget in (
            self._thermal_name,
            self._thermal_targets,
            self._thermal_prefix,
            self._thermal_hold_seconds,
            self._thermal_max_temp,
            self._thermal_ambient,
            self._thermal_cool_to,
            self._thermal_ramp_rate,
            self._thermal_context,
            self._thermal_operator,
            self._thermal_artifact_dir,
            self._run_thermal_btn,
        ):
            widget.setEnabled(is_thermal)
        if not is_irm and not is_thermal:
            self._on_mode_changed()

    def _on_mode_changed(self) -> None:
        if self._procedure.currentData() in {"irm_voltage", "thermal_routine"}:
            return
        self._run_auto_btn.setEnabled(self._mode.currentData() == "automated")
        self._run_manual_btn.setEnabled(self._mode.currentData() == "manual")
        self._samples.setEnabled(self._mode.currentData() == "automated")
        self._manual_value.setEnabled(self._mode.currentData() == "manual")

    def set_procedure(self, procedure_id: str) -> None:
        for idx in range(self._procedure.count()):
            if self._procedure.itemData(idx) == procedure_id:
                self._procedure.setCurrentIndex(idx)
                return

    def set_mode(self, mode: str) -> None:
        for idx in range(self._mode.count()):
            if self._mode.itemData(idx) == mode:
                self._mode.setCurrentIndex(idx)
                return

    def _run_automated(self) -> None:
        if self._backend_provider is None:
            self._status.setText("No backend available for automated calibration.")
            return

        backend = self._backend_provider()
        procedure_id = self._procedure.currentData() or "gaussmeter_baseline"
        self._status.setText("Running automated calibration...")
        try:
            result = run_automated_squid_baseline_calibration(
                backend,
                procedure_id=f"calibration/{procedure_id}",
                expected=self._expected.value(),
                samples=self._samples.value(),
                tolerance=self._tolerance.value(),
                procedure_name="SQUID / Gaussmeter Baseline",
            )
            self._apply_result(result)
        except Exception as exc:
            self._status.setText(f"Calibration failed: {exc}")

    def _run_manual(self) -> None:
        procedure_id = self._procedure.currentData() or "gaussmeter_baseline"
        result = run_manual_calibration(
            self._manual_value.value(),
            procedure_id=f"calibration/{procedure_id}",
            expected=self._expected.value(),
            tolerance=self._tolerance.value(),
            procedure_name="SQUID / Gaussmeter Baseline",
            notes="Manual entry from Calibration Center.",
        )
        self._apply_result(result)

    def _run_irm_voltage(self) -> None:
        try:
            fit = fit_irm_voltage_calibration(
                [
                    IrmVoltageCalibrationPoint(
                        self._irm_field_1.value(),
                        self._irm_voltage_1.value(),
                    ),
                    IrmVoltageCalibrationPoint(
                        self._irm_field_2.value(),
                        self._irm_voltage_2.value(),
                    ),
                ],
                max_voltage_v=self._irm_max_voltage.value(),
                force_zero_intercept=self._irm_force_zero.isChecked(),
            )
            artifact_dir = Path(self._irm_artifact_dir.text()).expanduser()
            artifact_path = artifact_dir / f"irm_voltage_calibration_{QtCore.QDateTime.currentDateTimeUtc().toString('yyyyMMddTHHmmssZ')}.json"
            written = write_irm_voltage_calibration_artifact(
                artifact_path,
                fit,
                run_context=self._irm_context.text(),
                operator=self._irm_operator.text(),
                notes="Operator-recorded IRM voltage fit from Calibration Center.",
            )
        except Exception as exc:
            self._status.setText(f"IRM voltage calibration failed: {exc}")
            return

        text = (
            "IRM Voltage Calibration [RECORDED]\n"
            f"Slope: {fit.slope_v_per_mT:.9g} V/mT\n"
            f"Intercept: {fit.intercept_v:.9g} V\n"
            f"R^2: {fit.r_squared:.9g}\n"
            f"Max voltage: {fit.max_voltage_v:.9g} V\n"
            f"Artifact: {written}"
        )
        self._status.setText("IRM voltage calibration artifact recorded.")
        self._result.setPlainText(text)
        self._history.appendleft(
            f"{QtCore.QDateTime.currentDateTimeUtc().toString(QtCore.Qt.ISODate)} | "
            f"IRM Voltage Calibration | slope={fit.slope_v_per_mT:.6g} "
            f"intercept={fit.intercept_v:.6g}"
        )
        self._history_view.setPlainText("\n".join(self._history))

    def _thermal_target_values(self) -> list[float]:
        text = self._thermal_targets.text().replace(";", ",")
        values: list[float] = []
        for chunk in text.split(","):
            item = chunk.strip()
            if not item:
                continue
            values.append(float(item))
        if not values:
            raise ValueError("enter at least one thermal target")
        return values

    def _run_thermal_routine(self) -> None:
        try:
            limits = ThermalSafetyLimits(
                max_temperature_c=self._thermal_max_temp.value(),
                ambient_temperature_c=self._thermal_ambient.value(),
                cool_to_c=self._thermal_cool_to.value(),
                ramp_rate_c_per_min=self._thermal_ramp_rate.value(),
                default_hold_seconds=self._thermal_hold_seconds.value(),
            )
            plan = compile_thermal_routine(
                self._thermal_target_values(),
                name=self._thermal_name.text() or "Thermal routine",
                prefix=self._thermal_prefix.currentText(),
                limits=limits,
                hold_seconds=self._thermal_hold_seconds.value(),
            )
            artifact_dir = Path(self._thermal_artifact_dir.text()).expanduser()
            artifact_path = artifact_dir / f"thermal_routine_plan_{QtCore.QDateTime.currentDateTimeUtc().toString('yyyyMMddTHHmmssZ')}.json"
            written = write_thermal_routine_artifact(
                artifact_path,
                plan,
                run_context=self._thermal_context.text(),
                operator=self._thermal_operator.text(),
                notes=(
                    "Operator-recorded thermal planning artifact from Calibration Center. "
                    "Hardware furnace execution, abort, alarms, and temperature readback "
                    "remain hardware-only acceptance gates."
                ),
            )
        except Exception as exc:
            self._status.setText(f"Thermal routine planning failed: {exc}")
            return

        text = (
            "Thermal Routine Planning [PLANNED]\n"
            f"Routine: {plan.name}\n"
            f"Labels: {', '.join(plan.labels)}\n"
            f"Max temp: {plan.max_temperature_c:.9g} C\n"
            f"Estimated time: {plan.estimated_seconds()} s\n"
            f"Cooldown required: {'yes' if plan.requires_cooldown else 'no'}\n"
            f"Artifact: {written}\n"
            "Hardware acceptance still required for live furnace execution, safe abort, "
            "alarm handling, and temperature readback."
        )
        self._status.setText("Thermal routine planning artifact recorded.")
        self._result.setPlainText(text)
        self._history.appendleft(
            f"{QtCore.QDateTime.currentDateTimeUtc().toString(QtCore.Qt.ISODate)} | "
            f"Thermal Routine Planning | labels={','.join(plan.labels)} "
            f"estimate={plan.estimated_seconds()}s"
        )
        self._history_view.setPlainText("\n".join(self._history))

    def _apply_result(self, result: CalibrationResult) -> None:
        artifact_line = ""
        artifact_written = False
        try:
            artifact_dir = Path(self._baseline_artifact_dir.text()).expanduser()
            safe_name = str(result.procedure_id).split("/")[-1].replace("-", "_")
            stamp = QtCore.QDateTime.currentDateTimeUtc().toString("yyyyMMddTHHmmssZ")
            artifact_path = artifact_dir / f"{safe_name}_{stamp}.json"
            written = write_calibration_result_artifact(
                artifact_path,
                result,
                run_context=self._baseline_context.text(),
                operator=self._baseline_operator.text(),
                notes=result.notes,
            )
            artifact_line = f"\nArtifact: {written}"
            artifact_written = True
        except Exception as exc:
            artifact_line = f"\nArtifact write failed: {exc}"
        self._status.setText(
            f"{result.procedure_name} {result.status}: "
            f"{'accepted' if result.passed else 'rejected'}; "
            f"{'artifact recorded' if artifact_written else 'artifact write failed'}"
        )
        self._result.setPlainText(format_calibration_result(result) + artifact_line)
        self._history.appendleft(
            f"{result.timestamp_iso} | {result.procedure_name} | "
            f"{result.status} | mean={result.mean:.6g} "
            f"residual={result.max_abs_residual:.6g}"
        )
        self._history_view.setPlainText("\n".join(self._history))
