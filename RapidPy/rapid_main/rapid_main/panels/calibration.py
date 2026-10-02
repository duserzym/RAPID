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
from rapid_main.calibration_registry import (
    CalibrationRegistry,
    CalibrationRegistryError,
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

    thermal_plan_recorded = QtCore.Signal(object)

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
        registry: CalibrationRegistry | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend_provider = backend_provider
        self._registry = registry or CalibrationRegistry.default()
        self._last_calibration_artifact: Path | None = None
        self._history: deque[str] = deque(maxlen=8)
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("calibrationScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        body = QtWidgets.QWidget()
        root = QtWidgets.QVBoxLayout(body)
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
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)

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
        irm_form.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)
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

        self._thermal_box = QtWidgets.QGroupBox("Thermal planning — manual/external treatment only")
        thermal_form = QtWidgets.QFormLayout(self._thermal_box)
        thermal_form.setContentsMargins(12, 10, 12, 10)
        thermal_form.setSpacing(8)
        thermal_form.setLabelAlignment(QtCore.Qt.AlignRight)
        thermal_form.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)
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
        boundary = QtWidgets.QLabel(
            "RapidPy does not control a specimen furnace. This records and loads labels for "
            "operator-managed external treatment; live hardware mode remains blocked."
        )
        boundary.setWordWrap(True)
        boundary.setObjectName("statusWarning")
        boundary.setAccessibleName("Thermal automation boundary")
        thermal_form.addRow("Automation:", boundary)
        self._run_thermal_btn = QtWidgets.QPushButton("Record + Load Manual-External Plan")
        self._run_thermal_btn.setAccessibleName("Record and load manual-external thermal plan")
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
        fl.addRow("Status:", self._status)
        root.addWidget(card)

        self._result = QtWidgets.QTextEdit()
        self._result.setReadOnly(True)
        self._result.setMinimumHeight(180)
        self._result.setPlaceholderText("Run a calibration to see a result...")
        root.addWidget(self._result)

        self._build_registry_ui(root)

        history_label = QtWidgets.QLabel("Recent results")
        history_label.setObjectName("sectionHdr")
        root.addWidget(history_label)

        self._history_view = QtWidgets.QTextEdit()
        self._history_view.setReadOnly(True)
        self._history_view.setMaximumHeight(160)
        self._history_view.setPlaceholderText("No recent results")
        root.addWidget(self._history_view)

        root.addStretch()
        scroll.setWidget(body)
        outer.addWidget(scroll)
        self._on_procedure_changed()
        self._refresh_registry()

    def _build_registry_ui(self, root: QtWidgets.QVBoxLayout) -> None:
        lifecycle = QtWidgets.QGroupBox("Approval, validity, and rollback")
        layout = QtWidgets.QVBoxLayout(lifecycle)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        fields = QtWidgets.QFormLayout()
        fields.setLabelAlignment(QtCore.Qt.AlignRight)
        fields.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)
        self._registry_operator = QtWidgets.QLineEdit()
        self._registry_operator.setPlaceholderText("Required for every lifecycle event")
        self._registry_valid_days = QtWidgets.QSpinBox()
        self._registry_valid_days.setRange(0, 3650)
        self._registry_valid_days.setValue(365)
        self._registry_valid_days.setSpecialValueText("No expiry")
        self._registry_reason = QtWidgets.QLineEdit()
        self._registry_reason.setPlaceholderText("Approval note or required change reason")
        fields.addRow("Operator / approver:", self._registry_operator)
        fields.addRow("Validity:", self._registry_valid_days)
        fields.addRow("Reason / notes:", self._registry_reason)
        layout.addLayout(fields)

        actions = QtWidgets.QVBoxLayout()
        self._approve_artifact_btn = QtWidgets.QPushButton("Approve Last Artifact")
        self._approve_artifact_btn.setEnabled(False)
        self._approve_artifact_btn.clicked.connect(self._approve_last_artifact)
        self._activate_record_btn = QtWidgets.QPushButton("Activate Selected / Roll Back")
        self._activate_record_btn.clicked.connect(self._activate_selected_record)
        self._invalidate_record_btn = QtWidgets.QPushButton("Invalidate Selected")
        self._invalidate_record_btn.setProperty("danger", True)
        self._invalidate_record_btn.clicked.connect(self._invalidate_selected_record)
        actions.addWidget(self._approve_artifact_btn)
        actions.addWidget(self._activate_record_btn)
        actions.addWidget(self._invalidate_record_btn)
        layout.addLayout(actions)

        self._registry_table = QtWidgets.QTableWidget(0, 7)
        self._registry_table.setHorizontalHeaderLabels(
            ["Current", "Procedure", "Version", "Status", "Approved", "Expires", "Record ID"]
        )
        self._registry_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._registry_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._registry_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._registry_table.verticalHeader().setVisible(False)
        self._registry_table.setMinimumWidth(0)
        self._registry_table.setSizePolicy(
            QtWidgets.QSizePolicy.Ignored,
            QtWidgets.QSizePolicy.Preferred,
        )
        header = self._registry_table.horizontalHeader()
        header.setStretchLastSection(False)
        for column in range(6):
            header.setSectionResizeMode(column, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QtWidgets.QHeaderView.Stretch)
        self._registry_table.setMinimumHeight(170)
        self._registry_table.itemSelectionChanged.connect(self._sync_registry_actions)
        layout.addWidget(self._registry_table)

        self._registry_status = QtWidgets.QLabel("No approved calibration records.")
        self._registry_status.setWordWrap(True)
        layout.addWidget(self._registry_status)
        root.addWidget(lifecycle)

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

    def _set_last_calibration_artifact(
        self,
        path: Path | None,
        *,
        suggested_operator: str = "",
    ) -> None:
        self._last_calibration_artifact = path
        self._approve_artifact_btn.setEnabled(path is not None)
        if suggested_operator.strip() and not self._registry_operator.text().strip():
            self._registry_operator.setText(suggested_operator.strip())

    def _selected_record_id(self) -> str:
        row = self._registry_table.currentRow()
        if row < 0:
            return ""
        item = self._registry_table.item(row, 0)
        return str(item.data(QtCore.Qt.UserRole)) if item is not None else ""

    def _sync_registry_actions(self) -> None:
        selected = bool(self._selected_record_id())
        self._activate_record_btn.setEnabled(selected)
        self._invalidate_record_btn.setEnabled(selected)

    def _refresh_registry(self) -> None:
        try:
            states = self._registry.states()
        except CalibrationRegistryError as exc:
            self._registry_table.setRowCount(0)
            self._registry_status.setText(f"Registry error: {exc}")
            self._sync_registry_actions()
            return
        self._registry_table.setRowCount(len(states))
        for row, state in enumerate(states):
            record = state.record
            values = (
                "ACTIVE" if state.active else "",
                record.procedure_name,
                str(record.version),
                state.status.upper(),
                f"{record.approved_at_iso} · {record.approved_by}",
                record.expires_at_iso or "No expiry",
                record.record_id,
            )
            for column, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                item.setData(QtCore.Qt.UserRole, record.record_id)
                self._registry_table.setItem(row, column, item)
        active_count = sum(1 for state in states if state.active)
        self._registry_status.setText(
            f"{len(states)} approved record(s); {active_count} active and eligible "
            "for measurement provenance. Records and events are append-only."
            if states
            else "No approved calibration records. Measurement bundles will report an empty registry."
        )
        self._sync_registry_actions()

    def _approve_last_artifact(self) -> None:
        if self._last_calibration_artifact is None:
            self._registry_status.setText("Run and pass a calibration before approval.")
            return
        valid_days = self._registry_valid_days.value() or None
        try:
            record = self._registry.approve(
                self._last_calibration_artifact,
                approved_by=self._registry_operator.text(),
                valid_days=valid_days,
                notes=self._registry_reason.text(),
            )
        except CalibrationRegistryError as exc:
            self._registry_status.setText(f"Approval failed: {exc}")
            return
        self._set_last_calibration_artifact(None)
        self._refresh_registry()
        self._registry_status.setText(
            f"Approved and activated {record.record_id}. The artifact snapshot and "
            "lifecycle events are immutable."
        )

    def _activate_selected_record(self) -> None:
        record_id = self._selected_record_id()
        if not record_id:
            self._registry_status.setText("Select a calibration record to activate.")
            return
        try:
            record = self._registry.activate(
                record_id,
                operator=self._registry_operator.text(),
                reason=self._registry_reason.text(),
            )
        except CalibrationRegistryError as exc:
            self._registry_status.setText(f"Activation failed: {exc}")
            return
        self._refresh_registry()
        self._registry_status.setText(
            f"Activated {record.record_id}. This rollback added an event; no history was overwritten."
        )

    def _invalidate_selected_record(self) -> None:
        record_id = self._selected_record_id()
        if not record_id:
            self._registry_status.setText("Select a calibration record to invalidate.")
            return
        try:
            record = self._registry.invalidate(
                record_id,
                operator=self._registry_operator.text(),
                reason=self._registry_reason.text(),
            )
        except CalibrationRegistryError as exc:
            self._registry_status.setText(f"Invalidation failed: {exc}")
            return
        self._refresh_registry()
        self._registry_status.setText(
            f"Invalidated {record.record_id}. It will not appear in new measurement provenance."
        )

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

        self._set_last_calibration_artifact(
            written,
            suggested_operator=self._irm_operator.text(),
        )

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
        self._set_last_calibration_artifact(None)
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
                    "No specimen furnace is controlled. Live automation remains blocked until "
                    "the controller protocol, interlocks, and hardware acceptance exist."
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
            "MANUAL/EXTERNAL ONLY: RapidPy did not control a furnace. Live automation remains "
            "blocked pending a known controller protocol, interlocks, and hardware acceptance."
        )
        self._status.setText(
            "Thermal planning artifact recorded and manual-external plan loaded into Sequence."
        )
        self._result.setPlainText(text)
        self._history.appendleft(
            f"{QtCore.QDateTime.currentDateTimeUtc().toString(QtCore.Qt.ISODate)} | "
            f"Thermal Routine Planning | labels={','.join(plan.labels)} "
            f"estimate={plan.estimated_seconds()}s"
        )
        self._history_view.setPlainText("\n".join(self._history))
        self.thermal_plan_recorded.emit(plan)

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
            self._set_last_calibration_artifact(
                written if result.passed else None,
                suggested_operator=self._baseline_operator.text(),
            )
        except Exception as exc:
            artifact_line = f"\nArtifact write failed: {exc}"
            self._set_last_calibration_artifact(None)
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
