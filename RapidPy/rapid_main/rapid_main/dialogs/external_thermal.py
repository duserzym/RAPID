"""Record the furnace log for one externally treated specimen step."""
from datetime import datetime, timezone

from PySide6 import QtCore, QtWidgets

from rapid_main.thermal import ExternalThermalTreatment, bind_external_thermal_treatment


class ExternalThermalDialog(QtWidgets.QDialog):
    def __init__(self, parent, plan, specimen, operator):
        super().__init__(parent)
        self.setWindowTitle("Measure Externally Treated Specimen")
        self.plan = plan
        self.specimen = specimen
        self.operator = operator
        self.treatment = None
        layout = QtWidgets.QVBoxLayout(self)
        heading = QtWidgets.QLabel(f"{specimen} · {plan.name}")
        heading.setTextFormat(QtCore.Qt.TextFormat.PlainText)
        layout.addWidget(heading)
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)
        self.step = QtWidgets.QComboBox()
        self.step.setPlaceholderText("Select the finished step")
        self.step.addItems(plan.labels)
        self.step.setCurrentIndex(-1)
        self.step.currentIndexChanged.connect(self._review_step)
        form.addRow("Reviewed step", self.step)
        self.expected = QtWidgets.QLabel()
        form.addRow("Plan limits", self.expected)
        self.log = QtWidgets.QLineEdit()
        form.addRow("Furnace-log reference", self.log)
        self.peak = self._temperature()
        form.addRow("Recorded peak", self.peak)
        self.hold = QtWidgets.QSpinBox()
        self.hold.setRange(-1, 2_147_483_647)
        self.hold.setSpecialValueText("Enter dwell time")
        self.hold.setSuffix(" s")
        self.hold.setValue(-1)
        form.addRow("Recorded dwell", self.hold)
        self.handling = self._temperature()
        form.addRow("Handling temperature", self.handling)
        self.completed = QtWidgets.QDateTimeEdit(QtCore.QDateTime.currentDateTime())
        self.completed.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.completed.setCalendarPopup(True)
        form.addRow("Furnace completion (local time)", self.completed)
        self.confirmed = QtWidgets.QCheckBox("I verified the log, specimen identity, and cooled specimen mount")
        layout.addWidget(self.confirmed)
        self.error = QtWidgets.QLabel()
        self.error.setWordWrap(True)
        layout.addWidget(self.error)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Cancel)
        self.measure = buttons.addButton("Start SQUID Measurement", QtWidgets.QDialogButtonBox.ButtonRole.AcceptRole)
        self.measure.setEnabled(False)
        self.confirmed.toggled.connect(self.measure.setEnabled)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @staticmethod
    def _temperature():
        field = QtWidgets.QDoubleSpinBox()
        field.setRange(-1, 10_000)
        field.setValue(-1)
        field.setSpecialValueText("Enter recorded temperature")
        field.setSuffix(" °C")
        return field

    def _review_step(self, index):
        if index < 0:
            return
        step = self.plan.steps[index]
        dwell = self.plan.limits.default_hold_seconds if step.hold_seconds is None else step.hold_seconds
        self.expected.setText(f"{step.temperature_c:g} °C target, {dwell} s dwell, "
                              f"cool to {self.plan.limits.cool_to_c:g} °C")
        self.confirmed.setChecked(False)

    def accept(self):
        if not self.confirmed.isChecked():
            return
        try:
            treatment = ExternalThermalTreatment(
                self.specimen, self.step.currentText(), self.operator, self.log.text().strip(),
                self.peak.value(), self.hold.value(), self.handling.value(),
                datetime.fromtimestamp(self.completed.dateTime().toSecsSinceEpoch(), timezone.utc).isoformat(),
            )
            bind_external_thermal_treatment(self.plan.to_artifact(), treatment,
                                            specimen=self.specimen, operator=self.operator, labels=[treatment.label])
        except ValueError as exc:
            self.error.setText(str(exc))
            return
        self.treatment = treatment
        super().accept()
