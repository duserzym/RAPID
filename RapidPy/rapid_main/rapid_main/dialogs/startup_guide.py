from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class StartupGuideDialog(QtWidgets.QDialog):
    """First-run guidance replacing the VB6 splash and tip entry points."""

    show_at_startup_changed = QtCore.Signal(bool)
    open_settings_requested = QtCore.Signal()
    open_queue_requested = QtCore.Signal()

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        *,
        show_at_startup: bool = True,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("RAPID Quick Start")
        self.setAccessibleName("RAPID quick start guide")
        self.setMinimumSize(340, 280)
        self.resize(500, 320)
        self.setModal(False)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui(show_at_startup)

    def set_show_at_startup(self, enabled: bool) -> None:
        self._show_at_startup.blockSignals(True)
        self._show_at_startup.setChecked(enabled)
        self._show_at_startup.blockSignals(False)

    def _build_ui(self, show_at_startup: bool) -> None:
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 18)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("Ready to begin a RAPID session?")
        title.setObjectName("dialogTitle")
        title.setAccessibleName("RAPID quick start title")
        layout.addWidget(title)

        guidance = QtWidgets.QLabel(
            "Confirm instrument settings, prepare the sample queue, then start the "
            "measurement sequence. Use No-Comm mode when training or reviewing data "
            "without connected hardware."
        )
        guidance.setWordWrap(True)
        guidance.setObjectName("guidanceText")
        layout.addWidget(guidance)

        steps = QtWidgets.QLabel(
            "1. Review ports and hardware status in Settings.\n"
            "2. Load specimens and organize the measurement queue.\n"
            "3. Select a sequence and begin the controlled run."
        )
        steps.setObjectName("dialogSubtitle")
        steps.setWordWrap(True)
        steps.setAccessibleName("Quick start steps")
        layout.addWidget(steps)

        self._show_at_startup = QtWidgets.QCheckBox("Show this guide when RAPID starts")
        self._show_at_startup.setAccessibleName("Show quick start guide at startup")
        self._show_at_startup.setChecked(show_at_startup)
        self._show_at_startup.toggled.connect(self.show_at_startup_changed)
        layout.addWidget(self._show_at_startup)

        button_row = QtWidgets.QGridLayout()
        self._settings_button = QtWidgets.QPushButton("Open Settings")
        self._queue_button = QtWidgets.QPushButton("Open Sample Queue")
        self._close_button = QtWidgets.QPushButton("Close")
        self._settings_button.setAccessibleName("Open RAPID settings")
        self._queue_button.setAccessibleName("Open sample queue")
        self._close_button.setAccessibleName("Close quick start guide")
        self._close_button.setDefault(True)
        self._settings_button.clicked.connect(lambda: self.open_settings_requested.emit())
        self._queue_button.clicked.connect(lambda: self.open_queue_requested.emit())
        self._close_button.clicked.connect(self.close)
        button_row.addWidget(self._settings_button, 0, 0)
        button_row.addWidget(self._queue_button, 0, 1)
        button_row.addWidget(self._close_button, 1, 1, alignment=QtCore.Qt.AlignRight)
        button_row.setColumnStretch(0, 1)
        layout.addLayout(button_row)
