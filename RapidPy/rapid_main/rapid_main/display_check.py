"""Offline layout check on the current virtual Windows work area."""
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6 import QtCore, QtWidgets

from rapid_main.app import _apply_desktop_theme
from rapid_main.dialogs.sample_select import SampleSelectDialog


def run_display_check() -> int:
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    _apply_desktop_theme(app)
    with TemporaryDirectory(prefix="rapidpy-display-") as directory:
        root = Path(directory)
        settings = QtCore.QSettings(str(root / "registry.ini"), QtCore.QSettings.Format.IniFormat)
        index = root / "display.sam"
        index.write_text("display-specimen\n", encoding="ascii")
        dialog = SampleSelectDialog(settings=settings)
        window = QtWidgets.QMainWindow()
        try:
            if not dialog._open_source(index):
                return 2
            dialog._direction.setCurrentText("Down with saved Up")
            dialog._register_btn.click()
            dialog._table.selectRow(0)
            window.show()
            window.setWindowState(QtCore.Qt.WindowState.WindowMaximized)
            dialog.show()
            app.processEvents()
            app.processEvents()

            screen = dialog.screen()
            available = screen.availableGeometry()
            controls = {"registered_indexes": dialog._indexes, "save": dialog._register_btn,
                        "remove": dialog._remove_btn, "browse": dialog._load_btn,
                        "direction": dialog._direction, "blocks": dialog._blocks,
                        "select": dialog._select_btn, "cancel": dialog._cancel_btn}
            clipped = [name for name, control in controls.items()
                       if not control.isVisible() or not dialog.rect().contains(
                           QtCore.QRect(control.mapTo(dialog, QtCore.QPoint()), control.size()))]
            table = dialog._table
            specimen_visible = table.viewport().rect().intersects(table.visualItemRect(table.item(0, 0)))
            direction = dialog._direction
            option = QtWidgets.QStyleOptionComboBox()
            direction.initStyleOption(option)
            label_area = direction.style().subControlRect(QtWidgets.QStyle.ComplexControl.CC_ComboBox, option,
                                                         QtWidgets.QStyle.SubControl.SC_ComboBoxEditField, direction)
            label_width = direction.fontMetrics().horizontalAdvance(direction.currentText())
            label_fits = label_area.width() >= label_width
            if not label_fits and "direction" not in clipped:
                clipped.append("direction")
            report = {"schema": "rapidpy.display_check.v1", "device_pixel_ratio": screen.devicePixelRatio(),
                      "work_area": [available.width(), available.height()],
                      "dialog_size": [dialog.width(), dialog.height()], "clipped_controls": clipped,
                      "direction_width": direction.width(), "direction_text_area": label_area.width(),
                      "direction_label_width": label_width,
                      "dialog_in_work_area": available.contains(dialog.frameGeometry()),
                      "registry_controls_visible": not clipped and specimen_visible and dialog._select_btn.isEnabled(),
                      "maximized_preserved": window.isMaximized()}
            report["ok"] = all(report[key] for key in (
                "dialog_in_work_area", "registry_controls_visible", "maximized_preserved",
            ))
            print(json.dumps(report, sort_keys=True))
            return 0 if report["ok"] else 2
        finally:
            dialog.close()
            window.close()
