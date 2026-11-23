"""Offline layout check on the current virtual Windows work area."""
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6 import QtCore, QtWidgets

from rapid_main.app import _apply_desktop_theme
from rapid_main.dialogs.sample_select import SampleSelectDialog
from rapidpy_common.ui import apply_liquid_glass_theme, apply_window_bounds_guard


def _combo_label_widths(combo: QtWidgets.QComboBox) -> tuple[int, int]:
    option = QtWidgets.QStyleOptionComboBox()
    combo.initStyleOption(option)
    field = combo.style().subControlRect(QtWidgets.QStyle.ComplexControl.CC_ComboBox, option,
                                        QtWidgets.QStyle.SubControl.SC_ComboBoxEditField, combo)
    return field.width(), combo.fontMetrics().horizontalAdvance(combo.currentText())


def _check_data_review_display(app: QtWidgets.QApplication) -> dict:
    from data_viewer.app import ZijderveldWindow

    apply_liquid_glass_theme(app)
    apply_window_bounds_guard(app)
    viewer = ZijderveldWindow()
    try:
        viewer.show()
        app.processEvents()
        app.processEvents()
        scroll = viewer._body
        viewport = scroll.viewport()
        toolbar = viewer._toolbar
        clipped = []
        for control_type in (QtWidgets.QPushButton, QtWidgets.QCheckBox, QtWidgets.QComboBox, QtWidgets.QLabel):
            for control in toolbar.findChildren(control_type):
                name = control.accessibleName()
                if isinstance(control, QtWidgets.QComboBox):
                    field, label = _combo_label_widths(control)
                    full_text = field >= label
                    name = name or control.currentText()
                else:
                    full_text = control.width() >= control.sizeHint().width()
                    name = name or control.text()
                if (not control.isVisible() or not full_text
                        or control.visibleRegion().boundingRect() != control.rect()):
                    clipped.append(name)

        plot = viewer._directional_stack
        scroll.ensureWidgetVisible(plot, 0, 0)
        plot_rect = QtCore.QRect(plot.mapTo(viewport, QtCore.QPoint()), plot.size())
        visible_plot = plot_rect.intersected(viewport.rect())

        def first_row_visible(widget, item) -> bool:
            row = widget.visualItemRect(item)
            center = widget.viewport().mapTo(scroll.widget(), row.center())
            scroll.ensureVisible(center.x(), center.y(), row.width() // 2 + 3, row.height() // 2 + 3)
            point = widget.viewport().mapTo(viewport, row.topLeft())
            return (row.isValid() and widget.viewport().rect().contains(row)
                    and viewport.rect().contains(QtCore.QRect(point, row.size())))

        table = viewer._table
        specimens = viewer._specimen_list
        report = {
            "window_size": [viewer.width(), viewer.height()],
            "window_in_work_area": viewer.screen().availableGeometry().contains(viewer.frameGeometry()),
            "clipped_controls": clipped,
            "plot_reachable": visible_plot.width() >= 240 and visible_plot.height() >= 160,
            "measurement_row_reachable": first_row_visible(table, table.item(0, 0)),
            "specimen_row_reachable": first_row_visible(specimens, specimens.item(0)),
        }
        report["ok"] = not clipped and all(report[key] for key in (
            "window_in_work_area", "plot_reachable", "measurement_row_reachable", "specimen_row_reachable",
        ))
        return report
    finally:
        viewer.close()


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
            label_area, label_width = _combo_label_widths(direction)
            label_fits = label_area >= label_width
            if not label_fits and "direction" not in clipped:
                clipped.append("direction")
            report = {"schema": "rapidpy.display_check.v1", "device_pixel_ratio": screen.devicePixelRatio(),
                      "work_area": [available.width(), available.height()],
                      "dialog_size": [dialog.width(), dialog.height()], "clipped_controls": clipped,
                      "direction_width": direction.width(), "direction_text_area": label_area,
                      "direction_label_width": label_width,
                      "dialog_in_work_area": available.contains(dialog.frameGeometry()),
                      "registry_controls_visible": not clipped and specimen_visible and dialog._select_btn.isEnabled(),
                      "maximized_preserved": window.isMaximized()}
            report["data_review"] = _check_data_review_display(app)
            report["ok"] = report["data_review"]["ok"] and all(report[key] for key in (
                "dialog_in_work_area", "registry_controls_visible", "maximized_preserved",
            ))
            print(json.dumps(report, sort_keys=True))
            return 0 if report["ok"] else 2
        finally:
            dialog.close()
            window.close()
