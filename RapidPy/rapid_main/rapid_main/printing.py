from __future__ import annotations

from collections.abc import Callable

from PySide6 import QtCore, QtGui, QtPrintSupport, QtWidgets


PrintDialogFactory = Callable[[QtPrintSupport.QPrinter, QtWidgets.QWidget], QtWidgets.QDialog]


def print_widget_snapshot(
    parent: QtWidgets.QWidget,
    widget: QtWidgets.QWidget,
    *,
    title: str,
    dialog_factory: PrintDialogFactory | None = None,
) -> bool:
    """Open the native print dialog and print a scaled snapshot of the widget.

    Returning a boolean keeps cancel/failure behavior explicit and makes the
    formerly disconnected print actions deterministic in tests.
    """

    printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.PrinterMode.HighResolution)
    printer.setDocName(title)
    factory = dialog_factory or (lambda p, owner: QtPrintSupport.QPrintDialog(p, owner))
    dialog = factory(printer, parent)
    if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
        return False

    snapshot = widget.grab()
    if snapshot.isNull():
        QtWidgets.QMessageBox.warning(parent, "Print", "The current view could not be captured.")
        return False

    page_rect = printer.pageLayout().paintRectPixels(printer.resolution())
    target = snapshot.size()
    target.scale(page_rect.size(), QtCore.Qt.AspectRatioMode.KeepAspectRatio)
    x = page_rect.x() + max(0, (page_rect.width() - target.width()) // 2)
    y = page_rect.y() + max(0, (page_rect.height() - target.height()) // 2)

    painter = QtGui.QPainter()
    if not painter.begin(printer):
        QtWidgets.QMessageBox.warning(parent, "Print", "The selected printer could not be opened.")
        return False
    try:
        painter.setRenderHint(QtGui.QPainter.RenderHint.SmoothPixmapTransform, True)
        painter.drawPixmap(QtCore.QRect(x, y, target.width(), target.height()), snapshot)
    finally:
        painter.end()
    return True
