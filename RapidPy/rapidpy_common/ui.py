from __future__ import annotations

import inspect
import sys
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from .palette import GOLD, MAROON


MIN_WINDOW_WIDTH = 300
MIN_WINDOW_HEIGHT = 240


def clamp_window_geometry(available: QtCore.QRect, requested: tuple[int, int]) -> tuple[int, int]:
    """Return a geometry size that fits within the configured working area."""
    requested_width, requested_height = requested
    area_width = max(1, available.width())
    area_height = max(1, available.height())

    # Constrain wide windows to keep a compact working footprint on modern
    # mixed-resolution desktop layouts while still keeping enough room for
    # data-dense workflows.
    # Cap width growth on very large monitors so full-width startup windows never
    # occur on high-resolution rigs.
    max_w = min(int(area_width * 0.35), area_width, 960)
    max_h = min(int(area_height * 0.80), area_height)
    min_max_w = min(MIN_WINDOW_WIDTH, area_width)
    min_max_h = min(MIN_WINDOW_HEIGHT, area_height)
    max_w = max(max_w, min_max_w)
    max_h = max(max_h, min_max_h)

    min_w = min(int(area_width * 0.26), max_w)
    min_h = min(int(area_height * 0.24), max_h)
    min_w = max(MIN_WINDOW_WIDTH, min_w)
    min_h = max(MIN_WINDOW_HEIGHT, min_h)
    # Guard tiny or oddly shaped viewports where "compact" minimums can exceed
    # the usable area when the monitor is very narrow.
    min_w = min(min_w, area_width)
    min_h = min(min_h, area_height)

    width = max(min_w, min(max(1, int(requested_width)), max_w))
    height = max(min_h, min(max(1, int(requested_height)), max_h))
    width = min(width, area_width)
    height = min(height, area_height)
    return width, height


def _screen_area_for_widget(window: QtWidgets.QWidget) -> QtCore.QRect | None:
    """Return the widget's current available screen working area."""
    handle = window.windowHandle()
    if handle is not None and handle.screen() is not None:
        return handle.screen().availableGeometry()

    app = QtWidgets.QApplication.instance()
    if app is None:
        return None

    widget_rect = window.frameGeometry()
    candidate_screens = []
    for screen in app.screens():
        screen_geo = screen.geometry()
        intersection = screen_geo.intersected(widget_rect)
        overlap = intersection.width() * intersection.height()
        if overlap > 0:
            candidate_screens.append((overlap, screen))

    if candidate_screens:
        _, screen = max(candidate_screens, key=lambda entry: entry[0])
        return screen.availableGeometry()

    # If the widget has not been placed yet, fall back to the screen at (0, 0),
    # then finally choose nearest screen geometry before using primary/first screen.
    for screen in app.screens():
        if screen.geometry().contains(widget_rect.topLeft()):
            return screen.availableGeometry()

    # Widgets can start at synthetic off-screen coordinates before the window system
    # assigns a native handle. In that case, choose the nearest screen center to
    # avoid snapping to an unexpectedly unrelated primary display.
    if widget_rect.isValid():
        candidate_center = widget_rect.center()
        nearest: QtGui.QScreen | None = None
        nearest_distance = None
        for screen in app.screens():
            screen_center = screen.geometry().center()
            dx = screen_center.x() - candidate_center.x()
            dy = screen_center.y() - candidate_center.y()
            distance = (dx * dx) + (dy * dy)
            if nearest_distance is None or distance < nearest_distance:
                nearest_distance = distance
                nearest = screen
        if nearest is not None:
            return nearest.availableGeometry()

    if app.primaryScreen() is not None:
        return app.primaryScreen().availableGeometry()

    if not app.screens():
        return None
    return app.screens()[0].availableGeometry()


def _fit_window_to_screen(window: QtWidgets.QWidget) -> None:
    """Clamp and re-center a top-level window into the current work area."""
    available = _screen_area_for_widget(window)
    if available is None:
        return

    max_w, max_h = clamp_window_geometry(available, (window.width(), window.height()))

    if window.isMaximized():
        window.showNormal()

    min_size = window.minimumSize()
    if min_size.isValid() and not min_size.isNull():
        window.setMinimumSize(min(min_size.width(), max_w), min(min_size.height(), max_h))

    # Keep explicit runtime caps so window manager restore/maximize actions cannot
    # temporarily bypass our intended safe-boot envelope on this screen.
    window.setMaximumWidth(max_w)
    window.setMaximumHeight(max_h)

    frame = window.frameGeometry()
    if frame.width() > max_w or frame.height() > max_h:
        frame.setSize(QtCore.QSize(max_w, max_h))

    if frame.width() > available.width() or frame.height() > available.height():
        frame.moveCenter(available.center())
    else:
        new_x = max(available.left(), min(frame.left(), available.right() - frame.width() + 1))
        new_y = max(available.top(), min(frame.top(), available.bottom() - frame.height() + 1))
        frame.moveTopLeft(QtCore.QPoint(new_x, new_y))

    window.setGeometry(frame)
    window.resize(
        min(window.width(), frame.width()),
        min(window.height(), frame.height()),
    )


def _fit_window_with_widget_handler(window: QtWidgets.QWidget, screen: QtGui.QScreen | None = None) -> bool:
    """Allow window-specific fit handlers to opt into geometry enforcement."""
    method_names = (
        "_fit_to_screen",
        "_fit_window_to_current_screen",
        "fit_to_screen",
    )
    for method_name in method_names:
        method = getattr(window, method_name, None)
        if not callable(method):
            continue

        sig = None
        try:
            sig = inspect.signature(method)
        except (TypeError, ValueError):
            sig = None

        try:
            if sig is not None:
                params = [
                    p
                    for p in sig.parameters.values()
                    if p.kind in (
                        inspect.Parameter.POSITIONAL_ONLY,
                        inspect.Parameter.POSITIONAL_OR_KEYWORD,
                    )
                ]
                if len(params) == 0:
                    method()
                    return True
                if len(params) == 1:
                    method(screen)
                    return True
            method()
            return True
        except TypeError:
            # Fall back to best-effort invocation for optional, screen-aware handlers.
            if screen is not None:
                method(screen)
                return True
            raise
    return False


def _refit_top_level_windows(app: QtWidgets.QApplication) -> None:
    """Reapply bounds protection after a screen topology or DPI change."""

    for window in app.topLevelWidgets():
        try:
            if not window.isWindow():
                continue
            if not _fit_window_with_widget_handler(window):
                _fit_window_to_screen(window)
        except RuntimeError:
            # A window can be destroyed while Qt dispatches a topology change.
            continue


def _apply_window_bounds_guard(app: QtWidgets.QApplication) -> None:
    """Install a QApplication-level guard so main windows stay within visible screens."""
    if getattr(app, "_rapidpy_window_guard", None) is not None:
        return

    class _WindowBoundsGuard(QtCore.QObject):
        def __init__(self, parent_app: QtWidgets.QApplication) -> None:
            super().__init__(parent_app)
            self._app = parent_app
            self._watched_ids: set[int] = set()
            self._connected_ids: set[int] = set()
            self._connected_screen_metrics: set[tuple[int, int]] = set()
            parent_app.installEventFilter(self)
            parent_app.screenAdded.connect(lambda _screen: self._schedule_refit_all())
            parent_app.screenRemoved.connect(lambda _screen: self._schedule_refit_all())

        def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:  # type: ignore[override]
            if not isinstance(obj, QtWidgets.QWidget) or not obj.isWindow():
                return False
            if event.type() == QtCore.QEvent.Type.Show:
                QtCore.QTimer.singleShot(0, lambda: self._bind_window(obj))
            return False

        def _schedule_refit_all(self) -> None:
            QtCore.QTimer.singleShot(0, lambda: _refit_top_level_windows(self._app))

        def _bind_window(self, window: QtWidgets.QWidget, screen: QtGui.QScreen | None = None) -> None:
            try:
                if id(window) not in self._watched_ids:
                    self._watched_ids.add(id(window))
                if not _fit_window_with_widget_handler(window, screen=screen):
                    _fit_window_to_screen(window)
                self._watch_screen_changes(window)
            except RuntimeError:
                return

        def _watch_screen_changes(self, window: QtWidgets.QWidget) -> None:
            handle = window.windowHandle()
            if handle is None:
                QtCore.QTimer.singleShot(
                    75,
                    lambda _w=window: self._watch_screen_changes(_w),
                )
                return
            if id(window) in self._connected_ids:
                self._watch_screen_metrics(window, handle.screen())
                return
            self._connected_ids.add(id(window))
            handle.screenChanged.connect(
                lambda _s, _w=window: QtCore.QTimer.singleShot(
                    0, lambda: self._bind_window(_w, screen=_s)
                ),
            )
            self._watch_screen_metrics(window, handle.screen())

        def _watch_screen_metrics(
            self,
            window: QtWidgets.QWidget,
            screen: QtGui.QScreen | None,
        ) -> None:
            if screen is None:
                return
            key = (id(window), id(screen))
            if key in self._connected_screen_metrics:
                return
            self._connected_screen_metrics.add(key)
            for signal in (
                screen.availableGeometryChanged,
                screen.geometryChanged,
                screen.logicalDotsPerInchChanged,
            ):
                signal.connect(
                    lambda *_args, _w=window: QtCore.QTimer.singleShot(
                        0, lambda: self._bind_window(_w)
                    )
                )

    app._rapidpy_window_guard = _WindowBoundsGuard(app)  # type: ignore[attr-defined]
    QtCore.QTimer.singleShot(0, lambda: _refit_top_level_windows(app))


def apply_window_bounds_guard(app: QtWidgets.QApplication) -> None:
    """Public shim for installing the shared window/bounds guard."""
    _apply_window_bounds_guard(app)


def apply_liquid_glass_theme(app: QtWidgets.QApplication) -> None:
    """Apply the shared Apple-inspired liquid glass styling."""
    assets_dir = Path(__file__).resolve().parent / "assets"
    arrow_down = (assets_dir / "arrow_down.svg").as_posix()
    arrow_up = (assets_dir / "arrow_up.svg").as_posix()

    app.setStyle("Fusion")
    font = QtGui.QFont("SF Pro Text", 10)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Avenir Next", 10)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Segoe UI", 10)
    app.setFont(font)

    app.setStyleSheet(
        f"""
        QWidget {{
            background: #f3eee2;
            color: #2f2827;
        }}
        QFrame#card {{
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(122, 2, 25, 0.14);
            border-radius: 24px;
        }}
        QFrame#card QWidget {{
            background: transparent;
        }}
        QFrame#card QPushButton {{
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(122, 2, 25, 0.45);
            border-radius: 14px;
            padding: 9px 14px;
            color: #2f2827;
        }}
        QFrame#card QPushButton:hover {{
            background: rgba(255, 255, 255, 1.0);
        }}
        QFrame#card QPushButton:pressed {{
            background: rgba(244, 238, 231, 1.0);
        }}
        QFrame#card QPushButton#accent {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {MAROON}, stop:1 #5a0013);
            color: #ffffff;
            border: 1px solid rgba(122, 2, 25, 0.85);
            font-weight: 680;
        }}
        QFrame#card QPushButton#accent:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8a0220, stop:1 #650016);
        }}
        QFrame#card QPushButton#accent:pressed {{
            background: #5a0013;
        }}
        QFrame#livePanel {{
            background: rgba(255, 255, 255, 0.60);
            border: 1px solid rgba(122, 2, 25, 0.10);
            border-radius: 24px;
        }}
        QLabel#title {{
            font-size: 24px;
            font-weight: 760;
            color: {MAROON};
        }}
        QLabel#subtitle {{
            color: #61534d;
            margin-bottom: 4px;
        }}
        QLabel#valuePill {{
            background: rgba(255, 255, 255, 0.82);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 16px;
            padding: 8px 10px;
            font-weight: 650;
        }}
        QPlainTextEdit#statusPill {{
            background: rgba(255, 255, 255, 0.82);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 16px;
            padding: 8px 10px;
            font-weight: 650;
            color: #2f2827;
            selection-background-color: rgba(122, 2, 25, 0.18);
        }}
        QLabel#readingDisplay {{
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid rgba(122, 2, 25, 0.14);
            border-radius: 20px;
            padding: 16px 18px;
            color: {MAROON};
        }}
        QPlainTextEdit#console {{
            background: rgba(28, 20, 19, 0.88);
            color: #fff2c9;
            border-radius: 14px;
            border: 1px solid rgba(255, 205, 52, 0.32);
            padding: 8px;
            selection-background-color: {MAROON};
        }}
        QScrollArea#panelScroll {{
            background: transparent;
            border: none;
        }}
        QScrollArea#panelScroll > QWidget > QWidget {{
            background: transparent;
        }}
        QScrollBar:vertical {{
            background: transparent;
            width: 12px;
            margin: 4px 3px 4px 3px;
            border: none;
        }}
        QScrollBar:horizontal {{
            background: transparent;
            height: 12px;
            margin: 3px 4px 3px 4px;
            border: none;
        }}
        QScrollBar::handle:vertical,
        QScrollBar::handle:horizontal {{
            background: rgba(86, 72, 69, 0.36);
            border: 1px solid rgba(255, 255, 255, 0.46);
            border-radius: 6px;
        }}
        QScrollBar::handle:vertical {{
            min-height: 34px;
        }}
        QScrollBar::handle:horizontal {{
            min-width: 34px;
        }}
        QScrollBar::handle:vertical:hover,
        QScrollBar::handle:horizontal:hover {{
            background: rgba(86, 72, 69, 0.52);
        }}
        QScrollBar::handle:vertical:pressed,
        QScrollBar::handle:horizontal:pressed {{
            background: rgba(122, 2, 25, 0.54);
        }}
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {{
            border: none;
            background: transparent;
            width: 0px;
            height: 0px;
        }}
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical,
        QScrollBar::add-page:horizontal,
        QScrollBar::sub-page:horizontal {{
            background: transparent;
        }}
        QPushButton {{
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(122, 2, 25, 0.45);
            border-radius: 14px;
            padding: 9px 14px;
            color: #2f2827;
        }}
        QPushButton:hover {{
            background: rgba(255, 255, 255, 0.88);
        }}
        QPushButton:pressed {{
            background: rgba(255, 255, 255, 0.94);
        }}
        QPushButton#accent {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {MAROON}, stop:1 #5a0013);
            color: #fff9eb;
            border: 1px solid rgba(255, 255, 255, 0.26);
            font-weight: 680;
        }}
        QPushButton#accent:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8a0220, stop:1 #650016);
        }}
        QPushButton#accent:pressed {{
            background: #5a0013;
        }}
        QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {{
            border: 1px solid rgba(122, 2, 25, 0.35);
            background: #ffffff;
            border-radius: 12px;
            padding: 7px;
            selection-background-color: {MAROON};
            selection-color: #ffffff;
        }}
        QComboBox {{
            padding-right: 34px;
            min-width: 80px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            margin: 3px;
            border: none;
            border-radius: 10px;
            background: rgba(122, 2, 25, 0.12);
        }}
        QComboBox::drop-down:hover {{
            background: rgba(122, 2, 25, 0.2);
        }}
        QComboBox::drop-down:pressed {{
            background: rgba(122, 2, 25, 0.28);
        }}
        QComboBox::down-arrow {{
            image: url({arrow_down});
            width: 14px;
            height: 14px;
        }}
        QAbstractSpinBox {{
            padding-right: 30px;
            min-width: 64px;
        }}
        QAbstractSpinBox::up-button,
        QAbstractSpinBox::down-button {{
            width: 22px;
            border: none;
            border-radius: 5px;
            background: rgba(122, 2, 25, 0.12);
        }}
        QAbstractSpinBox::up-button {{
            subcontrol-origin: border;
            subcontrol-position: top right;
            margin: 5px 5px 1px 0px;
        }}
        QAbstractSpinBox::down-button {{
            subcontrol-origin: border;
            subcontrol-position: bottom right;
            margin: 1px 5px 5px 0px;
        }}
        QAbstractSpinBox::up-button:hover,
        QAbstractSpinBox::down-button:hover {{
            background: rgba(122, 2, 25, 0.2);
        }}
        QAbstractSpinBox::up-button:pressed,
        QAbstractSpinBox::down-button:pressed {{
            background: rgba(122, 2, 25, 0.28);
        }}
        QAbstractSpinBox::up-arrow {{
            image: url({arrow_up});
            width: 13px;
            height: 13px;
        }}
        QAbstractSpinBox::down-arrow {{
            image: url({arrow_down});
            width: 13px;
            height: 13px;
        }}
        QFrame#card QLineEdit,
        QFrame#card QComboBox,
        QFrame#card QDoubleSpinBox,
        QFrame#card QSpinBox {{
            background: #ffffff;
        }}
        QHeaderView::section {{
            background: rgba(255, 255, 255, 0.85);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 6px;
            padding: 6px;
            color: #4d3a39;
        }}
        QTableWidget {{
            background: rgba(255, 255, 255, 0.8);
            alternate-background-color: rgba(255, 255, 255, 0.65);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 12px;
            gridline-color: rgba(122, 2, 25, 0.12);
        }}
        QTableWidget::item:selected {{
            background: rgba(255, 205, 52, 0.38);
            color: #251f1e;
        }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {GOLD};
            border: 1px solid {MAROON};
        }}
        """
    )
    _apply_window_bounds_guard(app)


def set_app_icon(
    target: "QtWidgets.QApplication | QtWidgets.QWidget",
    icon_name: str,
    dev_assets_dir: Path,
) -> None:
    """Set window/application icon, resolving path for both dev and frozen (PyInstaller) runs."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        icon_path = Path(sys._MEIPASS) / "assets" / icon_name  # type: ignore[attr-defined]
    else:
        icon_path = dev_assets_dir / icon_name

    # Prefer platform-friendly .ico files when both are available.
    if not icon_path.exists():
        if icon_path.suffix.lower() == ".png":
            ico_path = icon_path.with_suffix(".ico")
            if ico_path.exists():
                icon_path = ico_path
        elif icon_path.suffix.lower() == ".ico":
            png_path = icon_path.with_suffix(".png")
            if png_path.exists():
                icon_path = png_path
    if not icon_path.exists():
        return

    target.setWindowIcon(QtGui.QIcon(str(icon_path)))

    if sys.platform != "win32":
        return

    # On Windows, multiple script-run windows can still appear under the Python
    # executable's taskbar group. A stable AppUserModelID helps taskbar identity
    # pick up each app's own branding when the shell honors the AUMID path.
    try:
        import ctypes

        app_id = f"RapidPy.{Path(icon_name).stem}"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(str(app_id))  # type: ignore[attr-defined]
    except Exception:
        pass


def apply_card_shadow(widget: QtWidgets.QWidget) -> None:
    shadow = QtWidgets.QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(34)
    shadow.setOffset(0, 10)
    shadow.setColor(QtGui.QColor(35, 25, 25, 48))
    widget.setGraphicsEffect(shadow)
