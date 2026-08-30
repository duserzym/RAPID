from __future__ import annotations

import unittest
from dataclasses import dataclass
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets

from rapid_main.app import (
    MainWindow,
    clamp_sidebar_target,
    clamp_window_size_for_screen,
    _AdaptivePanelStack,
    _MAIN_FIXED_MIN_WIDTH_THRESHOLD,
    _MAX_SIDEBAR_WIDTH,
    _MIN_SIDEBAR_WIDTH,
    _MAIN_MAX_WIDTH_RATIO,
    _MAIN_MIN_WIDTH,
    _SIDEBAR_RESTORE_RATIO,
    _clamp_main_window_size,
)
from rapidpy_common.ui import (
    MIN_WINDOW_WIDTH,
    MIN_WINDOW_HEIGHT,
    _fit_window_to_screen,
    _refit_top_level_windows,
    _screen_area_for_widget,
    clamp_window_geometry,
)


@dataclass(frozen=True)
class _FakeScreen:
    geo: QtCore.QRect
    avail: QtCore.QRect

    def geometry(self) -> QtCore.QRect:
        return self.geo

    def availableGeometry(self) -> QtCore.QRect:
        return self.avail


@dataclass(frozen=True)
class _FakeApplication:
    screens_list: list[_FakeScreen]
    primary: _FakeScreen

    def screens(self) -> list[_FakeScreen]:
        return self.screens_list

    def primaryScreen(self) -> _FakeScreen:
        return self.primary


class _TopologyApplication:
    def __init__(self, windows: list[QtWidgets.QWidget]) -> None:
        self._windows = windows

    def topLevelWidgets(self) -> list[QtWidgets.QWidget]:
        return self._windows


class _HeadlessWindow:
    def __init__(self, frame: QtCore.QRect) -> None:
        self._frame = frame

    def windowHandle(self) -> None:
        return None

    def frameGeometry(self) -> QtCore.QRect:
        return self._frame


class _WindowSettingsBridge:
    def __init__(self, values: dict[str, object] | None = None) -> None:
        self._values = dict(values or {})

    def value(self, key: str, default: object | None = None, type: type | None = None) -> object:
        if key not in self._values:
            return default
        value = self._values[key]
        if type is None or value is None:
            return value
        try:
            return type(value)
        except (TypeError, ValueError):
            return value

    def setValue(self, key: str, value: object) -> None:
        self._values[key] = value

    def remove(self, _key: str) -> None:
        self._values.pop(_key, None)


class TestWindowLayoutHelpers(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if QtWidgets.QApplication.instance() is None:
            cls._qt_app = QtWidgets.QApplication([])
        else:
            cls._qt_app = None

    @classmethod
    def tearDownClass(cls) -> None:
        if cls._qt_app is not None:
            cls._qt_app.quit()
            cls._qt_app = None

    def test_screen_area_for_window_prefers_largest_geometry_overlap(self) -> None:
        left = _FakeScreen(
            geo=QtCore.QRect(-1920, 0, 1920, 1080),
            avail=QtCore.QRect(-1920, 0, 1920, 1000),
        )
        right = _FakeScreen(
            geo=QtCore.QRect(0, 0, 1920, 1080),
            avail=QtCore.QRect(0, 0, 1920, 1000),
        )
        app = _FakeApplication(
            screens_list=[left, right],
            primary=right,
        )
        window = _HeadlessWindow(QtCore.QRect(-1860, 100, 200, 120))

        with patch("PySide6.QtWidgets.QApplication.instance", return_value=app):
            available = _screen_area_for_widget(window)

        self.assertEqual(available, left.availableGeometry())

    def test_screen_area_for_window_falls_back_to_primary_when_unpositioned(self) -> None:
        monitor = _FakeScreen(
            geo=QtCore.QRect(120, 40, 1600, 900),
            avail=QtCore.QRect(120, 40, 1600, 860),
        )
        app = _FakeApplication(screens_list=[monitor], primary=monitor)
        window = _HeadlessWindow(QtCore.QRect(4000, 4000, 300, 180))

        with patch("PySide6.QtWidgets.QApplication.instance", return_value=app):
            available = _screen_area_for_widget(window)

        self.assertEqual(available, monitor.availableGeometry())

    def test_screen_area_for_window_falls_back_to_nearest_when_unpositioned(self) -> None:
        left = _FakeScreen(
            geo=QtCore.QRect(-1600, 0, 1200, 900),
            avail=QtCore.QRect(-1600, 0, 1200, 860),
        )
        right = _FakeScreen(
            geo=QtCore.QRect(2400, 0, 1200, 900),
            avail=QtCore.QRect(2400, 0, 1200, 860),
        )
        app = _FakeApplication(
            screens_list=[left, right],
            primary=right,
        )
        window = _HeadlessWindow(QtCore.QRect(500, 100, 120, 80))

        with patch("PySide6.QtWidgets.QApplication.instance", return_value=app):
            available = _screen_area_for_widget(window)

        self.assertEqual(available, left.availableGeometry())

    def test_screen_area_for_window_uses_working_area_insets(self) -> None:
        monitor = _FakeScreen(
            geo=QtCore.QRect(0, 0, 1920, 1080),
            avail=QtCore.QRect(0, 40, 1920, 1040),
        )
        app = _FakeApplication(screens_list=[monitor], primary=monitor)
        window = _HeadlessWindow(QtCore.QRect(150, 150, 300, 180))

        with patch("PySide6.QtWidgets.QApplication.instance", return_value=app):
            available = _screen_area_for_widget(window)

        self.assertEqual(available, monitor.availableGeometry())
        self.assertEqual(available.height(), 1040)

    def test_app_window_minimum_constraints_for_small_displays(self) -> None:
        self.assertEqual(
            clamp_window_size_for_screen(
                QtCore.QRect(0, 0, 1024, 768),
                (160, 120),
            ),
            (MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1024, 768),
                (160, 120),
            ),
            (MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
        )

    def test_app_window_minimum_constraints_for_thin_viewports(self) -> None:
        self.assertEqual(
            clamp_window_size_for_screen(
                QtCore.QRect(10, 20, 1024, 680),
                (240, 180),
            ),
            (MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(10, 20, 1024, 680),
                (240, 180),
            ),
            (MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
        )

    def test_app_window_geometry_clamps_to_screen(self) -> None:
        self.assertEqual(
            clamp_window_size_for_screen(
                QtCore.QRect(0, 0, 1920, 1080),
                (4000, 3000),
            ),
            (672, 864),
        )

        self.assertEqual(
            clamp_window_size_for_screen(
                QtCore.QRect(0, 0, 1366, 768),
                (1200, 700),
            ),
            (478, 614),
        )

    def test_common_window_geometry_clamps_to_screen(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1920, 1080),
                (4000, 3000),
            ),
            (672, 864),
        )

        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1366, 768),
                (1200, 700),
            ),
            (478, 614),
        )

    def test_common_window_geometry_small_and_wide_profiles(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 2560, 1440),
                (5000, 3600),
            ),
            (896, 1152),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1280, 720),
                (600, 450),
            ),
            (448, 450),
        )

    def test_common_window_geometry_tiny_viewport_caps(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 300, 180),
                (2000, 1600),
            ),
            (300, 180),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 280, 160),
                (2000, 1600),
            ),
            (280, 160),
        )

    def test_common_window_geometry_dpi_scaled_profiles(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 2400, 1350),
                (5000, 4500),
            ),
            (840, 1080),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1500, 1000),
                (2000, 1600),
            ),
            (525, 800),
        )

    def test_common_window_geometry_ratio_and_bounds_profile(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 2000, 1000),
                (2000, 1000),
            ),
            (700, 800),
        )

    def test_common_window_geometry_tiny_viewport_profiles(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 1024, 576),
                (3000, 2200),
            ),
            (358, 460),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 768, 432),
                (1600, 1200),
            ),
            (300, 345),
        )

    def test_window_clamp_scales_with_high_resolution_profiles(self) -> None:
        self.assertEqual(
            clamp_window_size_for_screen(
                QtCore.QRect(0, 0, 3840, 2160),
                (5000, 4500),
            ),
            (960, 1728),
        )
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 3840, 2160),
                (5000, 4500),
            ),
            (960, 1728),
        )

    def test_common_window_geometry_caps_extreme_desktop_width(self) -> None:
        self.assertEqual(
            clamp_window_geometry(
                QtCore.QRect(0, 0, 7680, 4320),
                (8000, 8000),
            ),
            (960, 3456),
        )

    def test_main_window_width_caps_keep_compact_ratio(self) -> None:
        self.assertEqual(
            _clamp_main_window_size(
                QtCore.QRect(0, 0, 1920, 1080),
                (1280, 900),
            ),
            (
                min(
                    clamp_window_size_for_screen(
                        QtCore.QRect(0, 0, 1920, 1080),
                        (1280, 900),
                    )[0],
                    max(_MAIN_MIN_WIDTH, int(1920 * _MAIN_MAX_WIDTH_RATIO)),
                ),
                864,
            ),
        )

    def test_main_window_width_caps_large_displays_without_full_width_startup(self) -> None:
        high_resolution = QtCore.QRect(0, 0, 2560, 1440)
        compacted = _clamp_main_window_size(high_resolution, (5000, 4500))
        shared = clamp_window_size_for_screen(high_resolution, (5000, 4500))
        self.assertEqual(
            compacted,
            (
                min(
                    shared[0],
                    max(_MAIN_MIN_WIDTH, int(high_resolution.width() * _MAIN_MAX_WIDTH_RATIO)),
                ),
                shared[1],
            ),
        )
        self.assertEqual(shared, (896, 1152))
        self.assertLess(compacted[0], high_resolution.width())
        self.assertLessEqual(compacted[0], 960)

    def test_main_window_cap_never_exceeds_minimal_work_area(self) -> None:
        compact = _clamp_main_window_size(QtCore.QRect(0, 0, 160, 120), (5000, 3000))
        self.assertLessEqual(compact[0], 160)
        self.assertGreaterEqual(compact[0], _MIN_SIDEBAR_WIDTH)
        self.assertEqual(compact[1], clamp_window_size_for_screen(QtCore.QRect(0, 0, 160, 120), (5000, 3000))[1])

    def test_main_window_tiny_profiles_follow_compact_ratio_not_fixed_floor(self) -> None:
        for available in (
            QtCore.QRect(0, 0, 800, 480),
            QtCore.QRect(0, 0, 900, 560),
        ):
            with self.subTest(available=available):
                compact = _clamp_main_window_size(available, (5000, 3000))
                self.assertLess(available.width(), _MAIN_FIXED_MIN_WIDTH_THRESHOLD)
                self.assertLessEqual(compact[0], int(available.width() * _MAIN_MAX_WIDTH_RATIO))
                self.assertGreaterEqual(compact[0], _MIN_SIDEBAR_WIDTH)

    def test_shared_window_guard_reduces_oversized_minimums(self) -> None:
        available = QtCore.QRect(80, 40, 1024, 640)
        max_w, max_h = clamp_window_geometry(available, (1800, 1200))
        window = QtWidgets.QMainWindow()
        try:
            window.setMinimumSize(1600, 900)
            window.resize(1800, 1200)
            window.move(-5000, -4000)

            with patch("rapidpy_common.ui._screen_area_for_widget", return_value=available):
                _fit_window_to_screen(window)

            self.assertLessEqual(window.minimumWidth(), max_w)
            self.assertLessEqual(window.minimumHeight(), max_h)
            self.assertLessEqual(window.width(), max_w)
            self.assertLessEqual(window.height(), max_h)
            self.assertGreaterEqual(window.x(), available.left())
            self.assertGreaterEqual(window.y(), available.top())
            self.assertLessEqual(window.x() + window.width(), available.right() + 1)
            self.assertLessEqual(window.y() + window.height(), available.bottom() + 1)
        finally:
            window.deleteLater()

    def test_topology_refit_reclamps_all_top_level_windows(self) -> None:
        available = QtCore.QRect(-1600, 40, 1024, 640)
        max_w, max_h = clamp_window_geometry(available, (2200, 1600))
        window = QtWidgets.QMainWindow()
        try:
            window.setMinimumSize(2200, 1600)
            window.resize(2200, 1600)
            window.move(4000, 3000)

            with patch("rapidpy_common.ui._screen_area_for_widget", return_value=available):
                _refit_top_level_windows(_TopologyApplication([window]))  # type: ignore[arg-type]

            self.assertLessEqual(window.width(), max_w)
            self.assertLessEqual(window.height(), max_h)
            self.assertGreaterEqual(window.x(), available.left())
            self.assertGreaterEqual(window.y(), available.top())
            self.assertLessEqual(window.x() + window.width(), available.right() + 1)
            self.assertLessEqual(window.y() + window.height(), available.bottom() + 1)
        finally:
            window.deleteLater()

    def test_main_window_fit_reduces_oversized_restore_and_sidebar(self) -> None:
        available = QtCore.QRect(80, 40, 1024, 640)
        fitted_w, fitted_h = _clamp_main_window_size(available, (2200, 1600))
        window = QtWidgets.QMainWindow()
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        sidebar = QtWidgets.QWidget()
        content = QtWidgets.QWidget()
        splitter.addWidget(sidebar)
        splitter.addWidget(content)
        window.setCentralWidget(splitter)
        try:
            window._current_screen = lambda: available  # type: ignore[attr-defined]
            window._main_splitter = splitter  # type: ignore[attr-defined]
            window._sidebar = sidebar  # type: ignore[attr-defined]
            window._sidebar_min_width = _MIN_SIDEBAR_WIDTH  # type: ignore[attr-defined]
            window._sidebar_default_width = _MAX_SIDEBAR_WIDTH  # type: ignore[attr-defined]
            window.setMinimumSize(1600, 900)
            window.resize(2200, 1600)
            window.move(-5000, -4000)
            window.show()
            QtWidgets.QApplication.processEvents()

            MainWindow._fit_window_to_current_screen(window)  # type: ignore[arg-type]
            QtWidgets.QApplication.processEvents()

            self.assertLessEqual(window.minimumWidth(), fitted_w)
            self.assertLessEqual(window.minimumHeight(), fitted_h)
            self.assertLessEqual(window.maximumWidth(), fitted_w)
            self.assertLessEqual(window.maximumHeight(), fitted_h)
            self.assertLessEqual(window.width(), fitted_w)
            self.assertLessEqual(window.height(), fitted_h)
            geometry = window.geometry()
            self.assertGreaterEqual(geometry.left(), available.left())
            self.assertGreaterEqual(geometry.top(), available.top())
            self.assertLessEqual(geometry.right(), available.right())
            self.assertLessEqual(geometry.bottom(), available.bottom())
            self.assertLessEqual(
                splitter.sizes()[0],
                clamp_sidebar_target(
                    _MAX_SIDEBAR_WIDTH,
                    window_width=max(1, window.width()),
                    minimum=_MIN_SIDEBAR_WIDTH,
                    maximum=_MAX_SIDEBAR_WIDTH,
                    ratio=_SIDEBAR_RESTORE_RATIO,
                ),
            )
        finally:
            window.deleteLater()

    def test_main_window_restore_geometry_is_compacted_before_show(self) -> None:
        available = QtCore.QRect(96, 60, 1366, 768)
        saver = QtWidgets.QMainWindow()
        saver.resize(5000, 3300)
        saver.move(0, 0)
        oversized_geometry = saver.saveGeometry()

        window = MainWindow()
        try:
            window._current_screen = lambda: available  # type: ignore[method-assign]
            window._settings = _WindowSettingsBridge({"ui/window_geometry": oversized_geometry})  # type: ignore[assignment]
            window._restore_layout_state()
            fitted = _clamp_main_window_size(available, (5000, 3300))

            self.assertLessEqual(window.width(), fitted[0])
            self.assertLessEqual(window.height(), fitted[1])
            self.assertGreaterEqual(window.width(), _MIN_SIDEBAR_WIDTH)
            self.assertGreaterEqual(window.height(), 240)
            window.show()
            QtWidgets.QApplication.processEvents()
            self.assertLessEqual(window.geometry().right(), available.right())
            self.assertLessEqual(window.geometry().bottom(), available.bottom())
        finally:
            window.deleteLater()
            saver.deleteLater()

    def test_main_window_keeps_sidebar_scale_under_tight_display_caps(self) -> None:
        for available in (
            QtCore.QRect(0, 0, 1366, 768),
            QtCore.QRect(10, 0, 1024, 700),
            QtCore.QRect(0, 40, 800, 500),
        ):
            with self.subTest(available=available):
                compact = _clamp_main_window_size(available, (4000, 3200))
                self.assertLessEqual(compact[0], max(_MAIN_MIN_WIDTH, int(available.width() * _MAIN_MAX_WIDTH_RATIO)))
                self.assertLessEqual(compact[0], available.width())
                if available.width() < _MAIN_FIXED_MIN_WIDTH_THRESHOLD:
                    self.assertGreaterEqual(compact[0], _MIN_SIDEBAR_WIDTH)
                else:
                    self.assertGreaterEqual(compact[0], min(_MAIN_MIN_WIDTH, available.width()))
                self.assertLess(compact[0], available.width())

    def test_main_window_restore_geometry_can_not_expand_to_full_width(self) -> None:
        tiny_offset_screen = QtCore.QRect(1920, 20, 1280, 760)
        compact = _clamp_main_window_size(tiny_offset_screen, (3000, 2200))
        self.assertLess(compact[0], tiny_offset_screen.width())
        self.assertLess(compact[1], tiny_offset_screen.height())
        self.assertLessEqual(
            compact[0],
            max(_MAIN_MIN_WIDTH, int(tiny_offset_screen.width() * _MAIN_MAX_WIDTH_RATIO)),
        )
        self.assertEqual(compact[1], clamp_window_size_for_screen(tiny_offset_screen, (3000, 2200))[1])

    def test_sidebar_width_respects_compact_ratio_caps(self) -> None:
        self.assertEqual(
            clamp_sidebar_target(
                82,
                window_width=460,
                minimum=_MIN_SIDEBAR_WIDTH,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            ),
            _MIN_SIDEBAR_WIDTH,
        )
        self.assertEqual(
            clamp_sidebar_target(
                24,
                window_width=1200,
                minimum=_MIN_SIDEBAR_WIDTH,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            ),
            _MIN_SIDEBAR_WIDTH,
        )
        self.assertEqual(
            clamp_sidebar_target(
                128,
                window_width=2000,
                minimum=_MIN_SIDEBAR_WIDTH,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            ),
            _MAX_SIDEBAR_WIDTH,
        )

    def test_main_window_representation_across_representative_profiles(self) -> None:
        profile_data = [
            (QtCore.QRect(0, 0, 1366, 768), (5000, 4500)),
            (QtCore.QRect(10, 20, 1280, 720), (600, 450)),
            (QtCore.QRect(0, 0, 1920, 1200), (400, 240)),
            (QtCore.QRect(0, 0, 3840, 2160), (5000, 4500)),
            (QtCore.QRect(0, 0, 1024, 576), (3000, 2200)),
            (QtCore.QRect(0, 0, 768, 432), (1600, 1200)),
        ]
        for available, requested in profile_data:
            with self.subTest(available=available.size(), requested=requested):
                shared = clamp_window_size_for_screen(available, requested)
                compact = _clamp_main_window_size(available, requested)
                max_main_width = min(
                    available.width(),
                    960,
                    max(_MAIN_MIN_WIDTH, int(available.width() * _MAIN_MAX_WIDTH_RATIO)),
                )
                self.assertLessEqual(compact[0], max_main_width)
                self.assertLessEqual(compact[0], max_main_width)
                if available.width() < _MAIN_FIXED_MIN_WIDTH_THRESHOLD:
                    self.assertGreaterEqual(compact[0], _MIN_SIDEBAR_WIDTH)
                else:
                    self.assertGreaterEqual(compact[0], min(_MAIN_MIN_WIDTH, available.width()))
                self.assertLessEqual(compact[0], max(shared[0], _MAIN_MIN_WIDTH))
                self.assertLessEqual(compact[1], available.height())

    def test_clamp_window_handles_offscreen_and_scaled_profiles(self) -> None:
        for available in (
            QtCore.QRect(1920, 0, 2560, 1440),
            QtCore.QRect(0, 0, 3200, 1800),
            QtCore.QRect(-1920, 0, 1920, 1080),
            QtCore.QRect(0, 0, 1024, 768),
            QtCore.QRect(10, -240, 1280, 720),
        ):
            with self.subTest(available=available):
                clamped = clamp_window_geometry(available, (1000, 1000))
                self.assertLessEqual(
                    clamped[0],
                    MIN_WINDOW_WIDTH if available.width() < 914 else int(available.width() * 0.35),
                )
                self.assertGreaterEqual(clamped[0], MIN_WINDOW_WIDTH)
                self.assertLessEqual(clamped[1], max(int(available.height() * 0.80), 240))
                self.assertGreaterEqual(clamped[1], 240)
                self.assertLessEqual(clamped[0], available.width())
                self.assertLessEqual(clamped[1], available.height())

                compact = _clamp_main_window_size(available, (1000, 1000))
                self.assertLessEqual(
                    compact[0],
                    min(
                        available.width(),
                        960,
                        max(_MAIN_MIN_WIDTH, int(available.width() * _MAIN_MAX_WIDTH_RATIO)),
                    ),
                )
                if available.width() < _MAIN_FIXED_MIN_WIDTH_THRESHOLD:
                    self.assertGreaterEqual(compact[0], _MIN_SIDEBAR_WIDTH)
                else:
                    self.assertGreaterEqual(compact[0], min(_MAIN_MIN_WIDTH, available.width()))
                self.assertLessEqual(compact[1], max(240, int(available.height() * 0.80)))

                # Ensure requested values are never used when they exceed compact-safe bounds.
                self.assertLess(compact[0], 1000)
                self.assertLessEqual(compact[1], 1000)

    def test_sidebar_ratio_caps_track_profile(self) -> None:
        for window_width in (500, 900, 1200, 1800):
            with self.subTest(window_width=window_width):
                target = clamp_sidebar_target(
                    requested=64,
                    window_width=window_width,
                    minimum=_MIN_SIDEBAR_WIDTH,
                    maximum=_MAX_SIDEBAR_WIDTH,
                    ratio=_SIDEBAR_RESTORE_RATIO,
                )
                self.assertGreaterEqual(target, _MIN_SIDEBAR_WIDTH)
                self.assertLessEqual(target, _MAX_SIDEBAR_WIDTH)
                self.assertLessEqual(target, max(_MIN_SIDEBAR_WIDTH, int(window_width * _SIDEBAR_RESTORE_RATIO)))

    def test_adaptive_stack_minimum_size_hint_preserves_screen_capped_height(self) -> None:
        stack = _AdaptivePanelStack()
        for _ in range(3):
            panel = QtWidgets.QWidget()
            panel.setMinimumWidth(1600)
            panel.setMinimumHeight(1200)
            stack.addWidget(panel)

        hint = stack.minimumSizeHint()
        self.assertEqual(hint.width(), 0)
        self.assertGreater(hint.height(), 0)
        app = QtWidgets.QApplication.instance()
        screen = app.primaryScreen() if app is not None else None
        if screen is not None:
            self.assertLessEqual(hint.height(), screen.availableGeometry().height())


if __name__ == "__main__":
    unittest.main(verbosity=2)
