from __future__ import annotations

import unittest

from PySide6 import QtCore, QtWidgets
from rapid_main.app import (
    MainWindow,
    _MAIN_MAX_WIDTH_RATIO,
    _MAIN_MIN_WIDTH,
    _MAX_SIDEBAR_WIDTH,
    _MIN_SIDEBAR_WIDTH,
    _clamp_main_window_size,
)


class _TestAppMixin:
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


class SequenceLayoutSmokeTest(_TestAppMixin, unittest.TestCase):
    def _assert_spin_text_fits(self, spin: QtWidgets.QDoubleSpinBox) -> None:
        # Use a value near the configured max to exercise widest rendering.
        candidate = spin.maximum() / 2
        spin.setValue(candidate)
        text = spin.textFromValue(candidate)
        fm = spin.lineEdit().fontMetrics()
        available = spin.lineEdit().width() - 2
        self.assertGreater(
            available,
            0,
            "Spin line edit width must be positive before validating sequence layout",
        )
        self.assertLessEqual(
            fm.horizontalAdvance(text) + 4,
            available,
            f"Spin value '{text}' should fit in its field (need {fm.horizontalAdvance(text) + 4}, have {available})",
        )

        spin.setValue(spin.maximum())
        text = spin.textFromValue(spin.maximum())
        self.assertLessEqual(
            fm.horizontalAdvance(text) + 4,
            available,
            f"Spin max value '{text}' should fit in its field (need {fm.horizontalAdvance(text) + 4}, have {available})",
        )

    def test_sequence_panel_value_cells_fit_at_common_window_sizes(self) -> None:
        for width, height in ((1280, 720), (1920, 1080)):
            with self.subTest(size=f"{width}x{height}"):
                mw = MainWindow()
                mw.resize(width, height)
                mw._nav_select(2)  # Sequence panel
                mw.show()
                QtWidgets.QApplication.processEvents()

                sequence = mw._sequence
                spins = sequence.findChildren(QtWidgets.QDoubleSpinBox)
                self.assertGreater(len(spins), 0)
                for spin in spins:
                    self._assert_spin_text_fits(spin)

                nav_buttons = [
                    btn for btn in mw.findChildren(QtWidgets.QPushButton) if btn.objectName() == "navBtn"
                ]
                self.assertGreater(len(nav_buttons), 0)
                for btn in nav_buttons:
                    text_width = btn.fontMetrics().horizontalAdvance(btn.text())
                    self.assertGreaterEqual(
                        btn.width(),
                        text_width + 28,
                        f"Navigation button must display '{btn.text()}' fully",
                    )

                for area in mw.findChildren(QtWidgets.QAbstractScrollArea):
                    self.assertFalse(
                        area.horizontalScrollBar().isVisible(),
                        f"Horizontal scrollbar should stay hidden for {type(area).__name__}",
                    )

                mw.deleteLater()

    def test_stacked_panel_hints_are_clamped_for_startup_safety(self) -> None:
        mw = MainWindow()
        mw.show()
        QtWidgets.QApplication.processEvents()
        app = QtWidgets.QApplication.instance()
        screen = app.primaryScreen() if app is not None else None
        available = screen.availableGeometry() if screen is not None else None

        stack_hint = mw._stack.sizeHint()
        stack_min = mw._stack.minimumSizeHint()
        panel_min_widths = [mw._stack.widget(i).minimumWidth() for i in range(mw._stack.count())]
        self.assertTrue(
            stack_hint.width() > 0 and stack_hint.height() > 0,
            "Stacked panel sizeHint must be populated",
        )
        self.assertTrue(
            stack_min.width() == 0 and stack_min.height() > 0,
            "Stacked panel minimumSizeHint must preserve vertical minimum and permit compact startup width",
        )
        self.assertTrue(
            all(width == 0 for width in panel_min_widths),
            "Individual stacked pages must not enforce a minimum width.",
        )
        self.assertEqual(
            mw._stack.minimumWidth(),
            0,
            "Stacked panel widget must not force startup width.",
        )
        self.assertEqual(
            mw._stack.minimumHeight(),
            0,
            "Stacked panel widget must not force startup height.",
        )
        if available is not None:
            self.assertLessEqual(
                stack_min.width(),
                max(320, int(available.width() * _MAIN_MAX_WIDTH_RATIO)),
            )
            self.assertLessEqual(stack_min.height(), available.height())

        self.assertLessEqual(stack_hint.height(), stack_min.height() + 640)

        mw.deleteLater()

    def test_main_chrome_does_not_force_full_width_startup(self) -> None:
        mw = MainWindow()
        mw.show()
        QtWidgets.QApplication.processEvents()
        compact = QtCore.QRect(0, 0, 1280, 720)
        fitted = _clamp_main_window_size(compact, (3200, 1800))
        mw._current_screen = lambda: compact  # type: ignore[method-assign]

        try:
            mw._fit_window_to_current_screen()
            QtWidgets.QApplication.processEvents()

            self.assertLessEqual(mw.minimumSizeHint().width(), fitted[0])
            self.assertLessEqual(mw.width(), fitted[0])
            self.assertLess(mw.width(), compact.width())
        finally:
            mw.deleteLater()

    def test_main_window_fit_allows_tiny_profile_below_legacy_floor(self) -> None:
        mw = MainWindow()
        tiny = QtCore.QRect(0, 0, 800, 600)
        fitted = _clamp_main_window_size(tiny, (5000, 3000))
        mw._current_screen = lambda: tiny  # type: ignore[method-assign]

        try:
            mw.resize(5000, 3000)
            mw.show()
            QtWidgets.QApplication.processEvents()
            mw._fit_window_to_current_screen()
            QtWidgets.QApplication.processEvents()

            self.assertLess(fitted[0], _MAIN_MIN_WIDTH)
            self.assertLessEqual(fitted[0], int(tiny.width() * _MAIN_MAX_WIDTH_RATIO))
            self.assertLessEqual(mw.width(), fitted[0])
            self.assertLessEqual(mw.maximumWidth(), fitted[0])
            self.assertLess(mw.width(), tiny.width())
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
        finally:
            mw.deleteLater()

    def test_startup_window_restored_state_never_opens_full_width(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        app = QtWidgets.QApplication.instance()
        if app is None:
            raise AssertionError("QApplication must be initialized for startup-size test.")
        screen = app.primaryScreen()
        if screen is None:
            raise AssertionError("Primary screen is not available for startup-size test.")
        available = screen.availableGeometry()

        original_geometry = settings.value("ui/window_geometry")
        original_sidebar = settings.value("ui/sidebar_width")
        original_splitter = settings.value("ui/main_splitter_state")

        sentinel = QtWidgets.QMainWindow()
        sentinel.setGeometry(0, 0, 3000, 2200)
        settings.setValue("ui/window_geometry", sentinel.saveGeometry())
        settings.setValue("ui/sidebar_width", 512)
        settings.setValue("ui/main_splitter_state", QtCore.QByteArray())

        try:
            mw = MainWindow()
            mw.show()
            QtWidgets.QApplication.processEvents()

            compact_cap = _clamp_main_window_size(available, (3000, 2200))
            self.assertLessEqual(mw.width(), compact_cap[0])
            self.assertLessEqual(mw.height(), compact_cap[1])
            self.assertLess(mw.width(), available.width())
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
            self.assertLess(mw._sidebar.width(), available.width())
            mw.deleteLater()
        finally:
            if original_geometry is None:
                settings.remove("ui/window_geometry")
            else:
                settings.setValue("ui/window_geometry", original_geometry)

            if original_sidebar is None:
                settings.remove("ui/sidebar_width")
            else:
                settings.setValue("ui/sidebar_width", original_sidebar)

            if original_splitter is None:
                settings.remove("ui/main_splitter_state")
            else:
                settings.setValue("ui/main_splitter_state", original_splitter)

    def test_startup_with_extreme_saved_geometry_remains_compact(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        app = QtWidgets.QApplication.instance()
        if app is None:
            raise AssertionError("QApplication must be initialized for startup-size test.")
        screen = app.primaryScreen()
        if screen is None:
            raise AssertionError("Primary screen is not available for startup-size test.")
        available = screen.availableGeometry()

        original_geometry = settings.value("ui/window_geometry")
        original_sidebar = settings.value("ui/sidebar_width")
        original_splitter = settings.value("ui/main_splitter_state")

        sentinel = QtWidgets.QMainWindow()
        sentinel.setGeometry(0, 0, 16000, 9600)
        settings.setValue("ui/window_geometry", sentinel.saveGeometry())
        settings.setValue("ui/sidebar_width", 4096)
        settings.setValue("ui/main_splitter_state", QtCore.QByteArray())

        try:
            mw = MainWindow()
            mw.show()
            QtWidgets.QApplication.processEvents()

            compact_cap = _clamp_main_window_size(available, (16000, 9600))
            self.assertLessEqual(mw.width(), compact_cap[0])
            self.assertLess(mw.width(), available.width())
            self.assertLess(mw.height(), available.height())
            self.assertLessEqual(mw.width(), _MAIN_MAX_WIDTH_RATIO * available.width())
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
            mw.deleteLater()
        finally:
            if original_geometry is None:
                settings.remove("ui/window_geometry")
            else:
                settings.setValue("ui/window_geometry", original_geometry)

            if original_sidebar is None:
                settings.remove("ui/sidebar_width")
            else:
                settings.setValue("ui/sidebar_width", original_sidebar)

            if original_splitter is None:
                settings.remove("ui/main_splitter_state")
            else:
                settings.setValue("ui/main_splitter_state", original_splitter)

    def test_restored_layout_settings_can_never_reopen_full_screen(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        app = QtWidgets.QApplication.instance()
        if app is None:
            raise AssertionError("QApplication must be initialized for layout restoration test.")
        screen = app.primaryScreen()
        if screen is None:
            raise AssertionError("Primary screen is not available for layout restoration test.")
        available = screen.availableGeometry()

        sentinel = QtWidgets.QMainWindow()
        sentinel.setGeometry(0, 0, 3000, 2200)
        original_geometry = settings.value("ui/window_geometry")
        original_sidebar = settings.value("ui/sidebar_width", None, type=int)
        original_splitter = settings.value("ui/main_splitter_state")

        settings.setValue("ui/window_geometry", sentinel.saveGeometry())
        settings.setValue("ui/sidebar_width", 512)
        settings.setValue("ui/main_splitter_state", QtCore.QByteArray())

        try:
            mw = MainWindow()
            restored_cap = _clamp_main_window_size(available, (3000, 2200))
            self.assertLessEqual(mw.width(), restored_cap[0])
            self.assertLessEqual(mw.height(), restored_cap[1])
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
            mw.deleteLater()
        finally:
            if original_geometry is None:
                settings.remove("ui/window_geometry")
            else:
                settings.setValue("ui/window_geometry", original_geometry)
            if original_sidebar is None:
                settings.remove("ui/sidebar_width")
            else:
                settings.setValue("ui/sidebar_width", original_sidebar)
            if original_splitter is None:
                settings.remove("ui/main_splitter_state")
            else:
                settings.setValue("ui/main_splitter_state", original_splitter)

    def test_restored_geometry_respects_offscreen_like_available_area(self) -> None:
        settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
        app = QtWidgets.QApplication.instance()
        if app is None:
            raise AssertionError("QApplication must be initialized for layout restoration test.")

        original_geometry = settings.value("ui/window_geometry")
        original_sidebar = settings.value("ui/sidebar_width")
        original_splitter = settings.value("ui/main_splitter_state")

        sentinel = QtWidgets.QMainWindow()
        sentinel.setGeometry(0, 0, 3400, 2100)
        settings.setValue("ui/window_geometry", sentinel.saveGeometry())
        settings.setValue("ui/sidebar_width", 520)
        settings.setValue("ui/main_splitter_state", QtCore.QByteArray())

        # Simulate restore from a monitor positioned to the far left.
        available = QtCore.QRect(-1920, 0, 1280, 720)
        fitted = None

        class _FakeMonitorWindow(MainWindow):
            def _current_screen(self) -> QtCore.QRect:  # type: ignore[override]
                return available

        try:
            mw = _FakeMonitorWindow()
            mw.show()
            QtWidgets.QApplication.processEvents()
            mw._fit_window_to_current_screen()
            QtWidgets.QApplication.processEvents()
            fitted = _clamp_main_window_size(
                available,
                (3400, 2100),
            )

            self.assertLessEqual(mw.width(), fitted[0])
            self.assertLess(mw.width(), available.width())
            self.assertLessEqual(mw.height(), fitted[1])
            self.assertLess(mw.height(), available.height())
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
            self.assertLess(mw.x(), available.left() + available.width())
            self.assertGreaterEqual(mw.x(), available.left())
            self.assertLessEqual(fitted[0], max(_MAIN_MIN_WIDTH, int(available.width() * _MAIN_MAX_WIDTH_RATIO)))
            self.assertEqual(mw.x(), max(available.left(), min(mw.x(), available.right() - mw.width() + 1)))
            self.assertLess(fitted[1], available.height())
            mw.deleteLater()
        finally:
            if original_geometry is None:
                settings.remove("ui/window_geometry")
            else:
                settings.setValue("ui/window_geometry", original_geometry)

            if original_sidebar is None:
                settings.remove("ui/sidebar_width")
            else:
                settings.setValue("ui/sidebar_width", original_sidebar)

            if original_splitter is None:
                settings.remove("ui/main_splitter_state")
            else:
                settings.setValue("ui/main_splitter_state", original_splitter)

    def test_dynamic_screen_resize_keeps_sidebar_and_window_compact(self) -> None:
        mw = MainWindow()
        try:
            mw.resize(2600, 1600)
            mw.show()
            QtWidgets.QApplication.processEvents()

            compact_screen = QtCore.QRect(0, 0, 900, 560)
            mw._current_screen = lambda: compact_screen  # type: ignore[method-assign]
            mw._fit_window_to_current_screen()
            restored_cap = _clamp_main_window_size(compact_screen, (2600, 1600))

            self.assertLessEqual(mw.width(), restored_cap[0])
            self.assertLessEqual(mw.height(), restored_cap[1])
            self.assertLessEqual(mw._sidebar.width(), _MAX_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw._sidebar.width(), _MIN_SIDEBAR_WIDTH)
            self.assertGreaterEqual(mw.width(), _MIN_SIDEBAR_WIDTH)
            self.assertLessEqual(mw.width(), compact_screen.width())
            self.assertLessEqual(mw.height(), compact_screen.height())
        finally:
            mw.deleteLater()


if __name__ == "__main__":
    unittest.main()
