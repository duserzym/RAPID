"""Omarchy-style tiling canvas: dwindle tiling, focus/swap, workspaces, docking."""
from __future__ import annotations

import unittest

from PySide6 import QtCore, QtWidgets

from rapidpy_common.tiling import (
    Leaf,
    Split,
    TilingCanvas,
    WorkspaceBar,
    insert_beside,
    install_tiling_shortcuts,
    leaves,
    node_from_dict,
    remove_leaf,
)

_APP = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def _pump(times: int = 5) -> None:
    for _ in range(times):
        _APP.processEvents()
        QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.Type.DeferredDelete)


class TreeTests(unittest.TestCase):
    def test_insert_remove_collapse(self):
        tree = insert_beside(None, None, "a", "right")
        tree = insert_beside(tree, "a", "b", "right")
        tree = insert_beside(tree, "b", "c", "down")
        self.assertEqual(leaves(tree), ["a", "b", "c"])
        self.assertIsInstance(tree, Split)
        tree = remove_leaf(tree, "b")
        self.assertEqual(leaves(tree), ["a", "c"])
        tree = remove_leaf(tree, "a")
        self.assertEqual(tree, Leaf("c"))

    def test_restore_drops_unknown_tiles_and_normalises_ratios(self):
        payload = {"split": "h", "children": [{"leaf": "a"}, {"leaf": "gone"}, {"leaf": "b"}], "ratios": [2, 1, 2]}
        node = node_from_dict(payload, {"a", "b"})
        self.assertEqual(leaves(node), ["a", "b"])
        self.assertAlmostEqual(sum(node.ratios), 1.0)


class CanvasTests(unittest.TestCase):
    def setUp(self):
        self.window = QtWidgets.QMainWindow()
        self.canvas = TilingCanvas(workspaces=4)
        self.canvas.set_animation_duration(0)
        self.window.setCentralWidget(self.canvas)
        self.panels = {}
        for key in ("dash", "queue", "seq", "measure"):
            panel = QtWidgets.QLabel(key)
            panel.setMinimumSize(0, 0)
            self.panels[key] = panel
            self.canvas.register(key, panel, key.title())
        self.window.resize(1200, 800)
        self.window.show()
        _pump()

    def tearDown(self):
        self.window.close()
        self.window.deleteLater()
        _pump()

    def test_stacked_widget_facade(self):
        canvas = self.canvas
        self.assertEqual(canvas.count(), 4)
        self.assertIs(canvas.widget(2), self.panels["seq"])
        canvas.setCurrentIndex(2)
        self.assertIs(canvas.currentWidget(), self.panels["seq"])
        self.assertEqual(canvas.currentIndex(), 2)
        self.assertEqual(canvas.minimumSizeHint().width(), 0)
        self.assertGreater(canvas.minimumSizeHint().height(), 0)

    def test_dwindle_splits_focused_tile_and_closing_returns_space(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue")
        _pump()
        self.assertEqual(canvas.visible_keys(), ["dash", "queue"])
        dash, queue = canvas.frame("dash"), canvas.frame("queue")
        self.assertTrue(dash.isVisible() and queue.isVisible())
        self.assertLess(dash.geometry().right(), queue.geometry().left() + 20)  # side by side
        canvas.open("seq")  # splits the focused (queue) tile, which is now taller than wide
        _pump()
        self.assertEqual(set(canvas.visible_keys()), {"dash", "queue", "seq"})
        canvas.close_tile("queue")
        _pump()
        self.assertEqual(canvas.visible_keys(), ["dash", "seq"])
        self.assertFalse(canvas.frame("queue").isVisible())

    def test_directional_focus_and_swap(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue", side="right")
        _pump()
        canvas.focus("dash")
        canvas.focus_direction("right")
        self.assertEqual(canvas.focused_key(), "queue")
        canvas.swap_direction("left")
        _pump()
        self.assertEqual(leaves(canvas.tree()), ["queue", "dash"])
        self.assertEqual(canvas.focused_key(), "queue")

    def test_workspaces_switch_and_move(self):
        canvas = self.canvas
        workspaces = []
        canvas.workspaceChanged.connect(workspaces.append)
        canvas.open("dash")
        canvas.open("seq", workspace=2)
        self.assertEqual(canvas.active_workspace, 2)
        self.assertFalse(canvas.frame("dash").isVisible())
        canvas.setCurrentIndex(0)  # opening a placed panel jumps to its workspace
        self.assertEqual(canvas.active_workspace, 1)
        canvas.move_to_workspace(3, "dash")
        self.assertEqual(canvas.workspace_of("dash"), 3)
        self.assertEqual(canvas.occupied_workspaces(), [2, 3])
        self.assertIn(2, workspaces)

    def test_monocle_and_classic_modes(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue")
        canvas.toggle_monocle()
        _pump()
        self.assertEqual(canvas.visible_keys(), ["queue"])
        canvas.toggle_monocle()
        canvas.set_classic(True)
        canvas.setCurrentIndex(3)
        _pump()
        self.assertEqual(canvas.visible_keys(), ["measure"])
        self.assertIs(canvas.currentWidget(), self.panels["measure"])
        canvas.set_classic(False)
        _pump()
        self.assertIn("measure", canvas.visible_keys())

    def test_resize_toggle_balance(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue", side="right")
        tree = canvas.tree()
        canvas.resize_focused(0.2)
        self.assertAlmostEqual(tree.ratios[1], 0.7, places=5)
        canvas.toggle_split()
        self.assertEqual(canvas.tree().orientation, "v")
        canvas.balance()
        self.assertEqual(canvas.tree().ratios, [0.5, 0.5])

    def test_drop_docking_zones(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue", side="right")
        canvas.dock("seq", "dash", "up")
        self.assertEqual(leaves(canvas.tree()), ["seq", "dash", "queue"])
        canvas.dock("queue", "seq", "center")  # centre swaps
        self.assertEqual(leaves(canvas.tree()), ["queue", "dash", "seq"])

    def test_state_round_trip(self):
        canvas = self.canvas
        canvas.open("dash")
        canvas.open("queue", side="down")
        canvas.open("seq", workspace=2)
        state = canvas.save_state()
        other = TilingCanvas(workspaces=4)
        for key in ("dash", "queue", "seq", "measure"):
            other.register(key, QtWidgets.QLabel(key), key)
        self.assertTrue(other.restore_state(state))
        self.assertEqual(leaves(other.tree(1)), ["dash", "queue"])
        self.assertEqual(other.tree(1).orientation, "v")
        self.assertEqual(other.active_workspace, 2)
        self.assertFalse(other.restore_state({"version": 99}))
        other.deleteLater()

    def test_tool_window_docks_floats_and_closes_with_its_window(self):
        canvas = self.canvas
        canvas.open("dash")
        dialog = QtWidgets.QDialog()
        dialog.setAttribute(QtCore.Qt.WidgetAttribute.WA_DeleteOnClose, True)
        destroyed = []
        dialog.destroyed.connect(lambda *_: destroyed.append(True))
        canvas.dock_window(dialog, "vacuum", "Vacuum")
        _pump()
        self.assertIn("vacuum", canvas.visible_keys())
        self.assertFalse(dialog.isWindow())
        canvas.close_tile("vacuum")
        _pump(10)
        self.assertNotIn("vacuum", canvas.registered_keys())
        self.assertEqual(destroyed, [True])

        floater = QtWidgets.QDialog()
        canvas.dock_window(floater, "step", "Step monitor")
        canvas.float_tile("step")
        _pump()
        self.assertTrue(floater.isWindow())
        self.assertTrue(floater.isVisible())
        self.assertNotIn("step", canvas.registered_keys())
        floater.close()

    def test_dialog_finishing_removes_its_tile(self):
        canvas = self.canvas
        canvas.open("dash")
        dialog = QtWidgets.QDialog()
        canvas.dock_window(dialog, "debug", "Debug console")
        dialog.reject()
        _pump(10)
        self.assertNotIn("debug", canvas.registered_keys())
        self.assertEqual(canvas.visible_keys(), ["dash"])

    def test_busy_tool_window_that_refuses_to_close_keeps_its_tile(self):
        class Busy(QtWidgets.QDialog):
            def closeEvent(self, event):
                event.ignore()

        canvas = self.canvas
        canvas.open("dash")
        busy = Busy()
        canvas.dock_window(busy, "adwin", "ADwin")
        canvas.close_tile("adwin")
        _pump()
        self.assertIn("adwin", canvas.visible_keys())

    def test_workspace_bar_and_shortcuts(self):
        canvas = self.canvas
        canvas.open("dash")
        bar = WorkspaceBar(canvas)
        shortcuts = install_tiling_shortcuts(self.window, canvas)
        self.assertGreater(len(shortcuts), 20)
        canvas.switch_workspace(2)
        self.assertTrue(bar._buttons[1].isChecked())
        self.assertTrue(bar._buttons[0].property("occupied"))
        bar.deleteLater()

    def test_launcher_lists_panels_and_actions(self):
        canvas = self.canvas
        calls = []
        canvas.add_launcher_action("Webcam", lambda: calls.append("webcam"))
        titles = [title for title, _ in canvas.launcher_entries()]
        self.assertTrue(any("Measure" in title for title in titles))
        self.assertIn("Webcam", titles)
        launcher = canvas.show_launcher()
        launcher.search.setText("webcam")
        launcher._activate()
        self.assertEqual(calls, ["webcam"])


if __name__ == "__main__":
    unittest.main()


class MainWindowTilingTests(unittest.TestCase):
    def _window(self, *, fresh: bool = True):
        from rapid_main.app import MainWindow

        if fresh:
            settings = QtCore.QSettings("RAPID", "RapidPy-rapid_main")
            settings.remove("ui/tiling_state")
            settings.remove("ui/active_panel")
            settings.sync()
        window = MainWindow()
        window._stack.set_animation_duration(0)
        window.resize(1600, 900)
        window.show()
        _pump()
        return window

    def _dispose(self, window):
        window._confirm_shutdown = lambda **kwargs: True
        window.close()
        window.deleteLater()
        _pump()

    def test_default_workspaces_and_navigation_sync(self):
        window = self._window()
        try:
            canvas = window._stack
            self.assertEqual(canvas.occupied_workspaces(), [1, 2, 3, 4])
            self.assertEqual(sorted(canvas.visible_keys()), ["dashboard", "measure"])
            window._nav_select(2)
            self.assertEqual(canvas.active_workspace, 3)
            self.assertIs(window._stack.currentWidget(), window._sequence)
            canvas.focus_direction("left")  # nothing to the left: focus stays
            canvas.switch_workspace(1)
            canvas.focus("measure")
            self.assertTrue(window._nav_btns[3].isChecked())
        finally:
            self._dispose(window)

    def test_tool_windows_tile_and_layout_persists(self):
        window = self._window()
        try:
            window._launch_step_monitor()
            _pump()
            self.assertIn("tool:step", window._stack.visible_keys())
            self.assertFalse(window._step_dlg.isWindow())
            window._stack.open("queue", workspace=1)
            window._save_layout_state()
        finally:
            self._dispose(window)
        reopened = self._window(fresh=False)
        try:
            # Panels persist across restarts; tool windows are not resurrected.
            self.assertEqual(reopened._stack.workspace_of("queue"), 1)
            self.assertNotIn("tool:step", reopened._stack.registered_keys())
            reopened._settings.remove("ui/tiling_state")
        finally:
            self._dispose(reopened)

    def test_classic_pages_mode_shows_one_panel(self):
        window = self._window()
        try:
            window._classic_action.setChecked(True)
            window._nav_select(4)
            _pump()
            self.assertEqual(window._stack.visible_keys(), ["settings"])
            window._classic_action.setChecked(False)
            self.assertFalse(window._stack.classic)
        finally:
            self._dispose(window)

    def test_header_safety_buttons_are_never_truncated(self):
        window = self._window()
        try:
            window.resize(1000, 700)
            _pump()
            for button in (window._pause_btn, window._halt_btn):
                self.assertGreaterEqual(button.width(), button.sizeHint().width(), button.text())
                self.assertEqual(button.maximumWidth(), 16777215)
            self.assertGreater(window._flow_lbl.width(), 60)
        finally:
            self._dispose(window)
