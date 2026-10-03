from __future__ import annotations

import ast
from pathlib import Path
import unittest

from rapidpy_common.ui import _fit_window_with_widget_handler


class _HasSimpleFit:
    def __init__(self) -> None:
        self.calls: list[str] = []

    def _fit_to_screen(self) -> None:
        self.calls.append("fit_to_screen")


class _HasScreenAwareFit:
    def __init__(self) -> None:
        self.calls: list[object | None] = []

    def _fit_window_to_current_screen(self, _screen) -> None:
        self.calls.append(_screen)


class TestWindowBootstrapContracts(unittest.TestCase):
    _EXPECTED_APP_ENTRYPOINTS = 16
    _EXPECTED_MAIN_WRAPPERS = 16

    @staticmethod
    def _app_py_files() -> list[Path]:
        repo_root = Path(__file__).resolve().parents[2]
        return [
            path
            for path in sorted(repo_root.glob("*/*/app.py"))
            if "__pycache__" not in path.parts
        ]

    @staticmethod
    def _main_py_files() -> list[Path]:
        repo_root = Path(__file__).resolve().parents[2]
        return [
            path
            for path in sorted(repo_root.glob("*/main.py"))
            if "__pycache__" not in path.parts
        ]

    @staticmethod
    def _call_uses_apply_bounds_guard(node: ast.AST, target_var: str = "app") -> bool:
        for call in ast.walk(node):
            if not isinstance(call, ast.Call):
                continue
            if not isinstance(call.func, ast.Name):
                continue
            if call.func.id != "apply_window_bounds_guard":
                continue
            if not call.args:
                continue
            arg = call.args[0]
            if isinstance(arg, ast.Name) and arg.id == target_var:
                return True
        return False

    @staticmethod
    def _imports_apply_bounds_guard(node: ast.AST) -> bool:
        for stmt in ast.walk(node):
            if (
                isinstance(stmt, ast.ImportFrom)
                and stmt.module == "rapidpy_common.ui"
                and any(alias.name == "apply_window_bounds_guard" for alias in stmt.names)
            ):
                return True
            if (
                isinstance(stmt, ast.Import)
                and any(alias.name == "rapidpy_common.ui" for alias in stmt.names)
            ):
                return True
        return False

    @staticmethod
    def _call_system_exit_main(node: ast.AST) -> bool:
        for call in ast.walk(node):
            if not isinstance(call, ast.Call):
                continue
            if not isinstance(call.func, ast.Name) or call.func.id != "SystemExit":
                continue
            if not call.args:
                continue
            inner = call.args[0]
            if isinstance(inner, ast.Call) and isinstance(inner.func, ast.Name) and inner.func.id == "main":
                return True
        return False

    @staticmethod
    def _imports_local_app_main(tree: ast.AST, package_hint: str) -> bool:
        for stmt in tree.body:
            if not isinstance(stmt, ast.ImportFrom):
                continue
            if package_hint == "rapid_main" and stmt.module == "rapid_main.__main__":
                if any(alias.name == "console_main" and alias.asname == "main" for alias in stmt.names):
                    return True
            if stmt.module is None or not stmt.module.endswith(".app"):
                continue
            module_base = stmt.module.split(".")[-2]
            if module_base != package_hint and module_base != "app":
                continue
            if any(alias.name == "main" for alias in stmt.names):
                return True
        return False

    def test_all_app_entrypoints_use_window_bounds_guard(self) -> None:
        missing_import: list[str] = []
        missing_call: list[str] = []
        for path in self._app_py_files():
            tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
            imported = self._imports_apply_bounds_guard(tree)
            if not imported:
                missing_import.append(str(path))
                continue

            main_defs = [stmt for stmt in tree.body if isinstance(stmt, ast.FunctionDef) and stmt.name == "main"]
            if not main_defs:
                missing_call.append(str(path))
                continue

            for main_def in main_defs:
                if not self._call_uses_apply_bounds_guard(main_def):
                    missing_call.append(str(path))
                    break

        self.assertEqual(len(self._app_py_files()), self._EXPECTED_APP_ENTRYPOINTS)
        self.assertEqual(
            missing_import,
            [],
            msg=f"Missing apply_window_bounds_guard import: {missing_import}",
        )
        self.assertEqual(
            missing_call,
            [],
            msg=f"main() missing apply_window_bounds_guard(app): {missing_call}",
        )

    def test_all_main_wrappers_call_packaged_app_main(self) -> None:
        self.assertEqual(len(self._main_py_files()), self._EXPECTED_MAIN_WRAPPERS)
        missing_import: list[str] = []
        missing_dispatch: list[str] = []
        for path in self._main_py_files():
            tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
            package_hint = path.parent.name
            imports_main = False
            for stmt in tree.body:
                if isinstance(stmt, ast.ImportFrom):
                    if any(alias.name == "main" or (alias.name == "console_main" and alias.asname == "main") for alias in stmt.names):
                        imports_main = True
            if not imports_main:
                missing_import.append(str(path))
            if not self._imports_local_app_main(tree, package_hint):
                missing_import.append(str(path))
            if not self._call_system_exit_main(tree):
                missing_dispatch.append(str(path))

            text = path.read_text(encoding="utf-8", errors="replace")
            if (
                "QApplication" in text
                and "from PySide6 import QtWidgets" in text
            ):
                self.fail(
                    f"{path} should rely on app.main() for QApplication bootstrap "
                    "so shared sizing/screen guards are consistently applied."
                )

        self.assertEqual(
            missing_import,
            [],
            msg=f"main.py wrapper missing app.main import: {missing_import}",
        )
        self.assertEqual(
            missing_dispatch,
            [],
            msg=f"main.py wrapper missing SystemExit(main()) call: {missing_dispatch}",
        )

    def test_entrypoint_counts_match(self) -> None:
        self.assertEqual(len(self._app_py_files()), self._EXPECTED_APP_ENTRYPOINTS)
        self.assertEqual(len(self._main_py_files()), self._EXPECTED_MAIN_WRAPPERS)

    def test_no_entrypoint_maximizes_by_default(self) -> None:
        offenders: list[str] = []
        for path in self._app_py_files() + self._main_py_files():
            text = path.read_text(encoding="utf-8", errors="replace")
            if "showMaximized(" in text or "showFullScreen(" in text:
                offenders.append(str(path))

        self.assertEqual(
            offenders,
            [],
            msg=f"Found startup fullscreen/maximized usage that can force oversized windows: {offenders}",
        )

    def test_rapid_main_vrm_launch_contract_points_to_vrm_logger(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapid_main" / "rapid_main" / "app.py"
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
        launch_defs = [
            stmt for stmt in ast.walk(tree)
            if isinstance(stmt, ast.FunctionDef) and stmt.name == "_launch_vrm"
        ]
        self.assertEqual(len(launch_defs), 1)

        calls = [
            call for call in ast.walk(launch_defs[0])
            if (
                isinstance(call, ast.Call)
                and isinstance(call.func, ast.Attribute)
                and call.func.attr == "_launch_external_tool"
            )
        ]
        self.assertEqual(len(calls), 1)

        keywords = {kw.arg: kw.value for kw in calls[0].keywords}
        self.assertIsInstance(keywords.get("target_path"), ast.Constant)
        self.assertIsInstance(keywords.get("app_name"), ast.Constant)
        self.assertEqual(keywords["target_path"].value, "vrm_logger/main.py")
        self.assertEqual(keywords["app_name"].value, "VRM Logger")

    def test_rapid_main_irm_voltage_calibration_routes_to_calibration_center(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapid_main" / "rapid_main" / "app.py"
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
        launch_defs = [
            stmt
            for stmt in ast.walk(tree)
            if isinstance(stmt, ast.FunctionDef) and stmt.name == "_launch_irm_voltage_calibration"
        ]
        self.assertEqual(len(launch_defs), 1)

        nav_calls = [
            call
            for call in ast.walk(launch_defs[0])
            if (
                isinstance(call, ast.Call)
                and isinstance(call.func, ast.Attribute)
                and call.func.attr == "_nav_select"
            )
        ]
        procedure_calls = [
            call
            for call in ast.walk(launch_defs[0])
            if (
                isinstance(call, ast.Call)
                and isinstance(call.func, ast.Attribute)
                and call.func.attr == "set_procedure"
            )
        ]

        self.assertTrue(any(call.args and isinstance(call.args[0], ast.Constant) and call.args[0].value == 5 for call in nav_calls))
        self.assertTrue(any(call.args and isinstance(call.args[0], ast.Constant) and call.args[0].value == "irm_voltage" for call in procedure_calls))

    def test_rapid_main_thermal_planning_routes_to_calibration_center(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapid_main" / "rapid_main" / "app.py"
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
        launch_defs = [
            stmt
            for stmt in ast.walk(tree)
            if isinstance(stmt, ast.FunctionDef) and stmt.name == "_launch_thermal_routine_planning"
        ]
        self.assertEqual(len(launch_defs), 1)

        nav_calls = [
            call
            for call in ast.walk(launch_defs[0])
            if (
                isinstance(call, ast.Call)
                and isinstance(call.func, ast.Attribute)
                and call.func.attr == "_nav_select"
            )
        ]
        procedure_calls = [
            call
            for call in ast.walk(launch_defs[0])
            if (
                isinstance(call, ast.Call)
                and isinstance(call.func, ast.Attribute)
                and call.func.attr == "set_procedure"
            )
        ]

        self.assertTrue(any(call.args and isinstance(call.args[0], ast.Constant) and call.args[0].value == 5 for call in nav_calls))
        self.assertTrue(any(call.args and isinstance(call.args[0], ast.Constant) and call.args[0].value == "thermal_routine" for call in procedure_calls))

    def test_rapid_main_debug_snapshot_includes_af_backend(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapid_main" / "rapid_main" / "app.py"
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
        status_defs = [
            stmt
            for stmt in ast.walk(tree)
            if isinstance(stmt, ast.FunctionDef) and stmt.name == "_diagnostic_status_lines"
        ]
        toggle_defs = [
            stmt
            for stmt in ast.walk(tree)
            if isinstance(stmt, ast.FunctionDef)
            and stmt.name in ("_on_nocomm_toggled", "_rebuild_diagnostic_backends")
        ]
        self.assertEqual(len(status_defs), 1)
        self.assertEqual(len(toggle_defs), 2)

        status_constants = [
            node.value
            for node in ast.walk(status_defs[0])
            if isinstance(node, ast.Constant)
        ]
        toggle_call_names = [
            node.id
            for definition in toggle_defs
            for node in ast.walk(definition)
            if isinstance(node, ast.Name)
        ]

        self.assertIn("AF Demag", status_constants)
        self.assertIn("build_af_demag_backend", toggle_call_names)

    def test_window_bounds_guard_prefers_widget_fit_handler(self) -> None:
        probe = _HasSimpleFit()
        self.assertTrue(_fit_window_with_widget_handler(probe))
        self.assertEqual(probe.calls, ["fit_to_screen"])

    def test_window_bounds_guard_prefers_screen_aware_fit_handler(self) -> None:
        probe = _HasScreenAwareFit()
        self.assertTrue(_fit_window_with_widget_handler(probe))
        self.assertEqual(probe.calls, [None])

    def test_debug_console_uses_qtgui_screen_type_for_screen_change_fit(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapid_main" / "rapid_main" / "dialogs" / "debug_console.py"
        text = path.read_text(encoding="utf-8", errors="replace")
        self.assertIn("QtGui.QScreen", text)
        self.assertNotIn("QtCore.QScreen", text)

    def test_shared_window_guard_watches_topology_work_area_and_dpi_changes(self) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        path = repo_root / "rapidpy_common" / "ui.py"
        text = path.read_text(encoding="utf-8", errors="replace")

        for signal_name in (
            "screenAdded.connect",
            "screenRemoved.connect",
            "availableGeometryChanged",
            "geometryChanged",
            "logicalDotsPerInchChanged",
        ):
            self.assertIn(signal_name, text)
