from __future__ import annotations

import io
from pathlib import Path
import tempfile
import tomllib
import unittest
from unittest.mock import patch

from rapid_main.__main__ import main as module_main
from rapid_main.package_launch import ToolUnavailableError, resolve_tool_launch
from rapid_main.startup import collect_startup_environment, select_main_icon


class PackagingTests(unittest.TestCase):
    def test_distribution_manifest_exposes_main_and_helper_entry_points(self) -> None:
        pyproject = Path(__file__).resolve().parents[2] / "pyproject.toml"
        payload = tomllib.loads(pyproject.read_text(encoding="utf-8"))

        self.assertEqual(payload["project"]["name"], "berkeley-rapidpy")
        scripts = payload["project"]["scripts"]
        for name in (
            "rapid-main",
            "rapid-af-tuner",
            "rapid-data-viewer",
            "rapid-gaussmeter",
            "rapid-updown",
            "rapid-vrm",
            "rapid-webcam",
        ):
            self.assertIn(name, scripts)
        packages = payload["tool"]["setuptools"]["packages"]
        self.assertIn("rapid_main", packages)
        self.assertIn("rapid_main_assets", packages)
        self.assertIn("rapidpy_common", packages)

    def test_startup_report_finds_required_modules_and_packaged_icon(self) -> None:
        report = collect_startup_environment(optional_modules=())

        self.assertTrue(report.ok, report.blockers)
        self.assertTrue(all(report.required_modules.values()))
        self.assertTrue(Path(report.icon_path).is_file())
        icon_name, icon_path = select_main_icon()
        self.assertTrue(icon_name)
        self.assertTrue(icon_path.is_file())

    def test_module_startup_check_is_headless_and_machine_readable(self) -> None:
        output = io.StringIO()
        with patch("sys.stdout", output):
            exit_code = module_main(["--check-startup"])

        self.assertEqual(exit_code, 0)
        self.assertIn('"schema": "rapidpy.startup_environment.v1"', output.getvalue())
        self.assertIn('"ok": true', output.getvalue())

    def test_startup_check_blocks_a_malformed_existing_config(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            config_path = Path(td) / "config.json"
            config_path.write_text("not-json", encoding="utf-8")
            with patch.dict("os.environ", {"RAPID_CONFIG": str(config_path)}):
                report = collect_startup_environment(optional_modules=())

        self.assertFalse(report.ok)
        self.assertTrue(report.config_exists)
        self.assertFalse(report.config_valid)
        self.assertTrue(any("valid JSON object" in item for item in report.blockers))

    def test_startup_check_names_a_missing_required_dependency(self) -> None:
        report = collect_startup_environment(
            required_modules=("rapidpy_dependency_that_does_not_exist",),
            optional_modules=(),
        )

        self.assertFalse(report.ok)
        self.assertEqual(
            report.required_modules,
            {"rapidpy_dependency_that_does_not_exist": False},
        )
        self.assertIn(
            "Required Python module is unavailable: rapidpy_dependency_that_does_not_exist",
            report.blockers,
        )

    def test_tool_launcher_prefers_checkout_then_installed_module(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            script = root / "tool" / "main.py"
            script.parent.mkdir()
            script.write_text("# fixture\n", encoding="utf-8")
            source = resolve_tool_launch(
                module="example_tool",
                source_root=root,
                source_relative="tool/main.py",
                python_executable="python-test",
            )
            self.assertEqual(source.source, "source-checkout")
            self.assertEqual(source.command, ("python-test", str(script)))
            self.assertEqual(source.cwd, root)

            script.unlink()
            with patch("rapid_main.package_launch.find_spec", return_value=object()):
                installed = resolve_tool_launch(
                    module="example_tool",
                    source_root=root,
                    source_relative="tool/main.py",
                    python_executable="python-test",
                )
            self.assertEqual(installed.source, "installed-package")
            self.assertEqual(installed.command, ("python-test", "-m", "example_tool"))
            self.assertIsNone(installed.cwd)

    def test_tool_launcher_reports_a_precise_missing_component(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            with patch("rapid_main.package_launch.find_spec", return_value=None):
                with self.assertRaisesRegex(ToolUnavailableError, "installed module"):
                    resolve_tool_launch(
                        module="missing_tool",
                        source_root=td,
                        source_relative="missing/main.py",
                    )


if __name__ == "__main__":
    unittest.main(verbosity=2)

