"""Verify the release manifest and prelaunch checks using Windows PowerShell."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest


INSTALLER = Path(__file__).resolve().parents[3] / "installer"
POWERSHELL = shutil.which("powershell.exe") if sys.platform == "win32" else None


@unittest.skipUnless(POWERSHELL, "Windows PowerShell is required")
class PortableWindowsManifestTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory(prefix="rapidpy-portable-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.bundle = self.root / "RAPID Portable"
        self.bundle.mkdir()
        self.env = {**os.environ, "RAPID_TEST_BUNDLE": str(self.bundle),
                    "RAPID_TEST_INSTALLER": str(INSTALLER), "RAPID_TEST_EVIDENCE": str(self.root / "reports")}
        for name in ("RapidPyMain.exe", "RapidPyMainConsole.exe", "OperatorGuide.md",
                     "Verify-RapidPy.ps1", "portable_bundle.psm1", "_internal/runtime.dll"):
            path = self.bundle / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(name, encoding="utf-8")
        self.manifest = self.bundle / "build_manifest.json"
        inventory = self.powershell(
            "Get-RapidPyBundleInventory -Bundle $env:RAPID_TEST_BUNDLE | ConvertTo-Json -Depth 4",
        )
        self.assertEqual(inventory.returncode, 0, inventory.stderr)
        self.write_manifest({"schema": "rapidpy.portable_build.v2", "files": json.loads(inventory.stdout),
                             "source_revision": "a" * 40})

    def write_manifest(self, payload):
        self.manifest.write_text(json.dumps(payload), encoding="utf-8")

    def powershell(self, command):
        return subprocess.run(
            [POWERSHELL, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command",
             "$ErrorActionPreference='Stop'; Import-Module (Join-Path $env:RAPID_TEST_INSTALLER "
             "'portable_bundle.psm1') -Force; " + command],
            env=self.env, capture_output=True, text=True, timeout=30,
        )

    def verify_inventory(self, *, require_full=True):
        flag = "-RequireFullInventory" if require_full else ""
        return self.powershell(f"Test-RapidPyBundle -Bundle $env:RAPID_TEST_BUNDLE {flag} | ConvertTo-Json")

    def test_complete_inventory_is_verified_on_windows_powershell(self):
        result = self.verify_inventory()
        self.assertEqual(result.returncode, 0, result.stderr)
        report = json.loads(result.stdout)
        self.assertIs(report["full_inventory"], True)
        self.assertEqual(report["files_checked"], 6)
        self.assertEqual(report["source_revision"], "a" * 40)

    def test_same_size_corruption_extra_and_missing_files_fail(self):
        runtime = self.bundle / "_internal" / "runtime.dll"
        original = runtime.read_bytes()
        runtime.write_bytes(b"x" * len(original))
        result = self.verify_inventory()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("differs from its manifest", result.stderr)
        runtime.write_bytes(original)
        extra = self.bundle / "_internal" / "unlisted.dll"
        extra.write_bytes(b"extra")
        result = self.verify_inventory()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Unexpected portable file", result.stderr)
        extra.unlink()
        runtime.unlink()
        result = self.verify_inventory()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Missing portable files", result.stderr)

    def test_legacy_manifest_checks_executables_without_claiming_full_inventory(self):
        self.write_manifest({"schema": "rapidpy.portable_build.v1", "executables": {
            name: hashlib.sha256((self.bundle / name).read_bytes()).hexdigest()
            for name in ("RapidPyMain.exe", "RapidPyMainConsole.exe")
        }})
        result = self.verify_inventory(require_full=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIs(json.loads(result.stdout)["full_inventory"], False)
        result = self.verify_inventory()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("version 2 manifest", result.stderr)

    def test_verifier_writes_failed_inventory_evidence_before_launching(self):
        (self.bundle / "OperatorGuide.md").write_bytes(b"changed")
        result = self.powershell(
            "& (Join-Path $env:RAPID_TEST_INSTALLER 'verify_rapid_main.ps1') "
            "-Bundle $env:RAPID_TEST_BUNDLE -EvidenceDirectory $env:RAPID_TEST_EVIDENCE -RequireFullInventory"
        )
        self.assertNotEqual(result.returncode, 0)
        reports = list((self.root / "reports").glob("*/verification.json"))
        self.assertEqual(len(reports), 1, result.stderr)
        evidence = json.loads(reports[0].read_text(encoding="utf-8-sig"))
        self.assertIs(evidence["passed"], False)
        self.assertIn("OperatorGuide.md", evidence["error"])
        self.assertEqual(list(reports[0].parent.glob("*.stdout.log")), [])
