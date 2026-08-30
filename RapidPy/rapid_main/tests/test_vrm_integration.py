from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from rapid_main.vrm import VRM_CONTEXT_ENV, build_vrm_launch_context, write_vrm_launch_context
from vrm_logger.session_manifest import (
    build_vrm_output_manifest,
    load_handoff_context,
    vrm_manifest_path,
    write_vrm_output_manifest,
)


class VrmIntegrationTests(unittest.TestCase):
    def test_rapid_main_writes_launch_context_for_vrm_handoff(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            context = build_vrm_launch_context(
                active_automation=True,
                nocomm=True,
                metadata={"sample_id": "VRM_DEMO"},
            )
            path = write_vrm_launch_context(tmp, context=context)
            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["schema"], "rapidpy.vrm.launch_context.v1")
        self.assertTrue(payload["active_automation"])
        self.assertTrue(payload["nocomm"])
        self.assertEqual(payload["metadata"]["sample_id"], "VRM_DEMO")

    def test_vrm_logger_loads_handoff_context_and_writes_csv_sidecar(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            handoff_path = root / "launch.json"
            handoff_path.write_text(
                json.dumps({"schema": "rapidpy.vrm.launch_context.v1", "source_app": "rapid_main"}),
                encoding="utf-8",
            )
            csv_path = root / "run.csv"
            handoff = load_handoff_context({VRM_CONTEXT_ENV: str(handoff_path)})
            manifest_path = write_vrm_output_manifest(
                csv_path,
                session_start_epoch=1_700_000_000.0,
                interval_s=2.5,
                spacing_mode="Log",
                display_unit="Moment",
                baseline_volts=(0.1, 0.2, 0.3),
                calibration={"x": 1.0, "y": 2.0, "z": 3.0, "range_factor": 1e-6},
                handoff_context=handoff,
            )
            payload = json.loads(manifest_path.read_text(encoding="utf-8"))

        self.assertEqual(manifest_path.name, "run.vrm.json")
        self.assertEqual(payload["schema"], "rapidpy.vrm.session_manifest.v1")
        self.assertEqual(payload["output_csv"], str(csv_path))
        self.assertEqual(payload["handoff_context"]["source_app"], "rapid_main")
        self.assertEqual(payload["baseline_volts"], [0.1, 0.2, 0.3])

    def test_vrm_manifest_path_uses_csv_stem(self) -> None:
        self.assertEqual(
            vrm_manifest_path(Path("example.csv")),
            Path("example.vrm.json"),
        )

    def test_vrm_output_manifest_payload_is_json_ready(self) -> None:
        payload = build_vrm_output_manifest(
            output_csv=Path("run.csv"),
            session_start_epoch=1_700_000_000.0,
            interval_s=1.0,
            spacing_mode="Linear",
            display_unit="Volts",
            baseline_volts=None,
            calibration={"x": 1.0, "y": 1.0, "z": 1.0, "range_factor": 1.0},
            handoff_context={},
        )

        self.assertIsNone(payload["baseline_volts"])
        json.dumps(payload)
