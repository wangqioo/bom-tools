import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

from pstx_integrations.diagnostics import (
    build_diagnostics_status,
    diagnostics_export_bytes,
    sanitize_for_diagnostics,
    tail_diagnostics,
    write_diagnostic_event,
)
from pstx_integrations import diagnostics as integration_diagnostics
from pstx_integrations import diagnostics as pstx_diagnostics


class DiagnosticsTests(unittest.TestCase):
    def test_integration_diagnostics_entrypoint_exports_public_api(self):
        self.assertIs(integration_diagnostics.write_diagnostic_event, pstx_diagnostics.write_diagnostic_event)
        self.assertIs(integration_diagnostics.build_diagnostics_status, pstx_diagnostics.build_diagnostics_status)
        self.assertFalse(Path("pstx_diagnostics.py").exists())

    def test_write_event_redacts_sensitive_fields_and_tail_reads_jsonl(self):
        old_file = os.environ.get("PSTX_DIAGNOSTICS_LOG_FILE")
        old_enabled = os.environ.get("PSTX_DIAGNOSTICS_ENABLED")
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                log_file = str(Path(temp_dir) / "pstx_diagnostics.log")
                os.environ["PSTX_DIAGNOSTICS_LOG_FILE"] = log_file
                os.environ["PSTX_DIAGNOSTICS_ENABLED"] = "1"

                record = write_diagnostic_event("unit.test", {
                    "apiKey": "secret-key-value",
                    "nested": {"Authorization": "Bearer hidden-token"},
                    "message": "upstream failed Authorization: Bearer embedded-token apiKey=embedded-key appSecret=embedded-secret",
                    "normal": "visible",
                }, request_id="req-unit")
                tailed = tail_diagnostics(limit=5)

                self.assertEqual("req-unit", record["request_id"])
                self.assertEqual(1, tailed["count"])
                payload = tailed["records"][0]
                payload_text = json.dumps(payload, ensure_ascii=False)
                self.assertIn("visible", payload_text)
                self.assertNotIn("secret-key-value", payload_text)
                self.assertNotIn("hidden-token", payload_text)
                self.assertNotIn("embedded-token", payload_text)
                self.assertNotIn("embedded-key", payload_text)
                self.assertNotIn("embedded-secret", payload_text)
                self.assertTrue(payload["apiKey"]["redacted"])
        finally:
            if old_file is None:
                os.environ.pop("PSTX_DIAGNOSTICS_LOG_FILE", None)
            else:
                os.environ["PSTX_DIAGNOSTICS_LOG_FILE"] = old_file
            if old_enabled is None:
                os.environ.pop("PSTX_DIAGNOSTICS_ENABLED", None)
            else:
                os.environ["PSTX_DIAGNOSTICS_ENABLED"] = old_enabled

    def test_status_and_export_bundle_include_log_without_raw_secret(self):
        old_file = os.environ.get("PSTX_DIAGNOSTICS_LOG_FILE")
        old_parse_file = os.environ.get("PSTX_FEISHU_PARSE_LOG_FILE")
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                os.environ["PSTX_DIAGNOSTICS_LOG_FILE"] = str(Path(temp_dir) / "pstx_diagnostics.log")
                os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = str(Path(temp_dir) / "feishu_bom_parse_debug.log")
                write_diagnostic_event("unit.export", {"appSecret": "super-secret"})
                write_diagnostic_event(
                    "feishu_bom_parse.unit",
                    {"sheet_id": "sh1", "hq_no": "HQ001"},
                    log_file=os.environ["PSTX_FEISHU_PARSE_LOG_FILE"],
                )

                status = build_diagnostics_status()
                data, filename = diagnostics_export_bytes()
                bundle = Path(temp_dir) / filename
                bundle.write_bytes(data)

                self.assertTrue(status["ok"])
                self.assertTrue(status["log_file"]["exists"])
                self.assertTrue(status["feishu_parse_log_file"]["exists"])
                self.assertTrue(filename.endswith(".zip"))
                with zipfile.ZipFile(bundle) as archive:
                    names = archive.namelist()
                    self.assertIn("diagnostics_status.json", names)
                    self.assertIn("logs/pstx_diagnostics.log", names)
                    self.assertIn("logs/feishu_bom_parse_debug.log", names)
                    log_text = archive.read("logs/pstx_diagnostics.log").decode("utf-8")
                    parse_text = archive.read("logs/feishu_bom_parse_debug.log").decode("utf-8")
                    self.assertNotIn("super-secret", log_text)
                    self.assertIn("HQ001", parse_text)
        finally:
            if old_file is None:
                os.environ.pop("PSTX_DIAGNOSTICS_LOG_FILE", None)
            else:
                os.environ["PSTX_DIAGNOSTICS_LOG_FILE"] = old_file
            if old_parse_file is None:
                os.environ.pop("PSTX_FEISHU_PARSE_LOG_FILE", None)
            else:
                os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = old_parse_file

    def test_sanitize_truncates_large_payloads(self):
        payload = sanitize_for_diagnostics({
            "token": "abc123",
            "items": list(range(100)),
            "text": "x" * 1400,
        })

        self.assertTrue(payload["token"]["redacted"])
        self.assertLessEqual(len(payload["items"]), 81)
        self.assertLess(len(payload["text"]), 1300)


if __name__ == "__main__":
    unittest.main()
