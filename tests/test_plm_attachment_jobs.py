# -*- coding: utf-8 -*-
import sys
import time
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
WEB_APP = ROOT / "web_app2"
TESTS = ROOT / "tests"
if str(TESTS) not in sys.path:
    sys.path.insert(0, str(TESTS))
from test_env import configure_test_environment  # noqa: E402

configure_test_environment()
if str(WEB_APP) not in sys.path:
    sys.path.insert(0, str(WEB_APP))

from app import app  # noqa: E402


class PlmAttachmentJobTests(unittest.TestCase):
    def test_plm_attachment_job_reports_progress_and_download(self):
        from plm import _PLM_ATTACHMENT_JOBS

        out_path = WEB_APP / "outputs" / "mock_plm_attachment.zip"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(b"PK\x05\x06" + b"\x00" * 18)

        def fake_run(playwright, *, username, password, hqpn, output_dir, headless=False, log=None):
            log("Open PLM search page")
            log("Downloaded selected attachments: document_contents.zip")
            return out_path

        class FakePlaywrightContext:
            def __enter__(self):
                return object()

            def __exit__(self, exc_type, exc, tb):
                return False

        with app.test_client() as client:
            client.post("/api/login", json={"employee_id": "ADMIN"})
            with patch("plm.automation.run_hq_attachment_download", fake_run, create=True), patch(
                "playwright.sync_api.sync_playwright", return_value=FakePlaywrightContext()
            ):
                resp = client.post(
                    "/api/plm/auto_hq_attachments",
                    data={"username": "100448405", "password": "pw", "hqpn": "HQTEST"},
                )
                payload = resp.get_json()
                self.assertTrue(payload["success"])
                job_id = payload["job_id"]

                for _ in range(20):
                    status = client.get(f"/api/plm/auto_hq_attachments/status/{job_id}").get_json()
                    if status.get("status") == "done":
                        break
                    time.sleep(0.05)

                self.assertEqual(status["status"], "done")
                self.assertEqual(status["progress"], 100)
                self.assertEqual(status["download"], "/download/mock_plm_attachment.zip")
                self.assertIn("Downloaded selected attachments", status["log"])

            _PLM_ATTACHMENT_JOBS.pop(job_id, None)


if __name__ == "__main__":
    unittest.main()
