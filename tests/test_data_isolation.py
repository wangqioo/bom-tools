import os
import sys
import unittest
import uuid
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
from feishu import _hq_read_sheet, _read_cache, _write_cache  # noqa: E402
from shared import OUTPUT_DIR, _record_file_owner  # noqa: E402


def _login(client, employee_id, name):
    response = client.post("/api/login", json={"employee_id": employee_id, "display_name": name})
    payload = response.get_json()
    if not payload.get("success"):
        raise AssertionError(payload)
    return payload["user"]


class DataIsolationTests(unittest.TestCase):
    def setUp(self):
        self.old_auth_required = app.config.get("AUTH_REQUIRED")
        app.config["AUTH_REQUIRED"] = True

    def tearDown(self):
        app.config["AUTH_REQUIRED"] = self.old_auth_required

    def test_output_download_is_restricted_to_the_owner(self):
        user_a_id = "UA" + uuid.uuid4().hex[:10]
        user_b_id = "UB" + uuid.uuid4().hex[:10]
        filename = f"isolation_{uuid.uuid4().hex}.xlsx"
        path = Path(OUTPUT_DIR) / filename
        path.write_bytes(b"private export")
        try:
            client_a = app.test_client()
            user_a = _login(client_a, user_a_id, "Owner A")
            _record_file_owner(str(path), user_a["id"])
            owner_download = client_a.get(f"/download/{filename}")
            self.assertEqual(owner_download.status_code, 200)
            owner_download.close()

            client_b = app.test_client()
            _login(client_b, user_b_id, "Owner B")
            self.assertEqual(client_b.get(f"/download/{filename}").status_code, 404)
        finally:
            for candidate in (path, Path(f"{path}.owner")):
                try:
                    candidate.unlink()
                except FileNotFoundError:
                    pass

    def test_feishu_cache_cannot_be_read_by_another_user(self):
        user_a_id = "CA" + uuid.uuid4().hex[:10]
        user_b_id = "CB" + uuid.uuid4().hex[:10]
        token = f"token-{uuid.uuid4().hex}"
        sheet_id = f"sheet-{uuid.uuid4().hex}"
        client_a = app.test_client()
        client_b = app.test_client()
        with client_a:
            _login(client_a, user_a_id, "Cache A")
            key, _, _ = _write_cache(token, sheet_id, [["PN"], ["A"]])
            self.assertIsNotNone(_read_cache(key))
        with client_b:
            _login(client_b, user_b_id, "Cache B")
            self.assertIsNone(_read_cache(key))
        cache_path = WEB_APP / "cache" / f"feishu_{key}.json"
        try:
            cache_path.unlink()
        except FileNotFoundError:
            pass

    def test_second_feishu_block_keeps_its_first_data_row(self):
        class FakeResponse:
            def __init__(self, values):
                self.values = values

            def raise_for_status(self):
                return None

            def json(self):
                return {"code": 0, "data": {"valueRange": {"values": self.values}}}

        responses = {
            "sheet!A1:Z1": [["PN"]],
            "sheet!A1:A3": [["PN"], ["A"], ["B"]],
            "sheet!A4:A6": [["C"], ["D"], ["E"]],
        }

        def fake_get(url, params=None, timeout=None):
            return FakeResponse(responses.get((params or {}).get("range"), []))

        with patch("feishu._requests.get", side_effect=fake_get):
            rows = _hq_read_sheet("https://example.test", "origin", "user", "token", "sheet", row_count=6, col_count=26, batch_size=3)

        self.assertEqual(rows, [["PN"], ["A"], ["B"], ["C"], ["D"], ["E"]])

    def test_non_admin_cannot_change_global_manufacturer_aliases(self):
        client = app.test_client()
        _login(client, "MA" + uuid.uuid4().hex[:10], "Normal User")
        response = client.post(
            "/api/manufacturer_aliases",
            data={"canonical_name": "Maker", "alias": "Alias"},
        )
        self.assertEqual(response.status_code, 403)
