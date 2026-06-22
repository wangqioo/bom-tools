import io
import sys
import unittest
import uuid
from pathlib import Path

import openpyxl


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


def _xlsx_bytes(headers, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class AuthTests(unittest.TestCase):
    def setUp(self):
        self.old_auth_required = app.config.get("AUTH_REQUIRED")
        app.config["AUTH_REQUIRED"] = True

    def tearDown(self):
        app.config["AUTH_REQUIRED"] = self.old_auth_required

    def test_index_redirects_to_login_when_not_authenticated(self):
        resp = app.test_client().get("/")
        self.assertEqual(resp.status_code, 302)
        self.assertIn("/login", resp.headers["Location"])

    def test_api_requires_login_and_admin_can_login_with_employee_id(self):
        client = app.test_client()
        blocked = client.get("/api/bug_reports")
        self.assertEqual(blocked.status_code, 401)
        self.assertFalse(blocked.get_json()["success"])

        login = client.post("/api/login", json={"employee_id": "ADMIN"})
        payload = login.get_json()
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["user"]["employee_id"], "ADMIN")
        self.assertEqual(payload["user"]["role"], "admin")

        resp = client.get("/")
        self.assertEqual(resp.status_code, 200)
        self.assertIn("BOM Tools", resp.get_data(as_text=True))

    def test_unknown_employee_id_requires_name_then_registers_as_normal_user(self):
        employee_id = "U" + uuid.uuid4().hex[:10]
        client = app.test_client()
        missing_name_resp = client.post("/api/login", json={"employee_id": employee_id})
        missing_name = missing_name_resp.get_json()
        self.assertEqual(missing_name_resp.status_code, 409)
        self.assertFalse(missing_name["success"])
        self.assertTrue(missing_name["need_name"])

        first = client.post("/api/login", json={"employee_id": employee_id, "display_name": "张三"}).get_json()
        self.assertTrue(first["success"], first.get("error"))
        self.assertTrue(first["created"])
        self.assertEqual(first["user"]["employee_id"], employee_id.upper())
        self.assertEqual(first["user"]["display_name"], "张三")
        self.assertEqual(first["user"]["role"], "user")

        client.post("/api/logout")
        second = client.post("/api/login", json={"employee_id": employee_id}).get_json()
        self.assertTrue(second["success"], second.get("error"))
        self.assertFalse(second["created"])
        self.assertEqual(second["user"]["display_name"], "张三")
        self.assertEqual(second["user"]["role"], "user")

    def test_numeric_and_wb_employee_ids_are_allowed(self):
        client = app.test_client()
        numeric = "10" + uuid.uuid4().hex[:8]
        wb_id = "WB" + uuid.uuid4().hex[:8]
        for employee_id in (numeric, wb_id):
            payload = client.post(
                "/api/login",
                json={"employee_id": employee_id, "display_name": "工号格式测试"},
            ).get_json()
            self.assertTrue(payload["success"], payload.get("error"))
            self.assertEqual(payload["user"]["employee_id"], employee_id.upper())
            client.post("/api/logout")

    def test_status_update_requires_admin_role(self):
        client = app.test_client()
        login = client.post("/api/login", json={"employee_id": "U" + uuid.uuid4().hex[:10], "display_name": "普通用户"})
        self.assertTrue(login.get_json()["success"])

        resp = client.post(
            "/api/bug_reports/seed-bug-bom-header-detect/status",
            json={"status": "\u5904\u7406\u4e2d"},
        )
        self.assertEqual(resp.status_code, 403)
        self.assertIn("\u7ba1\u7406\u5458", resp.get_json()["error"])

    def test_admin_can_list_and_manage_users(self):
        employee_id = "U" + uuid.uuid4().hex[:10]
        user_client = app.test_client()
        created = user_client.post(
            "/api/login",
            json={"employee_id": employee_id, "display_name": "管理测试用户"},
        ).get_json()["user"]

        denied = user_client.get("/api/admin/users")
        self.assertEqual(denied.status_code, 403)

        admin_client = app.test_client()
        admin_login = admin_client.post("/api/login", json={"employee_id": "ADMIN"}).get_json()
        self.assertTrue(admin_login["success"])

        list_payload = admin_client.get(f"/api/admin/users?q={employee_id}").get_json()
        self.assertTrue(list_payload["success"])
        self.assertEqual(list_payload["summary"]["total"], 1)
        listed = list_payload["users"][0]
        self.assertEqual(listed["employee_id"], employee_id.upper())
        self.assertEqual(listed["display_name"], "管理测试用户")
        self.assertGreaterEqual(listed["login_count"], 1)

        role_payload = admin_client.post(
            f"/api/admin/users/{created['id']}/role",
            json={"role": "admin"},
        ).get_json()
        self.assertTrue(role_payload["success"], role_payload.get("error"))
        self.assertEqual(role_payload["user"]["role"], "admin")

        active_payload = admin_client.post(
            f"/api/admin/users/{created['id']}/active",
            json={"is_active": False},
        ).get_json()
        self.assertTrue(active_payload["success"], active_payload.get("error"))
        self.assertFalse(active_payload["user"]["is_active"])

        activity_payload = admin_client.get("/api/admin/activity?limit=10").get_json()
        self.assertTrue(activity_payload["success"])
        self.assertTrue(any(item["action"] == "admin_update_user_active" for item in activity_payload["activities"]))
        refreshed = admin_client.get(f"/api/admin/users?q={employee_id}").get_json()["users"][0]
        self.assertEqual(refreshed["activity_count"], listed["login_count"])
        self.assertEqual(refreshed["tool_run_count"], 0)
        self.assertEqual(refreshed["tool_export_count"], 0)
        self.assertNotIn("bug_submit_count", refreshed)
        self.assertNotIn("feature_submit_count", refreshed)
        self.assertNotIn("feature_like_count", refreshed)
        self.assertNotIn("status_update_count", refreshed)

    def test_tool_export_activity_records_tool_and_filename(self):
        client = app.test_client()
        employee_id = "WB" + uuid.uuid4().hex[:8]
        login = client.post(
            "/api/login",
            json={"employee_id": employee_id, "display_name": "工具日志用户"},
        ).get_json()
        self.assertTrue(login["success"])
        resp = client.post("/api/bom/convert", data={
            "file": (_xlsx_bytes(["厂家", "型号", "用量"], [["MakerA", "M1", 2]]), "bom.xlsx"),
            "sheet": "Sheet",
            "header_row": "1",
            "fmt": "B",
            "col_brand": "A",
            "col_model": "B",
            "col_qty": "C",
            "output_mode": "expand",
        }, content_type="multipart/form-data")
        payload = resp.get_json()
        self.assertTrue(payload["success"], payload.get("error"))

        admin = app.test_client()
        self.assertTrue(admin.post("/api/login", json={"employee_id": "ADMIN"}).get_json()["success"])
        activity = admin.get("/api/admin/activity?limit=20").get_json()["activities"]
        exports = [item for item in activity if item["employee_id"] == employee_id.upper() and item["action"] == "tool_export"]
        self.assertTrue(exports)
        detail = exports[0]["detail"]
        self.assertEqual(detail["tool"], "BOM格式转换")
        self.assertTrue(detail["filename"].endswith(".xlsx"))


if __name__ == "__main__":
    unittest.main()

