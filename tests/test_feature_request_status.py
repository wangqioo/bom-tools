import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WEB_APP = ROOT / "web_app2"
if str(WEB_APP) not in sys.path:
    sys.path.insert(0, str(WEB_APP))

from app import app  # noqa: E402


class FeatureRequestStatusTests(unittest.TestCase):
    def test_feature_request_status_can_be_updated(self):
        data = {
            "requester": "Status Tester",
            "employee_id": "100007",
            "title": "Status transition request",
            "requirement": "Verify feature request status transitions.",
        }
        create_resp = app.test_client().post("/api/feature_requests", data=data)
        created = create_resp.get_json()["request"]

        update_resp = app.test_client().post(
            f"/api/feature_requests/{created['id']}/status",
            json={"status": "\u5f00\u53d1\u4e2d"},
        )
        payload = update_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["request"]["status"], "\u5f00\u53d1\u4e2d")

        list_payload = app.test_client().get("/api/feature_requests").get_json()
        updated = next(item for item in list_payload["requests"] if item["id"] == created["id"])
        self.assertEqual(updated["status"], "\u5f00\u53d1\u4e2d")

        invalid_resp = app.test_client().post(
            f"/api/feature_requests/{created['id']}/status",
            json={"status": "Unexpected"},
        )
        self.assertFalse(invalid_resp.get_json()["success"])

        missing_resp = app.test_client().post(
            "/api/feature_requests/not-exists/status",
            json={"status": "\u5df2\u5b8c\u6210"},
        )
        self.assertFalse(missing_resp.get_json()["success"])


if __name__ == "__main__":
    unittest.main()
