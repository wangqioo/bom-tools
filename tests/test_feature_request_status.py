import sys
import unittest
from pathlib import Path


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


class FeatureRequestStatusTests(unittest.TestCase):
    def test_bug_report_status_and_filters_work(self):
        client = app.test_client()
        data = {
            "reporter": "Bug Filter Tester",
            "employee_id": "100010",
            "module": "BOM 格式转换",
            "severity": "严重",
            "title": "Unique bug filter title",
            "description": "Unique bug filter description",
        }
        create_resp = client.post("/api/bug_reports", data=data)
        created = create_resp.get_json()["report"]

        update_resp = client.post(
            f"/api/bug_reports/{created['id']}/status",
            json={"status": "\u5904\u7406\u4e2d"},
        )
        payload = update_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["report"]["status"], "\u5904\u7406\u4e2d")

        filtered = client.get(
            "/api/bug_reports?status=%E5%A4%84%E7%90%86%E4%B8%AD&q=Unique%20bug%20filter"
        ).get_json()["reports"]
        self.assertTrue(any(item["id"] == created["id"] for item in filtered))

        excluded = client.get(
            "/api/bug_reports?status=%E5%B7%B2%E4%BF%AE%E5%A4%8D&q=Unique%20bug%20filter"
        ).get_json()["reports"]
        self.assertFalse(any(item["id"] == created["id"] for item in excluded))

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

    def test_feature_request_like_is_deduplicated_by_employee_id(self):
        data = {
            "requester": "Like Tester",
            "employee_id": "100008",
            "title": "Deduplicate likes",
            "requirement": "The same employee should only like once.",
        }
        client = app.test_client()
        create_resp = client.post("/api/feature_requests", data=data)
        created = create_resp.get_json()["request"]

        first_resp = client.post(
            f"/api/feature_requests/{created['id']}/like",
            json={"employee_id": "100008"},
        )
        first_payload = first_resp.get_json()
        self.assertTrue(first_payload["success"])
        self.assertEqual(first_payload["request"]["likes"], 1)
        self.assertFalse(first_payload.get("already_liked", False))

        second_resp = client.post(
            f"/api/feature_requests/{created['id']}/like",
            json={"employee_id": "100008"},
        )
        second_payload = second_resp.get_json()
        self.assertTrue(second_payload["success"])
        self.assertTrue(second_payload["already_liked"])
        self.assertEqual(second_payload["request"]["likes"], 1)

        third_resp = client.post(
            f"/api/feature_requests/{created['id']}/like",
            json={"employee_id": "100009"},
        )
        self.assertEqual(third_resp.get_json()["request"]["likes"], 2)

    def test_feature_request_filters_and_likes_sort_work(self):
        client = app.test_client()
        low_resp = client.post("/api/feature_requests", data={
            "requester": "Filter Tester",
            "employee_id": "100011",
            "module": "平台通用",
            "title": "Unique filter low likes",
            "requirement": "Unique feature filter group",
        })
        high_resp = client.post("/api/feature_requests", data={
            "requester": "Filter Tester",
            "employee_id": "100012",
            "module": "平台通用",
            "title": "Unique filter high likes",
            "requirement": "Unique feature filter group",
        })
        low = low_resp.get_json()["request"]
        high = high_resp.get_json()["request"]

        client.post(
            f"/api/feature_requests/{low['id']}/status",
            json={"status": "\u5df2\u7eb3\u5165"},
        )
        client.post(
            f"/api/feature_requests/{high['id']}/status",
            json={"status": "\u5df2\u7eb3\u5165"},
        )
        client.post(f"/api/feature_requests/{low['id']}/like", json={"employee_id": "200001"})
        client.post(f"/api/feature_requests/{high['id']}/like", json={"employee_id": "200002"})
        client.post(f"/api/feature_requests/{high['id']}/like", json={"employee_id": "200003"})

        payload = client.get(
            "/api/feature_requests?status=%E5%B7%B2%E7%BA%B3%E5%85%A5"
            "&module=%E5%B9%B3%E5%8F%B0%E9%80%9A%E7%94%A8&q=Unique%20feature%20filter&sort=likes"
        ).get_json()
        ids = [item["id"] for item in payload["requests"]]
        self.assertIn(low["id"], ids)
        self.assertIn(high["id"], ids)
        self.assertLess(ids.index(high["id"]), ids.index(low["id"]))


if __name__ == "__main__":
    unittest.main()
