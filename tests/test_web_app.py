import io
import json
import sys
import unittest
from pathlib import Path
from urllib.parse import unquote

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
WEB_APP = ROOT / "web_app2"
if str(WEB_APP) not in sys.path:
    sys.path.insert(0, str(WEB_APP))

from app import app  # noqa: E402
from feishu import _is_preferred_level  # noqa: E402


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


def _hq_export_bytes(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.append(["料号", "PROJECT", "描述", "DESC", "项目配置名", "CFG"])
    ws.append(["版本", "I.1", "替代项", "", "BOM名称", "BOM"])
    ws.append(["序号", "料号", "型号", "物料描述", "单耗", "替代关系", "位号", "生产厂家"])
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class WebAppTests(unittest.TestCase):
    def test_index_renders_main_shell(self):
        resp = app.test_client().get("/")
        self.assertEqual(resp.status_code, 200)
        html = resp.get_data(as_text=True)
        self.assertIn("BOM Tools", html)
        self.assertIn("飞书优选库+关系库匹配", html)
        self.assertIn("BOM比对工具合集", html)
        self.assertIn("Bug提交栏目", html)
        self.assertIn("客户BOM对比HQ BOM", html)
        self.assertIn("同项目HQ BOM版本对比", html)
        self.assertIn("Cadence导出BOM对比HQ BOM", html)

    def test_download_rejects_missing_and_traversal_paths(self):
        client = app.test_client()
        self.assertEqual(client.get("/download/not-found.xlsx").status_code, 404)
        self.assertEqual(client.get("/download/../app.py").status_code, 404)

    def test_preferred_level_does_not_treat_non_preferred_as_preferred(self):
        self.assertTrue(_is_preferred_level("优选"))
        self.assertTrue(_is_preferred_level("7"))
        self.assertTrue(_is_preferred_level("Preferred"))
        self.assertFalse(_is_preferred_level("非优选"))
        self.assertFalse(_is_preferred_level("不优选"))
        self.assertFalse(_is_preferred_level("6"))
        self.assertFalse(_is_preferred_level(""))

    def test_pref_rate_rejects_invalid_header_row_before_excel_processing(self):
        data = {
            "file": (_xlsx_bytes(["HQ料号"], [["HQ-1"]]), "bom.xlsx"),
            "config": json.dumps({"header_row": "bad", "local_key_col": "HQ料号"}),
        }
        resp = app.test_client().post(
            "/api/feishu/pref_rate",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertEqual(resp.status_code, 200)
        self.assertFalse(payload["success"])
        self.assertIn("表头行", payload["error"])

    def test_hq_bom_version_compare_reports_added_removed_and_changed(self):
        base_bom = _hq_export_bytes(
            [["1", "A", "R1", "DESC-A", 1, "", "R1", "厂商A"],
             ["2", "B", "C1", "DESC-B", 2, "", "C1", "厂商B"],
             ["3", "C", "L1", "DESC-C", 3, "", "L1", "厂商C"]]
        )
        compare_bom = _hq_export_bytes(
            [["1", "A", "R1", "DESC-A", 1, "", "R1", "厂商A"],
             ["2", "B", "C2", "DESC-B", 2, "", "C1", "厂商B"],
             ["4", "D", "L2", "DESC-D", 4, "", "L2", "厂商D"]]
        )
        data = {
            "old_file": (base_bom, "base.xlsx"),
            "new_file": (compare_bom, "compare.xlsx"),
            "config": json.dumps({
                "header_row": 3,
                "key_col": "料号",
                "compare_cols": ["型号", "单耗"],
            }),
        }
        resp = app.test_client().post(
            "/api/bom_compare/hq_version",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(payload["success"])
        self.assertEqual(payload["added"], 1)
        self.assertEqual(payload["removed"], 1)
        self.assertEqual(payload["changed"], 1)
        self.assertEqual(payload["unchanged"], 1)
        self.assertTrue(payload["download"].startswith("/download/"))
        filename = unquote(payload["download"].split("/download/", 1)[1])
        report_path = WEB_APP / "outputs" / filename
        wb = openpyxl.load_workbook(report_path, data_only=True)
        self.assertEqual(
            wb.sheetnames,
            ["差异总览", "差异明细", "新增物料", "删除物料", "变更物料", "重复料号"],
        )
        detail = wb["差异明细"]
        self.assertEqual(detail["A1"].value, "差异类型")
        self.assertEqual(detail["B1"].value, "料号")
        self.assertIn("基准版本型号", [cell.value for cell in detail[1]])
        wb.close()

    def test_hq_bom_export_format_defaults_to_header_row_three(self):
        data = {
            "file": (_hq_export_bytes([["1", "HQ1", "M1", "DESC1", 1, "", "R1", "厂商"]]), "hq.xlsx"),
            "header_row": "3",
        }
        resp = app.test_client().post(
            "/api/bom_compare/local_sheets",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["current_sheet"], "BOM")
        self.assertEqual(payload["detected_key"], "料号")
        self.assertIn("单耗", payload["headers"])

    def test_hq_bom_compare_rejects_non_standard_format(self):
        data = {
            "file": (_xlsx_bytes(["HQ料号", "规格型号", "单耗"], [["A", "R1", 1]]), "not-standard.xlsx"),
            "header_row": "1",
        }
        resp = app.test_client().post(
            "/api/bom_compare/local_sheets",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertFalse(payload["success"])
        self.assertIn("不支持当前文件格式", payload["error"])

    def test_bug_report_submit_and_list(self):
        data = {
            "reporter": "张三",
            "employee_id": "100001",
            "module": "BOM比对工具合集",
            "severity": "一般",
            "title": "测试问题",
            "description": "这里是问题描述",
            "steps": "1. 打开页面",
            "expected": "可以正常使用",
        }
        resp = app.test_client().post("/api/bug_reports", data=data)
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["report"]["reporter"], "张三")

        list_resp = app.test_client().get("/api/bug_reports")
        list_payload = list_resp.get_json()
        self.assertTrue(list_payload["success"])
        self.assertTrue(any(item["id"] == payload["report"]["id"] for item in list_payload["reports"]))

    def test_bug_report_accepts_excel_attachment(self):
        data = {
            "reporter": "李四",
            "employee_id": "100002",
            "title": "附件测试",
            "description": "上传 Excel 附件",
            "images": (_xlsx_bytes(["A"], [["B"]]), "case.xlsx"),
        }
        resp = app.test_client().post(
            "/api/bug_reports",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["report"]["attachments"][0]["name"], "case.xlsx")

    def test_bug_report_rejects_unsupported_attachment(self):
        data = {
            "reporter": "王五",
            "employee_id": "100003",
            "title": "附件测试",
            "description": "上传不支持附件",
            "images": (io.BytesIO(b"binary"), "tool.exe"),
        }
        resp = app.test_client().post(
            "/api/bug_reports",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertFalse(payload["success"])
        self.assertIn("仅支持上传", payload["error"])


if __name__ == "__main__":
    unittest.main()
