import io
import json
import sys
import unittest
import uuid
from pathlib import Path
from urllib.parse import unquote

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
WEB_APP = ROOT / "web_app2"
if str(WEB_APP) not in sys.path:
    sys.path.insert(0, str(WEB_APP))

from app import app  # noqa: E402
from feishu import _is_preferred_level, _write_cache  # noqa: E402
from manufacturer_alias import normalize_manufacturer_name  # noqa: E402


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

    def test_bom_detect_returns_fifty_preview_rows(self):
        rows = [["PN-%02d" % i, "Model-%02d" % i] for i in range(1, 56)]
        data = {
            "file": (_xlsx_bytes(["PN", "Model"], rows), "preview.xlsx"),
            "header_row": "1",
        }
        resp = app.test_client().post(
            "/api/bom/detect",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(len(payload["preview"]), 50)
        self.assertEqual(payload["preview"][0], ["PN-01", "Model-01"])
        self.assertEqual(payload["preview"][-1], ["PN-50", "Model-50"])

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


    def test_feishu_match_deduplicates_identical_rows_and_merges_sources(self):
        key_a, _, _ = _write_cache("token-a", "sid-a", [
            ["PN", "HQ PN", "Model"],
            ["A", "HQ-1", "M1"],
            ["A", "HQ-1", "M1"],
            ["A", "HQ-2", "M2"],
        ])
        key_b, _, _ = _write_cache("token-b", "sid-b", [
            ["PN", "HQ PN", "Model"],
            ["A", "HQ-1", "M1"],
        ])
        fetch_map = [
            {"output": "HQ PN", "alias": "HQ PN"},
            {"output": "Model", "alias": "Model"},
        ]
        data = {
            "file": (_xlsx_bytes(["PN"], [["A"], ["B"]]), "local.xlsx"),
            "config": json.dumps({
                "sheet_name": "Sheet",
                "header_row": 1,
                "tables": [
                    {"name": "LibA", "token": "token-a", "sheets": [{
                        "enabled": True,
                        "sheet_id": "sid-a",
                        "sheet_name": "SheetA",
                        "local_key_names": ["PN"],
                        "feishu_key_names": ["PN"],
                        "fetch_col_map": fetch_map,
                        "cache_key": key_a,
                    }]},
                    {"name": "LibB", "token": "token-b", "sheets": [{
                        "enabled": True,
                        "sheet_id": "sid-b",
                        "sheet_name": "SheetB",
                        "local_key_names": ["PN"],
                        "feishu_key_names": ["PN"],
                        "fetch_col_map": fetch_map,
                        "cache_key": key_b,
                    }]},
                ],
            }),
        }
        resp = app.test_client().post(
            "/api/feishu/match",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["total"], 2)
        self.assertEqual(payload["matched"], 1)
        self.assertEqual(payload["unmatched"], 1)

        filename = unquote(payload["download"].split("/download/", 1)[1])
        wb = openpyxl.load_workbook(WEB_APP / "outputs" / filename, data_only=True)
        ws = wb.active
        self.assertEqual([cell.value for cell in ws[1]], ["PN", "HQ PN", "Model", "来源表格"])
        self.assertEqual([ws["A2"].value, ws["B2"].value, ws["C2"].value], ["A", "HQ-1", "M1"])
        self.assertEqual(ws["D2"].value, "LibA - SheetA；LibB - SheetB")
        self.assertEqual([ws["A3"].value, ws["B3"].value, ws["C3"].value, ws["D3"].value], [None, "HQ-2", "M2", "LibA - SheetA"])
        self.assertEqual([ws["A4"].value, ws["B4"].value, ws["C4"].value, ws["D4"].value], ["B", None, None, "未匹配"])
        wb.close()


    def test_manufacturer_alias_create_lookup_duplicate_and_delete(self):
        suffix = uuid.uuid4().hex[:8]
        canonical = f"HQ Maker {suffix}"
        alias = f"Maker-{suffix}"
        client = app.test_client()

        create_resp = client.post("/api/manufacturer_aliases", data={
            "canonical_name": canonical,
            "alias": alias,
            "source": "unit-test",
            "note": "case and punctuation normalization",
        })
        created_payload = create_resp.get_json()
        self.assertTrue(created_payload["success"])
        created = created_payload["alias"]
        self.assertEqual(created["canonical_name"], canonical)
        self.assertEqual(created["normalized_alias"], normalize_manufacturer_name(alias))

        lookup_resp = client.get(f"/api/manufacturer_aliases/lookup?name=maker_{suffix.upper()}")
        lookup_payload = lookup_resp.get_json()
        self.assertTrue(lookup_payload["success"])
        self.assertEqual(lookup_payload["match"]["canonical_name"], canonical)

        duplicate_resp = client.post("/api/manufacturer_aliases", data={
            "canonical_name": "Other HQ Name",
            "alias": f" maker.{suffix} ",
        })
        duplicate_payload = duplicate_resp.get_json()
        self.assertFalse(duplicate_payload["success"])
        self.assertEqual(duplicate_payload["existing"]["id"], created["id"])

        list_resp = client.get(f"/api/manufacturer_aliases?q={alias}")
        list_payload = list_resp.get_json()
        self.assertTrue(list_payload["success"])
        self.assertEqual(list_payload["match"]["id"], created["id"])
        self.assertTrue(any(item["id"] == created["id"] for item in list_payload["aliases"]))

        delete_resp = client.delete(f"/api/manufacturer_aliases/{created['id']}")
        self.assertTrue(delete_resp.get_json()["success"])
        missing_lookup = client.get(f"/api/manufacturer_aliases/lookup?name={alias}").get_json()
        self.assertIsNone(missing_lookup["match"])

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
        self.assertEqual(detail["C1"].value, "基准版本行号")
        self.assertEqual(detail["D1"].value, "对比版本行号")
        self.assertEqual(detail["E1"].value, "变更字段")
        self.assertEqual(detail["F1"].value, "基准值")
        self.assertEqual(detail["G1"].value, "对比值")
        wb.close()

    def test_hq_bom_version_compare_reports_refdes_delta_as_part_change(self):
        base_bom = _hq_export_bytes(
            [["1", "A", "M1", "DESC-A", 3, "", "R1,R2,R3", "厂商A"]]
        )
        compare_bom = _hq_export_bytes(
            [["1", "A", "M1", "DESC-A", 3, "", "R1,R3,R4", "厂商A"]]
        )
        data = {
            "old_file": (base_bom, "base.xlsx"),
            "new_file": (compare_bom, "compare.xlsx"),
            "config": json.dumps({
                "header_row": 3,
                "key_col": "料号",
                "compare_cols": ["型号", "单耗", "位号"],
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
        self.assertEqual(payload["added"], 0)
        self.assertEqual(payload["removed"], 0)
        self.assertEqual(payload["changed"], 1)
        self.assertEqual(payload["unchanged"], 0)

        filename = unquote(payload["download"].split("/download/", 1)[1])
        report_path = WEB_APP / "outputs" / filename
        wb = openpyxl.load_workbook(report_path, data_only=True)
        detail = wb["差异明细"]
        self.assertEqual(detail["C1"].value, "基准版本行号")
        self.assertEqual(detail["D1"].value, "对比版本行号")
        changed_rows = [row for row in detail.iter_rows(min_row=2, values_only=True) if row[0] == "变更"]
        self.assertEqual(len(changed_rows), 1)
        self.assertEqual(changed_rows[0][1], "A")
        self.assertEqual(changed_rows[0][4], "位号")
        self.assertEqual(changed_rows[0][5], "R2")
        self.assertEqual(changed_rows[0][6], "R4")
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


    def test_generic_customer_hq_compare_reports_differences(self):
        customer = _xlsx_bytes(
            ["客户料号", "型号", "数量"],
            [["A", "R1", 1], ["B", "C1", 2], ["C", "L1", 3]],
        )
        hq = _xlsx_bytes(
            ["料号", "型号", "数量"],
            [["A", "R1", 1], ["B", "C2", 2], ["D", "L2", 4]],
        )
        data = {
            "left_file": (customer, "customer.xlsx"),
            "right_file": (hq, "hq.xlsx"),
            "config": json.dumps({
                "compare_type": "customer_hq",
                "left_header_row": 1,
                "right_header_row": 1,
                "left_key_col": "客户料号",
                "right_key_col": "料号",
                "field_pairs": [{"left": "型号", "right": "型号"}, {"left": "数量", "right": "数量"}],
            }),
        }
        resp = app.test_client().post(
            "/api/bom_compare/generic",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["left_only"], 1)
        self.assertEqual(payload["right_only"], 1)
        self.assertEqual(payload["changed"], 1)
        self.assertEqual(payload["same"], 1)
        filename = unquote(payload["download"].split("/download/", 1)[1])
        wb = openpyxl.load_workbook(WEB_APP / "outputs" / filename, data_only=True)
        self.assertIn("差异明细", wb.sheetnames)
        detail = wb["差异明细"]
        self.assertEqual(detail["A1"].value, "差异类型")
        self.assertIn("客户BOM型号", [cell.value for cell in detail[1]])
        wb.close()

    def test_generic_sheets_detects_left_and_right_keys(self):
        cadence = _xlsx_bytes(["REFDES", "PART_NUMBER", "QTY"], [["R1", "A", 1]])
        hq = _xlsx_bytes(["位号", "料号", "单耗"], [["R1", "A", 1]])
        data = {
            "left_file": (cadence, "cadence.xlsx"),
            "right_file": (hq, "hq.xlsx"),
            "left_header_row": "1",
            "right_header_row": "1",
        }
        resp = app.test_client().post(
            "/api/bom_compare/generic_sheets",
            data=data,
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["detected_left_key"], "REFDES")
        self.assertEqual(payload["detected_right_key"], "位号")

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
        self.assertTrue(any(item["id"] == "seed-bug-bom-header-detect" for item in list_payload["reports"]))

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


    def test_bug_report_status_can_be_updated(self):
        data = {
            "reporter": "孙八",
            "employee_id": "100006",
            "title": "状态测试",
            "description": "用于验证状态修改",
        }
        create_resp = app.test_client().post("/api/bug_reports", data=data)
        created = create_resp.get_json()["report"]

        update_resp = app.test_client().post(
            f"/api/bug_reports/{created['id']}/status",
            json={"status": "处理中"},
        )
        payload = update_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["report"]["status"], "处理中")

        invalid_resp = app.test_client().post(
            f"/api/bug_reports/{created['id']}/status",
            json={"status": "随意状态"},
        )
        self.assertFalse(invalid_resp.get_json()["success"])


    def test_feature_request_submit_and_list(self):
        data = {
            "requester": "赵六",
            "employee_id": "100004",
            "module": "BOM比对工具合集",
            "priority": "较高",
            "request_type": "新功能",
            "title": "增加导出汇总",
            "background": "需要快速查看差异总数",
            "requirement": "生成差异汇总页",
            "value": "减少手工统计",
            "acceptance": "导出文件包含汇总页",
        }
        resp = app.test_client().post("/api/feature_requests", data=data)
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["request"]["requester"], "赵六")
        self.assertEqual(payload["request"]["status"], "待评估")
        self.assertEqual(payload["request"]["likes"], 0)

        list_resp = app.test_client().get("/api/feature_requests")
        list_payload = list_resp.get_json()
        self.assertTrue(list_payload["success"])
        self.assertTrue(any(item["id"] == payload["request"]["id"] for item in list_payload["requests"]))
        self.assertTrue(any(item["id"] == "seed-bom-compare-summary" for item in list_payload["requests"]))

    def test_feature_request_rejects_missing_required_fields(self):
        resp = app.test_client().post("/api/feature_requests", data={"requester": "赵六"})
        payload = resp.get_json()
        self.assertFalse(payload["success"])
        self.assertIn("需求标题", payload["error"])


    def test_feature_request_like_increments_count(self):
        data = {
            "requester": "钱七",
            "employee_id": "100005",
            "title": "点赞测试需求",
            "requirement": "需要可以点赞",
        }
        create_resp = app.test_client().post("/api/feature_requests", data=data)
        created = create_resp.get_json()["request"]

        like_resp = app.test_client().post(f"/api/feature_requests/{created['id']}/like")
        payload = like_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["request"]["likes"], 1)

        second_resp = app.test_client().post(f"/api/feature_requests/{created['id']}/like")
        self.assertEqual(second_resp.get_json()["request"]["likes"], 2)


if __name__ == "__main__":
    unittest.main()
