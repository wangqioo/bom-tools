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


def _plm_full_bom_bytes(sheet_rows):
    bom = 'BOM'
    dbg = 'DBG\u4e1a\u52a1BOM'
    ctrl = 'DBGBOM\u5236\u63a7\u4fe1\u606f'
    base_headers = [
        '\u5e8f\u53f7', '\u6599\u53f7', '\u4e0a\u9636BOM\u540d\u79f0', 'BOM\u5c42\u7ea7', '\u578b\u53f7', '\u7269\u6599\u63cf\u8ff0', '\u5355\u8017', '\u66ff\u4ee3\u5173\u7cfb', '\u4f4d\u53f7', '\u751f\u4ea7\u5382\u5bb6',
        '\u662f\u5426\u73af\u4fdd', '\u6e7f\u654f\u5c5e\u6027', '\u7269\u6599\u66f4\u65b0\u65f6\u95f4', 'MOQ', '\u63d0\u524d\u671f', 'ECO_NO', 'ECN\u751f\u6548\u65f6\u95f4', '\u8d28\u91cf\u6807\u51c6',
        '\u4e34\u65f6\u6280\u672f\u5c01\u6837\u65e5\u671f', '\u5b9e\u9645\u6280\u672f\u5c01\u6837\u65e5\u671f', '\u5c01\u6837\u5c5e\u6027', '\u5c01\u6837\u5907\u6ce8', '\u5de5\u7a0b\u5e08\u5907\u6ce8', '\u73af\u4fdd\u8d44\u6599', '\u73af\u4fdd\u9f50\u5957',
        '\u5ba2\u6237\u96f6\u4ef6\u7f16\u7801', '\u5ba2\u6237\u96f6\u4ef6\u540d\u79f0', '\u5ba2\u6237\u751f\u4ea7\u5de5\u5382\u7269\u6599\u6599\u53f7', '\u539f\u4ea7\u56fd', '\u96f6\u4ef6\u5236\u9020\u5546', '\u96f6\u4ef6\u5236\u9020\u5546\u5730\u5740', '\u5355\u4f4d\u51c0\u91cd',
        '\u5ba2\u6237\u6599\u53f7', '\u7814\u53d1\u5c5e\u6027', '\u6700\u5c0f\u5305\u88c5(\u5173\u8054\u4e0b\u5355)', '\u6700\u5c0f\u5305\u88c5(\u4e0d\u5173\u8054\u4e0b\u5355)', '\u6700\u5c0f\u8bf7\u8d2d\u91cf(\u5173\u8054\u4e0b\u5355)', '\u6700\u5c0f\u8bf7\u8d2d\u91cf(\u4e0d\u5173\u8054\u4e0b\u5355)',
        '\u751f\u4ea7\u5382\u5bb6\uff08\u82f1\u6587\uff09', '\u751f\u4ea7\u5382\u5bb6\uff08\u4e2d\u6587\uff09', '\u8fdb\u53e3\u5546', '\u6e7f\u654f\u7b49\u7ea7', '\u5ba2\u6237\u96f6\u4ef6\u540d\u79f0(\u4e2d\u6587)', '\u7269\u6599\u63cf\u8ff0-\u82f1\u6587', '\u5236\u9020\u5546-\u82f1\u6587',
        '\u6e7f\u654f\u5c5e\u6027-\u82f1\u6587', '\u4e0b\u5355\u6599\u53f7', '\u8f6f\u4ef6\u7248\u672c', '\u786c\u4ef6\u7248\u672c', 'MODEL', '\u8d27\u8fd0\u7ec4\u7ec7', '\u4e3b\u8f85BOM\u5b9a\u4e49', '\u6781\u6027&1\u811a\u4f4d\u7f6e',
        '\u5668\u4ef6\u95f4\u8ddd(P1)', '\u8f7d\u5e26\u5bbd\u5ea6(W)', '\u6761\u7801\u89c4\u52191', '\u6761\u7801\u89c4\u52192', '\u6761\u7801\u89c4\u52193', 'HW\u5ba2\u6237\u7269\u6599\u5c5e\u6027', '\u82f1\u6587\u63cf\u8ff0', '\u4f9b\u5e94\u5546',
        '\u6570\u636e\u6cbb\u7406\u6807\u8bb0', 'MBG\u4f18\u9009\u5c5e\u6027', 'CBG\u4f18\u9009\u5c5e\u6027', 'DBG\u4f18\u9009\u5c5e\u6027', '\u91c7\u8d2d\u6a21\u5f0f', '\u662f\u5426\u53ef\u91cf\u4ea7\u4e0b\u5355', 'ABG\u4f18\u9009\u5c5e\u6027', 'IFM_PART', 'PCD_PART', 'ECCN',
    ]
    extras = {
        bom: [],
        dbg: ['\u5355\u4f4d\u603b\u7528\u91cf', '\u9996\u5236\u7a0b', '\u6b21\u5236\u7a0b', '\u6b21\u5236\u7a0b\u5355\u8017', '\u6b21\u5236\u7a0b\u4f4d\u53f7'],
        ctrl: ['\u5355\u4f4d\u603b\u7528\u91cf', '\u635f\u8017\u7387', '\u56fa\u5b9a\u5907\u635f\u503c', '\u9996\u5236\u7a0b', '\u6b21\u5236\u7a0b', '\u6b21\u5236\u7a0b\u5355\u8017', '\u6b21\u5236\u7a0b\u4f4d\u53f7'],
    }
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name in [bom, dbg, ctrl]:
        ws = wb.create_sheet(sheet_name)
        ws.append(['BOM\u5386\u53f2\u4fee\u6539\u8bb0\u5f55\u62a5\u8868'])
        ws.append([])
        ws.append(['\u63d0\u4ea4\u65f6\u95f4', '2026-05-13 14:50:41'])
        ws.append(['\u91cf\u4ea7/\u8bd5\u4ea7', '\u8bd5\u4ea7'])
        ws.append([None] * 12)
        ws.append(['\u6599\u53f7', 'HQ31200063SB0', '\u63cf\u8ff0', 'Demo', '\u9879\u76ee\u914d\u7f6e\u540d', 'NS8551AAA', '\u5de5\u7a0b\u5e08', 'Tester'])
        ws.append(['\u7248\u672c', 'I.4', '\u66ff\u4ee3\u9879', '', 'BOM\u540d\u79f0', 'Demo full BOM', '\u5f52\u6863\u90e8\u95e8', 'SYSTEMADMIN'])
        headers = base_headers + extras[sheet_name]
        ws.append(headers)
        for row in sheet_rows.get(sheet_name, sheet_rows[bom]):
            data = {'\u5e8f\u53f7': row[0], '\u6599\u53f7': row[1], '\u4e0a\u9636BOM\u540d\u79f0': 'HQ31200063SB0', 'BOM\u5c42\u7ea7': '1', '\u578b\u53f7': row[2], '\u7269\u6599\u63cf\u8ff0': row[3], '\u5355\u8017': row[4], '\u66ff\u4ee3\u5173\u7cfb': row[5], '\u4f4d\u53f7': row[6], '\u751f\u4ea7\u5382\u5bb6': row[7]}
            data.update(row[8] if len(row) > 8 else {})
            ws.append([data.get(h, '') for h in headers])
        ws.append([])
        ws.append(['BOM\u7248\u672c', '\u65e5\u671f', '\u4fee\u8ba2\u88c5\u914d\u4ef6', '\u4fee\u8ba2\u7ec4\u4ef6', '\u7ec4\u4ef6\u578b\u53f7', '\u7ec4\u4ef6\u63cf\u8ff0', '\u52a8\u4f5c\uff08\u65b0\u589e\uff0c\u66f4\u6539\uff0c\u7981\u7528\uff09', '\u4fee\u8ba2\u66ff\u4ee3\u4ef6', '\u66ff\u4ee3\u4ef6\u578b\u53f7', '\u66ff\u4ee3\u4ef6\u63cf\u8ff0', '\u52a8\u4f5c\uff08\u65b0\u589e\uff0c\u7981\u7528\uff09', '\u65e7\u6570\u91cf', '\u65b0\u6570\u91cf', '\u65e7\u4f4d\u53f7', '\u65b0\u4f4d\u53f7'])
        ws.append(['I.4', '2026-05-12', 'HQ31200063SB0', 'HISTORY', 'HistoryModel', 'HistoryDesc', '\u65b0\u589e\u7269\u6599', '', '', '', '', '', '1', '', 'H1'])
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

    def test_plm_full_bom_version_compare_reads_three_sheets_and_ignores_history(self):
        bom = "BOM"
        dbg = "DBG\u4e1a\u52a1BOM"
        ctrl = "DBGBOM\u5236\u63a7\u4fe1\u606f"
        main = "\u4e3b\u6599"
        maker_a = "\u5382\u5546A"
        maker_b = "\u5382\u5546B"
        maker_c = "\u5382\u5546C"
        total_qty = "\u5355\u4f4d\u603b\u7528\u91cf"
        first_process = "\u9996\u5236\u7a0b"
        loss = "\u635f\u8017\u7387"
        fixed_loss = "\u56fa\u5b9a\u5907\u635f\u503c"
        old_rows = {
            bom: [["1", "A", "M1", "DESC-A", 1, main, "R1", maker_a], ["2", "B", "M2", "DESC-B", 2, main, "R2", maker_b]],
            dbg: [["1", "A", "M1", "DESC-A", 1, main, "R1", maker_a, {total_qty: 1, first_process: "SMT"}], ["2", "B", "M2", "DESC-B", 2, main, "R2", maker_b, {total_qty: 2, first_process: "SMT"}]],
            ctrl: [["1", "A", "M1", "DESC-A", 1, main, "R1", maker_a, {total_qty: 1, loss: "0.005", fixed_loss: "50", first_process: "SMT"}], ["2", "B", "M2", "DESC-B", 2, main, "R2", maker_b, {total_qty: 2, loss: "0.005", fixed_loss: "50", first_process: "SMT"}]],
        }
        new_rows = {
            bom: [["1", "A", "M1-NEW", "DESC-A", 1, main, "R1", maker_a], ["3", "C", "M3", "DESC-C", 3, main, "R3", maker_c]],
            dbg: [["1", "A", "M1-NEW", "DESC-A", 1, main, "R1", maker_a, {total_qty: 1, first_process: "SMT"}], ["3", "C", "M3", "DESC-C", 3, main, "R3", maker_c, {total_qty: 3, first_process: "SMT"}]],
            ctrl: [["1", "A", "M1-NEW", "DESC-A", 1, main, "R1", maker_a, {total_qty: 1, loss: "0.01", fixed_loss: "50", first_process: "SMT"}], ["3", "C", "M3", "DESC-C", 3, main, "R3", maker_c, {total_qty: 3, loss: "0.005", fixed_loss: "50", first_process: "SMT"}]],
        }
        data = {
            "old_file": (_plm_full_bom_bytes(old_rows), "old_plm.xlsx"),
            "new_file": (_plm_full_bom_bytes(new_rows), "new_plm.xlsx"),
            "config": json.dumps({"key_col": "\u6599\u53f7", "compare_cols": ["\u578b\u53f7", loss]}),
        }
        resp = app.test_client().post("/api/bom_compare/hq_version", data=data, content_type="multipart/form-data")
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["format"], "plm_full")
        self.assertEqual(payload["sheets"], [bom, dbg, ctrl])
        self.assertEqual(payload["old_total"], 6)
        self.assertEqual(payload["new_total"], 6)
        self.assertEqual(payload["added"], 3)
        self.assertEqual(payload["removed"], 3)
        self.assertEqual(payload["changed"], 3)
        self.assertEqual(payload["unchanged"], 0)

        filename = unquote(payload["download"].split("/download/", 1)[1])
        wb = openpyxl.load_workbook(WEB_APP / "outputs" / filename, data_only=True)
        self.assertIn("BOM\u5dee\u5f02", wb.sheetnames)
        self.assertIn("DBG\u4e1a\u52a1BOM\u5dee\u5f02", wb.sheetnames)
        self.assertIn("DBGBOM\u5236\u63a7\u4fe1\u606f\u5dee\u5f02", wb.sheetnames)
        self.assertIn("\u5168\u90e8\u5dee\u5f02\u660e\u7ec6", wb.sheetnames)
        all_detail = wb["\u5168\u90e8\u5dee\u5f02\u660e\u7ec6"]
        keys = [row[2] for row in all_detail.iter_rows(min_row=2, values_only=True)]
        self.assertNotIn("2026-05-12", keys)
        self.assertIn("A", keys)
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
        hq = _hq_export_bytes([
            ["1", "A", "R1", "", 1, "", "", ""],
            ["2", "B", "C2", "", 2, "", "", ""],
            ["3", "D", "L2", "", 4, "", "", ""],
        ])
        data = {
            "left_file": (customer, "customer.xlsx"),
            "right_file": (hq, "hq.xlsx"),
            "config": json.dumps({
                "compare_type": "customer_hq",
                "left_header_row": 1,
                "right_header_row": 3,
                "left_key_col": "客户料号",
                "right_key_col": "料号",
                "field_pairs": [{"left": "型号", "right": "型号"}, {"left": "数量", "right": "单耗"}],
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
        self.assertEqual([cell.value for cell in detail[1]], ["差异类型", "匹配键", "客户BOM行号", "HQ BOM行号", "差异字段", "客户BOM值", "HQ BOM值"])
        detail_types = [row[0] for row in detail.iter_rows(min_row=2, values_only=True)]
        self.assertIn("仅客户BOM存在", detail_types)
        self.assertIn("仅HQ BOM存在", detail_types)
        self.assertNotIn("仅左侧存在", detail_types)
        self.assertNotIn("仅右侧存在", detail_types)
        changed_rows = [row for row in detail.iter_rows(min_row=2, values_only=True) if row[0] == "字段变更"]
        self.assertEqual(len(changed_rows), 1)
        self.assertEqual(changed_rows[0][1], "B")
        self.assertEqual(changed_rows[0][4], "型号")
        self.assertNotIn("数量", [row[4] for row in changed_rows])
        self.assertIn("仅客户BOM存在", wb.sheetnames)
        self.assertIn("仅HQ BOM存在", wb.sheetnames)
        left_only_sheet = wb["仅客户BOM存在"]
        right_only_sheet = wb["仅HQ BOM存在"]
        self.assertEqual([cell.value for cell in left_only_sheet[1]][:3], ["客户料号", "型号", "数量"])
        self.assertEqual([cell.value for cell in left_only_sheet[2]][:3], ["C", "L1", "3"])
        self.assertEqual([cell.value for cell in right_only_sheet[1]][:8], ["序号", "料号", "型号", "物料描述", "单耗", "替代关系", "位号", "生产厂家"])
        self.assertEqual([cell.value for cell in right_only_sheet[2]][:3], ["3", "D", "L2"])
        wb.close()

    def test_generic_customer_compare_accepts_plm_full_hq_bom_on_right(self):
        customer_key = "\u5ba2\u6237\u6599\u53f7"
        model = "\u578b\u53f7"
        part_no = "\u6599\u53f7"
        main = "\u4e3b\u6599"
        maker_a = "\u5382\u5546A"
        maker_b = "\u5382\u5546B"
        dbg_sheet = "DBG\u4e1a\u52a1BOM"
        customer_bytes = _xlsx_bytes([customer_key, model], [["A", "M1-NEW"], ["X", "MX"]]).getvalue()
        hq_rows = {"BOM": [["1", "A", "M1", "DESC-A", 1, main, "R1", maker_a], ["2", "B", "M2", "DESC-B", 2, main, "R2", maker_b]]}
        hq_bytes = _plm_full_bom_bytes(hq_rows).getvalue()
        data = {
            "left_file": (io.BytesIO(customer_bytes), "customer.xlsx"),
            "right_file": (io.BytesIO(hq_bytes), "hq_plm.xlsx"),
            "left_header_row": "1",
            "right_header_row": "3",
        }
        sheets_resp = app.test_client().post("/api/bom_compare/generic_sheets", data=data, content_type="multipart/form-data")
        sheets_payload = sheets_resp.get_json()
        self.assertTrue(sheets_payload["success"])
        self.assertEqual(sheets_payload["right_format"], "plm_full")
        self.assertEqual(sheets_payload["right_header_row"], 8)
        self.assertEqual(sheets_payload["right_current_sheet"], "BOM")
        self.assertIn(dbg_sheet, sheets_payload["right_bom_sheets"])

        compare_data = {
            "left_file": (io.BytesIO(customer_bytes), "customer.xlsx"),
            "right_file": (io.BytesIO(hq_bytes), "hq_plm.xlsx"),
            "config": json.dumps({
                "compare_type": "customer_hq",
                "left_header_row": 1,
                "right_header_row": 3,
                "left_key_col": customer_key,
                "right_key_col": part_no,
                "field_pairs": [{"left": model, "right": model}],
            }),
        }
        compare_resp = app.test_client().post("/api/bom_compare/generic", data=compare_data, content_type="multipart/form-data")
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["right_format"], "plm_full")
        self.assertEqual(payload["left_only"], 1)
        self.assertEqual(payload["right_only"], 1)
        self.assertEqual(payload["changed"], 1)

    def test_generic_cadence_standard_export_auto_header_and_refdes_expand(self):
        refdes = "位号"
        part_no = "料号"
        model = "型号"
        main = "主料"
        maker = "厂商A"
        left_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 2, main, "R1,R2", maker],
            ["2", "B", "M2", "", 1, main, "R3", maker],
        ]).getvalue()
        right_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 1, main, "R1", maker],
            ["2", "A", "M1", "", 1, main, "R2", maker],
            ["3", "C", "M3", "", 1, main, "R3", maker],
        ]).getvalue()

        sheets_resp = app.test_client().post("/api/bom_compare/generic_sheets", data={
            "left_file": (io.BytesIO(left_bytes), "cadence.xlsx"),
            "right_file": (io.BytesIO(right_bytes), "hq.xlsx"),
            "compare_type": "cadence_hq",
            "left_header_row": "1",
            "right_header_row": "3",
        }, content_type="multipart/form-data")
        sheets_payload = sheets_resp.get_json()
        self.assertTrue(sheets_payload["success"])
        self.assertEqual(sheets_payload["left_format"], "cadence_standard")
        self.assertEqual(sheets_payload["left_header_row"], 3)
        self.assertEqual(sheets_payload["detected_left_key"], part_no)
        self.assertEqual(sheets_payload["detected_right_key"], part_no)

        compare_resp = app.test_client().post("/api/bom_compare/generic", data={
            "left_file": (io.BytesIO(left_bytes), "cadence.xlsx"),
            "right_file": (io.BytesIO(right_bytes), "hq.xlsx"),
            "config": json.dumps({
                "compare_type": "cadence_hq",
                "left_header_row": 1,
                "right_header_row": 3,
                "left_key_col": refdes,
                "right_key_col": refdes,
                "field_pairs": [{"left": part_no, "right": part_no}, {"left": model, "right": model}],
            }),
        }, content_type="multipart/form-data")
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertTrue(payload["expanded_refdes"])
        self.assertEqual(payload["left_header_row"], 3)
        self.assertEqual(payload["left_total"], 3)
        self.assertEqual(payload["right_total"], 3)
        self.assertEqual(payload["left_only"], 0)
        self.assertEqual(payload["right_only"], 0)
        self.assertEqual(payload["changed"], 1)
        self.assertEqual(payload["same"], 2)

    def test_generic_cadence_treats_refdes_order_as_same(self):
        part_no = "料号"
        refdes = "位号"
        main = "主料"
        maker = "厂商A"
        left_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 3, main, "R1,R2,R3", maker],
        ]).getvalue()
        right_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 3, main, "R3,R1,R2", maker],
        ]).getvalue()
        compare_resp = app.test_client().post("/api/bom_compare/generic", data={
            "left_file": (io.BytesIO(left_bytes), "cadence.xlsx"),
            "right_file": (io.BytesIO(right_bytes), "hq.xlsx"),
            "config": json.dumps({
                "compare_type": "cadence_hq",
                "left_header_row": 1,
                "right_header_row": 3,
                "left_key_col": part_no,
                "right_key_col": part_no,
                "field_pairs": [{"left": refdes, "right": refdes}],
            }),
        }, content_type="multipart/form-data")
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["changed"], 0)
        self.assertEqual(payload["same"], 1)
        filename = unquote(payload["download"].split("/download/", 1)[1])
        wb = openpyxl.load_workbook(WEB_APP / "outputs" / filename, data_only=True)
        detail = wb["差异明细"]
        self.assertEqual(detail.max_row, 1)
        wb.close()

    def test_generic_cadence_ignores_empty_columns_by_default(self):
        part_no = "料号"
        model = "型号"
        empty_col = "物料描述"
        main = "主料"
        maker = "厂商A"
        left_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 1, main, "R1", maker],
        ]).getvalue()
        right_bytes = _hq_export_bytes([
            ["1", "A", "M1", "", 1, main, "R1", maker],
        ]).getvalue()

        sheets_resp = app.test_client().post("/api/bom_compare/generic_sheets", data={
            "left_file": (io.BytesIO(left_bytes), "cadence.xlsx"),
            "right_file": (io.BytesIO(right_bytes), "hq.xlsx"),
            "compare_type": "cadence_hq",
            "left_header_row": "1",
            "right_header_row": "3",
        }, content_type="multipart/form-data")
        sheets_payload = sheets_resp.get_json()
        self.assertTrue(sheets_payload["success"])
        self.assertNotIn(empty_col, sheets_payload["left_headers"])
        self.assertIn(empty_col, sheets_payload["left_ignored_headers"])

        compare_resp = app.test_client().post("/api/bom_compare/generic", data={
            "left_file": (io.BytesIO(left_bytes), "cadence.xlsx"),
            "right_file": (io.BytesIO(right_bytes), "hq.xlsx"),
            "config": json.dumps({
                "compare_type": "cadence_hq",
                "left_header_row": 1,
                "right_header_row": 3,
                "left_key_col": part_no,
                "right_key_col": part_no,
                "field_pairs": [{"left": empty_col, "right": model}, {"left": model, "right": model}],
            }),
        }, content_type="multipart/form-data")
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["changed"], 0)
        self.assertEqual(payload["same"], 1)
        self.assertIn(empty_col, payload["left_ignored_headers"])

    def test_generic_cadence_compare_accepts_plm_full_hq_bom_on_right(self):
        main = "\u4e3b\u6599"
        maker_a = "\u5382\u5546A"
        maker_b = "\u5382\u5546B"
        refdes = "\u4f4d\u53f7"
        cadence = _xlsx_bytes(["REFDES", "PART_NUMBER"], [["R1", "A"], ["R9", "Z"]])
        hq = _plm_full_bom_bytes({"BOM": [["1", "A", "M1", "DESC-A", 1, main, "R1", maker_a], ["2", "B", "M2", "DESC-B", 2, main, "R2", maker_b]]})
        data = {
            "left_file": (cadence, "cadence.xlsx"),
            "right_file": (hq, "hq_plm.xlsx"),
            "compare_type": "cadence_hq",
            "left_header_row": "1",
            "right_header_row": "3",
        }
        resp = app.test_client().post("/api/bom_compare/generic_sheets", data=data, content_type="multipart/form-data")
        payload = resp.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["right_format"], "plm_full")
        self.assertEqual(payload["right_header_row"], 8)
        self.assertEqual(payload["detected_left_key"], "PART_NUMBER")
        self.assertEqual(payload["detected_right_key"], "料号")


    def test_generic_sheets_detects_left_and_right_keys(self):
        cadence = _xlsx_bytes(["REFDES", "PART_NUMBER", "QTY"], [["R1", "A", 1]])
        hq = _hq_export_bytes([["1", "A", "M1", "", 1, "", "R1", ""]])
        data = {
            "left_file": (cadence, "cadence.xlsx"),
            "right_file": (hq, "hq.xlsx"),
            "left_header_row": "1",
            "right_header_row": "3",
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
