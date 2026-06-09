# -*- coding: utf-8 -*-
import json
import sys
import unittest
from io import BytesIO
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
from openpyxl import Workbook, load_workbook  # noqa: E402


def xlsx_bytes(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


class FreeBomCompareTests(unittest.TestCase):
    def test_free_bom_compare_accepts_two_arbitrary_excels(self):
        client = app.test_client()
        sheets_resp = client.post(
            "/api/bom_compare/free_sheets",
            data={
                "left_file": (xlsx_bytes(["Item", "MPN", "Qty", "Remark"], [["A1", "R-10K", 2, "old"]]), "left.xlsx"),
                "right_file": (xlsx_bytes(["Code", "Model", "Quantity", "Remark"], [["A1", "R-10K", 3, "new"]]), "right.xlsx"),
                "left_header_row": "1",
                "right_header_row": "1",
            },
            content_type="multipart/form-data",
        )
        sheets_payload = sheets_resp.get_json()
        self.assertTrue(sheets_payload["success"], sheets_payload)
        self.assertEqual(sheets_payload["left_format"], "generic")
        self.assertEqual(sheets_payload["right_format"], "generic")

        compare_resp = client.post(
            "/api/bom_compare/free",
            data={
                "left_file": (xlsx_bytes(["Item", "MPN", "Qty", "Remark"], [["A1", "R-10K", 2, "old"], ["A2", "C-1U", 1, "same"]]), "left.xlsx"),
                "right_file": (xlsx_bytes(["Code", "Model", "Quantity", "Remark"], [["A1", "R-10K", 3, "new"], ["A3", "L-2U2", 1, "added"]]), "right.xlsx"),
                "config": json.dumps({
                    "left_header_row": 1,
                    "right_header_row": 1,
                    "left_key_col": "Item",
                    "right_key_col": "Code",
                    "field_pairs": [
                        {"left": "MPN", "right": "Model"},
                        {"left": "Qty", "right": "Quantity"},
                        {"left": "Remark", "right": "Remark"},
                    ],
                }),
            },
            content_type="multipart/form-data",
        )
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"], payload)
        self.assertEqual(payload["left_only"], 1)
        self.assertEqual(payload["right_only"], 1)
        self.assertEqual(payload["changed"], 1)
        self.assertEqual(payload["same"], 0)
        self.assertTrue(payload["download"].endswith(".xlsx"))

    def test_free_bom_report_keeps_all_field_changes_on_one_row_per_item(self):
        client = app.test_client()
        compare_resp = client.post(
            "/api/bom_compare/free",
            data={
                "left_file": (xlsx_bytes(["Item", "MPN", "Qty", "Remark"], [["A1", "R-10K", 2, "old"]]), "left.xlsx"),
                "right_file": (xlsx_bytes(["Code", "Model", "Quantity", "Remark"], [["A1", "R-22K", 3, "new"]]), "right.xlsx"),
                "config": json.dumps({
                    "left_header_row": 1,
                    "right_header_row": 1,
                    "left_key_col": "Item",
                    "right_key_col": "Code",
                    "field_pairs": [
                        {"left": "MPN", "right": "Model"},
                        {"left": "Qty", "right": "Quantity"},
                        {"left": "Remark", "right": "Remark"},
                    ],
                }),
            },
            content_type="multipart/form-data",
        )
        payload = compare_resp.get_json()
        self.assertTrue(payload["success"], payload)

        report_name = payload["download"].rsplit("/", 1)[-1]
        report_path = WEB_APP / "outputs" / report_name
        wb = load_workbook(report_path, data_only=True)
        try:
            detail = wb["\u5dee\u5f02\u660e\u7ec6"]
            rows = list(detail.iter_rows(values_only=True))
            self.assertEqual(len(rows), 2)
            data = rows[1]
            self.assertEqual(data[0], "\u5b57\u6bb5\u53d8\u66f4")
            self.assertEqual(data[1], "A1")
            self.assertEqual(data[4], 3)
            self.assertEqual(data[5], "MPN <-> Model: R-10K -> R-22K")
            self.assertEqual(data[6], "Qty <-> Quantity: 2 -> 3")
            self.assertEqual(data[7], "Remark: old -> new")

            field_detail = wb["\u5b57\u6bb5\u53d8\u5316\u660e\u7ec6"]
            field_rows = list(field_detail.iter_rows(min_row=2, values_only=True))
            self.assertEqual(len(field_rows), 3)
            self.assertEqual(field_rows[0], ("A1", "MPN <-> Model", "R-10K", "R-22K", 2, 2))
            self.assertEqual(field_rows[1], ("A1", "Qty <-> Quantity", "2", "3", 2, 2))
            self.assertEqual(field_rows[2], ("A1", "Remark", "old", "new", 2, 2))
        finally:
            wb.close()

    def test_free_bom_preview_returns_rows_before_header_selection(self):
        client = app.test_client()
        resp = client.post(
            "/api/bom_compare/free_preview",
            data={
                "left_file": (xlsx_bytes(["note", "not header"], [["Item", "Qty"], ["A1", 2]]), "left.xlsx"),
            },
            content_type="multipart/form-data",
        )
        payload = resp.get_json()
        self.assertTrue(payload["success"], payload)
        self.assertIn("left", payload)
        self.assertEqual(payload["left"]["current_sheet"], "Sheet")
        rows = payload["left"]["rows"]
        self.assertEqual(rows[0]["row_number"], 1)
        self.assertEqual(rows[1]["values"][:2], ["Item", "Qty"])
        self.assertEqual(rows[2]["values"][:2], ["A1", "2"])

    def test_bom_compare_template_defaults_to_free_bom_first(self):
        html = (WEB_APP / "templates" / "partials" / "tools" / "bom-compare.html").read_text(encoding="utf-8")
        first_button = html.index('class="bomcmp-tab-btn active"')
        self.assertIn('data-bomcmp-tab="free-bom"', html[first_button:first_button + 140])
        self.assertLess(html.index('id="bomcmp-tab-free-bom"'), html.index('id="bomcmp-tab-customer-hq"'))
        self.assertIn('id="freeLeftPreviewRows"', html)
        self.assertIn('id="freeRightPreviewRows"', html)


if __name__ == "__main__":
    unittest.main()
