import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from pstx_integrations.feishu import gateway as pstx_feishu_bom
from pstx_integrations.feishu import gateway as integration_feishu_gateway


def build_feishu_data_dir() -> Path:
    root = Path(tempfile.mkdtemp())
    (root / "feishu_libraries.json").write_text(
        json.dumps({
            "base_url": "https://mcenter.example.local",
            "origin": "cli_demo",
            "user_id": "100001",
            "libraries": [{"id": "lib1", "name": "优选库"}],
        }, ensure_ascii=False),
        encoding="utf-8",
    )
    conn = sqlite3.connect(root / "feishu_cache.db")
    conn.execute(
        """
        CREATE TABLE materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lib_id TEXT,
            lib_name TEXT,
            sheet_name TEXT,
            key_value TEXT COLLATE NOCASE,
            hq_no TEXT,
            brand TEXT,
            spec TEXT,
            description TEXT,
            raw_data TEXT,
            synced_at TEXT
        )
        """
    )
    conn.executemany(
        "INSERT INTO materials(lib_id,lib_name,sheet_name,key_value,hq_no,brand,spec,description,raw_data,synced_at) "
        "VALUES(?,?,?,?,?,?,?,?,?,?)",
        [
            ("lib1", "优选库", "Sheet1", "ABC-123", "HQ001", "ACME", "ABC-123", "demo material", "{}", "2026-04-26"),
            ("lib1", "优选库", "Sheet1", "DEF-456", "HQ002", "ACME", "DEF-456", "other material", "{}", "2026-04-26"),
        ],
    )
    conn.commit()
    conn.close()
    return root


class FeishuBomAdapterTests(unittest.TestCase):
    def test_integration_feishu_gateway_entrypoint_exports_public_api(self):
        self.assertIs(integration_feishu_gateway.fetch_feishu_sheet_list, pstx_feishu_bom.fetch_feishu_sheet_list)
        self.assertIs(integration_feishu_gateway.preview_feishu_sheet, pstx_feishu_bom.preview_feishu_sheet)
        self.assertIs(integration_feishu_gateway.sync_feishu_library, pstx_feishu_bom.sync_feishu_library)
        self.assertFalse(hasattr(integration_feishu_gateway, "_normalize_rows"))

    def add_temp_cleanup(self, root: Path) -> None:
        def cleanup():
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            if root.exists():
                root.rmdir()

        self.addCleanup(cleanup)

    def test_status_reports_missing_data_dir_without_raising(self):
        status = pstx_feishu_bom.build_feishu_bom_status(
            data_dir="/tmp/not-exists-feishu-bom-data",
        )
        self.assertFalse(status["available"])
        self.assertIn("未找到飞书 BOM 数据目录", status["error"])

    def test_status_reads_config_and_cache_without_external_matcher(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)
        status = pstx_feishu_bom.build_feishu_bom_status(data_dir=str(data_dir))
        self.assertTrue(status["ok"])
        self.assertTrue(status["available"])
        self.assertTrue(status["configured"])
        self.assertEqual(1, status["library_count"])
        self.assertEqual(2, status["cache_count"])
        self.assertEqual("优选库", status["cache_stats"][0]["lib_name"])

    def test_match_rows_with_cache_uses_key_field(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)
        result = pstx_feishu_bom.match_rows_with_feishu_cache(
            [
                {"位号": "R1", "厂家型号": "ABC-123"},
                {"位号": "R2", "厂家型号": "NOPE"},
                {"位号": "R3", "厂家型号": ""},
            ],
            "厂家型号",
            data_dir=str(data_dir),
        )
        self.assertTrue(result["ok"])
        self.assertEqual(1, result["matched_count"])
        self.assertEqual(1, result["unmatched_count"])
        self.assertEqual(1, result["skipped_count"])
        self.assertEqual("HQ001", result["rows"][0]["HQ料号"])
        self.assertEqual("未匹配", result["rows"][1]["匹配状态"])
        self.assertEqual("跳过：关键值为空", result["rows"][2]["匹配状态"])

    def test_match_rows_with_cache_can_match_by_hq_no_and_return_standard_fields(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)
        pstx_feishu_bom.get_feishu_cache_rows(data_dir=str(data_dir), limit=1)
        conn = sqlite3.connect(data_dir / "feishu_cache.db")
        conn.execute(
            "UPDATE materials SET pi=?, selection_order=? WHERE hq_no=?",
            ("LiXinYu", "1", "HQ001"),
        )
        conn.commit()
        conn.close()

        result = pstx_feishu_bom.match_rows_with_feishu_cache(
            [{"位号": "U1", "料号": "HQ001", "描述": "IC_CPU", "封装": "BGA", "类型": "IC 芯片"}],
            "料号",
            data_dir=str(data_dir),
            match_mode="hq_no",
        )

        self.assertTrue(result["ok"])
        self.assertEqual("hq_no", result["match_mode"])
        self.assertEqual(1, result["matched_count"])
        row = result["rows"][0]
        self.assertEqual("HQ001", row["项目HQ料号"])
        self.assertEqual("HQ001", row["飞书HQ料号"])
        self.assertEqual("ABC-123", row["飞书规格型号"])
        self.assertEqual("LiXinYu", row["PI"])
        self.assertEqual("1", row["选型顺序"])

    def test_extract_spreadsheet_token_supports_feishu_urls(self):
        self.assertEqual(
            "abc123XYZ",
            pstx_feishu_bom.extract_spreadsheet_token("https://example.feishu.cn/sheets/abc123XYZ?sheet=foo"),
        )
        self.assertEqual(
            "baseToken9",
            pstx_feishu_bom.extract_spreadsheet_token("https://example.feishu.cn/base/baseToken9"),
        )
        self.assertEqual("plain_token", pstx_feishu_bom.extract_spreadsheet_token("plain_token"))

    def test_default_column_range_uses_sheet_column_count(self):
        self.assertEqual("A:Z", pstx_feishu_bom._default_column_range(26))
        self.assertEqual("A:AD", pstx_feishu_bom._default_column_range(30))
        self.assertEqual("A:AZ", pstx_feishu_bom._default_column_range(52))
        self.assertEqual("A:Z", pstx_feishu_bom._default_column_range(0))

    def test_safe_cell_str_formats_integer_float_without_decimal_suffix(self):
        self.assertEqual("1", pstx_feishu_bom._safe_cell_str(1.0))
        self.assertEqual("1.5", pstx_feishu_bom._safe_cell_str(1.5))

    def test_client_fetches_sheets_and_reads_values_via_gateway(self):
        calls = []

        def fake_transport(url, params, timeout):
            calls.append((url, params, timeout))
            if url.endswith("/spreadsheetsMetainfo"):
                return {
                    "code": 0,
                    "data": {
                        "sheets": [
                            {"sheetId": "sh1", "title": "优选库", "rowCount": 8},
                            {"sheetId": "empty-title", "title": ""},
                        ]
                    },
                }
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 0,
                    "data": {
                        "valueRange": {
                            "values": [
                                ["厂家型号", "HQ料号"],
                                ["ABC-123", "HQ001"],
                                ["", ""],
                            ]
                        }
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        sheets = client.get_sheets("https://feishu.example/sheets/token123")
        values = client.read_sheet("token123", "sh1", row_count=9, column_range="A:AZ")

        self.assertEqual(1, len(sheets))
        self.assertEqual("sh1", sheets[0]["sheet_id"])
        self.assertEqual(["厂家型号", "HQ料号"], values[0][:2])
        self.assertEqual(["ABC-123", "HQ001"], values[1][:2])
        self.assertTrue(all(len(row) == 52 for row in values))
        self.assertEqual("token123", calls[0][1]["spreadsheetToken"])
        self.assertEqual("sh1!A1:AZ50", calls[1][1]["range"])

    def test_client_skips_block_sheets_and_normalizes_rich_cells_from_value_ranges(self):
        calls = []

        def fake_transport(url, params, timeout):
            calls.append((url, params, timeout))
            if url.endswith("/spreadsheetsMetainfo"):
                return {
                    "code": 0,
                    "data": {
                        "properties": {"revision": 23},
                        "sheets": [
                            {
                                "sheetId": "sh1",
                                "title": "电容库",
                                "index": 0,
                                "rowCount": 6,
                                "columnCount": 30,
                                "frozenRowCount": 1,
                                "frozenColCount": 0,
                            },
                            {
                                "sheetId": "doc1",
                                "title": "说明文档",
                                "blockInfo": {"blockToken": "blk", "blockType": "doc"},
                            },
                        ],
                    },
                }
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 200,
                    "data": {
                        "valueRanges": [
                            {
                                "range": "sh1!A1:D6",
                                "values": [
                                    [
                                        {"text": "规格型号"},
                                        [{"text": "HQ编码"}],
                                        "PI",
                                        {"link": "https://example.local/spec", "type": "url"},
                                    ],
                                    ["CAP-100N", "HQ17101005", "LiXinYu", ""],
                                    ["", "", "", ""],
                                ],
                            }
                        ]
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        sheets = client.get_sheets("token123")
        values = client.read_sheet("token123", "sh1", row_count=6, column_range="A:D")

        self.assertEqual(1, len(sheets))
        self.assertEqual("sh1", sheets[0]["sheet_id"])
        self.assertEqual(23, sheets[0]["revision"])
        self.assertEqual(30, sheets[0]["column_count"])
        self.assertEqual("A:AD", sheets[0]["column_range"])
        self.assertEqual(
            [
                ["规格型号", "HQ编码", "PI", "https://example.local/spec"],
                ["CAP-100N", "HQ17101005", "LiXinYu", ""],
            ],
            values,
        )
        self.assertEqual("cli_demo", calls[0][1]["origin"])
        value_call = [call for call in calls if call[0].endswith("/getSheetsValue")][0]
        self.assertEqual("100001", value_call[1]["userId"])

    def test_client_aligns_values_when_returned_range_starts_after_requested_range(self):
        def fake_transport(url, params, timeout):
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 200,
                    "data": {
                        "valueRange": {
                            "range": "sh1!B1:D3",
                            "values": [
                                ["规格型号", "HQ料号", "PI"],
                                ["ABC-123", "HQ001", "LiXinYu"],
                            ],
                        }
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        values = client.read_sheet("token123", "sh1", row_count=3, column_range="A:D")

        self.assertEqual(
            [
                ["", "规格型号", "HQ料号", "PI"],
                ["", "ABC-123", "HQ001", "LiXinYu"],
            ],
            values,
        )

    def test_client_prefers_v2_formatted_values_over_formula_text(self):
        calls = []

        def fake_transport(url, params, timeout):
            calls.append((url, params, timeout))
            if url.endswith("/fs/sheets"):
                return {
                    "code": 0,
                    "data": {
                        "valueRanges": [
                            {
                                "range": "sh1!A1:E2",
                                "values": [
                                    ["分组", "HQ编码", "规格型号", "描述", "选型顺序"],
                                    ["1", "HQ1710101A3Q0", "0201N0R5A500CT", "MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", "5"],
                                ],
                            }
                        ]
                    },
                }
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 200,
                    "data": {
                        "valueRange": {
                            "range": "sh1!A1:E2",
                            "values": [
                                ["分组", "HQ编码", "规格型号", "描述", "选型顺序"],
                                ["1", "HQ1710101A3Q0", "0201N0R5A500CT", '"MLCC"&"_"&F2', "5"],
                            ],
                        }
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        values = client.read_sheet("token123", "sh1", row_count=2, column_range="A:E")

        self.assertEqual("MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", values[1][3])
        self.assertEqual("/fs/sheets", calls[0][0].rsplit("https://mcenter.example.local", 1)[-1])
        self.assertEqual("FormattedValue", calls[0][1]["valueRenderOption"])
        self.assertFalse(any(call[0].endswith("/getSheetsValue") for call in calls))

    def test_online_client_writes_debug_log_without_raw_spreadsheet_token(self):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        old_log = os.environ.get("PSTX_FEISHU_LOG_FILE")
        old_parse_log = os.environ.get("PSTX_FEISHU_PARSE_LOG_FILE")
        old_payload = os.environ.get("PSTX_FEISHU_LOG_PAYLOAD")
        log_path = Path(temp_dir.name) / "feishu.log"
        parse_log_path = Path(temp_dir.name) / "feishu_parse.log"
        os.environ["PSTX_FEISHU_LOG_FILE"] = str(log_path)
        os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = str(parse_log_path)
        os.environ.pop("PSTX_FEISHU_LOG_PAYLOAD", None)

        def cleanup_env():
            if old_log is None:
                os.environ.pop("PSTX_FEISHU_LOG_FILE", None)
            else:
                os.environ["PSTX_FEISHU_LOG_FILE"] = old_log
            if old_parse_log is None:
                os.environ.pop("PSTX_FEISHU_PARSE_LOG_FILE", None)
            else:
                os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = old_parse_log
            if old_payload is None:
                os.environ.pop("PSTX_FEISHU_LOG_PAYLOAD", None)
            else:
                os.environ["PSTX_FEISHU_LOG_PAYLOAD"] = old_payload

        self.addCleanup(cleanup_env)

        def fake_transport(url, params, timeout):
            if url.endswith("/spreadsheetsMetainfo"):
                return {
                    "code": 0,
                    "data": {
                        "sheets": [
                            {"sheetId": "sh1", "title": "优选库", "rowCount": 3, "columnCount": 4},
                            {"sheetId": "doc1", "title": "说明", "blockInfo": {"blockType": "doc"}},
                        ]
                    },
                }
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 0,
                    "data": {
                        "valueRange": {
                            "values": [
                                ["规格型号", "HQ料号"],
                                ["ABC-123", "HQ001"],
                            ]
                        }
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        client.get_sheets("token123")
        client.read_sheet("token123", "sh1", row_count=3, column_range="A:D")

        raw_lines = log_path.read_text(encoding="utf-8")
        self.assertIn("feishu_bom.metainfo.parsed", raw_lines)
        self.assertIn("feishu_bom.read_sheet.parsed", raw_lines)
        self.assertIn("skipped_sheet_count", raw_lines)
        self.assertIn("row_preview", raw_lines)
        self.assertNotIn('"spreadsheetToken": "token123"', raw_lines)
        self.assertNotIn('"spreadsheet_token": "token123"', raw_lines)

        records = [json.loads(line) for line in raw_lines.splitlines() if line.strip()]
        parsed = [record for record in records if record["event"] == "feishu_bom.read_sheet.parsed"][0]
        self.assertEqual(2, parsed["normalized_shape"]["row_count"])
        self.assertEqual("ABC-123", parsed["row_preview"][1][0])
        parse_lines = parse_log_path.read_text(encoding="utf-8")
        self.assertIn("feishu_bom_parse.read_sheet.rows_aligned", parse_lines)
        self.assertIn("first_rows_cells", parse_lines)

    def test_sync_library_writes_compatible_materials_cache(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                self.last_request = {
                    "token": token,
                    "sheet_id": sheet_id,
                    "row_count": row_count,
                    "column_range": column_range,
                }
                return [
                    ["厂家型号", "HQ料号", "制造商", "规格型号", "描述"],
                    ["ABC-123", "HQ001", "ACME", "ABC-123", "demo material"],
                    ["", "HQ-EMPTY", "ACME", "", ""],
                    ["DEF-456", "HQ002", "ACME", "DEF-456", "other material"],
                ]

        fake_client = FakeClient()
        result = pstx_feishu_bom.sync_feishu_library(
            library_name="优选库",
            spreadsheet_token_or_url="https://example.feishu.cn/sheets/token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="lib1",
            sheets=[
                {
                    "sheet_id": "sh1",
                    "title": "Sheet1",
                    "key_col": "厂家型号",
                    "hq_no_col": "HQ料号",
                    "brand_col": "制造商",
                    "spec_col": "规格型号",
                    "desc_col": "描述",
                    "row_count": 99,
                    "column_range": "A:Z",
                },
                {"sheet_id": "sh2", "title": "未配置", "key_col": "", "enabled": True},
            ],
            client=fake_client,
        )

        self.assertTrue(result["ok"])
        self.assertEqual(3, result["synced_rows"])
        self.assertEqual(1, result["skipped_sheets"])
        self.assertEqual("token123", fake_client.last_request["token"])
        self.assertTrue((data_dir / "feishu_libraries.json").is_file())

        match = pstx_feishu_bom.match_rows_with_feishu_cache(
            [{"位号": "R1", "厂家型号": "ABC-123"}],
            "厂家型号",
            data_dir=str(data_dir),
        )
        self.assertTrue(match["ok"])
        self.assertEqual(1, match["matched_count"])
        self.assertEqual("HQ001", match["rows"][0]["HQ料号"])
        self.assertEqual("ACME", match["rows"][0]["HQ制造商"])

    def test_sync_library_writes_dedicated_parse_log_with_row_field_diagnostics(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        old_parse_log = os.environ.get("PSTX_FEISHU_PARSE_LOG_FILE")
        old_parse_rows = os.environ.get("PSTX_FEISHU_PARSE_LOG_ROWS")
        parse_log_path = Path(temp_dir.name) / "feishu_parse.log"
        os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = str(parse_log_path)
        os.environ["PSTX_FEISHU_PARSE_LOG_ROWS"] = "20"

        def cleanup_env():
            if old_parse_log is None:
                os.environ.pop("PSTX_FEISHU_PARSE_LOG_FILE", None)
            else:
                os.environ["PSTX_FEISHU_PARSE_LOG_FILE"] = old_parse_log
            if old_parse_rows is None:
                os.environ.pop("PSTX_FEISHU_PARSE_LOG_ROWS", None)
            else:
                os.environ["PSTX_FEISHU_PARSE_LOG_ROWS"] = old_parse_rows

        self.addCleanup(cleanup_env)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["分组", "HQ编码", "规格型号", "PI", "选型顺序"],
                    ["", "HQ001", "CAP-100N", "LiXinYu", "5"],
                    ["", "", "", "", ""],
                    ["", "HQ002", "", "", "6"],
                ]

        result = pstx_feishu_bom.sync_feishu_library(
            library_name="电容优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="cap_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "CAP",
                "header_row": 1,
                "hq_code_col": "HQ编码",
                "spec_model_col": "规格型号",
                "pi_col": "PI",
                "selection_order_col": "选型顺序",
            }],
            client=FakeClient(),
        )

        self.assertEqual(2, result["synced_rows"])
        parse_lines = parse_log_path.read_text(encoding="utf-8")
        self.assertIn("feishu_bom_parse.sync.sheet.mapping_resolved", parse_lines)
        self.assertIn("feishu_bom_parse.sync.sheet.row_parse_summary", parse_lines)
        self.assertIn('"a_column_empty": true', parse_lines)
        self.assertIn("missing_spec_and_hq", parse_lines)
        self.assertIn("HQ001", parse_lines)

    def test_preview_sheet_returns_text_rows_and_mapping_suggestion(self):
        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["备注", "", ""],
                    ["厂家型号", "HQ料号", "制造商"],
                    ["ABC-123", "HQ001", "ACME"],
                ]

        result = pstx_feishu_bom.preview_feishu_sheet(
            spreadsheet_token_or_url="token123",
            sheet_id="sh1",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            row_count=3,
            header_row=2,
            client=FakeClient(),
        )

        self.assertTrue(result["ok"])
        self.assertEqual(["厂家型号", "HQ料号", "制造商"], result["headers"])
        self.assertEqual(2, result["mapping_suggestion"]["header_row"])
        self.assertEqual("厂家型号", result["mapping_suggestion"]["mapping"]["key_col"])
        self.assertEqual("厂家型号", result["mapping_suggestion"]["mapping"]["spec_model_col"])
        self.assertEqual("HQ料号", result["mapping_suggestion"]["mapping"]["hq_code_col"])

    def test_mapping_suggestion_honors_explicit_header_row(self):
        rows = [
            ["HQ编码", "规格型号", "PI", "选型顺序", "制造商"],
            ["备注列", "申请HQ编码", "规格型号", "PI", "选型顺序"],
            ["忽略", "HQ001", "ABC-123", "PI-A", "1"],
        ]

        result = pstx_feishu_bom.suggest_feishu_mapping_from_preview(rows, header_row=2)

        self.assertEqual(2, result["header_row"])
        self.assertEqual(["备注列", "申请HQ编码", "规格型号", "PI", "选型顺序"], result["headers"])
        self.assertEqual("申请HQ编码", result["mapping"]["hq_code_col"])
        self.assertEqual("规格型号", result["mapping"]["spec_model_col"])

    def test_mapping_prefers_exact_selection_order_over_adjust_flag(self):
        result = pstx_feishu_bom.build_feishu_mapping_from_headers([
            "分组",
            "HQ编码",
            "规格型号",
            "是否调整选型顺序",
            "PI",
            "选型顺序",
        ])

        mapping = result["mapping"]
        self.assertEqual("选型顺序", mapping["selection_order_col"])
        optional_columns = [field["column"] for field in mapping["optional_fields"]]
        self.assertIn("是否调整选型顺序", optional_columns)
        self.assertNotIn("选型顺序", optional_columns)

    def test_sync_library_honors_header_row(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["说明", "", ""],
                    ["厂家型号", "HQ料号", "制造商"],
                    ["ABC-123", "HQ001", "ACME"],
                ]

        result = pstx_feishu_bom.sync_feishu_library(
            library_name="优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="lib1",
            sheets=[{
                "sheet_id": "sh1",
                "title": "Sheet1",
                "header_row": 2,
                "key_col": "厂家型号",
                "hq_no_col": "HQ料号",
                "brand_col": "制造商",
            }],
            client=FakeClient(),
        )

        self.assertEqual(1, result["synced_rows"])
        self.assertEqual(2, result["per_sheet"][0]["header_row"])
        match = pstx_feishu_bom.match_rows_with_feishu_cache(
            [{"位号": "R1", "厂家型号": "ABC-123"}],
            "厂家型号",
            data_dir=str(data_dir),
        )
        self.assertEqual("HQ001", match["rows"][0]["HQ料号"])

    def test_sync_library_aligns_trimmed_returned_range_before_header_row(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        def fake_transport(url, params, timeout):
            if url.endswith("/getSheetsValue"):
                return {
                    "code": 200,
                    "data": {
                        "valueRange": {
                            "range": "sh1!B2:E3",
                            "values": [
                                ["规格型号", "HQ料号", "PI", "封装"],
                                ["CAP-100N", "HQ17101005", "LiXinYu", "0402"],
                            ],
                        }
                    },
                }
            raise AssertionError(url)

        client = pstx_feishu_bom.FeishuBomClient(
            "https://mcenter.example.local",
            "cli_demo",
            "100001",
            transport=fake_transport,
        )

        result = pstx_feishu_bom.sync_feishu_library(
            library_name="电容优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="cap_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "CAP",
                "header_row": 2,
                "column_range": "A:E",
                "spec_model_col": "规格型号",
                "hq_code_col": "HQ料号",
                "pi_col": "PI",
                "optional_fields": [{"label": "封装", "column": "封装"}],
            }],
            client=client,
        )

        self.assertTrue(result["ok"])
        self.assertEqual(1, result["synced_rows"])
        self.assertEqual(["", "规格型号", "HQ料号", "PI", "封装"], result["per_sheet"][0]["headers"])
        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="cap_lib",
            query="HQ17101005",
            data_dir=str(data_dir),
        )
        self.assertEqual(1, rows["total"])
        self.assertEqual("CAP-100N", rows["rows"][0]["spec"])
        self.assertEqual("HQ17101005", rows["rows"][0]["hq_no"])
        self.assertEqual("LiXinYu", rows["rows"][0]["pi"])
        self.assertEqual({"封装": "0402"}, rows["rows"][0]["extra_field_values"])

    def test_sync_library_writes_standard_fields_and_optional_fields(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["规格型号", "HQ编码", "PI", "选型顺序", "封装", "耐压"],
                    ["CAP-100N", "HQ17101005", "LiXinYu", "1", "0402", "50V"],
                ]

        result = pstx_feishu_bom.sync_feishu_library(
            library_name="电容优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="cap_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "CAP",
                "spec_model_col": "规格型号",
                "hq_code_col": "HQ编码",
                "pi_col": "PI",
                "selection_order_col": "选型顺序",
                "optional_fields": [
                    {"label": "封装", "column": "封装"},
                    {"label": "耐压", "column": "耐压"},
                ],
            }],
            client=FakeClient(),
        )

        self.assertTrue(result["ok"])
        self.assertEqual(1, result["synced_rows"])
        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="cap_lib",
            query="LiXinYu",
            data_dir=str(data_dir),
        )
        self.assertEqual(1, rows["total"])
        self.assertEqual("CAP-100N", rows["rows"][0]["key_value"])
        self.assertEqual("HQ17101005", rows["rows"][0]["hq_no"])
        self.assertEqual("LiXinYu", rows["rows"][0]["pi"])
        self.assertEqual("1", rows["rows"][0]["selection_order"])
        self.assertEqual({"封装": "0402", "耐压": "50V"}, rows["rows"][0]["extra_field_values"])

    def test_sync_library_repairs_rows_when_leading_blank_a_column_is_omitted(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["分组", "HQ编码", "制造商", "规格型号", "描述", "PI", "选型顺序"],
                    ["HQ1710101A3Q0", "Walsin(华新)", "0201N0R5A500CT", "MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", "LiXinYu", "5"],
                ]

        result = pstx_feishu_bom.sync_feishu_library(
            library_name="MLCC优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="mlcc_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "MLCC",
                "spec_model_col": "规格型号",
                "hq_code_col": "HQ编码",
                "brand_col": "制造商",
                "desc_col": "描述",
                "pi_col": "PI",
                "selection_order_col": "选型顺序",
            }],
            client=FakeClient(),
        )

        self.assertTrue(result["ok"])
        self.assertEqual(1, result["per_sheet"][0]["row_alignment_adjusted"])
        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="mlcc_lib",
            query="HQ1710101A3Q0",
            data_dir=str(data_dir),
        )
        self.assertEqual(1, rows["total"])
        row = rows["rows"][0]
        self.assertEqual("HQ1710101A3Q0", row["hq_no"])
        self.assertEqual("Walsin(华新)", row["brand"])
        self.assertEqual("0201N0R5A500CT", row["spec"])
        self.assertEqual("MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", row["description"])
        self.assertEqual("LiXinYu", row["pi"])
        self.assertEqual("5", row["selection_order"])

    def test_sync_library_writes_real_mlcc_display_value_row(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)
        headers = [
            "分组", "HQ编码", "制造商", "规格型号", "描述", "封装类型", "耐压值", "容值", "精度", "温度特性",
            "封装尺寸长(公差)", "封装尺寸宽(公差)", "封装尺寸高(公差)", "耐久性电压", "级别", "PLM描述",
            "特殊标识", "现有优选属性", "待调整系统状态", "是否调整选型顺序", "PI", "选型顺序",
            "供应商备注", "LT备注", "变更时间", "是否涉及其他客户料号", "是否需要申请新料号", "项目",
            "是否导入线上替代组", "参数是否确认", "参数核查日期",
        ]

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    headers,
                    [
                        "1", "HQ1710101A3Q0", "Walsin(华新)", "0201N0R5A500CT",
                        "MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", "0201", "50V", "0.5pF", "A(±0.05pF)", "C0G",
                        "0.60±0.03mm", "0.30±0.03mm", "0.30±0.03mm", "不区分",
                        "MLCC_0201_50 V_0.5pF_A(±0.05pF)_NP0", "优选", "优选", "", "", "",
                        "5", "5", "HQ1170345500B 切换新料号 HQ1710101A3Q0", "", "2024/9/11",
                        "", "已申请新料号并更新优选库", "", "", "已上传", "",
                    ],
                ]

        mapping = pstx_feishu_bom.build_feishu_mapping_from_headers(headers)["mapping"]
        result = pstx_feishu_bom.sync_feishu_library(
            library_name="MLCC优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="mlcc_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "工业级和不区分优选库(用于选型)",
                **mapping,
            }],
            client=FakeClient(),
        )

        self.assertTrue(result["ok"])
        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="mlcc_lib",
            query="0201N0R5A500CT",
            data_dir=str(data_dir),
        )
        self.assertEqual(1, rows["total"])
        row = rows["rows"][0]
        self.assertEqual("HQ1710101A3Q0", row["hq_no"])
        self.assertEqual("0201N0R5A500CT", row["spec"])
        self.assertEqual("MLCC_0201_50V_0.5pF_A(±0.05pF)_C0G", row["description"])
        self.assertEqual("5", row["pi"])
        self.assertEqual("5", row["selection_order"])
        self.assertEqual("0201", row["extra_field_values"]["封装类型"])

    def test_sync_library_uses_actual_selection_order_column_from_suggestion(self):
        data_dir = Path(tempfile.mkdtemp())
        self.add_temp_cleanup(data_dir)

        class FakeClient:
            def read_sheet(self, token, sheet_id, *, row_count=5000, column_range="A:Z"):
                return [
                    ["规格型号", "HQ编码", "是否调整选型顺序", "PI", "选型顺序"],
                    ["CAP-100N", "HQ17101005", "是", "LiXinYu", "1"],
                ]

        preview_rows = FakeClient().read_sheet("token123", "sh1")
        mapping = pstx_feishu_bom.suggest_feishu_mapping_from_preview(preview_rows)["mapping"]
        result = pstx_feishu_bom.sync_feishu_library(
            library_name="电容优选库",
            spreadsheet_token_or_url="token123",
            base_url="https://mcenter.example.local",
            origin="cli_demo",
            user_id="100001",
            data_dir=str(data_dir),
            library_id="cap_lib",
            sheets=[{
                "sheet_id": "sh1",
                "title": "CAP",
                **mapping,
            }],
            client=FakeClient(),
        )

        self.assertTrue(result["ok"])
        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="cap_lib",
            query="HQ17101005",
            data_dir=str(data_dir),
        )
        self.assertEqual(1, rows["total"])
        self.assertEqual("1", rows["rows"][0]["selection_order"])
        self.assertEqual("是", rows["rows"][0]["extra_field_values"]["是否调整选型顺序"])

    def test_database_overview_rows_and_delete_library(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)

        overview = pstx_feishu_bom.build_feishu_database_overview(data_dir=str(data_dir))
        self.assertTrue(overview["ok"])
        self.assertEqual(1, len(overview["libraries"]))
        self.assertEqual("lib1", overview["libraries"][0]["lib_id"])
        self.assertEqual(2, overview["libraries"][0]["cache_count"])

        rows = pstx_feishu_bom.get_feishu_cache_rows(
            lib_id="lib1",
            query="ABC",
            data_dir=str(data_dir),
        )
        self.assertTrue(rows["ok"])
        self.assertEqual(1, rows["total"])
        self.assertEqual("ABC-123", rows["rows"][0]["key_value"])
        self.assertIn("raw_fields", rows["rows"][0])

        deleted = pstx_feishu_bom.delete_feishu_cache_library("lib1", data_dir=str(data_dir))
        self.assertTrue(deleted["ok"])
        self.assertEqual(2, deleted["deleted_rows"])
        after = pstx_feishu_bom.build_feishu_database_overview(data_dir=str(data_dir))
        self.assertEqual(0, after["cache_count"])

    def test_create_and_update_feishu_cache_row(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)

        created = pstx_feishu_bom.create_feishu_cache_row(
            {
                "lib_id": "lib1",
                "lib_name": "优选库",
                "sheet_name": "手工维护",
                "key_value": "CAP-100N",
                "hq_no": "HQ17101005",
                "pi": "LiXinYu",
                "selection_order": "1",
                "extra_fields": {"封装": "0402"},
            },
            data_dir=str(data_dir),
        )

        self.assertTrue(created["ok"])
        self.assertEqual("CAP-100N", created["row"]["key_value"])
        self.assertEqual({"封装": "0402"}, created["row"]["extra_field_values"])

        updated = pstx_feishu_bom.update_feishu_cache_row(
            created["row_id"],
            {
                "key_value": "CAP-220N",
                "hq_no": "HQ17101006",
                "selection_order": "2",
                "extra_fields": '{"封装":"0201","耐压":"6.3V"}',
            },
            data_dir=str(data_dir),
        )

        self.assertTrue(updated["ok"])
        self.assertEqual("CAP-220N", updated["row"]["key_value"])
        self.assertEqual("HQ17101006", updated["row"]["hq_no"])
        self.assertEqual("2", updated["row"]["selection_order"])
        self.assertEqual({"封装": "0201", "耐压": "6.3V"}, updated["row"]["extra_field_values"])

    def test_get_feishu_cache_rows_reports_pagination_state(self):
        data_dir = build_feishu_data_dir()
        self.add_temp_cleanup(data_dir)

        rows = pstx_feishu_bom.get_feishu_cache_rows(data_dir=str(data_dir), limit=1)

        self.assertTrue(rows["ok"])
        self.assertEqual(2, rows["total"])
        self.assertEqual(1, rows["limit"])
        self.assertTrue(rows["has_more"])
        self.assertEqual(1, rows["next_offset"])


if __name__ == "__main__":
    unittest.main()
