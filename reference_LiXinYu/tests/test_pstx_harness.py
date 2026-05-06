import copy
import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from pstx_harness import HarnessError, HarnessRunRequest, run_harness_review
from pstx_harness import compare_agent as harness_compare_agent
from pstx_harness import compare_tools as harness_compare_tools
from pstx_harness import model as harness_model_entry
from pstx_harness import report_agent as harness_report_agent
from pstx_harness import report_tools as harness_report_tools
from pstx_harness import review as harness_review_entry
from pstx_harness.model import AsterHarnessModelProvider
from pstx_harness.report_tools import HarnessToolContext, build_default_harness_registry
from pstx_harness.tool_core import HarnessToolError
from pstx_knowledge import component_identity as knowledge_component_identity
from pstx_knowledge import datasheet_review_templates as knowledge_datasheet_templates
from pstx_knowledge import datasheets as knowledge_datasheets
from pstx_knowledge import document_search as knowledge_document_search
from pstx_knowledge import feishu_cache as knowledge_feishu_cache
from pstx_knowledge import reference_library as knowledge_reference_library
from pstx_knowledge import topology as knowledge_topology
from pstx_integrations.feishu import gateway as pstx_feishu_bom


def sample_report():
    return {
        "project_name": "demo_board",
        "ratio_limit": 70,
        "include_depop": False,
        "metrics": [
            {"label": "DEPOP 总数", "value": 2},
            {"label": "BOM圈问题", "value": 1},
            {"label": "降额不合格", "value": 1},
        ],
        "sections": [
            {
                "id": "drc",
                "title": "设计检查",
                "total_rows": 2,
                "tables": [
                    {
                        "id": "bom_option_components",
                        "title": "BOM_OPTION 元件",
                        "count": 1,
                        "rows": [{"位号": "R1", "BOM_OPTION": "DEPOP", "页码": "12"}],
                    },
                    {
                        "id": "bom_option_circle_issues",
                        "title": "BOM_OPTION 打圈覆盖问题",
                        "count": 1,
                        "rows": [{"位号": "R1", "状态": "未打圈"}],
                    },
                    {
                        "id": "missing_value",
                        "title": "缺少 VALUE",
                        "count": 1,
                        "rows": [{"位号": "C1", "页码": "12"}],
                    },
                ],
            },
            {
                "id": "resistor",
                "title": "电阻检查",
                "total_rows": 1,
                "tables": [
                    {
                        "id": "chip_pin_rows",
                        "title": "芯片 Pin 电阻状态",
                        "count": 1,
                        "rows": [{"芯片位号": "U1", "引脚": "GPIO1", "状态": "候选判断"}],
                    }
                ],
            },
            {
                "id": "network",
                "title": "网络分析",
                "total_rows": 4,
                "tables": [
                    {
                        "id": "page_rows",
                        "title": "页码元件分布",
                        "count": 4,
                        "columns": ["页码", "元件数"],
                        "rows": [
                            {"页码": "PAGE1", "元件数": 3},
                            {"页码": "PAGE2", "元件数": 5},
                            {"页码": "PAGE2", "元件数": 7},
                            {"页码": "PAGE10", "元件数": 1},
                        ],
                    }
                ],
            },
            {
                "id": "derating",
                "title": "电容降额",
                "total_rows": 1,
                "tables": [
                    {
                        "id": "derating",
                        "title": "电容降额结果",
                        "count": 1,
                        "rows": [{"位号": "C1", "状态": "❌ 不合格"}],
                    }
                ],
            },
        ],
    }


def sample_dfmea_bundle():
    return {
        "components": {
            "U1": {
                "HQ_CODE": "HQ100",
                "CDS_PART_NAME": "GPU_CORE_TEST_IC",
                "PACKAGE": "BGA",
                "page_submodule_mapped": "12",
                "nets": {"A1": "P3V3", "B1": "I2C_SCL"},
            },
            "PU2": {
                "CDS_PART_NAME": "POWER_MANAGER",
                "PACKAGE": "QFN",
                "page_submodule_mapped": "18",
                "nets": {"1": "VIN_12V", "2": "VOUT_1V8"},
            },
            "J1": {
                "HQ_CODE": "HQ999",
                "CDS_PART_NAME": "USB_CONNECTOR",
                "page_submodule_mapped": "20",
                "nets": {"1": "USB_DP", "2": "USB_DN"},
            },
            "R1": {
                "VALUE": "10K",
                "PACKAGE": "0402",
                "page_real": "3",
                "nets": {"1": "P3V3", "2": "GPIO_BOOT"},
            },
        },
        "nets": {
            "P3V3": [{"refdes": "U1", "pin": "A1", "pin_name": "VDD"}, {"refdes": "R1", "pin": "1", "pin_name": "1"}],
            "I2C_SCL": [{"refdes": "U1", "pin": "B1", "pin_name": "SCL"}],
            "VIN_12V": [{"refdes": "PU2", "pin": "1", "pin_name": "VIN"}],
            "VOUT_1V8": [{"refdes": "PU2", "pin": "2", "pin_name": "VOUT"}],
            "USB_DP": [{"refdes": "J1", "pin": "1", "pin_name": "DP"}],
            "USB_DN": [{"refdes": "J1", "pin": "2", "pin_name": "DN"}],
            "GPIO_BOOT": [{"refdes": "R1", "pin": "2", "pin_name": "2"}],
        },
    }


def sample_chip_topology_bundle():
    return {
        "components": {
            "U1": {
                "HQ_CODE": "HQ-FPGA",
                "CDS_PART_NAME": "GPU_FPGA_CORE",
                "PACKAGE": "BGA",
                "page_submodule_mapped": "12",
                "nets": {"A1": "I2C_SCL", "A2": "I2C_SDA", "VDD": "P3V3"},
            },
            "U2": {
                "HQ_CODE": "HQ-LS",
                "CDS_PART_NAME": "TXS0108_LEVEL_TRANSLATOR",
                "PACKAGE": "QFN",
                "page_submodule_mapped": "14",
                "nets": {"A1": "I2C_SCL", "A2": "I2C_SDA", "B1": "I2C_SCL_1V8", "B2": "I2C_SDA_1V8", "VCCA": "P1V8", "VCCB": "P3V3"},
            },
            "PU1": {
                "CDS_PART_NAME": "LDO_POWER",
                "page_submodule_mapped": "9",
                "nets": {"IN": "P3V3", "OUT": "P1V8"},
            },
            "R1": {
                "VALUE": "22R",
                "page_submodule_mapped": "14",
                "nets": {"1": "I2C_SCL", "2": "I2C_SCL_FPGA"},
            },
        },
        "nets": {
            "I2C_SCL": [
                {"refdes": "U1", "pin": "A1", "pin_name": "SCL"},
                {"refdes": "U2", "pin": "A1", "pin_name": "SCL_A"},
                {"refdes": "R1", "pin": "1", "pin_name": "1"},
            ],
            "I2C_SDA": [
                {"refdes": "U1", "pin": "A2", "pin_name": "SDA"},
                {"refdes": "U2", "pin": "A2", "pin_name": "SDA_A"},
            ],
            "I2C_SCL_1V8": [{"refdes": "U2", "pin": "B1", "pin_name": "SCL_B"}],
            "I2C_SDA_1V8": [{"refdes": "U2", "pin": "B2", "pin_name": "SDA_B"}],
            "P3V3": [
                {"refdes": "U1", "pin": "VDD", "pin_name": "VDD"},
                {"refdes": "U2", "pin": "VCCB", "pin_name": "VCCB"},
                {"refdes": "PU1", "pin": "IN", "pin_name": "IN"},
            ],
            "P1V8": [
                {"refdes": "U2", "pin": "VCCA", "pin_name": "VCCA"},
                {"refdes": "PU1", "pin": "OUT", "pin_name": "OUT"},
            ],
        },
    }


def sample_passive_bridge_topology_bundle():
    return {
        "components": {
            "U1": {
                "CDS_PART_NAME": "GPU_FPGA_CORE",
                "page_submodule_mapped": "12",
                "nets": {"A1": "BOOT_CFG_A", "VDD": "P3V3"},
            },
            "U3": {
                "CDS_PART_NAME": "GPIO_EXPANDER",
                "page_submodule_mapped": "16",
                "nets": {"B1": "BOOT_CFG_B", "VDD": "P1V8"},
            },
            "R10": {
                "VALUE": "22R",
                "PACKAGE": "0402",
                "page_submodule_mapped": "15",
                "nets": {"1": "BOOT_CFG_A", "2": "BOOT_CFG_B"},
            },
        },
        "nets": {
            "BOOT_CFG_A": [
                {"refdes": "U1", "pin": "A1", "pin_name": "BOOT_CFG"},
                {"refdes": "R10", "pin": "1", "pin_name": "1"},
            ],
            "BOOT_CFG_B": [
                {"refdes": "U3", "pin": "B1", "pin_name": "BOOT_CFG_IN"},
                {"refdes": "R10", "pin": "2", "pin_name": "2"},
            ],
            "P3V3": [{"refdes": "U1", "pin": "VDD", "pin_name": "VDD"}],
            "P1V8": [{"refdes": "U3", "pin": "VDD", "pin_name": "VDD"}],
        },
    }


class HarnessTests(unittest.TestCase):
    def test_harness_package_entrypoints_export_public_api(self):
        self.assertIs(harness_review_entry.run_harness_review, run_harness_review)
        self.assertIs(harness_review_entry.HarnessRunRequest, HarnessRunRequest)
        self.assertTrue(callable(harness_model_entry.MockHarnessModelProvider))
        self.assertTrue(callable(harness_report_tools.build_default_harness_registry))
        self.assertTrue(callable(harness_report_agent.run_harness_agent))
        self.assertTrue(callable(harness_compare_tools.build_compare_tool_registry))
        self.assertTrue(callable(harness_compare_agent.run_compare_agent))
        self.assertFalse(Path("pstx_harness.py").exists())
        self.assertFalse(Path("pstx_harness_agent.py").exists())
        self.assertFalse(Path("pstx_harness_tools.py").exists())
        self.assertFalse(Path("pstx_harness_model.py").exists())
        self.assertFalse(Path("pstx_compare_agent.py").exists())
        self.assertFalse(Path("pstx_compare_tools.py").exists())

    def test_knowledge_package_entrypoints_export_public_api(self):
        self.assertIs(knowledge_feishu_cache.get_feishu_cache_rows, pstx_feishu_bom.get_feishu_cache_rows)
        self.assertTrue(callable(knowledge_component_identity.build_component_identity_cards))
        self.assertTrue(callable(knowledge_component_identity.classify_refdes))
        self.assertTrue(callable(knowledge_datasheets.search_datasheet_chunks))
        self.assertTrue(callable(knowledge_datasheet_templates.list_datasheet_review_templates))
        self.assertTrue(callable(knowledge_reference_library.search_agent_ref))
        self.assertTrue(callable(knowledge_document_search.search_documents))
        self.assertTrue(callable(knowledge_topology.build_chip_topology))
        self.assertFalse(Path("pstx_component_identity.py").exists())
        self.assertFalse(Path("pstx_datasheet_index.py").exists())
        self.assertFalse(Path("pstx_agent_ref.py").exists())
        self.assertFalse(Path("pstx_document_search.py").exists())
        self.assertFalse(Path("pstx_chip_topology.py").exists())
        self.assertFalse(hasattr(knowledge_feishu_cache, "fetch_feishu_sheet_list"))

    def test_datasheet_review_template_tools_are_readonly_and_llm_readable(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_dfmea_bundle(), request=HarnessRunRequest())

        listed = registry.run("list_datasheet_review_templates", context, {"category": "complex_chip"})
        self.assertTrue(listed["readonly"])
        self.assertEqual("pstx-datasheet-review-template.v1", listed["schema_version"])
        self.assertEqual(1, listed["template_count"])
        self.assertEqual("complex_chip", listed["templates"][0]["template_id"])
        self.assertIn("review_playbook", listed["templates"][0])

        detail = registry.run("get_datasheet_review_template", context, {"template_id": "power_regulator"})
        self.assertTrue(detail["readonly"])
        self.assertEqual("power_regulator", detail["template"]["template_id"])
        self.assertIn("required_evidence", detail["template"])

    def make_feishu_cache(self) -> Path:
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        old_env = os.environ.get("PSTX_FEISHU_DATA_DIR")
        self.addCleanup(
            lambda: os.environ.pop("PSTX_FEISHU_DATA_DIR", None)
            if old_env is None else os.environ.__setitem__("PSTX_FEISHU_DATA_DIR", old_env)
        )
        os.environ["PSTX_FEISHU_DATA_DIR"] = str(root)
        (root / "feishu_libraries.json").write_text(
            json.dumps({
                "base_url": "https://mcenter.example.local",
                "origin": "cli_demo",
                "user_id": "100001",
                "libraries": [{
                    "id": "lib1",
                    "name": "优选库",
                    "token": "secret-token-should-not-leak",
                    "sheets": [{
                        "sheet_id": "sh1",
                        "title": "电容",
                        "header_row": 1,
                        "hq_code_col": "HQ料号",
                        "spec_model_col": "规格型号",
                        "pi_col": "PI",
                        "selection_order_col": "选型顺序",
                    }],
                }],
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
                pi TEXT,
                selection_order TEXT,
                extra_fields TEXT,
                raw_data TEXT,
                synced_at TEXT
            )
            """
        )
        conn.executemany(
            "INSERT INTO materials(lib_id,lib_name,sheet_name,key_value,hq_no,brand,spec,description,pi,selection_order,extra_fields,raw_data,synced_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [
                (
                    "lib1", "优选库", "电容", "CAP-100N", "HQ100", "ACME", "CAP-100N",
                    "100nF capacitor", "LiXinYu", "1",
                    json.dumps({"封装": "0402", "耐压": "50V"}, ensure_ascii=False),
                    json.dumps({"封装": "0402", "耐压": "50V", "备注": "preferred"}, ensure_ascii=False),
                    "2026-04-27",
                ),
                (
                    "lib1", "优选库", "电阻", "RES-10K", "HQ200", "ACME", "RES-10K",
                    "10K resistor", "ZhangSan", "2",
                    json.dumps({"封装": "0402", "精度": "1%"}, ensure_ascii=False),
                    json.dumps({"封装": "0402", "精度": "1%"}, ensure_ascii=False),
                    "2026-04-27",
                ),
            ],
        )
        conn.commit()
        conn.close()
        return root

    def make_harness_doc_dir(self) -> Path:
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        old_env = os.environ.get("PSTX_HARNESS_DOC_DIR")
        self.addCleanup(
            lambda: os.environ.pop("PSTX_HARNESS_DOC_DIR", None)
            if old_env is None else os.environ.__setitem__("PSTX_HARNESS_DOC_DIR", old_env)
        )
        os.environ["PSTX_HARNESS_DOC_DIR"] = str(root)
        (root / "review_notes.md").write_text(
            "# Review Notes\n\nU46 多 symbol 芯片需要检查每个 SECTION_NUMBER 的 HQ_CODE 和页码。\n"
            "电平转换芯片需要关注 I2C_SCL/I2C_SDA 的方向和电压域。\n",
            encoding="utf-8",
        )
        return root

    def test_mock_provider_completes_full_review(self):
        payload = run_harness_review(sample_report(), {}, HarnessRunRequest())

        self.assertTrue(payload["ok"])
        self.assertEqual("local-harness", payload["mode"])
        self.assertEqual("full_review", payload["task"])
        self.assertEqual("local-harness-mock", payload["model_metadata"]["provider"])
        self.assertTrue(payload["evidence_packs"])
        self.assertTrue(any(pack["id"] == "bom_depop" for pack in payload["evidence_packs"]))
        self.assertTrue(all(run["ok"] for run in payload["tool_runs"]))
        self.assertTrue(payload["review_checklist"])

    def test_registry_rejects_unknown_tool(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())

        with self.assertRaises(ValueError):
            registry.run("unknown_tool", context)

    def test_project_memory_evidence_tools_are_read_only(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(
            report=sample_report(),
            bundle={},
            request=HarnessRunRequest(),
            project_context={
                "evidence_memory_cards": [{
                    "id": "ev-u1",
                    "type": "component_identity",
                    "title": "U1 身份卡",
                    "summary": "U1 HQ=HQ100。",
                    "locator": {"refdes": "U1"},
                    "detail_tool": {"name": "get_component_identity_card", "args": {"refdes": "U1"}},
                }]
            },
        )

        listed = registry.run("list_project_memory_evidence", context, {"query": "HQ100"})
        detail = registry.run("get_project_memory_evidence", context, {"evidence_id": "ev-u1"})
        batch = registry.run("batch_get_project_memory_evidence", context, {"evidence_ids": ["ev-u1", "missing"]})

        self.assertEqual(1, listed["total_matches"])
        self.assertTrue(detail["found"])
        self.assertEqual("found", batch["items"][0]["status"])
        self.assertEqual("missing", batch["items"][1]["status"])

    def test_harness_skill_tools_are_read_only_and_return_skill_cards(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())
        tools = {item["name"]: item for item in registry.list_tools()}

        for name in ("list_harness_skills", "select_harness_skills", "get_harness_skill"):
            self.assertIn(name, tools)
            self.assertTrue(tools[name]["readonly"])
            self.assertEqual("harness_skill", tools[name]["evidence_kind"])
            self.assertEqual("none", tools[name]["approval_scope"])

        listed = registry.run("list_harness_skills", context, {"limit": 5})
        self.assertEqual("pstx-harness-skills/v1", listed["harness_skills"]["schema_version"])
        self.assertTrue(listed["harness_skills"]["skills"])
        self.assertIn("guidance only", listed["harness_skills"]["guidance_note"])

        selected = registry.run(
            "select_harness_skills",
            context,
            {
                "query": "请按 MinerU 读取 64144 datasheet 关键参数",
                "capability_profiles": ["datasheet_qa"],
                "include_body": True,
                "max_body_chars": 1200,
            },
        )
        selected_ids = [card["id"] for card in selected["harness_skills"]["skills"]]
        self.assertIn("datasheet-key-info", selected_ids)

        selected_connection = registry.run(
            "select_harness_skills",
            context,
            {
                "query": "请根据 MinerU datasheet 反查 U1 和 U2 的接口电平连接是否有问题",
                "capability_profiles": ["connection_datasheet_review"],
                "include_body": True,
                "max_body_chars": 1600,
            },
        )
        connection_ids = [card["id"] for card in selected_connection["harness_skills"]["skills"]]
        self.assertIn("schematic-datasheet-connection-review", connection_ids)
        connection_body = "\n".join(card.get("body", "") for card in selected_connection["harness_skills"]["skills"])
        self.assertIn("原理图/网表 evidence", connection_body)
        self.assertIn("MinerU-backed datasheet evidence", connection_body)

        detail = registry.run(
            "get_harness_skill",
            context,
            {"skill_id": "datasheet-key-info", "max_body_chars": 1600},
        )
        self.assertEqual("datasheet-key-info", detail["skill"]["id"])
        self.assertIn("已确认的 datasheet 事实", detail["skill"]["body"])
        self.assertIn("search_datasheet_chunks", detail["recommended_next_tools"])

        with self.assertRaises(HarnessToolError):
            registry.run("get_harness_skill", context, {"skill_id": "missing-skill"})

    def test_feishu_cache_tools_list_search_and_read_rows_without_leaking_token(self):
        root = self.make_feishu_cache()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())

        listed = registry.run("list_feishu_cache_libraries", context, {"include_sheets": True})
        self.assertTrue(listed["available"])
        self.assertEqual(1, listed["library_count"])
        self.assertEqual(2, listed["cache_count"])
        self.assertNotIn("secret-token", json.dumps(listed, ensure_ascii=False))

        by_hq = registry.run("search_feishu_cache_rows", context, {"query": "HQ100", "limit": 10})
        by_pi = registry.run("search_feishu_cache_rows", context, {"query": "LiXinYu", "limit": 10})
        by_extra = registry.run("search_feishu_cache_rows", context, {"query": "50V", "limit": 10})
        self.assertEqual(1, by_hq["total_rows"])
        self.assertEqual("HQ100", by_hq["rows"][0]["hq_no"])
        self.assertEqual(1, by_pi["total_rows"])
        self.assertEqual(1, by_extra["total_rows"])
        self.assertNotIn("raw_data", by_hq["rows"][0])

        detail = registry.run("get_feishu_cache_row", context, {"row_id": by_hq["rows"][0]["id"]})
        self.assertTrue(detail["ok"])
        self.assertEqual("CAP-100N", detail["row"]["spec"])
        self.assertEqual("1", detail["row"]["selection_order"])

        missing_detail = registry.run("get_feishu_cache_row", context, {"row_id": 999})
        self.assertFalse(missing_detail["ok"])
        self.assertIn("未找到缓存行", missing_detail["summary"])

        conn = sqlite3.connect(root / "feishu_cache.db")
        try:
            count_after = conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(2, count_after)

    def test_feishu_cache_search_reports_missing_or_empty_results(self):
        self.make_feishu_cache()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())

        result = registry.run("search_feishu_cache_rows", context, {"query": "NO_SUCH_MATERIAL"})

        self.assertTrue(result["ok"])
        self.assertEqual(0, result["total_rows"])
        self.assertIn("无命中", result["summary"])

    def test_component_identity_tools_enrich_and_summarize_dfmea_inputs(self):
        root = self.make_feishu_cache()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_dfmea_bundle(), request=HarnessRunRequest())

        listed = registry.run("list_component_identity_cards", context, {"category": "chip", "limit": 10})
        self.assertEqual(1, listed["total_cards"])
        self.assertEqual("U1", listed["cards"][0]["refdes"])
        self.assertEqual("matched", listed["cards"][0]["feishu_match"]["status"])
        self.assertEqual("LiXinYu", listed["cards"][0]["pi"])
        self.assertEqual("1", listed["cards"][0]["selection_order"])

        detail = registry.run("get_component_identity_card", context, {"refdes": "U1"})
        self.assertEqual("component.nets", detail["card"]["pin_net_summary"][0]["source"])
        self.assertIn("P3V3", detail["card"]["power_nets"])
        self.assertIn("I2C_SCL", detail["card"]["interface_nets"])
        self.assertNotIn("hq_no", detail["card"]["missing_fields"])

        searched = registry.run("search_component_identity_cards", context, {"query": "I2C_SCL", "limit": 5})
        self.assertEqual(1, searched["total_cards"])
        self.assertEqual("U1", searched["cards"][0]["refdes"])

        readiness = registry.run("summarize_dfmea_readiness", context, {})
        self.assertEqual(4, readiness["total_components"])
        self.assertGreaterEqual(readiness["ready_count"], 1)
        self.assertTrue(any(card["refdes"] == "PU2" for card in readiness["needs_context_cards"]))
        self.assertGreaterEqual(readiness["missing_counts"].get("hq_no", 0), 1)

        conn = sqlite3.connect(root / "feishu_cache.db")
        try:
            count_after = conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(2, count_after)

    def test_chip_topology_tools_build_fuzzy_chip_level_edges(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_chip_topology_bundle(), request=HarnessRunRequest())

        topology = registry.run("summarize_chip_topology", context, {"limit": 10})

        self.assertTrue(topology["ok"])
        self.assertEqual("llm-topology.v1", topology["schema_version"])
        self.assertIn("summary_layer", topology)
        self.assertIn("evidence_cards", topology)
        self.assertIn("raw_layer", topology)
        self.assertEqual(3, topology["node_count"])
        self.assertEqual(1, topology["edge_count"])
        refs = {node["refdes"] for node in topology["nodes"]}
        self.assertIn("U1", refs)
        self.assertIn("U2", refs)
        self.assertNotIn("R1", refs)
        node_u1 = next(node for node in topology["nodes"] if node["refdes"] == "U1")
        self.assertEqual("llm-topology-node-u1", node_u1["evidence_id"])
        self.assertEqual("get_llm_topology_node", node_u1["detail_tool"]["name"])
        self.assertIn("P3V3", node_u1["power_nets"])
        self.assertEqual(["3V3"], node_u1["voltage_domains"])
        self.assertEqual("server_hardware", node_u1["llm_device_identity_hint"]["domain_context"])
        self.assertIn("processor_or_fpga", node_u1["llm_device_identity_hint"]["server_device_taxonomy"])
        edge = topology["edges"][0]
        self.assertEqual("llm-topology-edge-chip-edge-u1-u2", edge["evidence_id"])
        self.assertEqual("芯片到电平转换连接", edge["relation_label"])
        self.assertEqual(2, edge["shared_net_count"])
        self.assertIn("i2c", edge["interface_groups"])
        self.assertTrue(edge["interface_completeness"])
        self.assertEqual("observed_required", edge["interface_completeness"][0]["status"])
        self.assertEqual("high", edge["confidence"])
        self.assertEqual("high", edge["review_priority"])
        self.assertIn("voltage_domain_transition", edge["risk_tags"])
        self.assertTrue(edge["interface_summary"])
        self.assertEqual("get_llm_topology_edge", edge["detail_tool"]["name"])
        self.assertTrue(edge["review_hints"])
        self.assertTrue(topology["role_links"])
        self.assertGreaterEqual(topology["supply_edge_count"], 1)
        self.assertTrue(topology["evidence_cards"]["supply_edges"])
        self.assertEqual("llm-topology-business-view.v1", topology["business_view"]["schema_version"])
        self.assertIn("pcie", topology["business_view"]["dictionary"]["interface_aliases"])
        self.assertIn("coverage_gaps", {item["partition_id"] for item in topology["business_view"]["review_partitions"]})
        self.assertEqual(topology["counts"], topology["summary_layer"]["counts"])
        self.assertEqual(topology["node_count"], topology["counts"]["total_node_count"])
        self.assertGreaterEqual(topology["review_task_count"], 1)
        self.assertTrue(topology["review_tasks"])
        self.assertTrue(topology["evidence_cards"]["review_tasks"])

        llm_topology = registry.run("summarize_llm_topology_netlist", context, {"limit": 10})
        self.assertEqual("summarize_llm_topology_netlist", llm_topology["id"])
        self.assertEqual("llm-topology.v1", llm_topology["schema_version"])
        self.assertEqual(1, llm_topology["evidence_cards"]["edges"][0]["detail_tool"]["args"]["edge_id"].count("chip-edge"))
        self.assertIn("business_view", llm_topology)

        dictionary = registry.run("list_business_dictionary", context, {})
        self.assertTrue(dictionary["ok"])
        self.assertIn("PCE", dictionary["dictionary"]["interface_aliases"]["pcie"])

        queried = registry.run("query_chip_topology", context, {"query": "U1 电平转换", "limit": 5})
        self.assertGreaterEqual(queried["total_matches"], 1)
        self.assertIn("edge", {item["kind"] for item in queried["items"]})

        llm_queried = registry.run("query_llm_topology_netlist", context, {"query": "U1 I2C", "limit": 5})
        self.assertEqual("llm-topology.v1", llm_queried["schema_version"])
        self.assertGreaterEqual(llm_queried["total_matches"], 1)

        node_detail = registry.run("get_llm_topology_node", context, {"refdes": "U1"})
        self.assertEqual("U1", node_detail["node"]["refdes"])
        self.assertTrue(node_detail["pin_nets"])

        detail = registry.run("get_chip_topology_edge", context, {"edge_id": edge["edge_id"]})
        self.assertEqual(edge["edge_id"], detail["edge"]["edge_id"])
        self.assertEqual(2, len(detail["edge"]["shared_nets"]))

        llm_detail = registry.run("get_llm_topology_edge", context, {"edge_id": edge["edge_id"]})
        self.assertEqual(edge["edge_id"], llm_detail["edge"]["edge_id"])
        self.assertEqual("llm-topology.v1", llm_detail["schema_version"])

        review_tasks = registry.run("summarize_topology_review_tasks", context, {"limit": 10})
        self.assertTrue(review_tasks["ok"])
        self.assertEqual("llm-topology-review-task.v1", review_tasks["schema_version"])
        self.assertGreaterEqual(review_tasks["total_count"], 1)
        task_id = review_tasks["tasks"][0]["task_id"]
        task_detail = registry.run("get_topology_review_task", context, {"task_id": task_id})
        self.assertEqual(task_id, task_detail["task"]["task_id"])
        self.assertIn(task_detail["task"]["source_kind"], {"signal_edge", "supply_edge", "node"})
        task_batch = registry.run("batch_expand_topology_review_tasks", context, {"task_ids": [task_id, "NO_SUCH_TASK"]})
        self.assertEqual(2, task_batch["query_count"])
        self.assertEqual("found", task_batch["items"][0]["status"])
        self.assertEqual("missing", task_batch["items"][1]["status"])

        batch = registry.run("batch_query_chip_topology", context, {"queries": ["U1", "NO_SUCH"], "limit_per_query": 3})
        self.assertEqual(2, batch["query_count"])
        self.assertEqual("found", batch["items"][0]["status"])
        self.assertEqual("missing", batch["items"][1]["status"])

        llm_batch = registry.run("batch_query_llm_topology_netlist", context, {"queries": ["U1", "NO_SUCH"], "limit_per_query": 3})
        self.assertEqual("llm-topology.v1", llm_batch["schema_version"])
        self.assertEqual("found", llm_batch["items"][0]["status"])

    def test_llm_topology_keeps_passive_bridge_as_edge_evidence_not_node(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_passive_bridge_topology_bundle(), request=HarnessRunRequest())

        topology = registry.run("summarize_llm_topology_netlist", context, {"limit": 10})

        self.assertTrue(topology["ok"])
        self.assertEqual(2, topology["node_count"])
        self.assertEqual(1, topology["edge_count"])
        self.assertNotIn("R10", {node["refdes"] for node in topology["nodes"]})
        edge = topology["edges"][0]
        self.assertEqual(0, edge["shared_net_count"])
        self.assertEqual(1, edge["passive_bridge_count"])
        self.assertEqual("R10", edge["passive_bridges"][0]["refdes"])
        self.assertEqual("resistive_or_series_path", edge["passive_bridges"][0]["bridge_semantics"])
        self.assertTrue(edge["passive_bridges"][0]["dc_conductive"])
        self.assertIn("一跳无源桥", " ".join(edge["review_hints"]))

    def test_llm_topology_skips_common_power_names_and_exposes_supply_edges(self):
        bundle = {
            "components": {
                "PU1": {"CDS_PART_NAME": "BUCK_REGULATOR", "nets": {"OUT": "VCORE", "GND": "GNDA", "EN": "PWR_EN"}},
                "U1": {"CDS_PART_NAME": "GPU_CORE_TEST_IC", "nets": {"VDD": "VCORE", "AVDD": "AVDD", "GND": "GNDA"}},
                "U2": {"CDS_PART_NAME": "AUXILIARY_IC", "nets": {"VDD": "P0V8"}},
                "U3": {"CDS_PART_NAME": "AUXILIARY_IC", "nets": {"VDD": "VCC3V3"}},
            },
            "nets": {
                "VCORE": [{"refdes": "PU1", "pin": "OUT"}, {"refdes": "U1", "pin": "VDD"}],
                "AVDD": [{"refdes": "U1", "pin": "AVDD"}],
                "GNDA": [{"refdes": "PU1", "pin": "GND"}, {"refdes": "U1", "pin": "GND"}],
                "PWR_EN": [{"refdes": "PU1", "pin": "EN"}],
                "P0V8": [{"refdes": "U2", "pin": "VDD"}],
                "VCC3V3": [{"refdes": "U3", "pin": "VDD"}],
            },
        }

        topology = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, limit=20)
        nodes_by_ref = {node["refdes"]: node for node in topology["nodes"]}

        self.assertEqual(0, topology["edge_count"])
        self.assertEqual(1, topology["supply_edge_count"])
        self.assertEqual({"VCORE", "AVDD", "GNDA", "P0V8", "VCC3V3"}, set(topology["skipped_power_nets_sample"]))
        self.assertEqual("supply", topology["supply_edges"][0]["edge_kind"])
        self.assertEqual("PU1", topology["supply_edges"][0]["source_refdes"])
        self.assertEqual("U1", topology["supply_edges"][0]["target_refdes"])
        self.assertIn("PWR_EN", topology["supply_edges"][0]["source_control_nets"])
        self.assertIn("上电时序", topology["supply_edges"][0]["review_focus"])
        self.assertEqual(["0V8"], nodes_by_ref["U2"]["voltage_domains"])
        self.assertEqual(["3V3"], nodes_by_ref["U3"]["voltage_domains"])

    def test_llm_topology_groups_large_supply_fanout_by_default_and_keeps_full_detail(self):
        components = {
            "PU1": {"CDS_PART_NAME": "BUCK_REGULATOR", "nets": {"OUT": "P3V3", "EN": "PWR_EN"}},
        }
        nets = {
            "P3V3": [{"refdes": "PU1", "pin": "OUT"}],
            "PWR_EN": [{"refdes": "PU1", "pin": "EN"}],
        }
        for index in range(120):
            ref = f"U{index:03d}"
            components[ref] = {"CDS_PART_NAME": "AUXILIARY_TEST_IC", "nets": {"VDD": "P3V3"}}
            nets["P3V3"].append({"refdes": ref, "pin": "VDD"})
        bundle = {"components": components, "nets": nets}

        topology = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, limit=20, supply_limit=5)

        self.assertTrue(topology["truncated"])
        self.assertEqual(120, topology["supply_edge_count"])
        self.assertEqual(5, topology["returned_supply_edge_count"])
        self.assertEqual(1, topology["returned_supply_group_count"])
        self.assertEqual(1, len(topology["supply_edge_groups"]))
        self.assertEqual(120, topology["supply_edge_groups"][0]["target_count"])
        self.assertLess(len(topology["supply_edges"]), topology["supply_edge_count"])
        self.assertEqual(6, topology["counts"]["visual_edge_count"])

        full = knowledge_topology.build_llm_topology_netlist(
            sample_report(),
            bundle,
            view="full",
            return_all_edges=True,
            supply_mode="details",
            supply_limit=5,
        )
        self.assertEqual(120, len(full["supply_edges"]))
        detail = knowledge_topology.get_llm_topology_edge(sample_report(), bundle, "supply-edge-pu1-u119-p3v3")
        self.assertTrue(detail["ok"])
        self.assertEqual("U119", detail["edge"]["target_refdes"])

    def test_llm_topology_uses_derived_cache_and_invalidates_by_params(self):
        old_cache = os.environ.get("PSTX_ANALYSIS_CACHE_DIR")
        old_disable = os.environ.get("PSTX_DISABLE_ANALYSIS_CACHE")
        cache_dir = tempfile.mkdtemp()
        os.environ["PSTX_ANALYSIS_CACHE_DIR"] = cache_dir
        os.environ.pop("PSTX_DISABLE_ANALYSIS_CACHE", None)
        try:
            bundle = {
                "project_root": "/tmp/topology-cache-demo",
                "components": {
                    "PU1": {"CDS_PART_NAME": "BUCK_REGULATOR", "nets": {"OUT": "P1V8"}},
                    "U1": {"CDS_PART_NAME": "AUXILIARY_TEST_IC", "nets": {"VDD": "P1V8"}},
                },
                "nets": {
                    "P1V8": [{"refdes": "PU1", "pin": "OUT"}, {"refdes": "U1", "pin": "VDD"}],
                },
            }
            first = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, supply_limit=4)
            second = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, supply_limit=4)
            changed = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, supply_limit=5)
            os.environ["PSTX_DISABLE_ANALYSIS_CACHE"] = "1"
            disabled = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, supply_limit=4)

            self.assertFalse(first["topology_cache_status"]["hit"])
            self.assertTrue(second["topology_cache_status"]["hit"])
            self.assertFalse(changed["topology_cache_status"]["hit"])
            self.assertEqual("disabled", disabled["topology_cache_status"]["status"])
        finally:
            if old_cache is None:
                os.environ.pop("PSTX_ANALYSIS_CACHE_DIR", None)
            else:
                os.environ["PSTX_ANALYSIS_CACHE_DIR"] = old_cache
            if old_disable is None:
                os.environ.pop("PSTX_DISABLE_ANALYSIS_CACHE", None)
            else:
                os.environ["PSTX_DISABLE_ANALYSIS_CACHE"] = old_disable

    def test_llm_topology_interface_grouping_avoids_common_false_matches(self):
        bundle = {
            "components": {
                "U1": {
                    "CDS_PART_NAME": "GPU_FPGA_CORE",
                    "nets": {
                        "A1": "SPI_SCLK", "A2": "ADC_CS_N", "A3": "DP_AUX_P", "A4": "OPEN_DRAIN",
                        "A5": "PCIE_TX_P", "A6": "RGMII_TXD0", "A7": "JTAG_TCK", "A8": "SDIO_CMD",
                        "A9": "I2S_BCLK", "A10": "ADC_SENSE", "A11": "PCE_RX_P", "A12": "P5E_REFCLK_P",
                    },
                },
                "U2": {"CDS_PART_NAME": "SPI_FLASH", "nets": {"B1": "SPI_SCLK", "B2": "ADC_CS_N"}},
                "U3": {"CDS_PART_NAME": "TYPEC_MUX", "nets": {"C1": "DP_AUX_P"}},
                "U4": {"CDS_PART_NAME": "GPIO_EXPANDER", "nets": {"D1": "OPEN_DRAIN"}},
                "U5": {"CDS_PART_NAME": "PCIE_SWITCH", "nets": {"E1": "PCIE_TX_P"}},
                "U6": {"CDS_PART_NAME": "ETH_PHY", "nets": {"F1": "RGMII_TXD0"}},
                "U7": {"CDS_PART_NAME": "DEBUG_MCU", "nets": {"G1": "JTAG_TCK"}},
                "U8": {"CDS_PART_NAME": "EMMC_MEMORY", "nets": {"H1": "SDIO_CMD"}},
                "U9": {"CDS_PART_NAME": "AUDIO_CODEC", "nets": {"I1": "I2S_BCLK"}},
                "U10": {"CDS_PART_NAME": "SENSOR_ADC", "nets": {"J1": "ADC_SENSE"}},
                "U11": {"CDS_PART_NAME": "PCIE_ENDPOINT", "nets": {"K1": "PCE_RX_P"}},
                "U12": {"CDS_PART_NAME": "PCIE_CLOCK_BUFFER", "nets": {"L1": "P5E_REFCLK_P"}},
            },
            "nets": {
                "SPI_SCLK": [{"refdes": "U1", "pin": "A1"}, {"refdes": "U2", "pin": "B1"}],
                "ADC_CS_N": [{"refdes": "U1", "pin": "A2"}, {"refdes": "U2", "pin": "B2"}],
                "DP_AUX_P": [{"refdes": "U1", "pin": "A3"}, {"refdes": "U3", "pin": "C1"}],
                "OPEN_DRAIN": [{"refdes": "U1", "pin": "A4"}, {"refdes": "U4", "pin": "D1"}],
                "PCIE_TX_P": [{"refdes": "U1", "pin": "A5"}, {"refdes": "U5", "pin": "E1"}],
                "RGMII_TXD0": [{"refdes": "U1", "pin": "A6"}, {"refdes": "U6", "pin": "F1"}],
                "JTAG_TCK": [{"refdes": "U1", "pin": "A7"}, {"refdes": "U7", "pin": "G1"}],
                "SDIO_CMD": [{"refdes": "U1", "pin": "A8"}, {"refdes": "U8", "pin": "H1"}],
                "I2S_BCLK": [{"refdes": "U1", "pin": "A9"}, {"refdes": "U9", "pin": "I1"}],
                "ADC_SENSE": [{"refdes": "U1", "pin": "A10"}, {"refdes": "U10", "pin": "J1"}],
                "PCE_RX_P": [{"refdes": "U1", "pin": "A11"}, {"refdes": "U11", "pin": "K1"}],
                "P5E_REFCLK_P": [{"refdes": "U1", "pin": "A12"}, {"refdes": "U12", "pin": "L1"}],
            },
        }

        topology = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, limit=20)
        by_pair = {
            (edge["source_refdes"], edge["target_refdes"]): set(edge["interface_groups"])
            for edge in topology["edges"]
        }

        self.assertIn("spi", by_pair[("U1", "U2")])
        self.assertNotIn("i2c", by_pair[("U1", "U2")])
        self.assertEqual({"high_speed"}, by_pair[("U1", "U3")])
        self.assertEqual({"misc_signal"}, by_pair[("U1", "U4")])
        self.assertEqual({"pcie"}, by_pair[("U1", "U5")])
        self.assertEqual({"ethernet"}, by_pair[("U1", "U6")])
        self.assertEqual({"jtag_debug"}, by_pair[("U1", "U7")])
        self.assertEqual({"storage_sdio"}, by_pair[("U1", "U8")])
        self.assertEqual({"audio"}, by_pair[("U1", "U9")])
        self.assertEqual({"analog_sense"}, by_pair[("U1", "U10")])
        self.assertEqual({"pcie"}, by_pair[("U1", "U11")])
        self.assertEqual({"pcie"}, by_pair[("U1", "U12")])
        pcie_edge = next(edge for edge in topology["edges"] if edge["target_refdes"] == "U5")
        self.assertIn("阻抗", pcie_edge["review_focus"])
        business_high_speed = next(item for item in topology["business_view"]["review_partitions"] if item["partition_id"] == "high_speed_interfaces")
        self.assertGreaterEqual(business_high_speed["item_count"], 3)
        alias_query = knowledge_topology.query_llm_topology_netlist(sample_report(), bundle, "U1 PCI-E", limit=5)
        self.assertGreaterEqual(alias_query["total_matches"], 1)

    def test_llm_topology_role_filter_keeps_related_peers(self):
        topology = knowledge_topology.build_llm_topology_netlist(
            sample_report(),
            sample_chip_topology_bundle(),
            role_filter="level_shifter",
            limit=10,
        )

        self.assertEqual(1, topology["edge_count"])
        self.assertTrue({"U1", "U2"}.issubset({node["refdes"] for node in topology["nodes"]}))
        self.assertEqual("芯片到电平转换连接", topology["edges"][0]["relation_label"])

    def test_llm_topology_query_and_detail_search_full_edge_index(self):
        components = {}
        nets = {}
        for index in range(120):
            left = f"U{index * 2:03d}"
            right = f"U{index * 2 + 1:03d}"
            net = f"SIG_{index:03d}"
            components[left] = {"CDS_PART_NAME": "GPU_CORE_TEST_IC", "nets": {"A1": net}}
            components[right] = {"CDS_PART_NAME": "PERIPHERAL_TEST_IC", "nets": {"B1": net}}
            nets[net] = [{"refdes": left, "pin": "A1"}, {"refdes": right, "pin": "B1"}]
        bundle = {"components": components, "nets": nets}

        summary = knowledge_topology.build_llm_topology_netlist(sample_report(), bundle, limit=20)
        queried = knowledge_topology.query_llm_topology_netlist(sample_report(), bundle, "U238", limit=5)
        detail = knowledge_topology.get_llm_topology_edge(sample_report(), bundle, "chip-edge-u238-u239")

        self.assertTrue(summary["truncated"])
        self.assertEqual("found" if queried["total_matches"] else "missing", "found")
        self.assertTrue(any(item.get("edge", {}).get("edge_id") == "chip-edge-u238-u239" for item in queried["items"]))
        self.assertTrue(detail["ok"])
        self.assertEqual("chip-edge-u238-u239", detail["edge"]["edge_id"])

    def test_document_search_tools_find_keyword_and_excerpt(self):
        self.make_harness_doc_dir()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())

        status = registry.run("list_document_search_sources", context, {})
        self.assertGreaterEqual(status["document_count"], 1)

        searched = registry.run("search_documents", context, {"query": "U46 SECTION_NUMBER", "limit": 10})
        self.assertGreaterEqual(searched["total_matches"], 1)
        match = next(item for item in searched["matches"] if item["title"] == "review_notes.md")
        self.assertEqual("review_notes.md", match["title"])
        self.assertIn("U46", match["snippet"])

        excerpt = registry.run(
            "get_document_excerpt",
            context,
            {"doc_id": match["doc_id"], "char_start": match["char_start"], "max_chars": 1000},
        )
        self.assertIn("SECTION_NUMBER", excerpt["excerpt"])

        batch = registry.run("batch_search_documents", context, {"queries": ["电平转换", "NO_SUCH"], "limit_per_query": 3})
        self.assertEqual("found", batch["items"][0]["status"])
        self.assertEqual("missing", batch["items"][1]["status"])

    def test_document_search_ignores_symlinks_outside_configured_root(self):
        root = self.make_harness_doc_dir()
        outside = root.parent / "outside_notes.md"
        outside.write_text("OUTSIDE_SECRET_U46", encoding="utf-8")
        try:
            (root / "linked_outside.md").symlink_to(outside)
        except (OSError, NotImplementedError):
            self.skipTest("filesystem does not support symlinks")

        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())
        searched = registry.run("search_documents", context, {"query": "OUTSIDE_SECRET_U46", "limit": 10})

        self.assertEqual(0, searched["total_matches"])

    def test_read_project_text_redacts_secret_like_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "runtime.txt").write_text(
                "apiKey=abc123 password=letmein normal=value",
                encoding="utf-8",
            )
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            result = registry.run("read_project_text", context, {"path": "packaged/runtime.txt"})

        self.assertIn("normal=value", result["content"])
        self.assertIn("apiKey=<redacted>", result["content"])
        self.assertIn("password=<redacted>", result["content"])
        self.assertNotIn("abc123", result["content"])
        self.assertNotIn("letmein", result["content"])

    def test_read_project_text_supports_line_and_query_excerpt(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text(
                "HEADER\nC1 VALUE='100NF'\napiKey=abc123\nU1 VALUE='SOC'\n",
                encoding="utf-8",
            )
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            by_line = registry.run(
                "read_project_text",
                context,
                {"path": "packaged/pstxprt.dat", "line_start": 2, "line_count": 2},
            )
            by_query = registry.run(
                "read_project_text",
                context,
                {"path": "packaged/pstxprt.dat", "query": "U1", "context_lines": 0},
            )

        self.assertEqual(2, by_line["line_start"])
        self.assertEqual(3, by_line["line_end"])
        self.assertIn("C1 VALUE", by_line["content"])
        self.assertIn("apiKey=<redacted>", by_line["content"])
        self.assertEqual("U1", by_query["query"])
        self.assertEqual(4, by_query["line_start"])
        self.assertIn("U1 VALUE", by_query["content"])

    def test_trace_project_source_finds_raw_file_excerpts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "sch_1").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text(
                "PART_NAME\nC1 'CAP':\n VALUE='100NF'\nU1 'SOC':\n VALUE='ASIC'\n",
                encoding="utf-8",
            )
            (root / "packaged" / "pstxnet.dat").write_text(
                "NET_NAME\n'P3V3'\nNODE_NAME U1 A1\nNET_NAME\n'GND'\nNODE_NAME C1 2\n",
                encoding="utf-8",
            )
            (root / "sch_1" / "page12.csa").write_text(
                "WIRE 0 0 100 0\nTEXT 10 10 C1\n",
                encoding="utf-8",
            )
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            from_refdes = registry.run(
                "trace_project_source",
                context,
                {"query": "C1", "kind": "refdes", "limit": 4},
            )
            from_net = registry.run(
                "trace_project_source",
                context,
                {"query": "P3V3", "kind": "net", "limit": 4},
            )
            from_row = registry.run(
                "trace_project_source",
                context,
                {"table_id": "missing_value", "row_index": 0, "limit": 6},
            )

        self.assertEqual("pstx-source-trace.v1", from_refdes["source_schema_version"])
        self.assertTrue(any(hit["path"] == "packaged/pstxprt.dat" for hit in from_refdes["source_hits"]))
        self.assertTrue(any("C1" in line["text"] for hit in from_refdes["source_hits"] for line in hit["excerpt"]))
        self.assertTrue(any(hit["path"] == "packaged/pstxnet.dat" for hit in from_net["source_hits"]))
        self.assertEqual({"table_id": "missing_value", "table_title": "缺少 VALUE", "row_index": 0, "row_number": 1}, from_row["derived_from"])
        self.assertIn(12, from_row["page_numbers"])
        self.assertTrue(any(hit["path"] == "sch_1/page12.csa" for hit in from_row["source_hits"]))

    def test_search_project_text_greps_allowed_project_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "sch_1").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text(
                "PART_NAME\nU1 VALUE='ASIC'\napiKey=abc123\nU2 VALUE='LEVEL'\n",
                encoding="utf-8",
            )
            (root / "packaged" / "pstxnet.dat").write_text(
                "NET_NAME\n'I2C_SCL'\nNODE_NAME U2 A1\n",
                encoding="utf-8",
            )
            (root / "sch_1" / "page12.csa").write_text(
                "WIRE 0 0 100 0\nTEXT 10 10 I2C_SCL\n",
                encoding="utf-8",
            )
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            result = registry.run(
                "search_project_text",
                context,
                {"query": "U1 I2C_SCL", "context_lines": 1, "limit": 10},
            )
            regex_result = registry.run(
                "search_project_text",
                context,
                {"query": r"NODE_NAME\s+U2", "mode": "regex", "file_glob": "pstxnet.dat", "limit": 5},
            )

        self.assertEqual("pstx-project-text-search.v1", result["source_schema_version"])
        self.assertEqual("search_project_text", result["id"])
        self.assertTrue(result["readonly"])
        self.assertIn("packaged/pstxprt.dat", result["candidate_files"])
        self.assertTrue(any(hit["path"] == "packaged/pstxprt.dat" for hit in result["source_hits"]))
        self.assertTrue(any(hit["path"] == "sch_1/page12.csa" for hit in result["source_hits"]))
        self.assertTrue(any("apiKey=<redacted>" in line["text"] for hit in result["source_hits"] for line in hit["excerpt"]))
        self.assertEqual("read_project_text", result["detail_tool"]["name"])
        self.assertEqual("packaged/pstxnet.dat", regex_result["source_hits"][0]["path"])
        self.assertIn("NODE_NAME U2", regex_result["source_hits"][0]["matched_terms"])

    def test_search_project_text_rejects_unsafe_filters(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text("U1 VALUE='ASIC'\n", encoding="utf-8")
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            with self.assertRaises(HarnessToolError):
                registry.run("search_project_text", context, {"query": "U1", "path_prefix": "../outside"})
            with self.assertRaises(HarnessToolError):
                registry.run("search_project_text", context, {"query": "U1", "suffixes": [".pdf"]})

    def test_batch_report_tools_return_per_item_status(self):
        self.make_feishu_cache()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_dfmea_bundle(), request=HarnessRunRequest())
        tools = {item["name"]: item for item in registry.list_tools()}
        self.assertEqual("v2", tools["get_table_rows"]["result_contract"]["version"])
        self.assertIn("completeness", tools["get_table_rows"]["result_contract"]["fields"])

        entity = registry.run(
            "batch_query_report_entities",
            context,
            {"queries": [{"refdes": "U1"}, {"net": "P3V3"}, "HQ100", "NO_SUCH"], "limit_per_query": 3},
        )
        self.assertEqual(4, len(entity["items"]))
        self.assertEqual("found", entity["items"][0]["status"])
        self.assertEqual("found", entity["items"][1]["status"])
        self.assertTrue(any(item["status"] == "missing" for item in entity["items"]))

        tables = registry.run(
            "batch_get_table_rows",
            context,
            {
                "requests": [{"table_id": "missing_value", "limit": 2}, {"table_id": "bad_table"}],
                "limit_per_request": 20,
            },
        )
        self.assertEqual(10, tables["limit_per_request"])
        self.assertEqual("found", tables["items"][0]["status"])
        self.assertEqual("error", tables["items"][1]["status"])
        self.assertEqual(1, len(tables["items"][0]["rows"]))

    def test_table_column_summary_counts_unique_values_without_full_rows(self):
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle={}, request=HarnessRunRequest())

        rows = registry.run("get_table_rows", context, {"table_id": "page_rows", "limit": 2})
        self.assertTrue(rows["has_more"])
        self.assertTrue(rows["truncated"])
        self.assertEqual(2, rows["next_offset"])
        self.assertIn("summarize_table_column_values", rows["aggregation_hint"])

        summary = registry.run(
            "summarize_table_column_values",
            context,
            {"table_id": "page_rows", "column": "页码", "limit_values": 10, "sample_per_value": 1},
        )
        self.assertEqual(4, summary["total_rows"])
        self.assertEqual(3, summary["unique_count"])
        self.assertEqual(["PAGE1", "PAGE2", "PAGE10"], summary["values"])
        self.assertEqual(2, next(item["count"] for item in summary["top_values"] if item["value"] == "PAGE2"))
        self.assertEqual("top", summary["operation"])
        self.assertEqual("complete", summary["completeness"])
        self.assertEqual("PAGE2", summary["value_counts"][0]["value"])
        self.assertEqual(2, summary["value_counts"][0]["count"])
        self.assertEqual("get_table_rows", summary["detail_tool"]["name"])
        self.assertTrue(summary["sample_rows_by_value"])
        self.assertIn("row_number", summary["sample_rows_by_value"][0]["samples"][0])

        count_summary = registry.run(
            "summarize_table_column_values",
            context,
            {"table_id": "page_rows", "column": "页码", "operation": "count", "limit_values": 2},
        )
        self.assertEqual(["PAGE1", "PAGE2"], [item["value"] for item in count_summary["value_counts"]])
        self.assertTrue(count_summary["truncated"])
        self.assertEqual("partial", count_summary["completeness"])

        with self.assertRaises(HarnessToolError):
            registry.run(
                "summarize_table_column_values",
                context,
                {"table_id": "page_rows", "column": "不存在的列"},
            )

    def test_schematic_page_count_uses_module_order_extent_not_page_rows(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "module_order.dat").write_text(
                "START_MODULEORDER\n"
                "@TOP_LIB.TOP(SCH_1):PAGE1_I1@LIB.BLOCK_A(SCH_1) 0 1 1 20 1\n"
                "@TOP_LIB.TOP(SCH_1):PAGE21_I2@LIB.BLOCK_B(SCH_1) 0 1 21 30 1\n"
                "@TOP_LIB.TOP(SCH_1):PAGE300_I3@LIB.EMPTY_TAIL(SCH_1) 0 1 300 24 1\n"
                "END_MODULEORDER\n",
                encoding="utf-8",
            )
            registry = build_default_harness_registry()
            context = HarnessToolContext(
                report=sample_report(),
                bundle={"project_root": str(root)},
                request=HarnessRunRequest(),
            )

            result = registry.run("summarize_schematic_page_count", context, {})

        self.assertTrue(result["available"])
        self.assertEqual(323, result["total_pages"])
        self.assertEqual("PAGE323", result["last_page"])
        self.assertEqual("PAGE300", result["last_entry"]["start_real_page"])
        self.assertEqual(24, result["last_entry"]["page_count"])
        self.assertIn("page_rows 只统计", result["scope_note"])

    def test_batch_feishu_identity_and_datasheet_tools_are_readonly(self):
        root = self.make_feishu_cache()
        registry = build_default_harness_registry()
        context = HarnessToolContext(report=sample_report(), bundle=sample_dfmea_bundle(), request=HarnessRunRequest())

        feishu = registry.run(
            "batch_search_feishu_cache_rows",
            context,
            {"queries": ["HQ100", "LiXinYu", "NO_SUCH_MATERIAL"], "limit_per_query": 2},
        )
        self.assertEqual(["found", "found", "missing"], [item["status"] for item in feishu["items"]])
        self.assertEqual("HQ100", feishu["items"][0]["rows"][0]["hq_no"])

        identities = registry.run(
            "batch_get_component_identity_cards",
            context,
            {"refdes_list": ["U1", "PU2", "NO_SUCH"]},
        )
        self.assertEqual("found", identities["items"][0]["status"])
        self.assertEqual("needs_context", identities["items"][1]["status"])
        self.assertEqual("missing", identities["items"][2]["status"])

        datasheets = registry.run(
            "batch_match_component_datasheets",
            context,
            {"refdes_list": ["U1", "NO_SUCH"], "limit_per_component": 2},
        )
        self.assertEqual("missing", datasheets["items"][0]["status"])
        self.assertEqual("identity_card_not_found", datasheets["items"][1]["missing_reason"])

        conn = sqlite3.connect(root / "feishu_cache.db")
        try:
            count_after = conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(2, count_after)

    def test_harness_does_not_mutate_report_or_bundle(self):
        report = sample_report()
        bundle = {"net_analysis": {"total": 3}}
        before_report = copy.deepcopy(report)
        before_bundle = copy.deepcopy(bundle)

        run_harness_review(report, bundle, HarnessRunRequest(max_rows_per_table=1, include_model=False))

        self.assertEqual(before_report, report)
        self.assertEqual(before_bundle, bundle)

    def test_aster_provider_uses_injected_model_interface(self):
        calls = []

        def fake_ask_model(prompt, *, inputs=None):
            calls.append({"prompt": prompt, "inputs": inputs})
            return {
                "answer": json.dumps({
                    "summary": "fake aster summary",
                    "priorities": [{"title": "优先项", "body": "证据", "target": "drc", "severity": "high"}],
                    "review_checklist": [{"item": "检查项", "status": "needs_review", "evidence": "证据", "target": "drc"}],
                    "manual_review": [{"topic": "人工确认", "reason": "边界", "target": "summary"}],
                }, ensure_ascii=False),
                "provider": "fake-aster",
                "mode": "live",
                "metadata": {"conversation_id": "conv-1"},
            }

        provider = AsterHarnessModelProvider(ask_model=fake_ask_model)
        payload = run_harness_review(sample_report(), {}, HarnessRunRequest(question="请重点看 DRC"), provider)

        self.assertEqual("fake aster summary", payload["summary"])
        self.assertEqual("fake-aster", payload["model_metadata"]["provider"])
        self.assertEqual("conv-1", payload["model_metadata"]["conversation_id"])
        self.assertEqual(1, len(calls))
        self.assertIn("不能请求或执行任何工具", calls[0]["prompt"])
        self.assertEqual("请重点看 DRC", calls[0]["inputs"]["question"])

    def test_request_validation_rejects_unknown_task_and_bad_limit(self):
        with self.assertRaises(HarnessError):
            HarnessRunRequest.from_mapping({"task": "write_files"})
        with self.assertRaises(HarnessError):
            HarnessRunRequest.from_mapping({"max_rows_per_table": 0})


if __name__ == "__main__":
    unittest.main()
