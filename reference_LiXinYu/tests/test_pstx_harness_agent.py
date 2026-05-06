import copy
import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pstx_harness.report_agent import HarnessAgentRequest, list_harness_agent_profiles, run_harness_agent
from pstx_harness.model import HarnessModelResponse, MockHarnessModelProvider
from pstx_harness.report_agent_topology_evidence import topology_evidence_nodes_from_tool_result
from pstx_knowledge.datasheets import reindex_datasheets
from pstx_knowledge.reference_library import reindex_agent_ref, reindex_review_checklists


def sample_report():
    return {
        "project_name": "agent_demo",
        "sections": [
            {
                "id": "drc",
                "title": "设计检查",
                "tables": [
                    {
                        "id": "missing_value",
                        "title": "缺少 VALUE",
                        "count": 1,
                        "columns": ["位号", "页码"],
                        "rows": [{"位号": "R1", "页码": "12"}],
                    }
                ],
            },
            {
                "id": "network",
                "title": "网络分析",
                "tables": [
                    {
                        "id": "page_rows",
                        "title": "页码元件分布",
                        "columns": ["页码", "元件数"],
                        "rows": [
                            {"页码": "PAGE1", "元件数": 3},
                            {"页码": "PAGE2", "元件数": 5},
                            {"页码": "PAGE10", "元件数": 1},
                        ],
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
                "nets": {"A1": "P3V3", "B1": "I2C_SCL", "B2": "I2C_SDA"},
            },
            "U2": {
                "HQ_CODE": "HQ200",
                "CDS_PART_NAME": "TXS0108_LEVEL_TRANSLATOR",
                "PACKAGE": "QFN",
                "page_submodule_mapped": "14",
                "nets": {"A1": "I2C_SCL", "A2": "I2C_SDA", "B1": "I2C_SCL_1V8", "B2": "I2C_SDA_1V8"},
            },
            "PU2": {
                "CDS_PART_NAME": "POWER_MANAGER",
                "PACKAGE": "QFN",
                "page_submodule_mapped": "18",
                "nets": {"1": "VIN_12V", "2": "VOUT_1V8"},
            },
        },
        "nets": {
            "P3V3": [{"refdes": "U1", "pin": "A1", "pin_name": "VDD"}],
            "I2C_SCL": [{"refdes": "U1", "pin": "B1", "pin_name": "SCL"}, {"refdes": "U2", "pin": "A1", "pin_name": "SCL_A"}],
            "I2C_SDA": [{"refdes": "U1", "pin": "B2", "pin_name": "SDA"}, {"refdes": "U2", "pin": "A2", "pin_name": "SDA_A"}],
            "I2C_SCL_1V8": [{"refdes": "U2", "pin": "B1", "pin_name": "SCL_B"}],
            "I2C_SDA_1V8": [{"refdes": "U2", "pin": "B2", "pin_name": "SDA_B"}],
            "VIN_12V": [{"refdes": "PU2", "pin": "1", "pin_name": "VIN"}],
            "VOUT_1V8": [{"refdes": "PU2", "pin": "2", "pin_name": "VOUT"}],
        },
    }


class SequenceProvider:
    provider = "sequence"
    mode = "mock"

    def __init__(self, answers):
        self.answers = list(answers)
        self.calls = []

    def generate_agent_step(self, prompt, *, inputs=None):
        self.calls.append({"prompt": prompt, "inputs": inputs or {}})
        answer = self.answers.pop(0) if self.answers else '{"final_answer":"done"}'
        return HarnessModelResponse(answer=answer, provider=self.provider, mode=self.mode)


class LoopProvider:
    provider = "loop"
    mode = "mock"

    def generate_agent_step(self, prompt, *, inputs=None):
        return HarnessModelResponse(
            answer=json.dumps({
                "tool_call": {
                    "name": "list_report_tables",
                    "args": {},
                    "reason": "持续读取表格清单",
                }
            }, ensure_ascii=False),
            provider=self.provider,
            mode=self.mode,
        )


class ExplodingProvider:
    provider = "exploding"
    mode = "mock"

    def generate_agent_step(self, prompt, *, inputs=None):
        raise RuntimeError("simulated subagent provider failure")


class CloneToExplodingSubagentProvider:
    provider = "parent-ok-child-fails"
    mode = "mock"

    def __init__(self):
        self.clone_count = 0

    def clone_for_subagent(self):
        self.clone_count += 1
        return ExplodingProvider()

    def generate_agent_step(self, prompt, *, inputs=None):
        return HarnessModelResponse(
            answer=json.dumps({"final_answer": "parent done"}, ensure_ascii=False),
            provider=self.provider,
            mode=self.mode,
        )


def make_feishu_cache(testcase: unittest.TestCase) -> Path:
    temp_dir = tempfile.TemporaryDirectory()
    testcase.addCleanup(temp_dir.cleanup)
    root = Path(temp_dir.name)
    old_env = os.environ.get("PSTX_FEISHU_DATA_DIR")
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_FEISHU_DATA_DIR", None)
        if old_env is None else os.environ.__setitem__("PSTX_FEISHU_DATA_DIR", old_env)
    )
    os.environ["PSTX_FEISHU_DATA_DIR"] = str(root)
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
            pi TEXT,
            selection_order TEXT,
            extra_fields TEXT,
            raw_data TEXT,
            synced_at TEXT
        )
        """
    )
    conn.execute(
        "INSERT INTO materials(lib_id,lib_name,sheet_name,key_value,hq_no,brand,spec,description,pi,selection_order,extra_fields,raw_data,synced_at) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (
            "lib1", "优选库", "电容", "CAP-100N", "HQ100", "ACME", "CAP-100N",
            "100nF capacitor", "LiXinYu", "1",
            json.dumps({"封装": "0402", "耐压": "50V"}, ensure_ascii=False),
            json.dumps({"封装": "0402", "耐压": "50V"}, ensure_ascii=False),
            "2026-04-27",
        ),
    )
    conn.commit()
    conn.close()
    return root


def make_datasheet_index(testcase: unittest.TestCase) -> Path:
    temp_dir = tempfile.TemporaryDirectory()
    testcase.addCleanup(temp_dir.cleanup)
    root = Path(temp_dir.name)
    source = root / "pdfs"
    source.mkdir()
    (source / "HQ100_GPU_CORE_TEST_IC.pdf").write_bytes(b"%PDF fake")
    old_dir = os.environ.get("PSTX_DATASHEET_DIR")
    old_data_dir = os.environ.get("PSTX_DATASHEET_DATA_DIR")
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_DATASHEET_DIR", None)
        if old_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DIR", old_dir)
    )
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_DATASHEET_DATA_DIR", None)
        if old_data_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DATA_DIR", old_data_dir)
    )
    os.environ["PSTX_DATASHEET_DIR"] = str(source)
    os.environ["PSTX_DATASHEET_DATA_DIR"] = str(root / "data")
    with mock.patch(
        "pstx_knowledge.datasheets._extract_pdf_pages",
        return_value=("indexed", ["HQ100 GPU_CORE_TEST_IC datasheet electrical limits"], "fake", ""),
    ):
        reindex_datasheets(force=True)
    return root


def make_harness_doc_dir(testcase: unittest.TestCase) -> Path:
    temp_dir = tempfile.TemporaryDirectory()
    testcase.addCleanup(temp_dir.cleanup)
    root = Path(temp_dir.name)
    old_env = os.environ.get("PSTX_HARNESS_DOC_DIR")
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_HARNESS_DOC_DIR", None)
        if old_env is None else os.environ.__setitem__("PSTX_HARNESS_DOC_DIR", old_env)
    )
    os.environ["PSTX_HARNESS_DOC_DIR"] = str(root)
    (root / "agent_notes.md").write_text(
        "U46 多 symbol 芯片需要按 SECTION_NUMBER 汇总页码和 HQ_CODE。\n",
        encoding="utf-8",
    )
    return root


def make_agent_ref_index(testcase: unittest.TestCase) -> Path:
    temp_dir = tempfile.TemporaryDirectory()
    testcase.addCleanup(temp_dir.cleanup)
    root = Path(temp_dir.name)
    source = root / "ref"
    source.mkdir()
    (source / "agent_capability_manual.pdf").write_bytes(b"%PDF fake")
    old_dir = os.environ.get("PSTX_AGENT_REF_DIR")
    old_data_dir = os.environ.get("PSTX_AGENT_REF_DATA_DIR")
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_AGENT_REF_DIR", None)
        if old_dir is None else os.environ.__setitem__("PSTX_AGENT_REF_DIR", old_dir)
    )
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_AGENT_REF_DATA_DIR", None)
        if old_data_dir is None else os.environ.__setitem__("PSTX_AGENT_REF_DATA_DIR", old_data_dir)
    )
    os.environ["PSTX_AGENT_REF_DIR"] = str(source)
    os.environ["PSTX_AGENT_REF_DATA_DIR"] = str(root / "data")
    with mock.patch(
        "pstx_knowledge.reference_library._extract_pdf_pages",
        return_value=("indexed", ["Agent Lab ref PDF explains boundary testing and citations."], "fake", ""),
    ):
        reindex_agent_ref(force=True)
    return root


def make_review_checklist_index(testcase: unittest.TestCase) -> Path:
    temp_dir = tempfile.TemporaryDirectory()
    testcase.addCleanup(temp_dir.cleanup)
    root = Path(temp_dir.name)
    source = root / "ref_checklist"
    source.mkdir()
    (source / "review_cases.md").write_text(
        "真实 review 问题：U46 多 symbol 芯片需要检查 HQ_CODE、Pin/Net 和页码。",
        encoding="utf-8",
    )
    old_dir = os.environ.get("PSTX_AGENT_CHECKLIST_REF_DIR")
    old_data_dir = os.environ.get("PSTX_AGENT_CHECKLIST_DATA_DIR")
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_AGENT_CHECKLIST_REF_DIR", None)
        if old_dir is None else os.environ.__setitem__("PSTX_AGENT_CHECKLIST_REF_DIR", old_dir)
    )
    testcase.addCleanup(
        lambda: os.environ.pop("PSTX_AGENT_CHECKLIST_DATA_DIR", None)
        if old_data_dir is None else os.environ.__setitem__("PSTX_AGENT_CHECKLIST_DATA_DIR", old_data_dir)
    )
    os.environ["PSTX_AGENT_CHECKLIST_REF_DIR"] = str(source)
    os.environ["PSTX_AGENT_CHECKLIST_DATA_DIR"] = str(root / "data")
    reindex_review_checklists(force=True)
    return root


class HarnessAgentTests(unittest.TestCase):
    def test_mock_provider_calls_tool_then_returns_final_answer(self):
        payload = run_harness_agent(sample_report(), {}, HarnessAgentRequest(), MockHarnessModelProvider())

        self.assertTrue(payload["ok"])
        self.assertEqual("local-agent-harness", payload["mode"])
        self.assertEqual("quick_scan", payload["profile"])
        self.assertTrue(payload["agent_run_id"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertEqual("get_table_rows", payload["tool_calls"][0]["tool"])
        self.assertTrue(payload["observations"])
        self.assertTrue(payload["raw_observations"])
        self.assertIn("evidence_layers", payload["observations"][0])
        self.assertIn("summary_layer", payload["observations"][0]["evidence_layers"])
        self.assertIn("evidence_card_layer", payload["observations"][0]["evidence_layers"])
        self.assertIn("raw_layer", payload["observations"][0]["evidence_layers"])
        self.assertIn("guidance_summary", payload)
        self.assertIn("selected_skills", payload)
        self.assertIn("effort_policy", payload)
        self.assertIn("task_memory_summary", payload)
        self.assertIn("raw_result", payload["raw_observations"][0])
        self.assertTrue(payload["final_evidence"])
        self.assertTrue(payload["citations"])
        self.assertTrue(payload["citations"][0]["valid"])
        self.assertIn("final_answer_quality_gate", payload)
        self.assertIn(payload["final_answer_quality_gate"]["status"], {"pass", "warn"})
        self.assertIn("repair_actions", payload["final_answer_quality_gate"])
        self.assertIn("repair_action_count", payload["final_answer_quality_gate"])
        self.assertIn("final_answer_quality_gate", payload["model_metadata"])
        self.assertIn("final_quality_status", payload["trace_summary"])
        self.assertIn("evidence_goal_contract", payload)
        self.assertEqual("agent-evidence-goal-contract/v1", payload["evidence_goal_contract"]["version"])
        self.assertIn("evidence_goal_status", payload["trace_summary"])
        self.assertIn("execution_journal", payload)
        self.assertIn("journal_summary", payload)
        self.assertGreaterEqual(payload["journal_summary"]["event_count"], 3)
        self.assertEqual("pstx-harness-turn-context.v1", payload["turn_context_snapshot"]["schema_version"])
        self.assertEqual("pstx-tool-dispatch-summary.v1", payload["tool_dispatch_summary"]["schema_version"])
        self.assertTrue(payload["tool_dispatch_trace"])
        self.assertGreaterEqual(payload["trace_summary"]["tool_dispatch_event_count"], 1)
        self.assertEqual("completed", payload["tool_dispatch_trace"][0]["status"])
        self.assertIn("continuation_pack", payload)
        self.assertEqual("agent-continuation-pack/v1", payload["continuation_pack"]["version"])
        self.assertTrue(payload["continuation_pack"]["continuation_brief"])
        self.assertIn("mock agent", payload["answer"])

    def test_report_agent_final_answer_can_write_scratch_files(self):
        old_workspace = os.environ.get("PSTX_AGENT_WORKSPACE_DIR")
        with tempfile.TemporaryDirectory() as tmp:
            os.environ["PSTX_AGENT_WORKSPACE_DIR"] = tmp
            try:
                provider = SequenceProvider([
                    json.dumps({
                        "tool_call": {
                            "name": "get_table_rows",
                            "args": {"table_id": "missing_value", "limit": 1},
                            "reason": "先取一个可引用证据。",
                        }
                    }, ensure_ascii=False),
                    json.dumps({
                        "final_answer": "已生成临时分析笔记。",
                        "citations": [{"id": "ev-1-row-1", "note": "临时笔记基于该行。"}],
                        "scratch_files": [
                            {
                                "filename": "../analysis-notes.md",
                                "content": "# 临时笔记\nU1 待复核。",
                                "content_type": "text/markdown",
                            }
                        ],
                    }, ensure_ascii=False),
                ])
                payload = run_harness_agent(
                    sample_report(),
                    {"run_id": "scratch-run"},
                    HarnessAgentRequest(profile="quick_scan", max_steps=2, max_tool_calls=2),
                    provider,
                    project_context={"run_id": "scratch-scope", "agent_workspace_agent_run_id": "scratch-agent"},
                )

                scratch = payload["scratch_files"]
                written = Path(scratch["files"][0]["path"])

                self.assertTrue(payload["ok"])
                self.assertEqual(1, scratch["file_count"])
                self.assertEqual("analysis-notes.md", scratch["files"][0]["name"])
                self.assertEqual("scratch-agent", scratch["files"][0]["agent_run_id"])
                self.assertTrue(written.is_file())
                self.assertIn("U1 待复核", written.read_text(encoding="utf-8"))
                self.assertEqual(1, payload["trace_summary"]["scratch_file_count"])
            finally:
                if old_workspace is None:
                    os.environ.pop("PSTX_AGENT_WORKSPACE_DIR", None)
                else:
                    os.environ["PSTX_AGENT_WORKSPACE_DIR"] = old_workspace

    def test_quality_gate_repair_executes_detail_tool_then_reanswers(self):
        report = sample_report()
        report["sections"][1]["tables"][0]["rows"] = [
            {"页面": "PAGE1", "元件数": 3},
            {"页面": "PAGE2", "元件数": 5},
        ]
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "get_table_rows", "args": {"table_id": "page_rows", "limit": 1}, "reason": "先看预览"}}),
            json.dumps({"final_answer": "先按预览回答，缺少引用。", "citations": []}),
            json.dumps({"final_answer": "已补齐第二行证据后回答。", "citations": [{"id": "ev-2-row-2"}]}),
        ])

        payload = run_harness_agent(
            report,
            {},
            HarnessAgentRequest(max_steps=4, max_tool_calls=5, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("已补齐第二行证据后回答。", payload["answer"])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertEqual(2, len(payload["tool_calls"]))
        self.assertEqual("get_table_rows", payload["tool_calls"][1]["tool"])
        self.assertEqual(1, payload["tool_calls"][1]["args"]["offset"])
        self.assertIn("quality_repair_tool_call", [item["type"] for item in payload["agent_steps"]])
        self.assertEqual(3, len(provider.calls))

    def test_agent_prompt_includes_runtime_protocol_contract(self):
        provider = SequenceProvider(['{"final_answer":"done","citations":[]}'])
        payload = run_harness_agent(sample_report(), {}, HarnessAgentRequest(), provider)

        self.assertTrue(payload["ok"])
        self.assertIn("pstx-agent-runtime/v1", provider.calls[0]["prompt"])
        self.assertIn("dispatch_tasks", provider.calls[0]["prompt"])
        self.assertIn("scratch_files", provider.calls[0]["prompt"])
        self.assertIn("ObservationBundle", provider.calls[0]["prompt"])
        self.assertIn("observation_bundle", provider.calls[0]["inputs"]["context_budget"])
        self.assertIn("runtime_state", provider.calls[0]["inputs"])
        self.assertEqual("pstx-agent-runtime/v1", provider.calls[0]["inputs"]["runtime_state"]["protocol_version"])
        self.assertEqual("agent-task-ledger/v1", provider.calls[0]["inputs"]["runtime_state"]["task_ledger"]["version"])
        self.assertEqual("agent-evidence-goal-contract/v1", provider.calls[0]["inputs"]["runtime_state"]["evidence_goal_contract"]["version"])
        self.assertIn("task_ledger", provider.calls[0]["inputs"])
        self.assertIn("session_state", provider.calls[0]["inputs"])
        self.assertEqual("pstx-agent-runtime/v1", provider.calls[0]["inputs"]["session_state"]["protocol_version"])
        self.assertEqual("agent-task-ledger/v1", provider.calls[0]["inputs"]["session_state"]["task_ledger"]["version"])
        self.assertIn("playbook_plan", provider.calls[0]["inputs"])
        self.assertIn("playbook_plan", provider.calls[0]["prompt"])
        self.assertIn("task_ledger", provider.calls[0]["prompt"])
        self.assertIn("runtime_state", payload)
        self.assertIn("session_state", payload)
        self.assertIn("task_ledger", payload["runtime_state"])
        self.assertIn("evidence_goal_contract", payload["runtime_state"])
        self.assertIn("task_ledger_open_count", payload["trace_summary"])
        self.assertIn("playbook_plan", payload)

    def test_agent_loop_dispatches_long_tasks_with_callback(self):
        provider = SequenceProvider([
            json.dumps({
                "dispatch_tasks": [
                    {
                        "task_id": "ds-u1",
                        "title": "U1 datasheet",
                        "profile": "datasheet_qa",
                        "question": "读取 U1 datasheet 供电和接口约束。",
                    }
                ],
                "reason": "规格书分支可后台执行。",
            }, ensure_ascii=False)
        ])
        seen = []

        def dispatch_callback(request):
            seen.append(request)
            return {
                "task_dispatch_summary": {"queue": "accepted"},
                "dispatched_tasks": [{
                    "task_id": "ds-u1",
                    "title": "U1 datasheet",
                    "profile": "datasheet_qa",
                    "question": "读取 U1 datasheet 供电和接口约束。",
                    "agent_run_id": "child-report-1",
                    "status": "queued",
                    "status_url": "/api/harness/agent-runs/child-report-1",
                }],
            }

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(max_steps=1, max_tool_calls=1, debug=True),
            provider,
            project_context={"run_id": "run-1"},
            dispatch_callback=dispatch_callback,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("completed", payload["status"])
        self.assertEqual("task_dispatched", payload["model_metadata"]["stopped_reason"])
        self.assertEqual("report", seen[0]["source"])
        self.assertEqual("ds-u1", seen[0]["tasks"][0]["task_id"])
        self.assertEqual("child-report-1", payload["dispatched_tasks"][0]["agent_run_id"])
        self.assertTrue(payload["task_dispatch_summary"]["available"])
        self.assertEqual("task_dispatch", payload["agent_steps"][0]["type"])

    def test_agent_loop_returns_dispatch_plan_without_callback(self):
        provider = SequenceProvider([
            json.dumps({
                "dispatch_tasks": [{
                    "task_id": "cad-114",
                    "title": "Cadence 114",
                    "profile": "cadence_pages",
                    "question": "复核第 114 页 Cadence 连接语义。",
                }]
            }, ensure_ascii=False)
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(max_steps=1, max_tool_calls=1),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertFalse(payload["task_dispatch_summary"]["available"])
        self.assertEqual("not_dispatched", payload["dispatched_tasks"][0]["status"])

    def test_unknown_tool_is_rejected(self):
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "unknown_tool", "args": {}, "reason": "try"}}),
        ])
        payload = run_harness_agent(sample_report(), {}, HarnessAgentRequest(max_steps=1), provider)

        self.assertFalse(payload["ok"])
        self.assertEqual("tool_error", payload["model_metadata"]["stopped_reason"])
        self.assertIn("Unknown harness tool", payload["answer"])
        self.assertEqual("blocked", payload["tool_dispatch_trace"][0]["status"])
        self.assertGreaterEqual(payload["tool_dispatch_summary"]["blocked_count"], 1)

    def test_schema_mismatch_is_rejected(self):
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "get_table_rows", "args": {}, "reason": "missing args"}}),
        ])
        payload = run_harness_agent(sample_report(), {}, HarnessAgentRequest(max_steps=1), provider)

        self.assertFalse(payload["ok"])
        self.assertIn("缺少参数", payload["answer"])

    def test_tool_error_can_recover_with_alternative_tool(self):
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "unknown_tool", "args": {}, "reason": "先试错"}}),
            json.dumps({"tool_call": {"name": "get_table_rows", "args": {"table_id": "missing_value", "limit": 1}, "reason": "改用表格取证"}}),
            json.dumps({"final_answer": "已从工具失败中恢复并完成取证。", "citations": [{"id": "ev-2-row-1"}]}),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(max_steps=4, max_tool_calls=4, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertEqual(1, payload["model_metadata"]["tool_error_recovery_count"])
        self.assertIn("tool_error_recovery", [item["type"] for item in payload["agent_steps"]])
        self.assertEqual(["unknown_tool", "get_table_rows"], [item["tool"] for item in payload["tool_calls"]])
        self.assertFalse(payload["observations"][0]["ok"])
        self.assertEqual("error", payload["observations"][0]["tool_result_contract"]["completeness"])
        self.assertIn("list_report_tables", payload["observations"][0]["tool_result_contract"]["recommended_next_tools"])
        self.assertNotIn("unknown_tool", payload["observations"][0]["tool_result_contract"]["recommended_next_tools"])
        self.assertEqual("unknown_tool", provider.calls[1]["inputs"]["observations"][0]["tool"])
        self.assertIn("Unknown harness tool", provider.calls[1]["inputs"]["observations"][0]["error"])
        self.assertIn("list_report_tables", provider.calls[1]["inputs"]["observations"][0]["tool_result_contract"]["recommended_next_tools"])

    def test_duplicate_tool_call_is_reexecuted_without_error_recovery(self):
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "list_report_tables", "args": {}, "reason": "先看表格"}}),
            json.dumps({"tool_call": {"name": "list_report_tables", "args": {}, "reason": "重复看表格"}}),
            json.dumps({"final_answer": "已避免重复工具调用。", "citations": []}),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(question="请快速扫描", max_steps=4, max_tool_calls=4, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertEqual(0, payload["model_metadata"].get("tool_error_recovery_count", 0))
        self.assertEqual(["list_report_tables", "list_report_tables"], [item["tool"] for item in payload["tool_calls"]])
        self.assertTrue(payload["tool_calls"][1]["duplicate"])
        self.assertTrue(payload["observations"][1]["ok"])
        self.assertEqual("complete", payload["observations"][1]["tool_result_contract"]["completeness"])
        self.assertIn("batch_get_table_rows", payload["observations"][1]["tool_result_contract"]["recommended_next_tools"])

    def test_tool_batch_call_runs_multiple_allowed_tools_in_one_step(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_batch_call": [
                    {"name": "list_report_tables", "args": {}, "reason": "先列表"},
                    {"name": "get_table_rows", "args": {"table_id": "missing_value", "limit": 1}, "reason": "再取行"},
                ]
            }, ensure_ascii=False),
            json.dumps({"final_answer": "batch done", "citations": []}, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(max_steps=2, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["list_report_tables", "get_table_rows"], [item["tool"] for item in payload["tool_calls"]])
        self.assertEqual("tool_batch_call", payload["agent_steps"][0]["type"])
        self.assertEqual(2, len(payload["observations"]))
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])

    def test_agent_can_read_harness_skill_during_run(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_harness_skill",
                    "args": {"skill_id": "datasheet-key-info", "max_body_chars": 1000},
                    "reason": "先读取 datasheet skill 的取证打法。",
                }
            }, ensure_ascii=False),
            json.dumps({"final_answer": "已读取 datasheet skill，并会按参数卡/detail chunk 取证。", "citations": []}, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(
                profile="datasheet_qa",
                question="怎么读 64144 这类 datasheet 的关键信息？",
                max_steps=2,
                max_tool_calls=2,
                debug=True,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("get_harness_skill", payload["tool_calls"][0]["tool"])
        self.assertIn("guidance only", payload["observations"][0]["summary"])
        self.assertIn("已确认的 datasheet 事实", payload["raw_observations"][0]["raw_result"]["skill"]["body"])

    def test_batch_domain_tool_creates_citable_evidence(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "batch_query_report_entities",
                    "args": {"queries": [{"refdes": "U1"}, {"net": "P3V3"}], "limit_per_query": 3},
                    "reason": "复合问题一次查询多个对象。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "batch evidence done",
                "citations": [{"id": "ev-1-batch-query-1-1", "note": "U1 查询证据"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="quick_scan", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("batch_query_report_entities", payload["tool_calls"][0]["tool"])
        self.assertTrue(any(node["id"] == "ev-1-batch-query-1-1" for node in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])

    def test_quality_gate_refocuses_when_final_answer_omits_batch_target(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U1", "U2"], "limit_per_query": 5},
                    "reason": "先批量查询两个目标位号。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 已检查，暂未发现明显异常。",
                "citations": [{"id": "ev-1-batch-query-1-1", "note": "U1 查询证据"}],
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 和 U2 均已检查，U2 与 U1 共享 I2C 网络，建议人工复核跨电平连接。",
                "citations": [
                    {"id": "ev-1-batch-query-1-1", "note": "U1 查询证据"},
                    {"id": "ev-2-batch-query-1-1", "note": "U2 聚焦补证据"},
                ],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="quick_scan", question="请同时检查 U1 和 U2", max_steps=3, max_tool_calls=3, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["batch_query_report_entities", "batch_query_report_entities"], [item["tool"] for item in payload["tool_calls"]])
        self.assertEqual(["U2"], payload["tool_calls"][1]["args"]["queries"])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertIn("U1 和 U2", payload["answer"])
        self.assertTrue(all(item["valid"] for item in payload["citations"]))

    def test_quality_gate_refocuses_when_final_answer_lacks_target_citation(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U1", "U2"], "limit_per_query": 5},
                    "reason": "先批量查询两个目标位号。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 和 U2 均已检查，暂未发现明显异常。",
                "citations": [{"id": "ev-1-batch-query-1-1", "note": "只有 U1 查询证据"}],
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 和 U2 均已检查，U2 需要结合 I2C 跨电平连接继续人工复核。",
                "citations": [
                    {"id": "ev-1-batch-query-1-1", "note": "U1 查询证据"},
                    {"id": "ev-2-batch-query-1-1", "note": "U2 聚焦补证据"},
                ],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="quick_scan", question="请同时检查 U1 和 U2", max_steps=3, max_tool_calls=3, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["batch_query_report_entities", "batch_query_report_entities"], [item["tool"] for item in payload["tool_calls"]])
        self.assertEqual(["U2"], payload["tool_calls"][1]["args"]["queries"])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertIn("quality_repair_tool_call", [item["type"] for item in payload["agent_steps"]])
        self.assertIn("citation 未覆盖", payload["agent_steps"][1]["summary"])
        self.assertIn("U1 和 U2", payload["answer"])
        self.assertTrue(all(item["valid"] for item in payload["citations"]))

    def test_table_column_summary_tool_creates_citable_evidence(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "summarize_table_column_values",
                    "args": {"table_id": "page_rows", "column": "页面", "limit_values": 20},
                    "reason": "统计 page_rows 的唯一页码数量。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "page count done",
                "citations": [{"id": "ev-1-table-aggregate-page-rows-页面", "note": "唯一页统计"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="page_mapping", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("summarize_table_column_values", payload["tool_calls"][0]["tool"])
        self.assertTrue(any(node["id"] == "ev-1-table-aggregate-page-rows-页面" for node in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])

    def test_schematic_page_count_tool_creates_citable_evidence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "module_order.dat").write_text(
                "START_MODULEORDER\n"
                "@TOP_LIB.TOP(SCH_1):PAGE300_I3@LIB.EMPTY_TAIL(SCH_1) 0 1 300 24 1\n"
                "END_MODULEORDER\n",
                encoding="utf-8",
            )
            provider = SequenceProvider([
                json.dumps({
                    "tool_call": {
                        "name": "summarize_schematic_page_count",
                        "args": {},
                        "reason": "按 module_order 统计原理图总页数。",
                    }
                }, ensure_ascii=False),
                json.dumps({
                    "final_answer": "总页数为 323 页。",
                    "citations": [{"id": "ev-1-schematic-page-count", "note": "module_order 页范围"}],
                }, ensure_ascii=False),
            ])

            payload = run_harness_agent(
                sample_report(),
                {"project_root": str(root)},
                HarnessAgentRequest(profile="page_mapping", question="我有多少页原理图？", max_steps=2, max_tool_calls=2),
                provider,
            )

        self.assertTrue(payload["ok"])
        self.assertEqual("summarize_schematic_page_count", payload["tool_calls"][0]["tool"])
        self.assertTrue(any(node["id"] == "ev-1-schematic-page-count" for node in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])
        self.assertIn("schematic_page_count", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])

    def test_truncated_table_observation_exposes_contract_and_playbook(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_table_rows",
                    "args": {"table_id": "page_rows", "limit": 1},
                    "reason": "先看页面表格。",
                }
            }, ensure_ascii=False),
            json.dumps({"final_answer": "需要聚合后再统计", "citations": []}, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="page_mapping", question="请统计 page_rows 总页数", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertIn("table_column_aggregation", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])
        self.assertIn("summarize_table_column_values", payload["playbook_plan"]["recommended_first_tools"])
        self.assertTrue(payload["tool_result_contracts"])
        contract = payload["tool_result_contracts"][0]
        self.assertEqual("truncated", contract["completeness"])
        self.assertIn("summarize_schematic_page_count", contract["recommended_next_tools"])
        self.assertEqual("summarize_schematic_page_count", contract["aggregation_tool"]["name"])
        self.assertEqual({}, contract["aggregation_tool"]["args"])
        self.assertEqual("truncated", payload["observations"][0]["tool_result_contract"]["completeness"])
        self.assertEqual("truncated", provider.calls[1]["inputs"]["observations"][0]["tool_result_contract"]["completeness"])

    def test_read_project_text_rejects_outside_project_root(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            outside = root.parent / "outside-secret.txt"
            outside.write_text("secret", encoding="utf-8")
            self.addCleanup(lambda: outside.exists() and outside.unlink())
            provider = SequenceProvider([
                json.dumps({
                    "tool_call": {
                        "name": "read_project_text",
                        "args": {"path": "../outside-secret.txt"},
                        "reason": "try outside",
                    }
                }),
            ])
            payload = run_harness_agent(
                sample_report(),
                {"project_root": str(root)},
                HarnessAgentRequest(profile="full_review", debug=True),
                provider,
            )

        self.assertFalse(payload["ok"])
        self.assertIn("项目根目录之外", payload["answer"])

    def test_source_trace_prefetch_creates_raw_file_evidence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text(
                "PART_NAME\nU1 'SOC':\n VALUE='ASIC'\n",
                encoding="utf-8",
            )

            class SourceTraceProvider:
                provider = "source-trace"
                mode = "mock"

                def __init__(self):
                    self.calls = []

                def generate_agent_step(self, prompt, *, inputs=None):
                    self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                    evidence_ids = []
                    for observation in (inputs or {}).get("observations") or []:
                        evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                    return HarnessModelResponse(
                        answer=json.dumps({
                            "final_answer": "已追溯到 U1 的原始文件片段。",
                            "citations": [{"id": evidence_ids[0], "note": "U1 原始 PSTX 文件片段。"}] if evidence_ids else [],
                        }, ensure_ascii=False),
                        provider=self.provider,
                        mode=self.mode,
                    )

            provider = SourceTraceProvider()
            payload = run_harness_agent(
                sample_report(),
                {"project_root": str(root)},
                HarnessAgentRequest(
                    profile="auto",
                    question="请把 U1 的分析结论追溯到原始文件级别",
                    max_steps=1,
                    max_tool_calls=4,
                ),
                provider,
            )

        self.assertTrue(payload["ok"])
        self.assertIn("source_file_drilldown", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])
        self.assertEqual("trace_project_source", payload["tool_calls"][0]["tool"])
        self.assertIn("source_trace", {item["type"] for item in payload["final_evidence"]})
        self.assertTrue(payload["citations"][0]["valid"])
        self.assertIn("trace_project_source", provider.calls[0]["inputs"]["playbook_plan"]["recommended_first_tools"])

    def test_auto_profile_greps_raw_project_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text(
                "PART_NAME\nU1 'SOC':\n VALUE='ASIC'\n",
                encoding="utf-8",
            )

            class ProjectGrepProvider:
                provider = "project-grep"
                mode = "mock"

                def __init__(self):
                    self.calls = []

                def generate_agent_step(self, prompt, *, inputs=None):
                    self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                    evidence_ids = []
                    for observation in (inputs or {}).get("observations") or []:
                        evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                    return HarnessModelResponse(
                        answer=json.dumps({
                            "final_answer": "已 grep 到 U1 的原始项目文件片段。",
                            "citations": [{"id": evidence_ids[0], "note": "U1 原始文件搜索片段。"}] if evidence_ids else [],
                        }, ensure_ascii=False),
                        provider=self.provider,
                        mode=self.mode,
                    )

            provider = ProjectGrepProvider()
            payload = run_harness_agent(
                sample_report(),
                {"project_root": str(root)},
                HarnessAgentRequest(
                    profile="auto",
                    question="请 grep 原始项目文件里的 U1",
                    max_steps=1,
                    max_tool_calls=4,
                ),
                provider,
            )

        self.assertTrue(payload["ok"])
        self.assertEqual("search_project_text", payload["tool_calls"][0]["tool"])
        self.assertIn("source_trace", {item["type"] for item in payload["final_evidence"]})
        self.assertTrue(payload["citations"][0]["valid"])
        self.assertIn("search_project_text", provider.calls[0]["inputs"]["playbook_plan"]["recommended_first_tools"])

    def test_auto_profile_routes_datasheet_connection_review_question(self):
        provider = SequenceProvider([
            json.dumps({
                "final_answer": "已按连接和 datasheet 证据链完成初步反查。",
                "citations": [{"id": "ev-1-llm-topology-summary", "note": "拓扑摘要。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(
                profile="auto",
                question="请根据 datasheet 反查 U1 和 U2 的 I2C 连接关系是否有问题，重点看接口电平兼容。",
                max_steps=1,
                max_tool_calls=1,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertIn("connection_datasheet_review", provider.calls[0]["inputs"]["capability_profiles"])
        self.assertIn("schematic_datasheet_connection_review", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])
        self.assertIn("batch_query_llm_topology_netlist", payload["playbook_plan"]["recommended_first_tools"])
        self.assertIn("batch_get_component_identity_cards", payload["playbook_plan"]["recommended_first_tools"])
        selected_skill_ids = [item["id"] for item in provider.calls[0]["inputs"]["selected_skills"]["selected_skills"]]
        self.assertIn("schematic-datasheet-connection-review", selected_skill_ids)
        task_ledger = provider.calls[0]["inputs"]["task_ledger"]
        ledger_item_ids = [item["id"] for item in task_ledger["items"]]
        self.assertIn("connection-review-schematic-evidence", ledger_item_ids)
        self.assertIn("connection-review-datasheet-detail", ledger_item_ids)
        next_action_tools = [item.get("tool") for item in task_ledger["next_actions"]]
        self.assertIn("batch_query_llm_topology_netlist", next_action_tools)
        self.assertIn("batch_get_component_identity_cards", next_action_tools)
        self.assertIn("batch_match_component_datasheets", next_action_tools)

    def test_auto_profile_prefetches_datasheet_connection_chain(self):
        make_datasheet_index(self)

        class PrefetchConnectionProvider:
            provider = "prefetch-connection"
            mode = "mock"

            def __init__(self):
                self.calls = []

            def generate_agent_step(self, prompt, *, inputs=None):
                self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                evidence_ids = []
                for observation in (inputs or {}).get("observations") or []:
                    evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                return HarnessModelResponse(
                    answer=json.dumps({
                        "final_answer": "已按拓扑、身份卡和 MinerU-backed datasheet 候选完成第一轮连接反查。",
                        "citations": [{"id": evidence_ids[0], "note": "引用预取证据。"}] if evidence_ids else [],
                    }, ensure_ascii=False),
                    provider=self.provider,
                    mode=self.mode,
                )

        provider = PrefetchConnectionProvider()
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(
                profile="auto",
                question="请根据 datasheet 反查 U1 和 U2 的 I2C 连接是否有问题，重点看接口电平。",
                max_steps=1,
                max_tool_calls=8,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(
            [
                "list_datasheet_sources",
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
            ],
            [item["tool"] for item in payload["tool_calls"][:4]],
        )
        self.assertEqual("runtime_prefetch", payload["agent_steps"][0]["type"])
        self.assertGreaterEqual(provider.calls[0]["inputs"]["tool_count"], 4)
        evidence_types = {item["type"] for item in payload["final_evidence"]}
        self.assertIn("component_identity", evidence_types)
        self.assertTrue({"datasheet_match", "datasheet_gap"} & evidence_types)
        self.assertIn("datasheet_document", evidence_types)
        self.assertIn("connection_datasheet_review", provider.calls[0]["inputs"]["capability_profiles"])

    def test_max_steps_stops_safely(self):
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(max_steps=1, max_tool_calls=8),
            LoopProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("max_steps", payload["model_metadata"]["stopped_reason"])
        self.assertIn("最大 agent 轮数", payload["answer"])

    def test_higher_agent_limits_are_allowed(self):
        request = HarnessAgentRequest.from_mapping({
            "profile": "full_review",
            "max_steps": 24,
            "max_tool_calls": 48,
        })

        self.assertEqual(24, request.max_steps)
        self.assertEqual(48, request.max_tool_calls)

        with self.assertRaises(Exception) as ctx:
            HarnessAgentRequest.from_mapping({"max_steps": 25})

        self.assertIn("max_steps", str(ctx.exception))

    def test_model_observations_are_compacted_before_next_model_call(self):
        report = sample_report()
        report["sections"][0]["tables"][0]["rows"] = [
            {
                "位号": f"R{index}",
                "页码": "12",
                "长字段": "X" * 2000,
            }
            for index in range(30)
        ]
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_table_rows",
                    "args": {"table_id": "missing_value", "limit": 30},
                    "reason": "读取大表验证上下文裁剪",
                }
            }),
            json.dumps({"final_answer": "done"}),
        ])

        payload = run_harness_agent(report, {}, HarnessAgentRequest(max_steps=2, max_tool_calls=2), provider)

        self.assertTrue(payload["ok"])
        self.assertEqual(2, len(provider.calls))
        second_inputs = provider.calls[1]["inputs"]
        observation_text = json.dumps(second_inputs["observations"], ensure_ascii=False)
        self.assertLess(len(observation_text), 36000)
        self.assertNotIn("X" * 1000, observation_text)
        self.assertIn("result_preview", second_inputs["observations"][0])
        self.assertIn("evidence_layers", second_inputs["observations"][0])
        self.assertTrue(second_inputs["observations"][0]["evidence_layers"]["raw_layer"]["preview_omitted_for_model"])
        self.assertNotIn("result", second_inputs["observations"][0])
        self.assertLess(provider.calls[1]["inputs"]["observations"][0]["result_json_chars"], 100000)
        self.assertTrue(second_inputs["context_budget"]["truncated"])
        self.assertEqual("harness-observation-bundle", second_inputs["observation_bundle"]["id"])
        self.assertTrue(second_inputs["observation_bundle"]["evidence_ids"])
        self.assertNotIn("observations", second_inputs["observation_bundle"])
        self.assertEqual(1, second_inputs["context_budget"]["source_observation_count"])
        self.assertTrue(payload["trace_summary"]["input_truncated"])
        self.assertTrue(payload["context_budget"]["truncated"])
        self.assertIn("last_context_budget", payload["model_metadata"])
        self.assertIn("last_runtime_state", payload["model_metadata"])
        self.assertIn("last_session_state", payload["model_metadata"])
        self.assertTrue(payload["runtime_state"]["memory_summary"]["evidence_ids"])
        self.assertTrue(payload["session_state"]["recent_evidence_ids"])
        self.assertGreaterEqual(payload["trace_summary"]["runtime_evidence_id_count"], 1)

    def test_subagents_run_parallel_focused_profiles(self):
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(
                profile="full_review",
                max_steps=6,
                max_tool_calls=10,
                enable_subagents=True,
                subagent_profiles=("bom_depop", "derating"),
                max_subagents=2,
            ),
            MockHarnessModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertTrue(payload["subagent_summary"]["enabled"])
        self.assertEqual(2, payload["subagent_summary"]["planned_count"])
        self.assertEqual(["bom_depop", "derating"], [item["profile"] for item in payload["subagents"]])
        self.assertTrue(all(item["agent_run_id"] for item in payload["subagents"]))
        self.assertEqual("pstx-agent-subagents.v1", payload["subagent_summary"]["schema_version"])
        self.assertEqual("pstx-agent-subagents.v1", payload["subagents"][0]["schema_version"])
        self.assertEqual("fresh_context", payload["subagents"][0]["isolation"])
        self.assertIn("definition", payload["subagents"][0])
        self.assertTrue(payload["subagent_summary"]["provider_parallel_safe"])
        self.assertEqual(2, payload["trace_summary"]["subagent_count"])

    def test_subagent_provider_error_degrades_parent_result(self):
        provider = CloneToExplodingSubagentProvider()
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(
                profile="full_review",
                max_steps=2,
                max_tool_calls=2,
                enable_subagents=True,
                subagent_profiles=("bom_depop",),
                max_subagents=1,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("completed", payload["status"])
        self.assertEqual(1, provider.clone_count)
        self.assertEqual(1, payload["subagent_summary"]["failed_count"])
        self.assertTrue(payload["subagent_summary"]["degraded"])
        self.assertEqual(["bom_depop"], payload["subagent_summary"]["failed_profiles"])
        self.assertEqual("model_error", payload["subagents"][0]["status"])
        self.assertIn("simulated subagent provider failure", payload["subagents"][0]["answer"])
        self.assertEqual(1, payload["trace_summary"]["subagent_failed_count"])

    def test_subagent_provider_without_clone_runs_serially(self):
        provider = SequenceProvider([
            json.dumps({"final_answer": "parent done", "citations": []}),
            json.dumps({"final_answer": "child one done", "citations": []}),
            json.dumps({"final_answer": "child two done", "citations": []}),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(
                profile="full_review",
                max_steps=1,
                max_tool_calls=2,
                enable_subagents=True,
                subagent_profiles=("bom_depop", "derating"),
                max_subagents=2,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertFalse(payload["subagent_summary"]["provider_parallel_safe"])
        self.assertEqual(1, payload["subagent_summary"]["max_workers"])
        self.assertEqual(["bom_depop", "derating"], [item["profile"] for item in payload["subagents"]])

    def test_subagent_default_concurrency_is_conservative(self):
        request = HarnessAgentRequest.from_mapping({"enable_subagents": True})

        self.assertEqual(2, request.max_subagents)
        self.assertLessEqual(len(request.subagent_profiles), 3)

    def test_subagent_profiles_reject_full_review_recursion(self):
        with self.assertRaises(Exception) as ctx:
            HarnessAgentRequest.from_mapping({
                "enable_subagents": True,
                "subagent_profiles": ["full_review"],
            })

        self.assertIn("full_review", str(ctx.exception))

    def test_profiles_list_contains_expected_profiles(self):
        profiles = {item["id"]: item for item in list_harness_agent_profiles()}

        self.assertIn("auto", profiles)
        self.assertIn("quick_scan", profiles)
        self.assertIn("bom_depop", profiles)
        self.assertIn("feishu_bom_qa", profiles)
        self.assertIn("datasheet_qa", profiles)
        self.assertIn("dfmea_prep", profiles)
        self.assertIn("connection_datasheet_review", profiles)
        self.assertIn("agent_ref_qa", profiles)
        self.assertIn("review_checklist_qa", profiles)
        self.assertIn("full_review", profiles)
        self.assertIn("get_evidence_pack", profiles["bom_depop"]["tools"])
        self.assertIn("batch_query_report_entities", profiles["quick_scan"]["tools"])
        self.assertIn("summarize_table_column_values", profiles["quick_scan"]["tools"])
        self.assertIn("search_project_text", profiles["quick_scan"]["tools"])
        self.assertIn("get_harness_skill", profiles["quick_scan"]["tools"])
        self.assertIn("summarize_table_column_values", profiles["page_mapping"]["tools"])
        self.assertIn("search_project_text", profiles["page_mapping"]["tools"])
        self.assertIn("search_feishu_cache_rows", profiles["feishu_bom_qa"]["tools"])
        self.assertIn("batch_search_feishu_cache_rows", profiles["feishu_bom_qa"]["tools"])
        self.assertIn("search_datasheet_chunks", profiles["datasheet_qa"]["tools"])
        self.assertIn("select_harness_skills", profiles["datasheet_qa"]["tools"])
        self.assertIn("list_datasheet_review_templates", profiles["datasheet_qa"]["tools"])
        self.assertIn("batch_search_datasheet_chunks", profiles["datasheet_qa"]["tools"])
        self.assertIn("search_datasheet_parameters", profiles["datasheet_qa"]["tools"])
        self.assertIn("get_datasheet_parameter", profiles["datasheet_qa"]["tools"])
        self.assertIn("get_datasheet_chunk", profiles["datasheet_qa"]["tools"])
        self.assertIn("summarize_dfmea_readiness", profiles["dfmea_prep"]["tools"])
        self.assertIn("batch_get_component_identity_cards", profiles["dfmea_prep"]["tools"])
        self.assertIn("search_component_identity_cards", profiles["dfmea_prep"]["tools"])
        self.assertIn("search_datasheet_chunks", profiles["dfmea_prep"]["tools"])
        self.assertIn("get_datasheet_review_template", profiles["dfmea_prep"]["tools"])
        self.assertIn("search_datasheet_parameters", profiles["dfmea_prep"]["tools"])
        self.assertIn("search_datasheets", profiles["dfmea_prep"]["tools"])
        self.assertIn("batch_match_component_datasheets", profiles["dfmea_prep"]["tools"])
        self.assertIn("summarize_dfmea_datasheet_coverage", profiles["dfmea_prep"]["tools"])
        self.assertIn("batch_query_llm_topology_netlist", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("batch_get_component_identity_cards", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("batch_match_component_datasheets", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("search_datasheet_parameters", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("get_datasheet_chunk", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("search_project_text", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("trace_project_source", profiles["connection_datasheet_review"]["tools"])
        self.assertIn("search_agent_ref_pdfs", profiles["agent_ref_qa"]["tools"])
        self.assertIn("get_agent_ref_pdf_excerpt", profiles["agent_ref_qa"]["tools"])
        self.assertIn("search_review_checklists", profiles["review_checklist_qa"]["tools"])
        self.assertIn("get_review_checklist_excerpt", profiles["review_checklist_qa"]["tools"])
        self.assertGreaterEqual(profiles["full_review"]["max_steps"], 12)
        self.assertIn("resistor_bias", profiles["full_review"]["subagent_profiles"])

    def test_auto_profile_combines_multiple_capability_toolsets(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_table_rows",
                    "args": {"table_id": "missing_value", "limit": 5},
                    "reason": "复合问题中 BOM 能力需要读取报告表。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "auto done",
                "citations": [{"id": "ev-1-row-1"}],
            }),
        ])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(
                profile="auto",
                question="请同时检查 DFMEA 准备度、BOM_OPTION 和页码映射。",
                max_steps=2,
                max_tool_calls=2,
            ),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("auto", payload["profile"])
        capability_ids = [item["id"] for item in payload["capability_plan"]]
        self.assertIn("dfmea_prep", capability_ids)
        self.assertIn("bom_depop", capability_ids)
        self.assertIn("page_mapping", capability_ids)
        self.assertEqual("runtime_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("batch_query_report_entities", payload["tool_calls"][0]["tool"])
        self.assertEqual("get_table_rows", payload["tool_calls"][1]["tool"])

    def test_invalid_profile_is_rejected(self):
        with self.assertRaises(Exception) as ctx:
            HarnessAgentRequest.from_mapping({"profile": "unsafe"})

        self.assertIn("未知 agent profile", str(ctx.exception))

    def test_profile_limits_tools(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "read_project_text",
                    "args": {"path": "packaged/pstxprt.dat"},
                    "reason": "bom profile should not read files",
                }
            }),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="bom_depop"),
            provider,
        )

        self.assertFalse(payload["ok"])
        self.assertIn("profile bom_depop", payload["answer"])

    def test_feishu_profile_limits_tools_to_cache_queries(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_table_rows",
                    "args": {"table_id": "missing_value"},
                    "reason": "feishu profile should not read report tables",
                }
            }),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="feishu_bom_qa"),
            provider,
        )

        self.assertFalse(payload["ok"])
        self.assertIn("profile feishu_bom_qa", payload["answer"])

    def test_feishu_profile_search_creates_material_evidence_and_citation(self):
        make_feishu_cache(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_feishu_cache_rows",
                    "args": {"query": "HQ100", "limit": 5},
                    "reason": "搜索本地飞书缓存。",
                }
            }),
            json.dumps({
                "final_answer": "HQ100 对应 CAP-100N，PI 为 LiXinYu，选型顺序为 1。",
                "citations": [{"id": "ev-1-feishu-row-1", "note": "搜索结果命中 HQ100。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="feishu_bom_qa", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("search_feishu_cache_rows", payload["tool_calls"][0]["tool"])
        self.assertEqual("feishu_material", payload["final_evidence"][0]["type"])
        self.assertEqual("HQ100", payload["final_evidence"][0]["payload_preview"]["hq_no"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_agent_ref_profile_limits_tools_to_ref_pdf_queries(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_table_rows",
                    "args": {"table_id": "missing_value"},
                    "reason": "ref profile should not read report tables",
                }
            }),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="agent_ref_qa"),
            provider,
        )

        self.assertFalse(payload["ok"])
        self.assertIn("profile agent_ref_qa", payload["answer"])

    def test_agent_ref_profile_search_creates_pdf_evidence_and_citation(self):
        make_agent_ref_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_agent_ref_pdfs",
                    "args": {"query": "boundary citations", "limit": 5},
                    "reason": "搜索 ref PDF 资料。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "ref PDF 命中了 Agent Lab 能力边界说明。",
                "citations": [{"id": "ev-1-agent-ref-1-1-1", "note": "搜索结果命中 ref PDF。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="agent_ref_qa", question="请查 ref PDF 能力边界", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("search_agent_ref_pdfs", payload["tool_calls"][0]["tool"])
        self.assertEqual("agent_ref_excerpt", payload["final_evidence"][0]["type"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_review_checklist_profile_search_creates_evidence_and_citation(self):
        make_review_checklist_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_review_checklists",
                    "args": {"query": "U46 多 symbol HQ_CODE", "limit": 5},
                    "reason": "搜索真实 review checklist。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "review checklist 命中了 U46 多 symbol 检查经验。",
                "citations": [{"id": "ev-1-review-checklist-1-1-1", "note": "搜索结果命中 checklist。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="review_checklist_qa", question="请参考 review checklist 检查 U46", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("search_review_checklists", payload["tool_calls"][0]["tool"])
        self.assertEqual("review_checklist_excerpt", payload["final_evidence"][0]["type"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_chip_topology_profile_creates_topology_evidence(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "summarize_llm_topology_netlist",
                    "args": {"limit": 10},
                    "reason": "先抽取 LLM 芯片级拓扑网表摘要。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 与 U2 存在芯片到电平转换连接，主要共享 I2C 信号。",
                "citations": [{"id": "ev-1-llm-edge-chip-edge-u1-u2", "note": "LLM 拓扑连接证据。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="chip_topology", question="请说明大芯片和电平转换芯片的连接关系", max_steps=2, max_tool_calls=2),
            provider,
        )

        evidence_types = {item["type"] for item in payload["final_evidence"]}
        self.assertTrue(payload["ok"])
        self.assertEqual("summarize_llm_topology_netlist", payload["tool_calls"][0]["tool"])
        self.assertIn("llm_topology_summary", evidence_types)
        self.assertIn("llm_topology_edge", evidence_types)
        self.assertIn("llm_topology_node", evidence_types)
        self.assertTrue(payload["citations"][0]["valid"])
        observation = provider.calls[1]["inputs"]["observations"][0]
        self.assertIn("芯片级节点", observation["summary"])

    def test_llm_topology_query_evidence_handles_single_query_and_supply_edges(self):
        evidence = topology_evidence_nodes_from_tool_result(
            "query_llm_topology_netlist",
            {
                "query": "VCORE",
                "items": [
                    {
                        "kind": "supply_edge",
                        "edge": {
                            "edge_id": "supply-edge-pu1-u1-vcore",
                            "edge_kind": "supply",
                            "source_refdes": "PU1",
                            "target_refdes": "U1",
                            "supply_net": "VCORE",
                            "voltage_domain": "CORE",
                            "summary": "PU1 通过 VCORE 给 U1 提供供电关系。",
                        },
                        "summary": "供电关系命中。",
                    }
                ],
            },
            call_index=7,
        )

        self.assertIsNotNone(evidence)
        self.assertEqual(["llm_topology_supply_edge"], [item["type"] for item in evidence])
        self.assertEqual("supply-edge-pu1-u1-vcore", evidence[0]["locator"]["edge_id"])
        self.assertNotIn("missing_context", {item["type"] for item in evidence})

    def test_auto_profile_routes_chip_topology_question(self):
        profiles = {item["id"]: item for item in list_harness_agent_profiles()}
        self.assertIn("chip_topology", profiles)
        provider = SequenceProvider([
            json.dumps({
                "final_answer": "已基于芯片级拓扑完成回答。",
                "citations": [{"id": "ev-1-llm-topology-summary", "note": "拓扑摘要。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="auto", question="大芯片 U1 和电平转换芯片连接关系是什么？", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertIn("chip_level_topology", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])
        self.assertIn("summarize_llm_topology_netlist", payload["playbook_plan"]["recommended_first_tools"])
        self.assertEqual("runtime_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("summarize_llm_topology_netlist", payload["tool_calls"][0]["tool"])
        self.assertGreater(provider.calls[0]["inputs"]["tool_count"], 0)

    def test_auto_profile_prefetch_opens_document_detail(self):
        make_harness_doc_dir(self)

        class PrefetchAwareProvider:
            provider = "prefetch-aware"
            mode = "mock"

            def __init__(self):
                self.calls = []

            def generate_agent_step(self, prompt, *, inputs=None):
                self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                evidence_ids = []
                for observation in (inputs or {}).get("observations") or []:
                    evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                    for node in observation.get("evidence_nodes", []) or []:
                        node_id = str(node.get("id") or "")
                        if node_id:
                            evidence_ids.append(node_id)
                return HarnessModelResponse(
                    answer=json.dumps({
                        "final_answer": "已读取文档命中详情。",
                        "citations": [{"id": evidence_ids[-1], "note": "引用自动展开的文档片段。"}] if evidence_ids else [],
                    }, ensure_ascii=False),
                    provider=self.provider,
                    mode=self.mode,
                )

        provider = PrefetchAwareProvider()
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="auto", question="请在文档中搜索 U46 多 symbol", max_steps=1, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["search_documents", "get_document_excerpt"], [item["tool"] for item in payload["tool_calls"]])
        self.assertEqual("runtime_prefetch_followup", payload["agent_steps"][1]["type"])
        self.assertEqual(2, provider.calls[0]["inputs"]["tool_count"])

    def test_document_search_profile_creates_match_and_excerpt_evidence(self):
        make_harness_doc_dir(self)
        # Fill the dynamic doc_id after the search result is known by using a
        # tiny provider that inspects previous call inputs.
        class DocumentProvider:
            provider = "document-sequence"
            mode = "mock"

            def __init__(self):
                self.calls = []

            def generate_agent_step(self, prompt, *, inputs=None):
                self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                if len(self.calls) == 1:
                    return HarnessModelResponse(
                        answer=json.dumps({
                            "tool_call": {
                                "name": "search_documents",
                                "args": {"query": "U46 SECTION_NUMBER", "limit": 5},
                                "reason": "先搜索本地文档。",
                            }
                        }, ensure_ascii=False),
                        provider=self.provider,
                        mode=self.mode,
                    )
                if len(self.calls) == 2:
                    observations = (inputs or {}).get("observations") or []
                    nodes = observations[0].get("evidence_nodes") or []
                    doc_id = nodes[0]["locator"]["doc_id"]
                    return HarnessModelResponse(
                        answer=json.dumps({
                            "tool_call": {
                                "name": "get_document_excerpt",
                                "args": {"doc_id": doc_id, "char_start": 0, "max_chars": 1000},
                                "reason": "读取命中文档上下文。",
                            }
                        }, ensure_ascii=False),
                        provider=self.provider,
                        mode=self.mode,
                    )
                if len(self.calls) == 3:
                    observations = (inputs or {}).get("observations") or []
                    excerpt_node = observations[-1].get("evidence_nodes", [{}])[0].get("id", "")
                    return HarnessModelResponse(
                        answer=json.dumps({
                            "final_answer": "文档中记录了 U46 多 symbol 需要按 SECTION_NUMBER 汇总。",
                            "citations": [{"id": excerpt_node, "note": "文档片段。"}],
                        }, ensure_ascii=False),
                        provider=self.provider,
                        mode=self.mode,
                    )
                return HarnessModelResponse(answer=json.dumps({"final_answer": "done"}, ensure_ascii=False), provider=self.provider, mode=self.mode)

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="document_search", question="请在文档中搜索 U46 多 symbol", max_steps=3, max_tool_calls=3),
            DocumentProvider(),
        )

        evidence_types = {item["type"] for item in payload["final_evidence"]}
        self.assertTrue(payload["ok"])
        self.assertEqual("search_documents", payload["tool_calls"][0]["tool"])
        self.assertEqual("get_document_excerpt", payload["tool_calls"][1]["tool"])
        self.assertIn("document_match", evidence_types)
        self.assertIn("document_excerpt", evidence_types)
        self.assertTrue(payload["citations"][0]["valid"])

    def test_dfmea_prep_profile_limits_tools_to_identity_and_cache_queries(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "read_project_text",
                    "args": {"path": "packaged/pstxprt.dat"},
                    "reason": "dfmea prep should not read project files in phase 1",
                }
            }),
        ])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep"),
            provider,
        )

        self.assertFalse(payload["ok"])
        self.assertIn("profile dfmea_prep", payload["answer"])

    def test_dfmea_prep_profile_creates_identity_evidence_and_citation(self):
        make_feishu_cache(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "summarize_dfmea_readiness",
                    "args": {},
                    "reason": "先看 DFMEA 输入准备度。",
                }
            }),
            json.dumps({
                "final_answer": "U1 已具备基础身份证据，PU2 仍需补 HQ 料号和飞书匹配。",
                "citations": [{"id": "ev-1-dfmea-readiness", "note": "准备度摘要。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", max_steps=2, max_tool_calls=2),
            provider,
        )

        evidence_types = {item["type"] for item in payload["final_evidence"]}
        self.assertTrue(payload["ok"])
        self.assertEqual("summarize_dfmea_readiness", payload["tool_calls"][0]["tool"])
        self.assertIn("dfmea_readiness", evidence_types)
        self.assertIn("component_identity", evidence_types)
        self.assertIn("material_match", evidence_types)
        self.assertIn("missing_context", evidence_types)
        self.assertTrue(payload["citations"][0]["valid"])
        first_model_observation = provider.calls[1]["inputs"]["observations"][0]
        self.assertIn("evidence_nodes", first_model_observation)
        self.assertLess(json.dumps(first_model_observation, ensure_ascii=False).count("pin_net_summary"), 4)

    def test_mock_provider_supports_dfmea_prep(self):
        make_feishu_cache(self)
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", max_steps=2, max_tool_calls=2),
            MockHarnessModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("waiting_for_user", payload["status"])
        self.assertEqual("summarize_dfmea_readiness", payload["tool_calls"][0]["tool"])
        self.assertIn("DFMEA 准备度", payload["observations"][0]["title"])
        self.assertEqual("needs_user_input", payload["model_metadata"]["stopped_reason"])
        self.assertTrue(payload["needs_user_input"]["questions"])

    def test_dfmea_prep_datasheet_search_creates_excerpt_evidence(self):
        make_datasheet_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_datasheets",
                    "args": {"query": "HQ100 GPU_CORE_TEST_IC", "limit": 5},
                    "reason": "搜索本地规格书证据。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "U1 命中本地规格书片段，后续 DFMEA 可引用该证据。",
                "citations": [{"id": "ev-1-datasheet-1-1-1", "note": "HQ100 规格书命中。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", question="请查 U1 的规格书证据", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("search_datasheets", payload["tool_calls"][0]["tool"])
        self.assertEqual("datasheet_excerpt", payload["final_evidence"][0]["type"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_datasheet_qa_searches_chunks_and_reads_detail(self):
        make_datasheet_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_datasheet_chunks",
                    "args": {"query": "HQ100 GPU_CORE_TEST_IC electrical limits", "limit": 5},
                    "reason": "先检索本地 datasheet chunk。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "tool_call": {
                    "name": "get_datasheet_chunk",
                    "args": {"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000},
                    "reason": "定量/电气限制类结论需要读取完整 chunk。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "HQ100 的规格书片段已读取，可作为参数核对证据。",
                "citations": [{"id": "ev-2-datasheet-chunk-1-p1-c1", "note": "已读取 detail chunk。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="datasheet_qa", max_steps=3, max_tool_calls=3),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["search_datasheet_chunks", "get_datasheet_chunk"], [call["tool"] for call in payload["tool_calls"]])
        self.assertIn("datasheet_chunk", {item["type"] for item in payload["final_evidence"]})
        self.assertTrue(payload["citations"][0]["valid"])

    def test_datasheet_search_citation_auto_opens_detail_before_final(self):
        make_datasheet_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_datasheet_chunks",
                    "args": {"query": "HQ100 GPU_CORE_TEST_IC electrical limits", "limit": 5},
                    "reason": "先检索本地 datasheet chunk。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "HQ100 的规格书搜索命中可用于参数判断。",
                "citations": [{"id": "ev-1-datasheet-chunk-1-p1-c1", "note": "搜索命中。"}],
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "已打开 detail chunk 后确认 HQ100 规格书证据。",
                "citations": [{"id": "ev-2-datasheet-chunk-1-p1-c1", "note": "detail chunk。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="datasheet_qa", max_steps=4, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["search_datasheet_chunks", "get_datasheet_chunk"], [call["tool"] for call in payload["tool_calls"]])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertEqual("已打开 detail chunk 后确认 HQ100 规格书证据。", payload["answer"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_quantitative_datasheet_answer_auto_opens_detail_before_final(self):
        make_datasheet_index(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "search_datasheet_chunks",
                    "args": {"query": "HQ100 GPU_CORE_TEST_IC recommended operating voltage", "limit": 5},
                    "reason": "先检索本地 datasheet chunk。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "HQ100 推荐工作电压为 3.3V。",
                "citations": [],
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "已打开 detail chunk 后确认 HQ100 推荐工作电压证据。",
                "citations": [{"id": "ev-2-datasheet-chunk-1-p1-c1", "note": "detail chunk 原文。"}],
            }, ensure_ascii=False),
        ])

        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="datasheet_qa", question="HQ100 推荐工作电压是多少？", max_steps=4, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["search_datasheet_chunks", "get_datasheet_chunk"], [call["tool"] for call in payload["tool_calls"]])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertEqual("已打开 detail chunk 后确认 HQ100 推荐工作电压证据。", payload["answer"])
        self.assertTrue(payload["citations"][0]["valid"])

    def test_mock_provider_supports_dfmea_datasheet_coverage_question(self):
        make_datasheet_index(self)
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", question="请检查规格书覆盖情况", max_steps=2, max_tool_calls=2),
            MockHarnessModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("summarize_dfmea_datasheet_coverage", payload["tool_calls"][0]["tool"])
        self.assertIn("datasheet_match", {item["type"] for item in payload["final_evidence"]})

    def test_context_answers_allow_dfmea_prep_to_continue(self):
        make_feishu_cache(self)
        request = HarnessAgentRequest(
            profile="dfmea_prep",
            max_steps=2,
            max_tool_calls=2,
            context_answers=({
                "question_id": "dfmea-missing-context-1",
                "answer": "PU2 暂无 HQ 料号，芯片类别为电源管理芯片，后续人工补规格。",
                "applies_to": {"refdes": "PU2", "field": "hq_no/spec/chip_type"},
            },),
        )
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            request,
            MockHarnessModelProvider(),
            project_context={
                "answers": list(request.context_answers),
                "pending_questions": [],
                "recent_agent_runs": [],
                "recent_evidence_ids": [],
            },
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("completed", payload["status"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertIn("不生成正式 DFMEA", payload["answer"])

    def test_continuation_pack_is_available_to_next_model_step(self):
        active_pack = {
            "version": "agent-continuation-pack/v1",
            "agent_run_id": "agent-prev",
            "profile": "dfmea_prep",
            "status": "waiting_for_user",
            "next_intent": "continue_dfmea_missing_context",
            "goal": "继续确认 PU2 的 DFMEA 准备输入。",
            "continuation_brief": "上一轮发现 PU2 缺少规格和芯片类别，需要接着追问或读取身份卡。",
            "evidence_ids": ["component_identity:PU2"],
            "pending_questions": [{"question_id": "q-pu2-spec", "question": "请补充 PU2 规格。"}],
            "open_ledger_items": [{"id": "ledger-1", "title": "确认 PU2 身份"}],
            "suggested_tool_calls": [{"name": "get_component_identity_card", "args": {"refdes": "PU2"}}],
            "quality_status": "needs-more-evidence",
            "quality_score": 62,
        }
        provider = SequenceProvider(['{"final_answer":"已接续上一轮任务。","confidence":"medium","citations":[]}'])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(
                profile="dfmea_prep",
                question="继续上一轮",
                max_steps=1,
                max_tool_calls=1,
                continue_agent_run_id="agent-prev",
            ),
            provider,
            project_context={
                "answers": [],
                "pending_questions": [],
                "recent_agent_runs": [],
                "recent_evidence_ids": [],
                "active_continuation_pack": active_pack,
                "latest_continuation_pack": {"agent_run_id": "agent-latest", "next_intent": "latest-only"},
            },
        )

        self.assertTrue(payload["ok"])
        model_context = provider.calls[0]["inputs"]["project_context"]
        self.assertEqual("agent-prev", model_context["active_continuation_pack"]["agent_run_id"])
        self.assertEqual("continue_dfmea_missing_context", model_context["active_continuation_pack"]["next_intent"])
        self.assertEqual("agent-latest", model_context["latest_continuation_pack"]["agent_run_id"])
        self.assertIn("active_continuation_pack", provider.calls[0]["prompt"])
        self.assertIn("continue_agent_run_id", provider.calls[0]["prompt"])

    def test_project_session_memory_is_available_to_model_step(self):
        provider = SequenceProvider(['{"final_answer":"已参考项目记忆。","confidence":"medium","citations":[]}'])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", question="继续分析 PU2", max_steps=1, max_tool_calls=1),
            provider,
            project_context={
                "answers": [],
                "pending_questions": [],
                "recent_agent_runs": [],
                "recent_evidence_ids": [],
                "evidence_memory_cards": [{
                    "id": "ev-pu2-memory",
                    "type": "component_identity",
                    "title": "PU2 身份卡",
                    "summary": "PU2 已确认是电源管理芯片。",
                    "locator": {"refdes": "PU2"},
                    "detail_tool": {"name": "get_component_identity_card", "args": {"refdes": "PU2"}},
                }],
                "session_memory_summary": {
                    "version": "agent-project-session-memory/v1",
                    "goal": "继续 DFMEA 准备",
                    "facts": ["PU2 已由用户确认是电源管理芯片"],
                    "open_questions": ["PU2 缺规格书"],
                    "next_actions": ["读取 PU2 身份卡"],
                    "evidence_ids": ["component_identity:PU2"],
                },
            },
        )

        self.assertTrue(payload["ok"])
        model_context = provider.calls[0]["inputs"]["project_context"]
        self.assertEqual(
            "agent-project-session-memory/v1",
            model_context["session_memory_summary"]["version"],
        )
        self.assertIn("PU2 已由用户确认是电源管理芯片", model_context["session_memory_summary"]["facts"])
        self.assertEqual("ev-pu2-memory", model_context["evidence_memory_cards"][0]["id"])
        self.assertIn("session_memory_summary", provider.calls[0]["prompt"])
        self.assertIn("get_project_memory_evidence", provider.calls[0]["prompt"])

    def test_auto_profile_prefetches_project_evidence_memory_before_model_step(self):
        provider = SequenceProvider(['{"final_answer":"已基于上一轮 U1 证据继续。","confidence":"medium","citations":[{"id":"ev-1-memory-ev-u1-memory"}]}'])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="auto", question="继续刚才 U1 的分析", max_steps=1, max_tool_calls=2),
            provider,
            project_context={
                "answers": [],
                "pending_questions": [],
                "recent_agent_runs": [],
                "recent_evidence_ids": ["ev-u1-memory"],
                "evidence_memory_cards": [{
                    "id": "ev-u1-memory",
                    "type": "component_identity",
                    "title": "U1 身份卡",
                    "summary": "U1 是 FPGA，HQ=HQ100。",
                    "locator": {"refdes": "U1"},
                    "detail_tool": {"name": "get_component_identity_card", "args": {"refdes": "U1"}},
                }],
            },
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("runtime_memory_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("list_project_memory_evidence", payload["tool_calls"][0]["tool"])
        self.assertEqual(1, provider.calls[0]["inputs"]["tool_count"])
        self.assertIn("memory_prefetch_plan", payload["model_metadata"])
        self.assertIn("component_identity", {item["type"] for item in payload["final_evidence"]})
        self.assertEqual(
            "get_component_identity_card",
            payload["final_evidence"][0]["detail_tool"]["name"],
        )
        self.assertTrue(payload["citations"][0]["valid"])

    def test_auto_profile_goal_prefetches_dfmea_readiness_without_entities(self):
        provider = SequenceProvider(['{"final_answer":"已先读取 DFMEA 准备度。","confidence":"medium","citations":[{"id":"ev-1-dfmea-readiness"}]}'])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="auto", question="请对这个项目做 DFMEA 准备度分析", max_steps=1, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("runtime_goal_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("summarize_dfmea_readiness", payload["tool_calls"][0]["tool"])
        self.assertEqual(1, provider.calls[0]["inputs"]["tool_count"])
        self.assertIn("goal_prefetch_plan", payload["model_metadata"])
        self.assertIn("dfmea_readiness", {item["type"] for item in payload["final_evidence"]})
        self.assertTrue(payload["citations"][0]["valid"])

    def test_auto_profile_routes_from_session_memory_when_user_says_continue(self):
        provider = SequenceProvider(['{"final_answer":"已按上一轮 DFMEA 任务继续。","confidence":"medium","citations":[]}'])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="auto", question="继续", max_steps=1, max_tool_calls=0),
            provider,
            project_context={
                "answers": [],
                "pending_questions": [],
                "recent_agent_runs": [],
                "recent_evidence_ids": [],
                "session_memory_summary": {
                    "version": "agent-project-session-memory/v1",
                    "goal": "DFMEA 准备：补齐 PU2 规格书和芯片类别",
                    "open_items": ["继续 DFMEA 准备，搜索 datasheet PDF，确认失效模式输入"],
                    "next_actions": ["调用 summarize_dfmea_readiness 和 match_component_datasheets"],
                    "evidence_ids": ["component_identity:PU2"],
                },
            },
        )

        self.assertTrue(payload["ok"])
        self.assertIn("dfmea_prep", [item["id"] for item in payload["capability_plan"]])
        self.assertIn("dfmea_prep", provider.calls[0]["inputs"]["capability_profiles"])
        self.assertIn("DFMEA 准备", provider.calls[0]["inputs"]["planning_context"])
        self.assertGreater(payload["trace_summary"]["planning_context_chars"], 0)

    def test_needs_user_input_protocol_is_normalized(self):
        provider = SequenceProvider([
            json.dumps({
                "needs_user_input": {
                    "reason": "需要补充规格。",
                    "missing_fields": ["spec"],
                    "related_evidence_ids": ["ev-not-yet"],
                    "questions": [{
                        "question_id": "q-spec-u1",
                        "question": "请补充 U1 的规格型号。",
                        "applies_to": {"refdes": "U1", "field": "spec"},
                        "missing_fields": ["spec"],
                    }],
                }
            }, ensure_ascii=False),
        ])
        payload = run_harness_agent(
            sample_report(),
            sample_dfmea_bundle(),
            HarnessAgentRequest(profile="dfmea_prep", max_steps=1, max_tool_calls=0),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("waiting_for_user", payload["status"])
        self.assertEqual("needs_user_input", payload["model_metadata"]["stopped_reason"])
        self.assertEqual("q-spec-u1", payload["needs_user_input"]["questions"][0]["question_id"])
        self.assertEqual("spec", payload["needs_user_input"]["questions"][0]["missing_fields"][0])

    def test_invalid_citation_is_marked_and_fallback_kept(self):
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "list_report_tables",
                    "args": {},
                    "reason": "list tables",
                }
            }),
            json.dumps({
                "final_answer": "引用了一个不存在的证据。",
                "citations": [{"id": "ev-not-exist", "note": "bad citation"}],
            }),
        ])
        payload = run_harness_agent(sample_report(), {}, HarnessAgentRequest(), provider)

        invalid = [item for item in payload["citations"] if not item["valid"]]
        fallback = [item for item in payload["citations"] if item.get("fallback")]
        self.assertTrue(invalid)
        self.assertTrue(fallback)
        self.assertEqual(1, payload["model_metadata"]["invalid_citation_count"])

    def test_agent_retries_when_model_gives_premature_refusal(self):
        provider = SequenceProvider([
            json.dumps({"final_answer": "无法回答，信息不足。"}, ensure_ascii=False),
            json.dumps({
                "tool_call": {
                    "name": "list_report_tables",
                    "args": {},
                    "reason": "先读取本地报告表格清单再判断。",
                }
            }, ensure_ascii=False),
            json.dumps({"final_answer": "已基于表格清单继续取证。", "citations": []}, ensure_ascii=False),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(profile="quick_scan", question="请统计当前报告表格情况", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("已基于表格清单继续取证。", payload["answer"])
        self.assertEqual("list_report_tables", payload["tool_calls"][0]["tool"])
        self.assertEqual(1, payload["model_metadata"]["perseverance_retry_count"])
        self.assertIn("perseverance_retry", provider.calls[1]["inputs"])
        self.assertTrue(provider.calls[1]["inputs"]["perseverance_retry"])

    def test_task_ledger_guard_retries_empty_final_before_evidence(self):
        provider = SequenceProvider([
            json.dumps({"final_answer": "已完成。"}, ensure_ascii=False),
            json.dumps({"final_answer": "已根据任务账本继续确认。", "citations": []}, ensure_ascii=False),
        ])
        payload = run_harness_agent(
            sample_report(),
            {},
            HarnessAgentRequest(question="请统计 page_rows 有多少页码", max_steps=1, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(2, len(provider.calls))
        self.assertEqual(1, payload["model_metadata"]["perseverance_retry_count"])
        self.assertIn("task_ledger", provider.calls[1]["inputs"])
        self.assertIn("task_ledger", provider.calls[1]["inputs"]["perseverance_retry_note"])

    def test_debug_false_hides_long_file_content(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "packaged").mkdir()
            (root / "packaged" / "pstxprt.dat").write_text("A" * 2000, encoding="utf-8")
            provider = SequenceProvider([
                json.dumps({
                    "tool_call": {
                        "name": "read_project_text",
                        "args": {"path": "packaged/pstxprt.dat", "max_chars": 2000},
                        "reason": "read allowed file",
                    }
                }),
                json.dumps({"final_answer": "done"}),
            ])
            payload = run_harness_agent(
                sample_report(),
                {"project_root": str(root)},
                HarnessAgentRequest(profile="full_review", debug=False),
                provider,
            )

        result = payload["observations"][0]["result"]
        self.assertNotIn("content", result)
        self.assertIn("content_preview", result)
        self.assertLessEqual(len(result["content_preview"]), 500)

    def test_agent_does_not_mutate_report_or_bundle(self):
        report = sample_report()
        bundle = {"project_root": ""}
        before_report = copy.deepcopy(report)
        before_bundle = copy.deepcopy(bundle)

        run_harness_agent(report, bundle, HarnessAgentRequest(), MockHarnessModelProvider())

        self.assertEqual(before_report, report)
        self.assertEqual(before_bundle, bundle)


if __name__ == "__main__":
    unittest.main()
