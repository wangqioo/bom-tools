import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pstx_core.cadence.page_model import build_cadence_page_payload, compare_page_models, load_cadence_page_model
from pstx_core.cadence import page_model as core_page_model
from pstx_harness.compare_agent import CompareAgentRequest, CompareMockModelProvider, list_compare_agent_profiles, run_compare_agent
from pstx_harness.compare_tools import CompareToolContext, build_compare_tool_registry
from pstx_knowledge.datasheets import reindex_datasheets
from pstx_harness.model import HarnessModelResponse


def sample_compare_payload():
    return {
        "left": {"run_id": "left1", "project_name": "alpha"},
        "right": {"run_id": "right1", "project_name": "beta"},
        "diff_totals": {"key_pin_nets": 1, "components": 1, "nets": 1},
        "compare_sections": [
            {
                "id": "key_pin_nets",
                "title": "关键器件 Pin/Net 连接差异",
                "lead": "芯片连接变化",
                "priority": "critical",
                "diff": {"added_count": 0, "removed_count": 0, "changed_count": 1, "total_rows": 1},
                "table": {
                    "id": "compare_key_pin_nets",
                    "columns": ["类型", "位号", "引脚", "左侧网络", "右侧网络", "左侧PI", "右侧PI"],
                    "rows": [
                        {
                            "类型": "变化",
                            "位号": "U1",
                            "引脚": "A1",
                            "左侧网络": "SMBALERT_N",
                            "右侧网络": "SMBALERT_ALT_N",
                            "左侧PI": "PI-A",
                            "右侧PI": "PI-B",
                        }
                    ],
                },
            },
            {
                "id": "components",
                "title": "元件属性差异",
                "lead": "全量属性",
                "priority": "normal",
                "diff": {"added_count": 0, "removed_count": 0, "changed_count": 1, "total_rows": 1},
                "table": {
                    "id": "compare_components",
                    "columns": ["类型", "位号", "变化字段"],
                    "rows": [{"类型": "变化", "位号": "U1", "变化字段": "PI"}],
                },
            },
        ],
    }


def sample_payloads(testcase: unittest.TestCase):
    left_temp = tempfile.TemporaryDirectory()
    right_temp = tempfile.TemporaryDirectory()
    testcase.addCleanup(left_temp.cleanup)
    testcase.addCleanup(right_temp.cleanup)
    left_root = Path(left_temp.name)
    right_root = Path(right_temp.name)
    (left_root / "packaged").mkdir()
    (right_root / "sch_1").mkdir()
    (left_root / "packaged" / "pstxprt.dat").write_text("PART_NAME U1\napiKey=SECRET_VALUE\n", encoding="utf-8")
    (right_root / "sch_1" / "page1.csv").write_text('"PAGE_NUMBER" = 1;\n', encoding="utf-8")
    return (
        {"bundle": {"project_root": str(left_root)}, "report": {"project_name": "alpha"}},
        {"bundle": {"project_root": str(right_root)}, "report": {"project_name": "beta"}},
    )


def sample_cadence_page_payloads(testcase: unittest.TestCase):
    left_temp = tempfile.TemporaryDirectory()
    right_temp = tempfile.TemporaryDirectory()
    testcase.addCleanup(left_temp.cleanup)
    testcase.addCleanup(right_temp.cleanup)
    left_root = Path(left_temp.name)
    right_root = Path(right_temp.name)
    (left_root / "sch_1").mkdir()
    (right_root / "sch_1").mkdir()
    (left_root / "sch_1" / "page1.csv").write_text('"PAGE_NUMBER" = 1;\nTITLE=MAIN;\n', encoding="utf-8")
    (right_root / "sch_1" / "page1.csv").write_text('"PAGE_NUMBER" = 1;\nTITLE=MAIN_B;\n', encoding="utf-8")
    (left_root / "sch_1" / "page1.csa").write_text(
        "\n".join([
            "PAGE_NUMBER = 1;",
            "WIRE 16 -1 (0 0)(100 0);",
            "FORCEPROP 2 LAST SIG_NAME SMB_ALERT;",
            "DOT 1 (50 0);",
            "CIRCLE 16 -1 (1000 1000) 100;",
            "MYSTERY_OBJECT A B C;",
        ]),
        encoding="utf-8",
    )
    (right_root / "sch_1" / "page1.csa").write_text(
        "\n".join([
            "MYSTERY_OBJECT A B D;",
            "DOT 1 (50 0);",
            "WIRE 16 -1 (0 0)(120 0);",
            "FORCEPROP 2 LAST SIG_NAME SMB_ALERT;",
            "CIRCLE 16 -1 (1000 1000) 100;",
            "PAGE_NUMBER = 1;",
        ]),
        encoding="utf-8",
    )
    return (
        {"bundle": {"project_root": str(left_root)}, "report": {"project_name": "alpha"}},
        {"bundle": {"project_root": str(right_root)}, "report": {"project_name": "beta"}},
    )


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


class CompareAgentTests(unittest.TestCase):
    def test_core_cadence_page_model_entrypoint_exports_public_api(self):
        self.assertTrue(callable(compare_page_models))
        self.assertTrue(callable(load_cadence_page_model))
        self.assertFalse(hasattr(core_page_model, "_parse_csa_file"))
        self.assertFalse(Path("pstx_cadence_page_model.py").exists())

    def test_compare_tools_read_sections_query_and_rows(self):
        left, right = sample_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())

        sections = registry.run("list_compare_sections", context, {})
        self.assertEqual(2, len(sections["sections"]))

        rows = registry.run("get_compare_section_rows", context, {"section_id": "key_pin_nets", "limit": 1})
        self.assertEqual("U1", rows["rows"][0]["位号"])

        query = registry.run("query_compare_diff", context, {"query": "PI-B", "limit": 5})
        self.assertEqual(1, len(query["matches"]))
        self.assertEqual("key_pin_nets", query["matches"][0]["section_id"])

        detail = registry.run("get_compare_row", context, {"section_id": "components", "row_index": 0})
        self.assertEqual("U1", detail["row"]["位号"])

    def test_compare_batch_tools_query_and_read_rows(self):
        left, right = sample_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())

        query = registry.run(
            "batch_query_compare_diff",
            context,
            {"queries": ["U1", "PI-B", "NO_SUCH"], "limit_per_query": 2},
        )
        self.assertEqual(["found", "found", "missing"], [item["status"] for item in query["items"]])
        self.assertEqual("key_pin_nets", query["items"][0]["matches"][0]["section_id"])

        rows = registry.run(
            "batch_get_compare_rows",
            context,
            {"items": [{"section_id": "key_pin_nets", "row_index": 0}, {"section_id": "bad", "row_index": 0}]},
        )
        self.assertEqual("found", rows["items"][0]["status"])
        self.assertEqual("U1", rows["items"][0]["row"]["位号"])
        self.assertEqual("error", rows["items"][1]["status"])

    def test_compare_harness_skill_tools_are_readonly_and_shared(self):
        left, right = sample_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())
        tools = {item["name"]: item for item in registry.list_tools()}

        for name in ("list_harness_skills", "select_harness_skills", "get_harness_skill"):
            self.assertIn(name, tools)
            self.assertTrue(tools[name]["readonly"])
            self.assertEqual("harness_skill", tools[name]["evidence_kind"])
            self.assertEqual("none", tools[name]["approval_scope"])

        selected = registry.run(
            "select_harness_skills",
            context,
            {
                "query": "对比 HQ11112042009 datasheet recommended operating 参数",
                "capability_profiles": ["compare_datasheet_qa"],
                "include_body": True,
                "max_body_chars": 1200,
            },
        )
        self.assertIn("datasheet-key-info", [card["id"] for card in selected["harness_skills"]["skills"]])

        detail = registry.run(
            "get_harness_skill",
            context,
            {"skill_id": "datasheet-key-info", "max_body_chars": 6000},
        )
        self.assertEqual("datasheet-key-info", detail["skill"]["id"])
        self.assertIn("Compare 场景", detail["skill"]["body"])

    def test_compare_harness_reuses_datasheet_chunk_tools(self):
        old_dir = os.environ.get("PSTX_DATASHEET_DIR")
        old_data_dir = os.environ.get("PSTX_DATASHEET_DATA_DIR")
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.addCleanup(lambda: os.environ.pop("PSTX_DATASHEET_DIR", None) if old_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DIR", old_dir))
        self.addCleanup(lambda: os.environ.pop("PSTX_DATASHEET_DATA_DIR", None) if old_data_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DATA_DIR", old_data_dir))
        root = Path(tmp.name)
        source = root / "pdfs"
        source.mkdir()
        (source / "HQ11112042009_LCMXO3LF.pdf").write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(source)
        os.environ["PSTX_DATASHEET_DATA_DIR"] = str(root / "data")
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQ11112042009 LCMXO3LF absolute maximum ratings and recommended operating conditions"], "fake", ""),
        ):
            reindex_datasheets(force=True)

        left, right = sample_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())

        documents = registry.run("list_datasheet_documents", context, {"limit": 10})
        self.assertEqual(1, documents["total_documents"])

        batch = registry.run(
            "batch_search_datasheet_chunks",
            context,
            {"queries": ["HQ11112042009", "NO_MATCH"], "limit_per_query": 2},
        )
        self.assertEqual(["found", "missing"], [item["status"] for item in batch["items"]])
        match = batch["items"][0]["matches"][0]
        chunk = registry.run(
            "get_datasheet_chunk",
            context,
            {"doc_id": match["doc_id"], "chunk_id": match["chunk_id"], "max_chars": 200},
        )
        self.assertIn("absolute maximum", chunk["content"])

    def test_compare_project_file_tools_are_readonly_and_bounded(self):
        left, right = sample_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())

        files = registry.run("list_compare_project_files", context, {"side": "both", "limit": 10})
        self.assertTrue(any(item["side"] == "left" and item["path"] == "packaged/pstxprt.dat" for item in files["files"]))
        self.assertTrue(any(item["side"] == "right" and item["path"] == "sch_1/page1.csv" for item in files["files"]))

        text = registry.run("read_compare_project_text", context, {"side": "left", "path": "packaged/pstxprt.dat", "max_chars": 80})
        self.assertIn("PART_NAME", text["content"])
        self.assertNotIn("SECRET_VALUE", text["content"])

        with self.assertRaises(Exception) as ctx:
            registry.run("read_compare_project_text", context, {"side": "left", "path": "../secret.txt"})
        self.assertIn("项目根目录之外", str(ctx.exception))

    def test_cadence_page_model_keeps_unknown_and_ignores_line_order(self):
        left_temp = tempfile.TemporaryDirectory()
        right_temp = tempfile.TemporaryDirectory()
        self.addCleanup(left_temp.cleanup)
        self.addCleanup(right_temp.cleanup)
        left_root = Path(left_temp.name)
        right_root = Path(right_temp.name)
        (left_root / "sch_1").mkdir()
        (right_root / "sch_1").mkdir()
        left_csa = "\n".join([
            "WIRE 16 -1 (0 0)(100 0);",
            "FORCEPROP 2 LAST SIG_NAME NET_A;",
            "DOT 1 (50 0);",
            "ARC 16 -1 (3000 3000)(3100 3000)(3050 3050);",
            "UNKNOWN_DEHL_LINE X;",
        ])
        right_csa = "\n".join([
            "UNKNOWN_DEHL_LINE X;",
            "DOT 1 (50 0);",
            "WIRE 16 -1 (0 0)(100 0);",
            "FORCEPROP 2 LAST SIG_NAME NET_A;",
            "ARC 16 -1 (3000 3000)(3100 3000)(3050 3050);",
        ])
        (left_root / "sch_1" / "page1.csa").write_text(left_csa, encoding="utf-8")
        (right_root / "sch_1" / "page1.csa").write_text(right_csa, encoding="utf-8")

        left_model = load_cadence_page_model(left_root, "left", 1)
        right_model = load_cadence_page_model(right_root, "right", 1)
        self.assertIn("UNKNOWN", left_model.counts())
        self.assertEqual(1, left_model.counts()["UNKNOWN"])
        self.assertEqual(1, left_model.counts()["ARC"])
        diff = compare_page_models(left_model, right_model)
        self.assertEqual("same", diff["status"])
        self.assertEqual(0, diff["diff_count"])

    def test_cadence_page_model_extracts_connection_semantics(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "sch_1").mkdir()
            (root / "sch_1" / "page1.csv").write_text('"PAGE_NUMBER" = 1;\n', encoding="utf-8")
            (root / "sch_1" / "page1.csa").write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME SMB_ALERT;",
                    "NET_LABEL 1 (50 0) SMB_ALERT;",
                    "PORT 1 (100 0) SMB_ALERT INPUT;",
                    "OFFPAGE 1 (0 0) SMB_ALERT_REMOTE;",
                    "BUS 1 (75 0) SMBUS[0..1];",
                    "NO_CONNECT 1 (200 200);",
                    "NET_LABEL 1 (300 300) FLOATING_LABEL;",
                    "MYSTERY_OBJECT A B C;",
                ]),
                encoding="utf-8",
            )

            model = load_cadence_page_model(root, "project", 1)

        self.assertEqual(2, model.counts()["NET_LABEL"])
        self.assertEqual(1, model.counts()["PORT"])
        self.assertEqual(1, model.counts()["OFFPAGE"])
        self.assertEqual(1, model.counts()["BUS"])
        self.assertEqual(1, model.counts()["NO_CONNECT"])
        self.assertEqual(2, len(model.unbound_semantics))
        conn = model.connectivity[0]
        self.assertEqual(["SMB_ALERT"], conn.signal_names)
        self.assertEqual(["SMB_ALERT"], conn.labels)
        self.assertEqual(["SMB_ALERT"], conn.ports)
        self.assertEqual(["SMB_ALERT_REMOTE"], conn.offpage_connectors)
        self.assertEqual(["SMBUS[0..1]"], conn.bus_names)
        self.assertEqual([], conn.no_connect_points)
        self.assertTrue(any(item.attributes["name"] == "FLOATING_LABEL" for item in model.unbound_semantics))
        self.assertTrue(any(item.object_type == "NO_CONNECT" for item in model.unbound_semantics))

    def test_cadence_page_payload_returns_object_detail(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "sch_1").mkdir()
            (root / "sch_1" / "page1.csa").write_text(
                "WIRE 16 -1 (0 0)(100 0);\nNET_LABEL 1 (50 0) NET_A;\n",
                encoding="utf-8",
            )

            payload = build_cadence_page_payload(root, 1, stdout="objects", limit=20)
            object_id = next(item["object_id"] for item in payload["objects"] if item["type"] == "NET_LABEL")
            detail = build_cadence_page_payload(root, 1, object_id=object_id)

        self.assertEqual("pstx-cadence-page.v1", payload["schema_version"])
        self.assertFalse(payload["truncated"])
        self.assertEqual(object_id, detail["object"]["object_id"])

    def test_cadence_page_payload_only_collects_junction_detail_for_full_mode(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "sch_1").mkdir()
            (root / "sch_1" / "page1.csa").write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "WIRE 16 -1 (50 -50)(50 50);",
                ]),
                encoding="utf-8",
            )

            objects_payload = build_cadence_page_payload(root, 1, stdout="objects", limit=20)
            full_payload = build_cadence_page_payload(root, 1, stdout="full", limit=20)
            conn_detail = build_cadence_page_payload(root, 1, object_id="conn-1", limit=20)

        self.assertEqual([], objects_payload["connectivity"][0]["junctions"])
        self.assertEqual([[50, 0]], full_payload["connectivity"][0]["junctions"])
        self.assertEqual([[50, 0]], conn_detail["object"]["junctions"])

    def test_compare_cadence_page_semantics_detects_page_file_changes(self):
        left, right = sample_cadence_page_payloads(self)
        registry = build_compare_tool_registry()
        context = CompareToolContext(sample_compare_payload(), left, right, CompareAgentRequest())

        page_range = registry.run("resolve_compare_page_range", context, {"page_range": "第1-1页"})
        self.assertEqual([1], page_range["pages"])
        self.assertIn("页码", page_range["summary"])

        diff = registry.run(
            "compare_cadence_page_semantics",
            context,
            {"page_start": 1, "page_end": 1, "include_raw_unknown": True, "max_diff_items": 20},
        )
        self.assertEqual([1], diff["changed_pages"])
        self.assertEqual("changed", diff["page_results"][0]["status"])
        item_types = {item["item_type"] for item in diff["page_results"][0]["diffs"]}
        self.assertIn("WIRE", item_types)
        self.assertIn("CSV_PROPERTY", item_types)
        self.assertIn("UNKNOWN", item_types)

        wire_id = next(
            item["left"]["object_id"]
            for item in diff["page_results"][0]["diffs"]
            if item["item_type"] == "WIRE" and item.get("left")
        )
        detail = registry.run("get_cadence_page_object", context, {"side": "left", "page": 1, "object_id": wire_id})
        self.assertEqual(wire_id, detail["object"]["object_id"])

        batch_detail = registry.run(
            "batch_get_cadence_page_objects",
            context,
            {"objects": [
                {"side": "left", "page": 1, "object_id": wire_id},
                {"side": "right", "page": 1, "object_id": "NO_SUCH"},
            ]},
        )
        self.assertEqual("found", batch_detail["items"][0]["status"])
        self.assertEqual(wire_id, batch_detail["items"][0]["object"]["object_id"])
        self.assertEqual("error", batch_detail["items"][1]["status"])

        raw = registry.run("get_cadence_page_raw_excerpt", context, {"side": "left", "page": 1, "file_type": "csa", "max_chars": 80})
        self.assertIn("WIRE", raw["content"])

    def test_mock_compare_agent_calls_tool_then_final_answer(self):
        left, right = sample_payloads(self)
        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(),
            CompareMockModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("local-compare-agent-harness", payload["mode"])
        self.assertEqual("compare_quick_scan", payload["profile"])
        self.assertEqual("summarize_compare_risks", payload["tool_calls"][0]["tool"])
        self.assertTrue(payload["raw_observations"])
        self.assertIn("evidence_layers", payload["observations"][0])
        self.assertIn("raw_result", payload["raw_observations"][0])
        self.assertIn("guidance_summary", payload)
        self.assertIn("selected_skills", payload)
        self.assertIn("effort_policy", payload)
        self.assertIn("task_memory_summary", payload)
        self.assertTrue(payload["final_evidence"])
        self.assertTrue(payload["citations"])
        self.assertTrue(payload["citations"][0]["valid"])
        self.assertTrue(payload["context_budget"]["truncated"])
        self.assertTrue(payload["trace_summary"]["input_truncated"])
        self.assertIn("last_context_budget", payload["model_metadata"])
        self.assertIn("runtime_state", payload)
        self.assertIn("last_runtime_state", payload["model_metadata"])
        self.assertIn("session_state", payload)
        self.assertIn("last_session_state", payload["model_metadata"])
        self.assertIn("evidence_goal_contract", payload)
        self.assertEqual("agent-evidence-goal-contract/v1", payload["evidence_goal_contract"]["version"])
        self.assertIn("evidence_goal_status", payload["trace_summary"])
        self.assertGreaterEqual(payload["trace_summary"]["runtime_evidence_id_count"], 1)

    def test_compare_agent_inputs_include_runtime_state(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([json.dumps({"final_answer": "done", "citations": []})])

        payload = run_compare_agent(sample_compare_payload(), left, right, CompareAgentRequest(), provider)

        self.assertTrue(payload["ok"])
        self.assertIn("dispatch_tasks", provider.calls[0]["prompt"])
        self.assertIn("runtime_state", provider.calls[0]["inputs"])
        self.assertEqual("pstx-agent-runtime/v1", provider.calls[0]["inputs"]["runtime_state"]["protocol_version"])
        self.assertEqual("agent-task-ledger/v1", provider.calls[0]["inputs"]["runtime_state"]["task_ledger"]["version"])
        self.assertEqual("agent-evidence-goal-contract/v1", provider.calls[0]["inputs"]["runtime_state"]["evidence_goal_contract"]["version"])
        self.assertIn("task_ledger", provider.calls[0]["inputs"])
        self.assertIn("session_state", provider.calls[0]["inputs"])
        self.assertEqual("pstx-agent-runtime/v1", provider.calls[0]["inputs"]["session_state"]["protocol_version"])
        self.assertEqual("agent-task-ledger/v1", provider.calls[0]["inputs"]["session_state"]["task_ledger"]["version"])
        self.assertEqual("pstx-agent-runtime/v1", payload["runtime_state"]["protocol_version"])
        self.assertEqual("pstx-agent-runtime/v1", payload["session_state"]["protocol_version"])
        self.assertIn("task_ledger", payload["runtime_state"])
        self.assertIn("evidence_goal_contract", payload["runtime_state"])
        self.assertIn("task_ledger_open_count", payload["trace_summary"])
        self.assertIn("final_answer_quality_gate", payload)
        self.assertIn("repair_actions", payload["final_answer_quality_gate"])
        self.assertIn("repair_action_count", payload["final_answer_quality_gate"])
        self.assertIn("final_answer_quality_gate", payload["model_metadata"])
        self.assertIn("final_quality_status", payload["trace_summary"])
        self.assertIn("execution_journal", payload)
        self.assertIn("journal_summary", payload)
        self.assertGreaterEqual(payload["journal_summary"]["event_count"], 2)
        self.assertEqual("pstx-harness-turn-context.v1", payload["turn_context_snapshot"]["schema_version"])
        self.assertEqual("pstx-tool-dispatch-summary.v1", payload["tool_dispatch_summary"]["schema_version"])
        self.assertIn("tool_dispatch_event_count", payload["trace_summary"])
        self.assertIn("continuation_pack", payload)
        self.assertEqual("agent-continuation-pack/v1", payload["continuation_pack"]["version"])

    def test_compare_agent_dispatches_long_tasks_with_callback(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "dispatch_tasks": [{
                    "task_id": "cmp-u1",
                    "title": "U1 差异",
                    "profile": "compare_datasheet_qa",
                    "question": "对比 U1 datasheet 关键参数差异。",
                }],
                "reason": "datasheet 对比可后台执行。",
            }, ensure_ascii=False)
        ])
        seen = []

        def dispatch_callback(request):
            seen.append(request)
            return {
                "task_dispatch_summary": {"queue": "accepted"},
                "dispatched_tasks": [{
                    "task_id": "cmp-u1",
                    "title": "U1 差异",
                    "profile": "compare_datasheet_qa",
                    "question": "对比 U1 datasheet 关键参数差异。",
                    "agent_run_id": "child-compare-1",
                    "status": "queued",
                    "status_url": "/api/harness/agent-runs/child-compare-1",
                }],
            }

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(max_steps=1, max_tool_calls=1, debug=True),
            provider,
            dispatch_callback=dispatch_callback,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("completed", payload["status"])
        self.assertEqual("task_dispatched", payload["model_metadata"]["stopped_reason"])
        self.assertEqual("compare", seen[0]["source"])
        self.assertEqual("cmp-u1", seen[0]["tasks"][0]["task_id"])
        self.assertEqual("child-compare-1", payload["dispatched_tasks"][0]["agent_run_id"])
        self.assertTrue(payload["task_dispatch_summary"]["available"])

    def test_compare_agent_returns_dispatch_plan_without_callback(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "dispatch_tasks": [{
                    "task_id": "cmp-page",
                    "title": "Cadence 差异",
                    "profile": "compare_cadence_pages",
                    "question": "对比第 1 页 Cadence 连接语义。",
                }]
            }, ensure_ascii=False)
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(max_steps=1, max_tool_calls=1),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertFalse(payload["task_dispatch_summary"]["available"])
        self.assertEqual("not_dispatched", payload["dispatched_tasks"][0]["status"])

    def test_compare_quality_gate_repair_executes_next_page_then_reanswers(self):
        left, right = sample_payloads(self)
        compare_payload = sample_compare_payload()
        compare_payload["compare_sections"][0]["table"]["rows"].append({
            "类型": "变化",
            "位号": "U2",
            "引脚": "B2",
            "左侧网络": "SCL",
            "右侧网络": "SCL_ALT",
            "左侧PI": "PI-C",
            "右侧PI": "PI-D",
        })
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "get_compare_section_rows", "args": {"section_id": "key_pin_nets", "limit": 1}, "reason": "先看首行"}}),
            json.dumps({"final_answer": "先按首行回答，缺少引用。", "citations": []}),
            json.dumps({"final_answer": "已补齐第二行差异后回答。", "citations": [{"id": "ev-2-compare-key_pin_nets-2"}]}),
        ])

        payload = run_compare_agent(
            compare_payload,
            left,
            right,
            CompareAgentRequest(max_steps=4, max_tool_calls=5, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("已补齐第二行差异后回答。", payload["answer"])
        self.assertEqual(1, payload["model_metadata"]["quality_repair_attempt_count"])
        self.assertEqual(2, len(payload["tool_calls"]))
        self.assertEqual("get_compare_section_rows", payload["tool_calls"][1]["tool"])
        self.assertEqual(1, payload["tool_calls"][1]["args"]["offset"])
        self.assertIn("quality_repair_tool_call", [item["type"] for item in payload["agent_steps"]])
        self.assertEqual(3, len(provider.calls))

    def test_mock_compare_agent_runs_cadence_page_semantic_profile(self):
        left, right = sample_cadence_page_payloads(self)
        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(
                profile="compare_cadence_pages",
                question="请比对两个项目第1-1页的区别",
                max_steps=4,
                max_tool_calls=4,
            ),
            CompareMockModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("compare_cadence_pages", payload["profile"])
        self.assertEqual("resolve_compare_page_range", payload["tool_calls"][0]["tool"])
        self.assertEqual("compare_cadence_page_semantics", payload["tool_calls"][1]["tool"])
        self.assertTrue(any(item["type"].startswith("cadence_") for item in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])

    def test_mock_compare_agent_runs_datasheet_profile(self):
        old_dir = os.environ.get("PSTX_DATASHEET_DIR")
        old_data_dir = os.environ.get("PSTX_DATASHEET_DATA_DIR")
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.addCleanup(lambda: os.environ.pop("PSTX_DATASHEET_DIR", None) if old_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DIR", old_dir))
        self.addCleanup(lambda: os.environ.pop("PSTX_DATASHEET_DATA_DIR", None) if old_data_dir is None else os.environ.__setitem__("PSTX_DATASHEET_DATA_DIR", old_data_dir))
        root = Path(tmp.name)
        source = root / "pdfs"
        source.mkdir()
        (source / "HQ11112042009_LCMXO3LF.pdf").write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(source)
        os.environ["PSTX_DATASHEET_DATA_DIR"] = str(root / "data")
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQ11112042009 LCMXO3LF recommended operating conditions"], "fake", ""),
        ):
            reindex_datasheets(force=True)

        left, right = sample_payloads(self)
        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(
                profile="compare_datasheet_qa",
                question="请对比 HQ11112042009 的 datasheet recommended operating 参数证据",
                max_steps=3,
                max_tool_calls=3,
            ),
            CompareMockModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("compare_datasheet_qa", payload["profile"])
        self.assertEqual("batch_search_datasheet_chunks", payload["tool_calls"][0]["tool"])
        self.assertTrue(any(item["type"] == "datasheet_chunk" for item in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])

    def test_compare_agent_tool_batch_call_runs_multiple_tools_in_one_step(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_batch_call": [
                    {"name": "list_compare_sections", "args": {}, "reason": "先看分区"},
                    {"name": "query_compare_diff", "args": {"query": "U1", "limit": 4}, "reason": "再查 U1"},
                ]
            }, ensure_ascii=False),
            json.dumps({"final_answer": "compare batch done", "citations": []}, ensure_ascii=False),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(max_steps=2, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(["list_compare_sections", "query_compare_diff"], [item["tool"] for item in payload["tool_calls"]])
        self.assertEqual("tool_batch_call", payload["agent_steps"][0]["type"])
        self.assertEqual(2, len(payload["observations"]))
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])

    def test_compare_agent_can_read_harness_skill_during_run(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "get_harness_skill",
                    "args": {"skill_id": "datasheet-key-info", "max_body_chars": 1000},
                    "reason": "先读取 compare datasheet skill 的取证打法。",
                }
            }, ensure_ascii=False),
            json.dumps({"final_answer": "已读取 compare datasheet skill。", "citations": []}, ensure_ascii=False),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(
                profile="compare_datasheet_qa",
                question="请说明 A/B datasheet 对比要怎么取证。",
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

    def test_compare_agent_tool_error_can_recover_with_alternative_tool(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "unknown_compare_tool", "args": {}, "reason": "先试错"}}),
            json.dumps({"tool_call": {"name": "list_compare_sections", "args": {}, "reason": "改用分区清单"}}),
            json.dumps({"final_answer": "已从 compare 工具失败中恢复。", "citations": []}),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(max_steps=4, max_tool_calls=4, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertEqual(1, payload["model_metadata"]["tool_error_recovery_count"])
        self.assertEqual(1, payload["trace_summary"]["tool_error_recovery_count"])
        self.assertIn("tool_error_recovery", [item["type"] for item in payload["agent_steps"]])
        self.assertEqual(["unknown_compare_tool", "list_compare_sections"], [item["tool"] for item in payload["tool_calls"]])
        self.assertFalse(payload["observations"][0]["ok"])
        self.assertEqual("error", payload["observations"][0]["tool_result_contract"]["completeness"])
        self.assertIn("list_compare_sections", payload["observations"][0]["tool_result_contract"]["recommended_next_tools"])
        self.assertNotIn("unknown_compare_tool", payload["observations"][0]["tool_result_contract"]["recommended_next_tools"])
        self.assertEqual("unknown_compare_tool", provider.calls[1]["inputs"]["observations"][0]["tool"])
        self.assertIn("Unknown harness tool", provider.calls[1]["inputs"]["observations"][0]["error"])
        self.assertIn("list_compare_sections", provider.calls[1]["inputs"]["observations"][0]["tool_result_contract"]["recommended_next_tools"])

    def test_compare_duplicate_tool_call_is_reexecuted_without_error_recovery(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({"tool_call": {"name": "list_compare_sections", "args": {}, "reason": "先看分区"}}),
            json.dumps({"tool_call": {"name": "list_compare_sections", "args": {}, "reason": "重复看分区"}}),
            json.dumps({"final_answer": "已避免重复 compare 工具调用。", "citations": []}),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(question="请快速扫描当前项目对比差异", max_steps=4, max_tool_calls=4, debug=True),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("final_answer", payload["model_metadata"]["stopped_reason"])
        self.assertEqual(0, payload["model_metadata"].get("tool_error_recovery_count", 0))
        self.assertEqual(["list_compare_sections", "list_compare_sections"], [item["tool"] for item in payload["tool_calls"]])
        self.assertTrue(payload["tool_calls"][1]["duplicate"])
        self.assertTrue(payload["observations"][1]["ok"])
        self.assertIn("batch_query_compare_diff", payload["observations"][1]["tool_result_contract"]["recommended_next_tools"])

    def test_compare_batch_domain_tool_creates_citable_evidence(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "tool_call": {
                    "name": "batch_query_compare_diff",
                    "args": {"queries": ["U1", "PI-B"], "limit_per_query": 2},
                    "reason": "复合问题一次查多个差异关键词。",
                }
            }, ensure_ascii=False),
            json.dumps({
                "final_answer": "compare batch evidence done",
                "citations": [{"id": "ev-1-batch-query-1-key_pin_nets-1", "note": "U1 差异"}],
            }, ensure_ascii=False),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(profile="compare_quick_scan", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("batch_query_compare_diff", payload["tool_calls"][0]["tool"])
        self.assertTrue(any(node["id"] == "ev-1-batch-query-1-key_pin_nets-1" for node in payload["final_evidence"]))
        self.assertTrue(payload["citations"][0]["valid"])

    def test_auto_compare_profile_combines_cadence_and_bom_capabilities(self):
        left, right = sample_cadence_page_payloads(self)
        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(
                profile="auto",
                question="请比对两个项目第1-1页，同时检查芯片 Pin/Net 和 PI 变化。",
                max_steps=4,
                max_tool_calls=6,
            ),
            CompareMockModelProvider(),
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("auto", payload["profile"])
        capability_ids = [item["id"] for item in payload["capability_plan"]]
        self.assertIn("compare_cadence_pages", capability_ids)
        self.assertIn("compare_pin_net", capability_ids)
        self.assertIn("compare_bom_feishu", capability_ids)
        self.assertEqual("runtime_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("compare_cadence_page_semantics", payload["tool_calls"][0]["tool"])
        self.assertIn("batch_query_compare_diff", [item["tool"] for item in payload["tool_calls"]])
        playbook_ids = [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]]
        self.assertIn("cadence_page_semantic_compare", playbook_ids)
        self.assertIn("compare_diff_batch_lookup", playbook_ids)
        seeded = {item["name"]: item for item in payload["playbook_plan"]["seeded_tool_calls"]}
        self.assertEqual(1, seeded["compare_cadence_page_semantics"]["args"]["page_start"])
        self.assertEqual(1, seeded["compare_cadence_page_semantics"]["args"]["page_end"])
        self.assertIn("batch_query_compare_diff", payload["playbook_plan"]["recommended_first_tools"])

    def test_compare_agent_exposes_tool_contracts_to_trace_and_model(self):
        left, right = sample_payloads(self)

        class PrefetchAwareProvider:
            provider = "prefetch-aware"
            mode = "mock"

            def __init__(self):
                self.calls = []

            def generate_agent_step(self, prompt, *, inputs=None):
                self.calls.append({"prompt": prompt, "inputs": inputs or {}})
                observations = list((inputs or {}).get("observations") or [])
                evidence_ids = []
                for observation in observations:
                    evidence_ids.extend(str(item) for item in observation.get("evidence_node_ids", []) if item)
                    for node in observation.get("evidence_nodes", []) or []:
                        node_id = str(node.get("id") or "")
                        if node_id:
                            evidence_ids.append(node_id)
                return HarnessModelResponse(
                    answer=json.dumps({
                        "final_answer": "done",
                        "citations": [{"id": evidence_ids[0], "note": "引用预取的对比差异证据。"}] if evidence_ids else [],
                    }, ensure_ascii=False),
                    provider=self.provider,
                    mode=self.mode,
                )

        provider = PrefetchAwareProvider()

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(profile="auto", question="请对比 U1 和 PI-B 的差异", max_steps=1, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertIn("compare_diff_batch_lookup", [item["id"] for item in payload["playbook_plan"]["selected_playbooks"]])
        self.assertEqual("runtime_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("batch_query_compare_diff", payload["tool_calls"][0]["tool"])
        self.assertEqual("runtime_prefetch_followup", payload["agent_steps"][1]["type"])
        self.assertEqual("get_compare_row", payload["tool_calls"][1]["tool"])
        self.assertTrue(payload["tool_result_contracts"])
        self.assertIn("get_compare_row", payload["tool_result_contracts"][0]["recommended_next_tools"])
        self.assertIn("tool_result_contract", payload["observations"][0])
        self.assertIn("playbook_plan", provider.calls[0]["inputs"])
        self.assertIn("tool_result_contract", provider.calls[0]["inputs"]["observations"][0])
        self.assertEqual(2, provider.calls[0]["inputs"]["tool_count"])

    def test_auto_compare_profile_goal_prefetches_risk_summary_without_entities(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({
                "final_answer": "已先读取项目对比风险总览。",
                "citations": [{"id": "ev-1-risk-1", "note": "风险总览。"}],
            }, ensure_ascii=False)
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(profile="auto", question="请快速扫描两个项目差异", max_steps=1, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("runtime_goal_prefetch", payload["agent_steps"][0]["type"])
        self.assertEqual("summarize_compare_risks", payload["tool_calls"][0]["tool"])
        self.assertEqual(1, provider.calls[0]["inputs"]["tool_count"])
        self.assertIn("goal_prefetch_plan", payload["model_metadata"])
        self.assertTrue({item["type"] for item in payload["final_evidence"]} & {"compare_component", "compare_net", "compare_diff"})
        self.assertTrue(payload["citations"][0]["valid"])

    def test_compare_agent_retries_premature_refusal_before_final_answer(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({"final_answer": "无法回答，信息不足。"}, ensure_ascii=False),
            json.dumps({
                "tool_call": {
                    "name": "list_compare_sections",
                    "args": {},
                    "reason": "先读取对比分区再判断。",
                }
            }, ensure_ascii=False),
            json.dumps({"final_answer": "已基于对比分区继续取证。", "citations": []}, ensure_ascii=False),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(profile="compare_quick_scan", question="请分析项目差异", max_steps=2, max_tool_calls=2),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual("已基于对比分区继续取证。", payload["answer"])
        self.assertEqual("list_compare_sections", payload["tool_calls"][0]["tool"])
        self.assertEqual(1, payload["model_metadata"]["perseverance_retry_count"])
        self.assertTrue(provider.calls[1]["inputs"]["perseverance_retry"])

    def test_compare_task_ledger_guard_retries_empty_final_before_evidence(self):
        left, right = sample_payloads(self)
        provider = SequenceProvider([
            json.dumps({"final_answer": "已完成。"}, ensure_ascii=False),
            json.dumps({"final_answer": "已根据任务账本继续确认。", "citations": []}, ensure_ascii=False),
        ])

        payload = run_compare_agent(
            sample_compare_payload(),
            left,
            right,
            CompareAgentRequest(question="请快速扫描当前项目对比差异", max_steps=1, max_tool_calls=4),
            provider,
        )

        self.assertTrue(payload["ok"])
        self.assertEqual(2, len(provider.calls))
        self.assertEqual(1, payload["model_metadata"]["perseverance_retry_count"])
        self.assertIn("task_ledger", provider.calls[1]["inputs"])
        self.assertIn("task_ledger", provider.calls[1]["inputs"]["perseverance_retry_note"])

    def test_compare_agent_rejects_unknown_tool_and_invalid_citation(self):
        left, right = sample_payloads(self)
        bad_tool = SequenceProvider([
            json.dumps({"tool_call": {"name": "read_project_text", "args": {"path": "../x"}, "reason": "try"}}),
        ])
        rejected = run_compare_agent(sample_compare_payload(), left, right, CompareAgentRequest(max_steps=1), bad_tool)
        self.assertFalse(rejected["ok"])
        self.assertIn("Unknown harness tool", rejected["answer"])

        invalid_citation = SequenceProvider([
            json.dumps({"final_answer": "done", "citations": [{"id": "ev-missing"}]}),
        ])
        payload = run_compare_agent(sample_compare_payload(), left, right, CompareAgentRequest(), invalid_citation)
        self.assertTrue(payload["ok"])
        self.assertEqual(1, payload["model_metadata"]["invalid_citation_count"])
        self.assertFalse(payload["citations"][0]["valid"])

    def test_compare_profiles_are_listed(self):
        profiles = {item["id"]: item for item in list_compare_agent_profiles()}
        self.assertIn("auto", profiles)
        self.assertIn("compare_full_review", profiles)
        self.assertIn("*", profiles["compare_full_review"]["tools"])
        self.assertIn("read_compare_project_text", profiles["compare_page_mapping"]["tools"])
        self.assertIn("batch_query_compare_diff", profiles["compare_quick_scan"]["tools"])
        self.assertIn("batch_get_compare_rows", profiles["compare_pin_net"]["tools"])
        self.assertIn("compare_cadence_pages", profiles)
        self.assertIn("compare_cadence_page_semantics", profiles["compare_cadence_pages"]["tools"])
        self.assertIn("batch_get_cadence_page_objects", profiles["compare_cadence_pages"]["tools"])
        self.assertIn("compare_datasheet_qa", profiles)
        self.assertIn("list_datasheet_review_templates", profiles["compare_datasheet_qa"]["tools"])
        self.assertIn("select_harness_skills", profiles["compare_datasheet_qa"]["tools"])
        self.assertIn("batch_search_datasheet_chunks", profiles["compare_datasheet_qa"]["tools"])
        self.assertIn("search_datasheet_parameters", profiles["compare_datasheet_qa"]["tools"])
        self.assertIn("get_datasheet_chunk", profiles["compare_datasheet_qa"]["tools"])
        self.assertIn("search_datasheet_chunks", profiles["compare_bom_feishu"]["tools"])
