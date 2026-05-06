import os
from pathlib import Path
import tempfile
import time
import unittest

from pstx_agent_runtime import (
    AgentBackgroundJob,
    AgentBackgroundRunner,
    AgentCheckpointReporter,
    AgentDurableRunStore,
    AgentProtocolError,
    AgentCapabilityProfile,
    AgentFinalizationResult,
    AgentTraceStore,
    AgentTodoList,
    PROJECT_SESSION_MEMORY_VERSION,
    TASK_DISPATCH_SCHEMA_VERSION,
    SUBAGENT_SCHEMA_VERSION,
    REPORT_AGENT_PLAYBOOKS,
    COMPARE_AGENT_PLAYBOOKS,
    MemorySummary,
    ObservationBundle,
    allowed_tool_names,
    build_agent_protocol_brief,
    build_agentic_envelope,
    build_agent_session_state,
    build_capability_plan,
    build_context_budget_summary,
    build_continuation_pack,
    build_evidence_goal_contract,
    build_evidence_layers,
    build_execution_journal,
    build_final_answer_quality_gate,
    build_harness_turn_context_snapshot,
    build_journal_summary,
    build_perseverance_retry_note,
    build_playbook_plan,
    build_project_evidence_memory,
    compact_project_evidence_memory,
    compact_subagent_result,
    build_quality_repair_tool_calls,
    build_runtime_state,
    build_task_ledger,
    build_tool_error_observation,
    build_tool_result_contract,
    build_trace_envelope,
    build_effort_policy_state,
    load_harness_skills,
    load_project_guidance,
    normalize_citations,
    normalize_needs_user_input,
    normalize_proposed_actions,
    status_from_stopped_reason,
    summarize_tool_dispatch_trace,
    execute_runtime_tool_calls,
    filtered_tool_list,
    fit_items_to_json_budget,
    get_project_evidence_memory_card,
    is_low_effort_answer,
    is_recoverable_tool_error,
    json_char_count,
    merge_project_session_memory,
    merge_runtime_tool_execution,
    normalize_tool_batch,
    normalize_dispatch_tasks,
    parse_agent_model_step,
    plan_subagents,
    recommended_tools_for_recovery,
    search_project_evidence_memory,
    select_goal_prefetch_tool_calls,
    select_prefetch_followup_tool_calls,
    select_project_memory_prefetch_tool_calls,
    select_seeded_prefetch_tool_calls,
    select_harness_skills,
    summarize_subagent_results,
    build_subagent_definition,
    build_subagent_question,
    tool_call_signature,
    read_task_memory,
    write_task_memory,
    append_workspace_log,
    list_workspace_artifacts,
    write_workspace_artifact,
    write_workspace_draft,
    write_workspace_scratch_files,
    workspace_status,
)


class AgentRuntimeProtocolTests(unittest.TestCase):
    def test_agent_workspace_store_writes_status_and_artifacts(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = AgentDurableRunStore(root=tmp)
            record = store.create_run(
                scope_id="run/a b",
                kind="report",
                request={"profile": "quick_scan", "question": "看 DRC"},
                agent_run_id="report-test-1",
            )

            self.assertEqual("queued", record["status"])
            self.assertIn("run_a_b", record["workspace"]["root"])

            finished = store.finish_record("report-test-1", {
                "ok": True,
                "agent_run_id": "report-test-1",
                "status": "completed",
                "answer": "完成。",
                "runtime_state": {"task_ledger": {"items": [{"title": "复核 DRC"}]}},
                "continuation_pack": {"version": "agent-continuation-pack/v1"},
            })
            status = store.public_status("report-test-1")
            artifacts = store.list_artifacts("report-test-1")
            workspace = workspace_status("run/a b", root=tmp)

            self.assertEqual("completed", finished["status"])
            self.assertTrue(status["ok"])
            self.assertTrue(status["result_available"])
            self.assertEqual("完成。", status["agent_run"]["answer"])
            self.assertGreaterEqual(artifacts["artifact_count"], 2)
            names = {item["name"] for item in artifacts["artifacts"]}
            self.assertIn("trace.json", names)
            self.assertIn("evidence_cards.json", names)
            self.assertIn("task_ledger.md", names)
            self.assertIn("review_draft.md", names)
            self.assertTrue(workspace["task_md"]["exists"])

    def test_durable_store_tracks_child_dispatch_runs(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = AgentDurableRunStore(root=tmp)
            store.create_run(
                scope_id="run-parent",
                kind="report",
                request={"profile": "auto", "question": "拆成长任务"},
                agent_run_id="parent-run",
            )
            child = store.create_run(
                scope_id="run-parent",
                kind="report",
                request={"profile": "datasheet_qa", "question": "查 U1 datasheet"},
                agent_run_id="child-run",
                parent_agent_run_id="parent-run",
                dispatch_task={"task_id": "task-1", "title": "U1 datasheet", "question": "查 U1 datasheet"},
            )
            store.append_child_runs(
                "parent-run",
                [{
                    "task_id": "task-1",
                    "title": "U1 datasheet",
                    "profile": "datasheet_qa",
                    "question": "查 U1 datasheet",
                    "agent_run_id": child["agent_run_id"],
                    "status": "queued",
                    "status_url": "/api/harness/agent-runs/child-run",
                }],
                task_dispatch_summary={"schema_version": TASK_DISPATCH_SCHEMA_VERSION, "task_count": 1},
            )
            finished = store.finish_record("parent-run", {
                "ok": True,
                "status": "completed",
                "answer": "已拆分后台子任务。",
                "task_dispatch_summary": {"schema_version": TASK_DISPATCH_SCHEMA_VERSION, "task_count": 1},
                "dispatched_tasks": [{
                    "task_id": "task-1",
                    "title": "U1 datasheet",
                    "profile": "datasheet_qa",
                    "question": "查 U1 datasheet",
                    "agent_run_id": "child-run",
                    "status": "queued",
                    "status_url": "/api/harness/agent-runs/child-run",
                }],
            })

            parent_status = store.public_status("parent-run")
            child_status = store.public_status("child-run")
            artifacts = store.list_artifacts("parent-run")

            self.assertEqual("parent-run", child_status["parent_agent_run_id"])
            self.assertEqual(["child-run"], parent_status["child_agent_run_ids"])
            self.assertEqual(1, parent_status["task_dispatch_summary"]["task_count"])
            self.assertEqual("child-run", parent_status["dispatch_tasks"][0]["agent_run_id"])
            self.assertEqual("completed", finished["status"])
            self.assertIn("task_dispatch.json", {item["name"] for item in artifacts["artifacts"]})

    def test_agent_workspace_artifact_path_is_sandboxed(self):
        with tempfile.TemporaryDirectory() as tmp:
            artifact = write_workspace_artifact(
                "scope",
                "run",
                "../danger.json",
                {"ok": True},
                root=tmp,
                content_type="application/json",
            )

            self.assertEqual("danger.json", artifact["name"])
            self.assertTrue(Path(artifact["path"]).is_file())
            self.assertIn(Path(tmp).resolve(), Path(artifact["path"]).resolve().parents)

            draft = write_workspace_draft("scope", "run", "../../review.md", "draft", root=tmp)
            self.assertEqual("review.md", draft["name"])
            self.assertTrue(Path(draft["path"]).is_file())
            log = append_workspace_log("scope", "run", {"phase": "planning"}, root=tmp)
            self.assertEqual("run.jsonl", log["name"])
            self.assertTrue(Path(log["path"]).is_file())

    def test_agent_workspace_scratch_files_are_run_scoped_and_listed(self):
        with tempfile.TemporaryDirectory() as tmp:
            scratch = write_workspace_scratch_files(
                "scope/a",
                "run/b",
                [
                    {"filename": "../notes.md", "content": "# Notes\n临时分析", "content_type": "text/markdown"},
                    {"filename": "data.json", "content": {"ok": True}, "content_type": "application/json"},
                ],
                root=tmp,
            )
            artifacts = list_workspace_artifacts("scope/a", "run/b", root=tmp)
            workspace = workspace_status("scope/a", root=tmp)

            self.assertEqual("pstx-agent-scratch-files/v1", scratch["version"])
            self.assertEqual(2, scratch["file_count"])
            self.assertEqual("notes.md", scratch["files"][0]["name"])
            self.assertTrue(Path(scratch["files"][0]["path"]).is_file())
            self.assertTrue(scratch["files"][0]["temporary"])
            names = {item["name"] for item in artifacts["artifacts"]}
            self.assertIn("notes.md", names)
            self.assertIn("data.json", names)
            self.assertTrue(any(item.get("temporary") for item in artifacts["artifacts"]))
            self.assertTrue(workspace["scratch"]["exists"])

    def test_checkpoint_reporter_updates_durable_progress_and_log(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = AgentDurableRunStore(root=tmp)
            store.create_run(scope_id="run-cp", kind="report", request={"question": "x", "max_steps": 3}, agent_run_id="cp-run")
            reporter = AgentCheckpointReporter(store, "cp-run", scope_id="run-cp", kind="report")

            reporter.emit({
                "phase": "tool_call",
                "step_index": 2,
                "max_steps": 3,
                "max_tool_calls": 5,
                "tool_calls": [{"tool": "get_table_rows", "ok": True}],
                "agent_steps": [{"index": 1, "type": "tool_call"}],
                "evidence_ids": ["ev-1", "ev-2"],
                "task_ledger": {"next_actions": [{"title": "继续查明细"}]},
            })

            status = store.public_status("cp-run")
            artifacts = store.list_artifacts("cp-run")
            self.assertEqual("running", status["status"])
            self.assertEqual("tool_call", status["current_phase"])
            self.assertEqual(2, status["progress"]["step_index"])
            self.assertEqual(1, status["progress"]["tool_call_count"])
            self.assertEqual(2, status["progress"]["evidence_count"])
            self.assertEqual("ev-1", status["partial_trace"]["evidence_ids"][0])
            self.assertEqual("cp-run.jsonl", {item["name"] for item in artifacts["artifacts"]}.pop())

    def test_durable_running_record_becomes_incomplete_after_stale_heartbeat(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_timeout = os.environ.get("PSTX_AGENT_HEARTBEAT_TIMEOUT_SECONDS")
            os.environ["PSTX_AGENT_HEARTBEAT_TIMEOUT_SECONDS"] = "1"
            try:
                store = AgentDurableRunStore(root=tmp)
                store.create_run(scope_id="run-stale", kind="report", request={"question": "x"}, agent_run_id="stale-run")
                store.update_record(
                    "stale-run",
                    status="running",
                    current_phase="model_call",
                    heartbeat_at="2000-01-01T00:00:00",
                )
                status = store.public_status("stale-run")
                self.assertEqual("incomplete", status["status"])
                self.assertTrue(status["can_continue"])
            finally:
                if old_timeout is None:
                    os.environ.pop("PSTX_AGENT_HEARTBEAT_TIMEOUT_SECONDS", None)
                else:
                    os.environ["PSTX_AGENT_HEARTBEAT_TIMEOUT_SECONDS"] = old_timeout

    def test_background_runner_executes_and_persists_checkpoint(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = AgentDurableRunStore(root=tmp)
            store.create_run(scope_id="run-bg", kind="report", request={"question": "x"}, agent_run_id="bg-run-1")
            runner = AgentBackgroundRunner(store, worker_count=1, queue_limit=2)
            runner.submit(AgentBackgroundJob(
                agent_run_id="bg-run-1",
                scope_id="run-bg",
                kind="report",
                run=lambda agent_run_id: {
                    "ok": True,
                    "agent_run_id": agent_run_id,
                    "status": "completed",
                    "answer": "后台完成",
                    "agent_steps": [{"index": 1, "type": "final_answer"}],
                },
            ))

            for _ in range(40):
                status = store.public_status("bg-run-1")
                if status.get("status") == "completed":
                    break
                time.sleep(0.05)

            status = store.public_status("bg-run-1")
            self.assertEqual("completed", status["status"])
            self.assertEqual("后台完成", status["agent_run"]["answer"])
            self.assertGreaterEqual(status["artifact_count"], 2)

    def test_background_runner_can_cancel_queued_record(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = AgentDurableRunStore(root=tmp)
            store.create_run(scope_id="run-cancel", kind="report", request={"question": "x"}, agent_run_id="cancel-run")
            runner = AgentBackgroundRunner(store, worker_count=1, queue_limit=2)

            cancelled = runner.cancel("cancel-run")

            self.assertEqual("cancelled", cancelled["status"])
            self.assertTrue(cancelled["cancel_requested"])

    def test_todo_list_and_memory_summary_are_serializable(self):
        todos = AgentTodoList.from_titles("检查 DFMEA 输入", ["识别芯片", "查规格书"])
        todos = todos.mark("todo-1", "completed", evidence_ids=["ev-u1"], note="U1 已有身份卡")
        memory = MemorySummary.from_parts(
            goal=todos.goal,
            facts=["U1 是 FPGA"],
            decisions=["先做准备度，不生成正式 DFMEA"],
            open_questions=["PU2 缺规格型号"],
            evidence_ids=["ev-u1"],
        )

        self.assertEqual(2, len(todos.to_dict()["items"]))
        self.assertEqual(1, todos.to_dict()["open_count"])
        self.assertEqual(["ev-u1"], memory.to_dict()["evidence_ids"])

    def test_tool_batch_normalizes_single_and_batch_calls(self):
        single = normalize_tool_batch(
            {"tool_call": {"name": "get_table_rows", "args": {"table_id": "bom"}, "reason": "查 BOM"}},
            allowed_tools={"get_table_rows"},
        )
        batch = normalize_tool_batch(
            [
                {"name": "search_component_identity_cards", "args": {"query": "U1"}},
                {"name": "search_feishu_cache_rows", "args": {"query": "HQ100"}},
            ],
            allowed_tools={"search_component_identity_cards", "search_feishu_cache_rows"},
            max_calls=2,
        )

        self.assertEqual("get_table_rows", single[0]["name"])
        self.assertEqual(["search_component_identity_cards", "search_feishu_cache_rows"], [item["name"] for item in batch])

    def test_tool_batch_rejects_unknown_tool_and_bad_args(self):
        with self.assertRaises(AgentProtocolError):
            normalize_tool_batch({"tool_call": {"name": "write_file", "args": {}}}, allowed_tools={"get_table_rows"})
        with self.assertRaises(AgentProtocolError):
            normalize_tool_batch({"tool_call": {"name": "get_table_rows", "args": []}}, allowed_tools={"get_table_rows"})

    def test_runtime_tool_execution_merge_updates_state_and_metadata(self):
        state = {
            "tool_calls": [],
            "tool_signatures": [],
            "tool_dispatch_trace": [],
            "tool_result_contracts": [],
            "observations_for_model": [],
            "public_observations": [],
            "raw_observations": [],
            "evidence_nodes": [],
            "metadata": {},
        }
        execution = {
            "ok": True,
            "tool_calls": [{"tool": "get_table_rows"}],
            "tool_signatures": ["get_table_rows::{}"],
            "tool_dispatch_trace": [{"tool": "get_table_rows", "status": "completed"}],
            "tool_result_contracts": [{"tool": "get_table_rows", "completeness": "complete"}],
            "observations_for_model": [{"tool": "get_table_rows"}],
            "public_observations": [{"tool": "get_table_rows", "result": {}}],
            "raw_observations": [{"tool": "get_table_rows", "raw_result": {}}],
            "evidence_nodes": [{"id": "ev-row", "type": "table_row"}],
        }

        counts = merge_runtime_tool_execution(
            execution=execution,
            tool_calls=state["tool_calls"],
            tool_signatures=state["tool_signatures"],
            tool_dispatch_trace=state["tool_dispatch_trace"],
            tool_result_contracts=state["tool_result_contracts"],
            observations_for_model=state["observations_for_model"],
            public_observations=state["public_observations"],
            raw_observations=state["raw_observations"],
            evidence_nodes=state["evidence_nodes"],
            metadata=state["metadata"],
            metadata_prefix="prefetch",
        )

        self.assertEqual(1, counts["tool_count"])
        self.assertEqual("get_table_rows", state["tool_calls"][0]["tool"])
        self.assertEqual("ev-row", state["evidence_nodes"][0]["id"])
        self.assertTrue(state["metadata"]["prefetch_ok"])
        self.assertEqual(1, state["metadata"]["prefetch_tool_count"])
        self.assertEqual(1, state["metadata"]["prefetch_observation_count"])

        merge_runtime_tool_execution(
            execution={"ok": False, "error": "boom", "tool_dispatch_trace": [{"status": "failed"}]},
            tool_calls=state["tool_calls"],
            tool_signatures=state["tool_signatures"],
            tool_dispatch_trace=state["tool_dispatch_trace"],
            tool_result_contracts=state["tool_result_contracts"],
            observations_for_model=state["observations_for_model"],
            public_observations=state["public_observations"],
            raw_observations=state["raw_observations"],
            evidence_nodes=state["evidence_nodes"],
            metadata=state["metadata"],
            metadata_prefix="repair",
        )
        self.assertFalse(state["metadata"]["repair_ok"])
        self.assertEqual("boom", state["metadata"]["repair_error"])
        self.assertEqual("failed", state["tool_dispatch_trace"][-1]["status"])

    def test_observation_bundle_keeps_recent_items_and_evidence_ids(self):
        observations = [
            {"tool": "t", "summary": f"obs-{index}", "evidence_node_ids": [f"ev-{index}"]}
            for index in range(10)
        ]
        bundle = ObservationBundle.from_observations(observations, max_items=3, max_chars=2000).to_dict()

        self.assertTrue(bundle["truncated"])
        self.assertEqual(7, bundle["omitted_count"])
        self.assertEqual(["ev-0", "ev-1", "ev-2"], bundle["evidence_ids"][:3])
        self.assertEqual(3, len(bundle["observations"]))

    def test_build_evidence_layers_keeps_summary_cards_and_raw_trace_hint(self):
        layers = build_evidence_layers(
            tool_name="get_table_rows",
            result={"id": "rows", "summary": "读取 2 行", "rows": [{"位号": "U1", "页码": "12"}]},
            observation={"id": "rows", "title": "表格行", "summary": "读取 2 行"},
            evidence_nodes=[{
                "id": "ev-u1",
                "type": "table_row",
                "title": "U1 行",
                "summary": "位号=U1",
                "source": {"tool": "get_table_rows", "tool_call_index": 1},
                "locator": {"table_id": "bom", "row_number": 2, "refdes": "U1", "user_visible_page": "12"},
                "detail_tool": {"name": "get_table_rows", "args": {"table_id": "bom", "offset": 1, "limit": 1}},
            }],
            tool_result_contract={
                "completeness": "partial",
                "recommended_next_tools": ["get_table_rows"],
            },
            include_raw_preview=False,
        )

        self.assertEqual("three-layer-evidence/v1", layers["version"])
        self.assertEqual("partial", layers["summary_layer"]["completeness"])
        self.assertEqual("ev-u1", layers["evidence_card_layer"][0]["id"])
        self.assertEqual("U1", layers["evidence_card_layer"][0]["refdes"])
        self.assertEqual("12", layers["evidence_card_layer"][0]["page"])
        self.assertTrue(layers["raw_layer"]["stored_in_trace"])
        self.assertTrue(layers["raw_layer"]["preview_omitted_for_model"])

    def test_project_evidence_memory_indexes_and_searches_cards(self):
        result = {
            "agent_run_id": "agent-ev-1",
            "profile": "dfmea_prep",
            "status": "completed",
            "finished_at": "2026-04-28T09:00:00",
            "citations": [{"id": "ev-u1", "valid": True}],
            "final_evidence": [{
                "id": "ev-u1",
                "type": "component_identity",
                "title": "U1 身份卡",
                "summary": "U1 是 FPGA，HQ=HQ100。",
                "locator": {"refdes": "U1", "user_visible_page": "12"},
                "detail_tool": {"name": "get_component_identity_card", "args": {"refdes": "U1"}},
            }],
        }

        cards = build_project_evidence_memory({}, result)
        search = search_project_evidence_memory(cards, query="HQ100")
        detail = get_project_evidence_memory_card(cards, "ev-u1")

        self.assertEqual(1, len(cards))
        self.assertTrue(cards[0]["cited"])
        self.assertEqual(1, search["total_matches"])
        self.assertTrue(detail["found"])
        self.assertEqual("get_component_identity_card", detail["card"]["detail_tool"]["name"])
        self.assertEqual({"refdes": "U1"}, detail["card"]["detail_tool"]["args"])

    def test_project_evidence_memory_keeps_detail_tool_args_structured(self):
        cards = build_project_evidence_memory({}, {
            "agent_run_id": "agent-ds",
            "final_evidence": [{
                "id": "ev-ds",
                "type": "datasheet_chunk",
                "title": "HQ100 datasheet chunk",
                "summary": "Recommended operating voltage is 3.3V.",
                "source": {"tool": "search_datasheet_chunks"},
                "detail_tool": {
                    "name": "get_datasheet_chunk",
                    "args": {"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000},
                },
            }],
        })
        compact = compact_project_evidence_memory(cards)

        self.assertIsInstance(cards[0]["detail_tool"]["args"], dict)
        self.assertEqual("p1-c1", cards[0]["detail_tool"]["args"]["chunk_id"])
        self.assertIsInstance(compact[0]["detail_tool"]["args"], dict)
        self.assertEqual("p1-c1", compact[0]["detail_tool"]["args"]["chunk_id"])

    def test_project_evidence_memory_preserves_same_local_id_across_runs(self):
        first = build_project_evidence_memory({}, {
            "agent_run_id": "agent-run-a",
            "final_evidence": [{"id": "ev-1", "type": "component", "title": "U1", "summary": "first"}],
        })
        second = build_project_evidence_memory({"evidence_memory_cards": first}, {
            "agent_run_id": "agent-run-b",
            "final_evidence": [{"id": "ev-1", "type": "component", "title": "U2", "summary": "second"}],
        })

        self.assertEqual(2, len(second))
        self.assertEqual(["agent-run-a:ev-1", "agent-run-b:ev-1"], [item["memory_id"] for item in second])
        self.assertEqual("second", get_project_evidence_memory_card(second, "ev-1")["card"]["summary"])
        self.assertEqual("first", get_project_evidence_memory_card(second, "agent-run-a:ev-1")["card"]["summary"])

    def test_project_memory_prefetch_reads_explicit_evidence_id(self):
        context = {
            "evidence_memory_cards": [{
                "id": "ev-u1",
                "type": "component_identity",
                "title": "U1 身份卡",
                "summary": "U1 是 FPGA。",
            }],
        }

        plan = select_project_memory_prefetch_tool_calls(
            "请继续看 ev-u1 这个证据",
            context,
            allowed_tools={"list_project_memory_evidence", "get_project_memory_evidence"},
            remaining_tool_calls=2,
        )

        self.assertEqual("agent-memory-prefetch-plan/v1", plan["version"])
        self.assertEqual(1, plan["selected_count"])
        self.assertEqual("get_project_memory_evidence", plan["tool_calls"][0]["name"])
        self.assertEqual({"evidence_id": "ev-u1"}, plan["tool_calls"][0]["args"])

    def test_project_memory_prefetch_searches_when_user_continues(self):
        context = {
            "evidence_memory_cards": [{
                "id": "ev-u1",
                "type": "component_identity",
                "title": "U1 身份卡",
                "summary": "U1 是 FPGA，HQ=HQ100。",
                "locator": {"refdes": "U1"},
            }],
        }

        plan = select_project_memory_prefetch_tool_calls(
            "继续刚才 U1 的分析",
            context,
            allowed_tools={"list_project_memory_evidence", "get_project_memory_evidence"},
            remaining_tool_calls=2,
        )

        self.assertEqual(1, plan["selected_count"])
        self.assertEqual("list_project_memory_evidence", plan["tool_calls"][0]["name"])
        self.assertEqual("U1", plan["tool_calls"][0]["args"]["query"])

    def test_protocol_brief_documents_runtime_boundaries(self):
        brief = build_agent_protocol_brief()

        self.assertIn("pstx-agent-runtime/v1", brief)
        self.assertIn("TodoList", brief)
        self.assertIn("ObservationBundle", brief)
        self.assertIn("当前每轮只允许一个 tool_call", brief)

    def test_parse_agent_model_step_supports_batch_tool_calls(self):
        step = parse_agent_model_step(
            """```json
            {"tool_batch_call":[
              {"name":"get_table_rows","args":{"table_id":"bom"},"reason":"查 BOM"},
              {"name":"query_report_entity","args":{"query":"U1"},"reason":"查 U1"}
            ]}
            ```""",
            allowed_tools={"get_table_rows", "query_report_entity"},
            max_batch_calls=2,
        )

        self.assertIsNotNone(step)
        self.assertEqual("tool_batch_call", step.type)
        self.assertEqual(["get_table_rows", "query_report_entity"], [item["name"] for item in step.tool_calls])
        self.assertEqual("tool_batch_call", step.to_legacy_dict()["type"])

    def test_parse_agent_model_step_rejects_unsafe_batch_tool(self):
        with self.assertRaises(AgentProtocolError):
            parse_agent_model_step(
                '{"tool_batch_call":[{"name":"write_file","args":{}}]}',
                allowed_tools={"get_table_rows"},
            )

    def test_parse_agent_model_step_keeps_single_tool_and_final_answer_compatibility(self):
        tool_step = parse_agent_model_step(
            '{"tool_call":{"name":"get_table_rows","args":{"table_id":"bom"}}}',
            allowed_tools={"get_table_rows"},
        )
        final_step = parse_agent_model_step('{"final_answer":"完成","citations":[]}')

        self.assertEqual("tool_call", tool_step.type)
        self.assertEqual("get_table_rows", tool_step.to_legacy_dict()["tool_call"]["name"])
        self.assertEqual("final_answer", final_step.type)
        self.assertEqual("完成", final_step.final_answer)

    def test_parse_agent_model_step_supports_dispatch_tasks(self):
        step = parse_agent_model_step(
            """{"dispatch_tasks":[
              {"task_id":"ds-u1","title":"U1 规格书","profile":"datasheet_qa","question":"读取 U1 datasheet 的供电和接口约束","max_steps":8},
              {"task_id":"cad-114","title":"第 114 页 Cadence","profile":"cadence_pages","question":"复核第 114 页连接语义","depends_on":["ds-u1"]}
            ],"reason":"两个分支可独立后台处理。"}""",
            max_dispatch_tasks=2,
        )

        self.assertIsNotNone(step)
        self.assertEqual("dispatch_tasks", step.type)
        self.assertEqual("pstx-agent-task-dispatch.v1", step.task_dispatch["schema_version"])
        self.assertEqual(["ds-u1", "cad-114"], [item["task_id"] for item in step.dispatch_tasks])
        legacy = step.to_legacy_dict()
        self.assertEqual("dispatch_tasks", legacy["type"])
        self.assertEqual(2, legacy["task_dispatch"]["task_count"])

    def test_normalize_dispatch_tasks_rejects_empty_or_too_many(self):
        with self.assertRaises(AgentProtocolError):
            normalize_dispatch_tasks({"dispatch_tasks": []})
        with self.assertRaises(AgentProtocolError):
            normalize_dispatch_tasks(
                {"dispatch_tasks": [
                    {"task_id": "a", "question": "a"},
                    {"task_id": "b", "question": "b"},
                ]},
                max_tasks=1,
            )

    def test_subagent_runtime_plans_definitions_and_compacts_results(self):
        profiles = {
            "bom": {
                "title": "BOM",
                "description": "BOM focused review",
                "default_question": "检查 BOM。",
                "tools": ["get_table_rows"],
                "max_steps": 6,
                "max_tool_calls": 10,
            },
            "full": {
                "title": "Full",
                "default_question": "完整审查。",
                "tools": ["*"],
                "max_steps": 12,
                "max_tool_calls": 24,
            },
        }

        definition = build_subagent_definition("bom", profiles["bom"], allowed_tools=["get_table_rows"])
        plan = plan_subagents(
            ["bom", "bom", "missing", "full"],
            profiles,
            max_subagents=2,
            disallowed_profiles=["full"],
            parent_profile="quick_scan",
        )
        question = build_subagent_question("父任务", definition)
        compact = compact_subagent_result({
            "ok": True,
            "profile": "bom",
            "agent_run_id": "child-1",
            "answer": "完成",
            "citations": [{"id": "ev-1"}],
            "final_evidence": [{"id": "ev-1"}],
            "proposed_actions": [{"title": "复核"}],
        }, definition=definition)
        summary = summarize_subagent_results([compact], plan=plan, max_workers=1, elapsed_ms=12, provider_parallel_safe=False)

        self.assertEqual(SUBAGENT_SCHEMA_VERSION, definition["schema_version"])
        self.assertEqual(["bom"], plan["profiles"])
        self.assertEqual(["duplicate", "unknown_profile", "disallowed"], [item["reason"] for item in plan["skipped"]])
        self.assertIn("父任务", question)
        self.assertEqual("child-1", compact["agent_run_id"])
        self.assertEqual(1, compact["citation_count"])
        self.assertFalse(summary["provider_parallel_safe"])
        self.assertEqual(1, summary["total_evidence_node_count"])

    def test_build_runtime_state_compacts_todos_memory_and_evidence(self):
        state = build_runtime_state(
            goal="做 DFMEA 准备",
            capability_plan=[
                {"id": "dfmea_prep", "title": "DFMEA 准备", "description": "身份卡和缺失项"},
                {"id": "feishu_bom_qa", "title": "飞书缓存问答"},
            ],
            observations=[
                {
                    "tool": "summarize_dfmea_readiness",
                    "summary": "U1 可分析，PU2 缺规格。",
                    "evidence_node_ids": ["ev-u1", "ev-pu2"],
                }
            ],
            project_context={
                "answers": [{"question_id": "q-spec", "answer": "PU2 是电源管理芯片"}],
                "pending_questions": [{"question": "请补充 U3 规格型号", "missing_fields": ["spec"]}],
            },
            truncated=True,
        )

        self.assertEqual("pstx-agent-runtime/v1", state["protocol_version"])
        self.assertGreaterEqual(len(state["todo_list"]["items"]), 2)
        self.assertEqual("completed", state["todo_list"]["items"][0]["status"])
        self.assertTrue(any(item["status"] == "blocked" for item in state["todo_list"]["items"]))
        self.assertIn("ev-u1", state["memory_summary"]["evidence_ids"])
        self.assertTrue(state["memory_summary"]["open_questions"])
        self.assertEqual("agent-task-ledger/v1", state["task_ledger"]["version"])
        self.assertTrue(state["task_ledger"]["items"])
        self.assertGreaterEqual(state["task_ledger"]["progress"]["open"], 1)
        self.assertTrue(state["truncated"])

    def test_project_session_memory_merges_result_and_feeds_runtime_state(self):
        result = {
            "agent_run_id": "agent-1",
            "status": "waiting_for_user",
            "answer": "PU2 缺少规格书证据。",
            "request": {"question": "做 DFMEA 准备"},
            "runtime_state": {
                "memory_summary": {
                    "goal": "做 DFMEA 准备",
                    "facts": ["U1 已有身份卡"],
                    "decisions": ["只输出准备度"],
                    "open_questions": ["PU2 缺规格"],
                    "evidence_ids": ["ev-u1"],
                },
                "task_ledger": {
                    "items": [
                        {"id": "i-1", "title": "确认 PU2 规格", "status": "in_progress", "note": "需要用户补充"}
                    ],
                    "next_actions": [
                        {"type": "tool_call", "title": "读取 PU2 身份卡", "tool": "get_component_identity_card", "args": {"refdes": "PU2"}, "reason": "补齐证据"}
                    ],
                    "evidence_ids": ["ev-ledger"],
                },
            },
            "continuation_pack": {
                "goal": "继续 DFMEA 准备",
                "continuation_brief": "继续处理 PU2 缺口。",
                "evidence_ids": ["ev-pack"],
                "pending_questions": [{"question": "请补充 PU2 规格"}],
                "open_ledger_items": [{"title": "补 PU2 规格"}],
                "suggested_tool_calls": [{"name": "get_component_identity_card", "args": {"refdes": "PU2"}}],
            },
            "needs_user_input": {
                "questions": [{"question": "请补充 PU2 规格型号", "missing_fields": ["spec"]}]
            },
            "final_answer_quality_gate": {"status": "warn", "score": 72},
            "final_evidence": [{"id": "ev-final"}],
            "model_metadata": {"stopped_reason": "needs_user_input"},
        }
        memory = merge_project_session_memory({}, result)

        self.assertEqual(PROJECT_SESSION_MEMORY_VERSION, memory["version"])
        self.assertIn("ev-u1", memory["evidence_ids"])
        self.assertIn("ev-pack", memory["evidence_ids"])
        self.assertTrue(memory["open_questions"])
        self.assertTrue(memory["next_actions"])
        self.assertTrue(any('"refdes": "PU2"' in item for item in memory["next_actions"]))

        state = build_runtime_state(
            goal="继续 DFMEA 准备",
            project_context={"session_memory_summary": memory},
        )
        self.assertIn("U1 已有身份卡", state["memory_summary"]["facts"])
        self.assertIn("PU2 缺规格", state["memory_summary"]["open_questions"])
        self.assertIn("ev-pack", state["memory_summary"]["evidence_ids"])

    def test_task_ledger_uses_project_session_memory_open_items(self):
        ledger = build_task_ledger(
            goal="继续上一轮审查",
            capability_plan=[{"id": "quick_scan", "title": "快速扫描"}],
            project_context={
                "session_memory_summary": {
                    "open_items": ["确认 PU2 规格书证据", "复核 U46 split symbol 料号"],
                    "next_actions": ["读取 PU2 身份卡", "搜索 ref_checklist 中类似问题"],
                    "evidence_ids": ["ev-memory-pu2"],
                }
            },
        )

        memory_items = [item for item in ledger["items"] if item.get("source") == "session_memory"]
        memory_actions = [item for item in ledger["next_actions"] if item.get("source") == "session_memory"]
        self.assertEqual(2, len(memory_items))
        self.assertTrue(all(item["status"] == "pending" for item in memory_items))
        self.assertIn("ev-memory-pu2", memory_items[0]["evidence_ids"])
        self.assertEqual(2, len(memory_actions))
        self.assertEqual("review_memory_next_action", memory_actions[0]["type"])

    def test_task_ledger_turns_truncated_contracts_into_next_actions(self):
        ledger = build_task_ledger(
            goal="统计 page_rows 有多少页码",
            capability_plan=[{"id": "page_mapping", "title": "页码映射"}],
            playbook_plan={
                "selected_playbooks": [{"id": "table_column_aggregation", "title": "表格统计聚合"}],
                "recommended_first_tools": ["summarize_table_column_values", "get_table_rows"],
            },
            observations=[{
                "tool": "get_table_rows",
                "summary": "preview 1/254",
                "evidence_node_ids": ["ev-page-preview"],
            }],
            tool_result_contracts=[{
                "completeness": "truncated",
                "recommended_next_tools": ["summarize_table_column_values"],
                "aggregation_tool": {
                    "name": "summarize_table_column_values",
                    "args": {"table_id": "page_rows", "column": "页码"},
                },
                "detail_tool": {
                    "name": "get_table_rows",
                    "args": {"table_id": "page_rows", "offset": 1, "limit": 50},
                },
                "scope_summary": "table_id=page_rows；total_rows=254",
            }],
        )

        self.assertEqual("agent-task-ledger/v1", ledger["version"])
        self.assertIn("ev-page-preview", ledger["evidence_ids"])
        self.assertTrue(any(item["source"] == "tool_result_contract" for item in ledger["items"]))
        self.assertEqual("summarize_table_column_values", ledger["next_actions"][0]["tool"])
        self.assertGreaterEqual(ledger["progress"]["in_progress"], 1)

    def test_task_ledger_tracks_pending_questions_as_blocked(self):
        ledger = build_task_ledger(
            goal="做 DFMEA 准备",
            capability_plan=[{"id": "dfmea_prep", "title": "DFMEA 准备"}],
            project_context={
                "pending_questions": [{
                    "question_id": "q-spec",
                    "question": "请补充 U1 规格型号",
                    "missing_fields": ["spec"],
                    "related_evidence_ids": ["ev-u1"],
                }]
            },
        )

        self.assertEqual(1, ledger["progress"]["blocked"])
        self.assertEqual("ask_user", ledger["next_actions"][0]["type"])
        self.assertIn("ev-u1", ledger["items"][-1]["evidence_ids"])

    def test_execute_runtime_tool_calls_runs_batch_and_builds_observations(self):
        class Registry:
            def get(self, name):
                if name not in {"a", "b"}:
                    raise ValueError(f"Unknown: {name}")
                return name

            def run(self, name, context, args=None):
                return {"id": name, "title": name.upper(), "summary": f"{name} ok", "value": args or {}}

        result = execute_runtime_tool_calls(
            tool_call_items=[
                {"name": "a", "args": {"x": 1}, "reason": "first"},
                {"name": "b", "args": {}, "reason": "second"},
            ],
            is_batch_call=True,
            registry=Registry(),
            context={},
            allowed_tools={"a", "b"},
            existing_tool_call_count=0,
            max_tool_calls=4,
            debug=False,
            profile_label="test",
            make_evidence_nodes=lambda name, _result, index, _args: [{"id": f"ev-{index}-{name}"}],
            summarize_observation=lambda name, result: {"tool": name, "summary": result["summary"]},
            make_model_observation=lambda name, result, nodes, observation: {**observation, "result": result, "evidence_nodes": nodes},
        )

        self.assertTrue(result["ok"])
        self.assertEqual("tool_batch_call", result["step_type"])
        self.assertEqual(["a", "b"], [item["tool"] for item in result["tool_calls"]])
        self.assertEqual(["ev-1-a", "ev-2-b"], [node["id"] for node in result["evidence_nodes"]])
        self.assertEqual(2, len(result["observations_for_model"]))
        self.assertEqual(2, len(result["raw_observations"]))
        self.assertIn("evidence_layers", result["observations_for_model"][0])
        self.assertIn("raw_result", result["raw_observations"][0])
        self.assertEqual(2, len(result["tool_result_contracts"]))
        self.assertEqual("complete", result["tool_result_contracts"][0]["completeness"])
        self.assertIn("tool_result_contract", result["tool_calls"][0])
        self.assertEqual(2, len(result["tool_dispatch_trace"]))
        self.assertEqual("pstx-tool-dispatch-trace.v1", result["tool_dispatch_trace"][0]["schema_version"])
        self.assertEqual("completed", result["tool_dispatch_trace"][0]["status"])
        self.assertEqual("tool-call-1", result["tool_dispatch_trace"][0]["call_id"])
        self.assertNotIn("args", result["tool_dispatch_trace"][0])
        self.assertEqual(["x"], result["tool_dispatch_trace"][0]["arg_keys"])
        self.assertTrue(result["tool_dispatch_trace"][0]["tool_boundary"]["readonly"])
        self.assertEqual("none", result["tool_dispatch_trace"][0]["tool_boundary"]["approval_scope"])
        self.assertEqual("passed", result["tool_dispatch_trace"][0]["preflight_status"])
        self.assertIn("duration_ms", result["tool_dispatch_trace"][0])
        summary = summarize_tool_dispatch_trace(result["tool_dispatch_trace"])
        self.assertEqual("pstx-tool-dispatch-summary.v1", summary["schema_version"])
        self.assertEqual(2, summary["preflight_status_counts"]["passed"])
        self.assertGreaterEqual(summary["duration_ms_total"], 0)
        self.assertIn(summary["slowest_tool"], {"a", "b"})
        self.assertTrue(summary["slowest_call_id"])

    def test_execute_runtime_tool_calls_rejects_non_readonly_tool_metadata(self):
        class Tool:
            readonly = False
            mutating = False
            file_access = False
            approval_scope = "none"
            evidence_kind = "unsafe"

        class Registry:
            def get(self, name):
                if name != "write_file":
                    raise ValueError(f"Unknown: {name}")
                return Tool()

            def run(self, name, context, args=None):
                raise AssertionError("non-readonly tool should be rejected before handler runs")

        result = execute_runtime_tool_calls(
            tool_call_items=[{"name": "write_file", "args": {"path": "x"}}],
            is_batch_call=False,
            registry=Registry(),
            context={},
            allowed_tools={"write_file"},
            existing_tool_call_count=0,
            max_tool_calls=2,
            debug=True,
            profile_label="test",
        )

        self.assertFalse(result["ok"])
        self.assertEqual("tool_error", result["stopped_reason"])
        self.assertIn("不是只读工具", result["error"])
        self.assertEqual("failed", result["tool_dispatch_trace"][0]["status"])
        self.assertEqual("failed", result["tool_dispatch_trace"][0]["preflight_status"])
        self.assertFalse(result["tool_dispatch_trace"][0]["tool_boundary"]["readonly"])

    def test_execute_runtime_tool_calls_rejects_limit_and_unknown_tool(self):
        class Registry:
            def get(self, name):
                if name != "a":
                    raise ValueError(f"Unknown: {name}")
                return name

            def run(self, name, context, args=None):
                return {"id": name}

        limited = execute_runtime_tool_calls(
            tool_call_items=[{"name": "a"}, {"name": "a"}],
            is_batch_call=True,
            registry=Registry(),
            context={},
            allowed_tools={"a"},
            existing_tool_call_count=1,
            max_tool_calls=2,
            debug=True,
            profile_label="test",
        )
        rejected = execute_runtime_tool_calls(
            tool_call_items=[{"name": "secret_tool"}],
            is_batch_call=False,
            registry=Registry(),
            context={},
            allowed_tools={"a"},
            existing_tool_call_count=0,
            max_tool_calls=2,
            debug=True,
            profile_label="test",
        )

        self.assertFalse(limited["ok"])
        self.assertEqual("max_tool_calls", limited["stopped_reason"])
        self.assertEqual("limit", limited["tool_dispatch_trace"][0]["status"])
        self.assertFalse(rejected["ok"])
        self.assertEqual("tool_error", rejected["stopped_reason"])
        self.assertTrue(rejected["tool_calls"][0]["error"])
        self.assertEqual("blocked", rejected["tool_dispatch_trace"][0]["status"])

    def test_execute_runtime_tool_calls_allows_duplicate_signature_without_stopping(self):
        class Registry:
            def __init__(self):
                self.run_count = 0

            def get(self, name):
                if name != "a":
                    raise ValueError(f"Unknown: {name}")
                return name

            def run(self, name, context, args=None):
                self.run_count += 1
                return {"id": name, "summary": "ok"}

        registry = Registry()
        duplicated = execute_runtime_tool_calls(
            tool_call_items=[{"name": "a", "args": {"x": 1}, "reason": "repeat"}],
            is_batch_call=False,
            registry=registry,
            context={},
            allowed_tools={"a"},
            existing_tool_call_count=1,
            max_tool_calls=4,
            debug=True,
            profile_label="test",
            previous_tool_signatures=[tool_call_signature("a", {"x": 1})],
        )
        different_args = execute_runtime_tool_calls(
            tool_call_items=[{"name": "a", "args": {"x": 2}, "reason": "new args"}],
            is_batch_call=False,
            registry=registry,
            context={},
            allowed_tools={"a"},
            existing_tool_call_count=1,
            max_tool_calls=4,
            debug=True,
            profile_label="test",
            previous_tool_signatures=[tool_call_signature("a", {"x": 1})],
        )

        self.assertTrue(duplicated["ok"])
        self.assertEqual("", duplicated["stopped_reason"])
        self.assertTrue(duplicated["tool_calls"][0]["duplicate"])
        self.assertTrue(duplicated["tool_calls"][0]["ok"])
        self.assertEqual("ok", duplicated["observations_for_model"][0]["summary"])
        self.assertEqual("completed", duplicated["tool_dispatch_trace"][0]["status"])
        self.assertTrue(duplicated["tool_dispatch_trace"][0]["duplicate"])
        self.assertTrue(different_args["ok"])
        self.assertEqual(2, registry.run_count)

    def test_tool_error_observation_contract_and_recoverability(self):
        recovery = build_tool_error_observation(
            execution={
                "tool_name": "bad_tool",
                "args": {"x": 1},
                "error": "Unknown harness tool: bad_tool",
            },
            call_index=1,
            debug=True,
            recommended_next_tools=["list_report_tables"],
            summarize_observation=lambda name, result: {"tool": name, "summary": result["summary"]},
        )

        self.assertTrue(is_recoverable_tool_error("Unknown harness tool: bad_tool"))
        self.assertFalse(is_recoverable_tool_error("项目根目录之外：../secret.txt"))
        self.assertEqual("bad_tool", recovery["model_observation"]["tool"])
        self.assertFalse(recovery["model_observation"]["ok"])
        self.assertEqual("error", recovery["contract"]["completeness"])
        self.assertEqual(["list_report_tables"], recovery["contract"]["recommended_next_tools"])
        self.assertEqual({"x": 1}, recovery["public_observation"]["result"]["args"])
        self.assertIn("evidence_layers", recovery["raw_observation"])

    def test_error_contract_guides_task_ledger_next_action(self):
        self.assertEqual(
            ["list_report_tables"],
            recommended_tools_for_recovery(
                {"recommended_first_tools": ["bad_tool", "list_report_tables", "read_project_text"]},
                allowed_tools={"bad_tool", "list_report_tables"},
                failed_tool="bad_tool",
            ),
        )
        ledger = build_task_ledger(
            goal="恢复失败工具调用",
            capability_plan=[{"id": "quick_scan", "title": "快速扫描"}],
            tool_result_contracts=[{
                "tool": "bad_tool",
                "completeness": "error",
                "recommended_next_tools": ["list_report_tables"],
                "scope_summary": "工具 bad_tool 调用失败。",
            }],
        )

        self.assertTrue(any(item["source"] == "tool_error_contract" for item in ledger["items"]))
        self.assertEqual("list_report_tables", ledger["next_actions"][0]["tool"])
        self.assertIn("替代工具", ledger["next_actions"][0]["title"])

    def test_trace_envelope_compacts_agent_run_payload(self):
        envelope = build_trace_envelope({
            "agent_run_id": "run-1",
            "ok": True,
            "mode": "local-agent-harness",
            "profile": "dfmea_prep",
            "answer": "完成" * 700,
            "trace_summary": {"step_count": 2},
            "runtime_state": {"protocol_version": "pstx-agent-runtime/v1"},
            "context_budget": {"truncated": True},
            "turn_context_snapshot": {"schema_version": "pstx-harness-turn-context.v1", "profile": "dfmea_prep"},
            "tool_dispatch_trace": [{"schema_version": "pstx-tool-dispatch-trace.v1", "status": "completed", "tool": "a", "event_index": 1}],
            "tool_calls": [{"tool": "a"}, {"tool": "b"}],
            "observations": [{"tool": "a"}],
            "agent_steps": [{"type": "tool_call"}],
            "final_evidence": [{"id": "ev-1"}],
            "citations": [{"id": "ev-1"}],
            "final_answer_quality_gate": {"status": "warn", "score": 70, "repair_action_count": 1},
            "request": {"profile": "dfmea_prep"},
        })

        self.assertEqual("run-1", envelope["agent_run_id"])
        self.assertEqual("dfmea_prep", envelope["profile"])
        self.assertEqual(2, envelope["tool_call_count"])
        self.assertEqual(1, envelope["observation_count"])
        self.assertEqual(1, envelope["evidence_node_count"])
        self.assertTrue(envelope["answer_preview"].endswith("…"))
        self.assertEqual("run-1", envelope["payload"]["agent_run_id"])
        self.assertEqual("pstx-harness-turn-context.v1", envelope["turn_context_snapshot"]["schema_version"])
        self.assertEqual("pstx-tool-dispatch-summary.v1", envelope["tool_dispatch_summary"]["schema_version"])
        self.assertEqual(1, envelope["tool_dispatch_summary"]["completed_count"])
        self.assertTrue(envelope["execution_journal"])
        self.assertEqual("agent-run-journal/v1", envelope["journal_summary"]["version"])
        self.assertGreaterEqual(envelope["journal_summary"]["tool_event_count"], 1)
        self.assertEqual("agent-continuation-pack/v1", envelope["continuation_pack"]["version"])
        self.assertTrue(envelope["continuation_pack"]["continuation_brief"])

    def test_execution_journal_summarizes_steps_tools_quality_and_ledger(self):
        journal = build_execution_journal({
            "agent_run_id": "run-2",
            "ok": True,
            "profile": "quick_scan",
            "answer": "完成",
            "agent_steps": [
                {"index": 1, "type": "tool_call", "tool": "get_table_rows", "summary": "读取表格", "ok": True},
                {"index": 2, "type": "quality_repair_tool_call", "tool": "get_table_rows", "summary": "补证据", "ok": True},
                {"index": 3, "type": "final_answer", "summary": "完成", "ok": True},
            ],
            "tool_calls": [
                {"index": 1, "tool": "get_table_rows", "ok": True, "evidence_node_ids": ["ev-1"]},
            ],
            "tool_dispatch_trace": [
                {"event_index": 1, "status": "completed", "tool": "get_table_rows", "reason": "读取表格", "evidence_node_ids": ["ev-1"]},
                {"event_index": 2, "status": "duplicate", "tool": "get_table_rows", "reason": "重复参数"},
            ],
            "final_answer_quality_gate": {"status": "warn", "score": 75, "reasons": [{"id": "incomplete", "message": "截断"}]},
            "runtime_state": {"task_ledger": {"progress": {"completed": 1, "open": 1, "blocked": 0}}},
        })
        summary = build_journal_summary(journal)

        self.assertEqual("run_started", journal[0]["type"])
        self.assertTrue(any(item["type"] == "quality_repair_tool_call" for item in journal))
        self.assertTrue(any(item["type"] == "final_answer_quality_gate" for item in journal))
        self.assertTrue(any(item["type"] == "tool_dispatch" and item["status"] == "warn" for item in journal))
        self.assertTrue(any(item.get("evidence_ids") == ["ev-1"] for item in journal))
        self.assertEqual("agent-run-journal/v1", summary["version"])
        self.assertGreaterEqual(summary["tool_event_count"], 2)

    def test_turn_context_snapshot_summarizes_runtime_boundaries(self):
        snapshot = build_harness_turn_context_snapshot(
            agent_run_id="run-ctx",
            mode="local-agent-harness",
            profile="auto",
            capability_profiles=["dfmea_prep", "datasheet_qa"],
            model_provider="MockProvider",
            model_mode="mock",
            guidance_summary={"source_count": 2, "sources": ["AGENTS.md"]},
            selected_skills={"selected_count": 1, "selected": [{"id": "datasheet-review"}]},
            playbook_plan={"playbook_ids": ["dfmea_prepare"], "recommended_first_tools": ["datasheet-status"]},
            allowed_tools={"datasheet-status", "read_project_text"},
            tool_list=[
                {
                    "name": "datasheet-status",
                    "readonly": True,
                    "file_access": False,
                    "mutating": False,
                    "supports_parallel": False,
                    "approval_scope": "none",
                    "evidence_kind": "datasheet",
                },
                {
                    "name": "read_project_text",
                    "readonly": True,
                    "file_access": True,
                    "mutating": False,
                    "supports_parallel": False,
                    "approval_scope": "read_project_file",
                    "evidence_kind": "project_file",
                },
            ],
            context_budget={"truncated": True, "model_observation_json_chars": 200, "source_observation_count": 4},
            runtime_state={
                "protocol_version": "pstx-agent-runtime/v1",
                "evidence_id_count": 3,
                "memory_summary": {"facts": ["U1"]},
                "task_ledger": {"progress": {"open": 1}},
                "evidence_goal_contract": {"status": "partial", "missing_evidence_types": ["datasheet_detail"]},
            },
            limits={"max_steps": 3},
            safeguards=["readonly"],
        )

        self.assertEqual("pstx-harness-turn-context.v1", snapshot["schema_version"])
        self.assertEqual(2, snapshot["tool_boundary"]["allowed_tool_count"])
        self.assertEqual(["read_project_text"], snapshot["tool_boundary"]["file_access_tools"])
        self.assertFalse(snapshot["tool_boundary"]["mutating"])
        self.assertEqual({"none": 1, "read_project_file": 1}, snapshot["tool_boundary"]["approval_scopes"])
        self.assertEqual({"datasheet": 1, "project_file": 1}, snapshot["tool_boundary"]["evidence_kinds"])
        self.assertTrue(snapshot["context_budget"]["truncated"])
        self.assertEqual("partial", snapshot["runtime_state"]["evidence_goal_status"])

    def test_continuation_pack_carries_open_work_evidence_and_questions(self):
        pack = build_continuation_pack({
            "agent_run_id": "run-3",
            "ok": True,
            "status": "waiting_for_user",
            "profile": "dfmea_prep",
            "answer": "需要补规格。",
            "request": {"question": "做 DFMEA 准备"},
            "final_evidence": [{"id": "ev-u1"}],
            "needs_user_input": {"questions": [{"question_id": "q1", "question": "请补 U1 规格", "missing_fields": ["spec"]}]},
            "final_answer_quality_gate": {
                "status": "warn",
                "score": 70,
                "repair_actions": [{"type": "tool_call", "tool": "get_component_identity_card", "title": "补身份", "source": "incomplete_tool_result-1"}],
            },
            "runtime_state": {
                "task_ledger": {
                    "items": [{
                        "id": "item-1",
                        "title": "补规格书",
                        "status": "in_progress",
                        "recommended_tools": ["search_datasheets"],
                        "note": "规格书缺失",
                    }],
                    "next_actions": [{
                        "type": "tool_call",
                        "title": "批量查身份卡",
                        "tool": "batch_get_component_identity_cards",
                        "args": {"refdes_list": ["U1", "PU2"]},
                        "reason": "继续上一轮已抽取的位号。",
                    }],
                }
            },
            "session_state": {"recent_evidence_ids": ["ev-old"]},
        })

        self.assertEqual("agent-continuation-pack/v1", pack["version"])
        self.assertEqual("await_user_input", pack["next_intent"])
        self.assertIn("ev-u1", pack["evidence_ids"])
        self.assertEqual("q1", pack["pending_questions"][0]["question_id"])
        self.assertEqual("item-1", pack["open_ledger_items"][0]["id"])
        self.assertEqual("get_component_identity_card", pack["suggested_tool_calls"][0]["tool"])
        self.assertEqual("batch_get_component_identity_cards", pack["suggested_tool_calls"][1]["tool"])
        self.assertEqual(["U1", "PU2"], pack["suggested_tool_calls"][1]["args"]["refdes_list"])

    def test_trace_store_preserves_legacy_payload_access_and_eviction(self):
        store = AgentTraceStore(max_items=2)
        first = store.remember({"agent_run_id": "a", "answer": "A"})
        store.remember({"agent_run_id": "b", "answer": "B"})
        self.assertIsNotNone(first)
        self.assertIn("a", store)
        self.assertEqual("A", store.get("a")["answer"])

        store.remember({"agent_run_id": "c", "answer": "C"})

        self.assertNotIn("a", store)
        self.assertIn("b", store)
        self.assertIn("c", store)
        self.assertEqual(["b", "c"], store.keys())
        self.assertEqual("C", store["c"]["answer"])

        envelope = store.get_envelope("c")
        self.assertEqual("C", envelope["payload"]["answer"])
        store.clear()
        self.assertEqual(0, len(store))

    def test_runtime_planner_combines_auto_profiles_and_allowed_tools(self):
        profiles = {
            "auto": {
                "title": "Auto",
                "description": "auto",
                "tools": [],
                "default_question": "自动",
                "max_steps": 4,
                "max_tool_calls": 8,
            },
            "quick": {
                "title": "Quick",
                "description": "quick",
                "tools": ["list"],
                "default_question": "快速",
                "max_steps": 4,
                "max_tool_calls": 8,
            },
            "bom": {
                "title": "BOM",
                "description": "bom",
                "tools": ["bom_search"],
                "default_question": "BOM",
                "max_steps": 6,
                "max_tool_calls": 10,
                "subagent_profiles": ["quick"],
            },
            "full": {
                "title": "Full",
                "description": "full",
                "tools": ["*"],
                "default_question": "完整",
                "max_steps": 8,
                "max_tool_calls": 12,
            },
        }
        registry_tools = [{"name": "list"}, {"name": "bom_search"}, {"name": "read_file"}]
        plan = build_capability_plan(
            requested_profile="auto",
            question="请检查 BOM 和飞书",
            profiles=profiles,
            default_profile="quick",
            quick_profile="quick",
            rules=[("bom", ["bom", "飞书"]), ("full", ["完整"])],
            registry_tools=registry_tools,
        )

        self.assertEqual(["bom"], list(plan.capability_profiles))
        self.assertEqual(["bom_search"], list(plan.allowed_tools))
        self.assertEqual("BOM", plan.plan_items[0]["title"])
        self.assertEqual(["list", "bom_search", "read_file"], allowed_tool_names(
            profile_ids=["full"],
            profiles=profiles,
            registry_tools=registry_tools,
        ))
        self.assertEqual(["bom_search"], [item["name"] for item in filtered_tool_list(
            profile_ids=["bom"],
            profiles=profiles,
            registry_tools=registry_tools,
        )])
        profile = AgentCapabilityProfile.from_mapping("bom", profiles["bom"]).to_public_dict(include_subagents=True)
        self.assertEqual(["quick"], profile["subagent_profiles"])

    def test_runtime_playbook_plan_combines_multiple_report_routes(self):
        plan = build_playbook_plan(
            question="请统计 page_rows 总页数，并查询 U1、U2 的 HQ 料号和飞书 PI",
            capability_profiles=["quick_scan", "feishu_bom_qa"],
            allowed_tools=[
                "summarize_table_column_values",
                "batch_query_report_entities",
                "batch_search_feishu_cache_rows",
                "get_table_rows",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("table_column_aggregation", playbook_ids)
        self.assertIn("report_entity_batch_lookup", playbook_ids)
        self.assertIn("feishu_material_qa", playbook_ids)
        self.assertEqual("summarize_table_column_values", plan["recommended_first_tools"][0])
        self.assertIn("batch_query_report_entities", plan["recommended_first_tools"])
        self.assertIn("feishu_material", plan["evidence_goals"])
        self.assertTrue(any("get_table_rows" in item for item in plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("batch_query_report_entities", seeded)
        self.assertIn("U1", seeded["batch_query_report_entities"]["args"]["queries"])
        self.assertIn("U2", seeded["batch_query_report_entities"]["args"]["queries"])
        self.assertIn("batch_search_feishu_cache_rows", seeded)

    def test_runtime_playbook_plan_prefers_module_order_for_schematic_page_count(self):
        plan = build_playbook_plan(
            question="我这个项目一共有多少页原理图？",
            capability_profiles=["page_mapping"],
            allowed_tools=[
                "summarize_schematic_page_count",
                "summarize_table_column_values",
                "get_table_rows",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("schematic_page_count", playbook_ids)
        self.assertEqual("summarize_schematic_page_count", plan["recommended_first_tools"][0])
        self.assertIn("不要用 page_rows", " ".join(plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("summarize_schematic_page_count", seeded)
        self.assertEqual({}, seeded["summarize_schematic_page_count"]["args"])

    def test_runtime_playbook_plan_supports_source_file_drilldown(self):
        plan = build_playbook_plan(
            question="请把 U1 的分析结论追溯到原始文件级别",
            capability_profiles=["page_mapping"],
            allowed_tools=["trace_project_source", "search_project_text", "read_project_text", "batch_query_report_entities"],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("source_file_drilldown", playbook_ids)
        self.assertEqual("trace_project_source", plan["recommended_first_tools"][0])
        self.assertIn("source_trace", plan["evidence_goals"])
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("trace_project_source", seeded)
        self.assertEqual("refdes", seeded["trace_project_source"]["args"]["kind"])
        self.assertIn("U1", seeded["trace_project_source"]["args"]["query"])
        self.assertIn("search_project_text", plan["recommended_first_tools"])

    def test_runtime_playbook_plan_seeds_project_grep_for_raw_search(self):
        plan = build_playbook_plan(
            question="请 grep 原始工程文件里的 U1 和 I2C_SCL",
            capability_profiles=["page_mapping"],
            allowed_tools=["search_project_text", "read_project_text"],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("source_file_drilldown", playbook_ids)
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("search_project_text", seeded)
        self.assertIn("U1", seeded["search_project_text"]["args"]["query"])
        self.assertIn("I2C_SCL", seeded["search_project_text"]["args"]["query"])
        self.assertEqual(12, seeded["search_project_text"]["args"]["limit"])

    def test_source_trace_tool_contract_points_to_raw_excerpt_detail(self):
        contract = build_tool_result_contract("trace_project_source", {
            "source_hits": [
                {
                    "path": "packaged/pstxprt.dat",
                    "line_start": 2,
                    "line_end": 4,
                    "detail_tool": {
                        "name": "read_project_text",
                        "args": {"path": "packaged/pstxprt.dat", "line_start": 2, "line_count": 3},
                    },
                }
            ],
            "detail_tool": {
                "name": "read_project_text",
                "args": {"path": "packaged/pstxprt.dat", "line_start": 2, "line_count": 3},
            },
            "completeness": "complete",
        }).to_dict()

        self.assertEqual("complete", contract["completeness"])
        self.assertIn("read_project_text", contract["recommended_next_tools"])
        self.assertEqual("read_project_text", contract["detail_tool"]["name"])

    def test_project_text_search_tool_contract_points_to_raw_excerpt_detail(self):
        contract = build_tool_result_contract("search_project_text", {
            "source_hits": [
                {
                    "path": "sch_1/page12.csa",
                    "line_start": 10,
                    "line_end": 12,
                    "detail_tool": {
                        "name": "read_project_text",
                        "args": {"path": "sch_1/page12.csa", "line_start": 10, "line_count": 3},
                    },
                }
            ],
            "detail_tool": {
                "name": "read_project_text",
                "args": {"path": "sch_1/page12.csa", "line_start": 10, "line_count": 3},
            },
            "completeness": "complete",
        }).to_dict()

        self.assertEqual("complete", contract["completeness"])
        self.assertIn("read_project_text", contract["recommended_next_tools"])
        self.assertEqual("read_project_text", contract["detail_tool"]["name"])

    def test_runtime_playbook_plan_does_not_seed_page_count_for_mapping_review(self):
        plan = build_playbook_plan(
            question="请检查页码映射是否正确。",
            capability_profiles=["page_mapping"],
            allowed_tools=["summarize_schematic_page_count", "summarize_table_column_values"],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertNotIn("summarize_schematic_page_count", seeded)

    def test_runtime_playbook_plan_supports_agent_ref_pdf_qa(self):
        plan = build_playbook_plan(
            question="请基于 ref 参考资料说明 Agent 能力边界",
            capability_profiles=["agent_ref_qa"],
            allowed_tools=[
                "list_agent_ref_sources",
                "search_agent_ref_pdfs",
                "get_agent_ref_pdf_excerpt",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("agent_ref_pdf_qa", playbook_ids)
        self.assertIn("search_agent_ref_pdfs", plan["recommended_first_tools"])
        self.assertIn("agent_ref_excerpt", plan["evidence_goals"])

    def test_runtime_playbook_plan_supports_datasheet_pdf_qa(self):
        plan = build_playbook_plan(
            question="请查 HQ100 GPU_CORE_TEST_IC 的 datasheet absolute maximum 参数",
            capability_profiles=["datasheet_qa"],
            allowed_tools=[
                "list_datasheet_documents",
                "search_datasheet_chunks",
                "batch_search_datasheet_chunks",
                "get_datasheet_chunk",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("datasheet_pdf_qa", playbook_ids)
        self.assertIn("batch_search_datasheet_chunks", plan["recommended_first_tools"])
        self.assertIn("datasheet_chunk", plan["evidence_goals"])
        self.assertIn("不要只凭搜索 snippet", " ".join(plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("batch_search_datasheet_chunks", seeded)
        self.assertIn("HQ100", seeded["batch_search_datasheet_chunks"]["args"]["queries"])

    def test_runtime_playbook_plan_supports_datasheet_connection_review(self):
        plan = build_playbook_plan(
            question="请根据 datasheet 反查 U1 和 U2 的 I2C 连接是否有问题，重点看接口电平和 reset timing。",
            capability_profiles=["connection_datasheet_review"],
            allowed_tools=[
                "list_datasheet_sources",
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
                "get_datasheet_chunk",
                "trace_project_source",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("schematic_datasheet_connection_review", playbook_ids)
        self.assertIn("batch_query_llm_topology_netlist", plan["recommended_first_tools"])
        self.assertIn("batch_get_component_identity_cards", plan["recommended_first_tools"])
        self.assertIn("batch_match_component_datasheets", plan["recommended_first_tools"])
        self.assertIn("datasheet_chunk", plan["evidence_goals"])
        self.assertIn("component_identity", plan["evidence_goals"])
        self.assertIn("llm_topology_edge", plan["evidence_goals"])
        self.assertIn("不要只看 datasheet 摘要", " ".join(plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("list_datasheet_sources", seeded)
        self.assertIn("batch_query_llm_topology_netlist", seeded)
        self.assertIn("batch_get_component_identity_cards", seeded)
        self.assertIn("batch_match_component_datasheets", seeded)
        self.assertIn("search_datasheet_parameters", seeded)
        self.assertIn("batch_search_datasheet_chunks", seeded)
        self.assertIn("trace_project_source", seeded)
        self.assertEqual(["U1", "U2"], seeded["batch_get_component_identity_cards"]["args"]["refdes_list"])
        self.assertIn("I2C", seeded["batch_query_llm_topology_netlist"]["args"]["queries"])
        self.assertIn("VIH VIL", seeded["batch_search_datasheet_chunks"]["args"]["queries"])

    def test_seeded_prefetch_prioritizes_datasheet_connection_chain(self):
        plan = {
            "seeded_tool_calls": [
                {"name": "trace_project_source", "args": {"query": "U1 U2", "kind": "refdes"}},
                {"name": "batch_search_datasheet_chunks", "args": {"queries": ["VIH VIL"], "limit_per_query": 6}},
                {"name": "batch_match_component_datasheets", "args": {"refdes_list": ["U1", "U2"], "limit_per_component": 4}},
                {"name": "batch_query_llm_topology_netlist", "args": {"queries": ["U1", "U2"], "limit_per_query": 8}},
                {"name": "batch_get_component_identity_cards", "args": {"refdes_list": ["U1", "U2"]}},
                {"name": "list_datasheet_sources", "args": {}},
            ]
        }

        prefetch = select_seeded_prefetch_tool_calls(
            plan,
            allowed_tools={
                "trace_project_source",
                "batch_search_datasheet_chunks",
                "batch_match_component_datasheets",
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "list_datasheet_sources",
            },
            max_calls=4,
            remaining_tool_calls=6,
        )

        self.assertEqual(
            [
                "list_datasheet_sources",
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
            ],
            [item["name"] for item in prefetch["tool_calls"]],
        )

    def test_runtime_playbook_plan_supports_chip_level_topology(self):
        plan = build_playbook_plan(
            question="请说明大芯片 U1 和电平转换芯片 U2 的芯片级拓扑连接关系",
            capability_profiles=["chip_topology"],
            allowed_tools=[
                "summarize_llm_topology_netlist",
                "batch_query_llm_topology_netlist",
                "get_llm_topology_node",
                "get_llm_topology_edge",
                "summarize_chip_topology",
                "batch_query_chip_topology",
                "get_chip_topology_edge",
                "batch_get_component_identity_cards",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("chip_level_topology", playbook_ids)
        self.assertEqual("batch_query_llm_topology_netlist", plan["recommended_first_tools"][0])
        self.assertIn("summarize_llm_topology_netlist", plan["recommended_first_tools"])
        self.assertIn("llm_topology_edge", plan["evidence_goals"])
        self.assertIn("llm_topology_node", plan["evidence_goals"])
        self.assertIn("不要把 R/C/L", " ".join(plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("summarize_llm_topology_netlist", seeded)
        self.assertIn("batch_query_llm_topology_netlist", seeded)
        self.assertIn("U1", seeded["batch_query_llm_topology_netlist"]["args"]["queries"])
        self.assertIn("U2", seeded["batch_query_llm_topology_netlist"]["args"]["queries"])

    def test_runtime_playbook_plan_supports_local_document_search(self):
        plan = build_playbook_plan(
            question="请在文档中搜索 U46 多 symbol 这一段内容",
            capability_profiles=["document_search"],
            allowed_tools=[
                "list_document_search_sources",
                "search_documents",
                "batch_search_documents",
                "get_document_excerpt",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("local_document_search", playbook_ids)
        self.assertEqual("batch_search_documents", plan["recommended_first_tools"][0])
        self.assertIn("document_match", plan["evidence_goals"])
        self.assertIn("不要只凭文件名", " ".join(plan["anti_patterns"]))
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("search_documents", seeded)
        self.assertIn("U46", seeded["search_documents"]["args"]["query"])

    def test_runtime_playbook_plan_connects_compare_to_datasheet_pdf_qa(self):
        plan = build_playbook_plan(
            question="请对比 U46 的 HQ11112042009 规格书 absolute maximum 差异",
            capability_profiles=["compare_datasheet_qa", "compare_bom_feishu"],
            allowed_tools=[
                "batch_query_compare_diff",
                "batch_search_datasheet_chunks",
                "query_compare_diff",
                "search_datasheet_chunks",
                "get_datasheet_chunk",
            ],
            playbooks=COMPARE_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("compare_datasheet_pdf_qa", playbook_ids)
        self.assertIn("compare_bom_feishu_material", playbook_ids)
        self.assertIn("batch_search_datasheet_chunks", plan["recommended_first_tools"])
        self.assertIn("batch_query_compare_diff", plan["recommended_first_tools"])
        self.assertIn("datasheet_chunk", plan["evidence_goals"])
        self.assertIn("compare_diff", plan["evidence_goals"])
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("batch_search_datasheet_chunks", seeded)
        self.assertIn("HQ11112042009", seeded["batch_search_datasheet_chunks"]["args"]["queries"])
        self.assertIn("batch_query_compare_diff", seeded)
        self.assertIn("U46", seeded["batch_query_compare_diff"]["args"]["queries"])

    def test_evidence_goal_contract_detects_missing_playbook_evidence(self):
        plan = build_playbook_plan(
            question="请对比 U46 的 HQ11112042009 规格书 absolute maximum 差异",
            capability_profiles=["compare_datasheet_qa", "compare_bom_feishu"],
            allowed_tools=[
                "batch_query_compare_diff",
                "batch_search_datasheet_chunks",
                "query_compare_diff",
                "get_datasheet_chunk",
            ],
            playbooks=COMPARE_AGENT_PLAYBOOKS,
        ).to_dict()
        contract = build_evidence_goal_contract(
            playbook_plan=plan,
            evidence_nodes=[{"id": "ev-diff", "type": "compare_diff"}],
        )

        self.assertEqual("agent-evidence-goal-contract/v1", contract["version"])
        self.assertEqual("partial", contract["status"])
        self.assertIn("compare_diff", contract["present_evidence_types"])
        self.assertIn("datasheet_chunk", contract["missing_evidence_types"])
        self.assertIn("batch_search_datasheet_chunks", contract["recommended_next_tools"])
        self.assertEqual("missing_evidence_goal", contract["repair_actions"][0]["source"])

        runtime_state = build_runtime_state(
            goal="对比规格书参数",
            playbook_plan=plan,
            observations=[{"tool": "batch_query_compare_diff", "evidence_nodes": [{"id": "ev-diff", "type": "compare_diff"}]}],
        )
        self.assertEqual("partial", runtime_state["evidence_goal_contract"]["status"])

    def test_evidence_goal_contract_tracks_connection_review_phases(self):
        plan = build_playbook_plan(
            question="请根据 datasheet 反查 U1 的 I2C 连接是否有问题，重点看接口电平。",
            capability_profiles=["connection_datasheet_review"],
            allowed_tools=[
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
                "get_datasheet_chunk",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()
        contract = build_evidence_goal_contract(
            playbook_plan=plan,
            evidence_nodes=[{
                "id": "ev-edge",
                "type": "llm_topology_edge",
                "title": "U1 I2C topology edge",
            }],
        )

        self.assertEqual("partial", contract["connection_review_phase_status"])
        missing_phase_ids = [item["id"] for item in contract["missing_connection_review_phases"]]
        self.assertIn("component_identity", missing_phase_ids)
        self.assertIn("datasheet_locator", missing_phase_ids)
        self.assertIn("datasheet_detail", missing_phase_ids)
        repair_tools = [item.get("tool") for item in contract["connection_review_repair_actions"]]
        self.assertIn("batch_get_component_identity_cards", repair_tools)
        self.assertIn("batch_match_component_datasheets", repair_tools)
        self.assertIn("search_datasheet_parameters", repair_tools)

    def test_evidence_goal_contract_opens_datasheet_detail_for_connection_review(self):
        plan = build_playbook_plan(
            question="请根据 datasheet 反查 U1 的 reset timing 连接。",
            capability_profiles=["connection_datasheet_review"],
            allowed_tools=[
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
                "get_datasheet_chunk",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()
        contract = build_evidence_goal_contract(
            playbook_plan=plan,
            evidence_nodes=[
                {"id": "ev-edge", "type": "llm_topology_edge", "title": "U1 reset edge"},
                {"id": "ev-id", "type": "component_identity", "title": "U1 identity"},
                {
                    "id": "ev-ds-hit",
                    "type": "datasheet_match",
                    "title": "U1 datasheet match",
                    "source": {"tool": "batch_match_component_datasheets"},
                    "detail_tool": {"name": "get_datasheet_chunk", "args": {"doc_id": 1, "chunk_id": "p3-c2", "max_chars": 4000}},
                },
            ],
        )

        self.assertEqual("partial", contract["connection_review_phase_status"])
        self.assertEqual(["datasheet_detail"], [item["id"] for item in contract["missing_connection_review_phases"]])
        action = contract["connection_review_repair_actions"][0]
        self.assertEqual("missing_connection_review_phase", action["source"])
        self.assertEqual("get_datasheet_chunk", action["tool"])
        self.assertEqual({"doc_id": 1, "chunk_id": "p3-c2", "max_chars": 4000}, action["args"])

    def test_runtime_playbook_plan_combines_compare_page_and_diff_routes(self):
        plan = build_playbook_plan(
            question="请比对第 1-30 页里 U46 pin/net 和飞书料号差异",
            capability_profiles=["compare_cadence_pages", "compare_pin_net", "compare_bom_feishu"],
            allowed_tools=[
                "resolve_compare_page_range",
                "compare_cadence_page_semantics",
                "batch_query_compare_diff",
                "batch_get_cadence_page_objects",
            ],
            playbooks=COMPARE_AGENT_PLAYBOOKS,
        ).to_dict()

        playbook_ids = [item["id"] for item in plan["selected_playbooks"]]
        self.assertIn("cadence_page_semantic_compare", playbook_ids)
        self.assertIn("compare_diff_batch_lookup", playbook_ids)
        self.assertIn("compare_bom_feishu_material", playbook_ids)
        self.assertIn("batch_query_compare_diff", plan["recommended_first_tools"])
        self.assertIn("cadence_topology_diff", plan["evidence_goals"])
        seeded = {item["name"]: item for item in plan["seeded_tool_calls"]}
        self.assertIn("compare_cadence_page_semantics", seeded)
        self.assertEqual(1, seeded["compare_cadence_page_semantics"]["args"]["page_start"])
        self.assertEqual(30, seeded["compare_cadence_page_semantics"]["args"]["page_end"])
        self.assertIn("batch_query_compare_diff", seeded)
        self.assertIn("U46", seeded["batch_query_compare_diff"]["args"]["queries"])

    def test_tool_result_contract_guides_truncated_page_rows_to_schematic_count(self):
        contract = build_tool_result_contract("get_table_rows", {
            "table_id": "page_rows",
            "total_rows": 254,
            "rows": [{"页码": "1"}],
            "limit": 1,
            "next_offset": 1,
            "has_more": True,
            "truncated": True,
            "aggregation_hint": "请调用 summarize_table_column_values。",
        }).to_dict()

        self.assertEqual("truncated", contract["completeness"])
        self.assertEqual("summarize_schematic_page_count", contract["recommended_next_tools"][0])
        self.assertEqual("summarize_schematic_page_count", contract["aggregation_tool"]["name"])
        self.assertEqual({}, contract["aggregation_tool"]["args"])
        self.assertEqual("get_table_rows", contract["detail_tool"]["name"])

    def test_tool_result_contract_keeps_column_aggregation_for_non_page_tables(self):
        contract = build_tool_result_contract("get_table_rows", {
            "table_id": "drc_missing_hq_code",
            "total_rows": 254,
            "rows": [{"位号": "U1"}],
            "limit": 1,
            "next_offset": 1,
            "has_more": True,
            "truncated": True,
        }).to_dict()

        self.assertEqual("truncated", contract["completeness"])
        self.assertEqual("summarize_table_column_values", contract["recommended_next_tools"][0])
        self.assertEqual("summarize_table_column_values", contract["aggregation_tool"]["name"])
        self.assertEqual("drc_missing_hq_code", contract["aggregation_tool"]["args"]["table_id"])

    def test_tool_result_contract_marks_partial_batch_table_reads(self):
        contract = build_tool_result_contract("batch_get_table_rows", {
            "items_truncated": True,
            "truncated": True,
            "items": [{
                "table_id": "page_rows",
                "total_rows": 254,
                "rows": [{"页码": "1"}],
                "has_more": True,
                "next_offset": 1,
                "truncated": True,
            }],
        }).to_dict()

        self.assertEqual("truncated", contract["completeness"])
        self.assertIn("get_table_rows", contract["recommended_next_tools"])
        self.assertIn("summarize_table_column_values", contract["recommended_next_tools"])
        self.assertEqual("get_table_rows", contract["detail_tool"]["name"])
        self.assertEqual("page_rows", contract["detail_tool"]["args"]["table_id"])
        self.assertEqual(1, contract["detail_tool"]["args"]["offset"])

    def test_task_ledger_preserves_seeded_tool_call_args(self):
        ledger = build_task_ledger(
            goal="查询多个器件",
            capability_plan=[{"id": "quick_scan", "title": "快速扫描"}],
            playbook_plan={
                "selected_playbooks": [{"id": "report_entity_batch_lookup", "title": "报告实体批量查询"}],
                "recommended_first_tools": ["batch_query_report_entities"],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U1", "U2", "HQ11112042009"], "limit_per_query": 10},
                    "reason": "本地提取实体。",
                    "source": "playbook_seed",
                }],
            },
        )

        action = ledger["next_actions"][0]
        self.assertEqual("batch_query_report_entities", action["tool"])
        self.assertEqual("playbook_seed", action["source"])
        self.assertEqual(["U1", "U2", "HQ11112042009"], action["args"]["queries"])

    def test_task_ledger_expands_datasheet_connection_review_phases(self):
        playbook_plan = build_playbook_plan(
            question="请根据 datasheet 反查 U1 和 U2 的 I2C 连接是否有问题，重点看接口电平。",
            capability_profiles=["connection_datasheet_review"],
            allowed_tools=[
                "list_datasheet_sources",
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
                "trace_project_source",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()

        ledger = build_task_ledger(
            goal="连接反查",
            capability_plan=[{"id": "connection_datasheet_review", "title": "连接 × Datasheet 反查"}],
            playbook_plan=playbook_plan,
        )

        item_by_id = {item["id"]: item for item in ledger["items"]}
        self.assertIn("connection-review-targets", item_by_id)
        self.assertEqual("completed", item_by_id["connection-review-targets"]["status"])
        self.assertEqual("pending", item_by_id["connection-review-schematic-evidence"]["status"])
        self.assertEqual("pending", item_by_id["connection-review-identity"]["status"])
        self.assertEqual("pending", item_by_id["connection-review-datasheet-detail"]["status"])
        tools = [item.get("tool") for item in ledger["next_actions"]]
        self.assertIn("batch_query_llm_topology_netlist", tools)
        self.assertIn("batch_get_component_identity_cards", tools)
        self.assertIn("batch_match_component_datasheets", tools)
        self.assertIn("search_datasheet_parameters", tools)
        self.assertIn("trace_project_source", tools)

    def test_task_ledger_marks_datasheet_connection_detail_gap_as_actionable_progress(self):
        ledger = build_task_ledger(
            goal="连接反查",
            capability_plan=[{"id": "connection_datasheet_review", "title": "连接 × Datasheet 反查"}],
            playbook_plan={
                "selected_playbooks": [{
                    "id": "schematic_datasheet_connection_review",
                    "title": "原理图连接 × Datasheet 反查",
                }],
                "seeded_tool_calls": [],
            },
            observations=[
                {
                    "tool": "batch_query_llm_topology_netlist",
                    "evidence_nodes": [{"id": "ev-edge", "type": "llm_topology_edge", "title": "U1 I2C edge"}],
                },
                {
                    "tool": "batch_get_component_identity_cards",
                    "evidence_nodes": [{"id": "ev-id", "type": "component_identity", "title": "U1 identity"}],
                },
                {
                    "tool": "batch_match_component_datasheets",
                    "evidence_nodes": [{"id": "ev-gap", "type": "datasheet_gap", "title": "U1 datasheet missing"}],
                },
            ],
        )

        item_by_id = {item["id"]: item for item in ledger["items"]}
        self.assertEqual("completed", item_by_id["connection-review-schematic-evidence"]["status"])
        self.assertEqual("completed", item_by_id["connection-review-identity"]["status"])
        self.assertEqual("completed", item_by_id["connection-review-datasheet-locator"]["status"])
        self.assertEqual("completed", item_by_id["connection-review-datasheet-detail"]["status"])
        self.assertIn("gap", item_by_id["connection-review-datasheet-detail"]["note"].lower())

    def test_seeded_prefetch_selects_safe_allowed_calls(self):
        plan = {
            "seeded_tool_calls": [
                {"name": "batch_query_report_entities", "args": {"queries": ["U1"], "limit_per_query": 10}},
                {"name": "unknown_tool", "args": {}},
                {"name": "search_documents", "args": {"query": "<需要用户补充>", "limit": 5}},
                {"name": "batch_query_report_entities", "args": {"queries": ["U1"], "limit_per_query": 10}},
            ]
        }

        prefetch = select_seeded_prefetch_tool_calls(
            plan,
            allowed_tools={"batch_query_report_entities", "search_documents"},
            max_calls=2,
            remaining_tool_calls=4,
        )

        self.assertEqual("agent-prefetch-plan/v1", prefetch["version"])
        self.assertTrue(prefetch["enabled"])
        self.assertEqual(1, prefetch["selected_count"])
        self.assertEqual("batch_query_report_entities", prefetch["tool_calls"][0]["name"])
        self.assertGreaterEqual(len(prefetch["skipped"]), 3)

    def test_prefetch_followup_selects_first_detail_hit(self):
        followup = select_prefetch_followup_tool_calls(
            [{
                "tool": "batch_search_datasheet_chunks",
                "raw_result": {
                    "items": [{
                        "query": "HQ100",
                        "status": "found",
                        "matches": [{"doc_id": 7, "chunk_id": "p1-c1", "page": 1}],
                    }]
                },
            }],
            allowed_tools={"get_datasheet_chunk"},
            max_calls=1,
            remaining_tool_calls=2,
        )

        self.assertEqual(1, followup["selected_count"])
        self.assertEqual("get_datasheet_chunk", followup["tool_calls"][0]["name"])
        self.assertEqual({"doc_id": 7, "chunk_id": "p1-c1", "max_chars": 4000}, followup["tool_calls"][0]["args"])

    def test_prefetch_followup_opens_component_datasheet_match(self):
        followup = select_prefetch_followup_tool_calls(
            [{
                "tool": "batch_match_component_datasheets",
                "raw_result": {
                    "items": [{
                        "refdes": "U1",
                        "status": "found",
                        "matches": [{"doc_id": 11, "chunk_id": "p3-c2", "page": 3}],
                    }]
                },
            }],
            allowed_tools={"get_datasheet_chunk"},
            max_calls=1,
            remaining_tool_calls=2,
        )

        self.assertEqual(1, followup["selected_count"])
        self.assertEqual("get_datasheet_chunk", followup["tool_calls"][0]["name"])
        self.assertEqual({"doc_id": 11, "chunk_id": "p3-c2", "max_chars": 4000}, followup["tool_calls"][0]["args"])

    def test_goal_prefetch_selects_safe_overview_tool(self):
        plan = {
            "recommended_first_tools": [
                "batch_search_datasheet_chunks",
                "summarize_dfmea_readiness",
                "search_component_identity_cards",
            ]
        }

        prefetch = select_goal_prefetch_tool_calls(
            plan,
            allowed_tools={"summarize_dfmea_readiness", "batch_search_datasheet_chunks"},
            remaining_tool_calls=2,
        )

        self.assertEqual("agent-goal-prefetch-plan/v1", prefetch["version"])
        self.assertEqual(1, prefetch["selected_count"])
        self.assertEqual("summarize_dfmea_readiness", prefetch["tool_calls"][0]["name"])
        self.assertEqual({}, prefetch["tool_calls"][0]["args"])
        self.assertGreaterEqual(len(prefetch["skipped"]), 1)

    def test_tool_result_contract_guides_agent_ref_search_to_excerpt(self):
        contract = build_tool_result_contract("search_agent_ref_pdfs", {
            "query": "能力边界",
            "matches": [{"doc_id": 1, "page": 2, "title": "manual.pdf"}],
        }).to_dict()

        self.assertEqual("complete", contract["completeness"])
        self.assertIn("get_agent_ref_pdf_excerpt", contract["recommended_next_tools"])

    def test_tool_result_contract_guides_memory_search_to_detail(self):
        contract = build_tool_result_contract("list_project_memory_evidence", {
            "cards": [{"id": "ev-u1"}],
        }).to_dict()

        self.assertIn("get_project_memory_evidence", contract["recommended_next_tools"])

    def test_tool_result_contract_guides_datasheet_search_to_chunk_detail(self):
        search_contract = build_tool_result_contract("search_datasheet_chunks", {
            "query": "HQ100 absolute maximum",
            "matches": [{"doc_id": 1, "page": 1, "chunk_id": "p1-c1", "title": "ds.pdf"}],
        }).to_dict()
        chunk_contract = build_tool_result_contract("get_datasheet_chunk", {
            "doc_id": 1,
            "chunk_id": "p1-c1",
            "truncated": False,
        }).to_dict()

        self.assertEqual("complete", search_contract["completeness"])
        self.assertIn("get_datasheet_chunk", search_contract["recommended_next_tools"])
        self.assertEqual("complete", chunk_contract["completeness"])

    def test_tool_result_contract_marks_complete_and_error_results(self):
        complete = build_tool_result_contract("summarize_table_column_values", {
            "table_id": "page_rows",
            "column": "页码",
            "total_rows": 254,
            "unique_count": 176,
            "values": ["1", "2"],
            "truncated": False,
        }).to_dict()
        error = build_tool_result_contract("get_table_rows", {"ok": False, "error": "missing table"}).to_dict()

        self.assertEqual("complete", complete["completeness"])
        self.assertIn("unique_count=176", complete["scope_summary"])
        self.assertEqual("error", error["completeness"])

    def test_perseverance_retry_note_rejects_premature_refusal(self):
        self.assertTrue(is_low_effort_answer("无法回答，信息不足。"))
        note = build_perseverance_retry_note(
            step_type="final_answer",
            answer="无法回答，信息不足。",
            tool_call_count=0,
            max_tool_calls=4,
            playbook_plan={"recommended_first_tools": ["list_report_tables"]},
            tool_result_contracts=[],
            evidence_node_count=0,
        )
        self.assertIn("未尝试本地只读工具", note)
        self.assertIn("list_report_tables", note)

        seeded_note = build_perseverance_retry_note(
            step_type="final_answer",
            answer="无法回答，信息不足。",
            tool_call_count=0,
            max_tool_calls=4,
            playbook_plan={
                "recommended_first_tools": ["batch_query_report_entities"],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U1", "HQ11112042009"], "limit_per_query": 10},
                }],
            },
            evidence_node_count=0,
        )
        self.assertIn("带参工具种子", seeded_note)
        self.assertIn("batch_query_report_entities", seeded_note)
        self.assertIn("U1", seeded_note)
        self.assertIn("HQ11112042009", seeded_note)

        ask_user_note = build_perseverance_retry_note(
            step_type="needs_user_input",
            tool_call_count=0,
            max_tool_calls=4,
            playbook_plan={
                "recommended_first_tools": ["batch_query_report_entities"],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U2"]},
                }],
            },
            evidence_node_count=0,
        )
        self.assertIn("U2", ask_user_note)

        continued_note = build_perseverance_retry_note(
            step_type="final_answer",
            answer="无法判断。",
            tool_call_count=1,
            max_tool_calls=4,
            playbook_plan={},
            tool_result_contracts=[{
                "completeness": "truncated",
                "recommended_next_tools": ["summarize_table_column_values"],
            }],
            evidence_node_count=1,
        )
        self.assertIn("继续取证", continued_note)

        ledger_note = build_perseverance_retry_note(
            step_type="final_answer",
            answer="已完成。",
            tool_call_count=0,
            max_tool_calls=4,
            task_ledger={
                "progress": {"in_progress": 1, "open": 1},
                "next_actions": [{
                    "type": "tool_call",
                    "tool": "summarize_table_column_values",
                    "title": "聚合 page_rows",
                }],
            },
            evidence_node_count=0,
        )
        self.assertIn("task_ledger", ledger_note)
        self.assertIn("summarize_table_column_values", ledger_note)

    def test_perseverance_does_not_block_when_tool_budget_is_exhausted(self):
        note = build_perseverance_retry_note(
            step_type="needs_user_input",
            answer="",
            tool_call_count=0,
            max_tool_calls=0,
            playbook_plan={"recommended_first_tools": ["summarize_dfmea_readiness"]},
            evidence_node_count=0,
        )
        self.assertEqual("", note)

    def test_runtime_compression_builds_context_budget_and_fits_items(self):
        observations = [
            {
                "tool": "tool",
                "summary": "x" * 2000,
                "evidence_node_ids": [f"ev-{index}"],
                "evidence_nodes": [{"id": f"ev-node-{index}", "summary": "y" * 500}],
                "result_preview_omitted": index % 2 == 0,
                "truncated_for_model": True,
            }
            for index in range(6)
        ]
        fitted = fit_items_to_json_budget(
            observations,
            json_budget_chars=1600,
            compact_item=lambda item: {
                "tool": item.get("tool"),
                "summary": str(item.get("summary"))[:20],
                "evidence_node_ids": item.get("evidence_node_ids", []),
                "truncated_for_model": True,
            },
        )
        budget = build_context_budget_summary(
            observations,
            fitted,
            json_budget_chars=1600,
            truncated_note="已压缩",
            ok_note="未压缩",
            include_observation_bundle=True,
            bundle_id="test-bundle",
        )

        self.assertLessEqual(json_char_count(fitted), 1600)
        self.assertTrue(budget["truncated"])
        self.assertEqual("已压缩", budget["notes"])
        self.assertEqual("test-bundle", budget["observation_bundle"]["id"])
        self.assertNotIn("observations", budget["observation_bundle"])

    def test_runtime_finalizer_normalizes_citations_actions_and_user_input(self):
        evidence_nodes = [{"id": "ev-1", "title": "证据", "type": "table_row", "locator": {"row": 1}, "source": {}}]
        raw = {
            "citations": [{"id": "ev-missing", "note": "bad"}],
            "proposed_actions": [{"title": "复核 U1", "reason": "证据不足"}],
        }
        citations, citation_meta = normalize_citations(raw, evidence_nodes, fallback_when_empty=True)
        actions = normalize_proposed_actions(raw)
        user_input = normalize_needs_user_input({
            "needs_user_input": {
                "reason": "缺规格",
                "missing_fields": ["spec"],
                "questions": [{"question_id": "q1", "question": "请补规格", "applies_to": {"refdes": "U1"}}],
            }
        }, evidence_nodes)
        result = AgentFinalizationResult(
            answer="需要补充",
            status=status_from_stopped_reason("needs_user_input"),
            stopped_reason="needs_user_input",
            citations=tuple(citations),
            proposed_actions=tuple(actions),
            invalid_citation_count=citation_meta["invalid_citation_count"],
            needs_user_input=user_input,
        ).to_dict()

        self.assertEqual(1, citation_meta["invalid_citation_count"])
        self.assertTrue(any(item.get("fallback") for item in citations))
        self.assertEqual("复核 U1", actions[0]["title"])
        self.assertEqual("waiting_for_user", result["status"])
        self.assertEqual("q1", result["needs_user_input"]["questions"][0]["question_id"])

    def test_final_answer_quality_gate_scores_evidence_and_ledger_risks(self):
        evidence_goal_contract = {
            "status": "partial",
            "missing_evidence_types": ["datasheet_chunk"],
            "repair_actions": [{
                "type": "tool_call",
                "tool": "batch_search_datasheet_chunks",
                "args": {"queries": ["HQ11112042009"], "limit_per_query": 8},
                "source": "missing_evidence_goal",
                "priority": 12,
            }],
        }
        gate = build_final_answer_quality_gate(
            answer="已完成。",
            citations=[{"id": "ev-missing", "valid": False}],
            evidence_nodes=[{"id": "ev-1", "title": "证据"}],
            tool_result_contracts=[{"completeness": "truncated", "recommended_next_tools": ["get_table_rows"]}],
            task_ledger={
                "progress": {"blocked": 1, "open": 2},
                "next_actions": [{"type": "tool_call", "tool": "get_table_rows"}],
            },
            evidence_goal_contract=evidence_goal_contract,
        )

        self.assertEqual("final-answer-quality-gate/v1", gate["version"])
        self.assertEqual("fail", gate["status"])
        self.assertLess(gate["score"], 100)
        reason_ids = [item["id"] for item in gate["reasons"]]
        self.assertIn("missing_valid_citation", reason_ids)
        self.assertIn("invalid_citation", reason_ids)
        self.assertIn("incomplete_tool_result", reason_ids)
        self.assertIn("blocked_ledger_item", reason_ids)
        self.assertIn("get_table_rows", gate["recommended_next_tools"])
        self.assertEqual("partial", gate["evidence_goal_contract"]["status"])
        self.assertGreaterEqual(gate["repair_action_count"], 4)
        repair_types = {item["type"] for item in gate["repair_actions"]}
        self.assertIn("revise_answer", repair_types)
        self.assertIn("tool_call", repair_types)
        self.assertIn("ask_user", repair_types)
        self.assertTrue(any(item.get("tool") == "get_table_rows" for item in gate["repair_actions"]))

    def test_quality_repair_tool_calls_skip_unsafe_actions(self):
        gate = {
            "repair_actions": [
                {"type": "tool_call", "tool": "get_table_rows", "args": {"table_id": "page_rows", "offset": 12, "limit": 12}, "priority": 10, "source": "incomplete_tool_result-1"},
                {"type": "tool_call", "tool": "summarize_table_column_values", "args": {"table_id": "page_rows", "column": "<需要统计的列名>"}, "priority": 20, "source": "incomplete_tool_result-1"},
                {"type": "tool_call", "tool": "get_component_identity_card", "priority": 30, "source": "incomplete_tool_result-2"},
                {"type": "tool_call", "tool": "read_project_text", "args": {"path": "x"}, "priority": 40, "source": "incomplete_tool_result-3"},
                {"type": "tool_call", "tool": "list_report_tables", "priority": 50, "source": "open_next_actions"},
                {"type": "tool_call", "tool": "list_datasheet_sources", "priority": 60, "source": "manual_review"},
            ],
        }
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"get_table_rows", "summarize_table_column_values", "get_component_identity_card"},
            max_calls=3,
        )

        self.assertEqual("quality-repair-plan/v1", plan["version"])
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("get_table_rows", plan["tool_calls"][0]["name"])
        skipped_reasons = {item["reason"] for item in plan["skipped_actions"]}
        self.assertIn("placeholder_args", skipped_reasons)
        self.assertIn("missing_args", skipped_reasons)
        self.assertIn("not_allowed", skipped_reasons)
        self.assertIn("non_evidence_repair", skipped_reasons)
        self.assertIn("open_next_not_needed", skipped_reasons)

    def test_quality_repair_tool_calls_respects_zero_budget(self):
        plan = build_quality_repair_tool_calls(
            {
                "repair_actions": [
                    {"type": "tool_call", "tool": "get_table_rows", "args": {"table_id": "page_rows"}, "source": "incomplete_tool_result-1"},
                ],
            },
            allowed_tools={"get_table_rows"},
            max_calls=0,
        )

        self.assertEqual(0, plan["selected_tool_call_count"])
        self.assertEqual([], plan["tool_calls"])
        self.assertEqual("max_calls_exhausted", plan["skipped_actions"][0]["reason"])

    def test_quality_repair_tool_calls_select_safe_open_next_actions(self):
        gate = {
            "reasons": [{"id": "low_effort_answer", "severity": "warn", "message": "过早放弃"}],
            "repair_actions": [
                {"type": "tool_call", "tool": "list_report_tables", "priority": 10, "source": "open_next_actions", "reason": "先列出可用表格"},
                {"type": "tool_call", "tool": "get_component_identity_card", "priority": 20, "source": "open_next_actions"},
                {"type": "tool_call", "tool": "summarize_table_column_values", "args": {"table_id": "page_rows", "column": "<需要列名>"}, "priority": 30, "source": "open_next_actions"},
                {"type": "tool_call", "tool": "list_report_tables", "priority": 40, "source": "open_next_actions"},
            ],
        }
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"list_report_tables", "get_component_identity_card", "summarize_table_column_values"},
            max_calls=3,
        )

        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("list_report_tables", plan["tool_calls"][0]["name"])
        self.assertEqual("open_next_actions", plan["tool_calls"][0]["source"])
        skipped_reasons = [item["reason"] for item in plan["skipped_actions"]]
        self.assertIn("missing_args", skipped_reasons)
        self.assertIn("placeholder_args", skipped_reasons)
        self.assertIn("duplicate", skipped_reasons)

    def test_quality_gate_open_next_actions_preserve_seeded_args(self):
        gate = build_final_answer_quality_gate(
            answer="无法判断，信息不足。",
            citations=[],
            evidence_nodes=[],
            task_ledger={
                "progress": {"open": 1, "in_progress": 1},
                "next_actions": [{
                    "type": "tool_call",
                    "tool": "batch_query_report_entities",
                    "args": {"queries": ["U1", "HQ11112042009"], "limit_per_query": 10},
                    "reason": "本地 playbook 已生成带参批量查询。",
                }],
            },
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_query_report_entities"},
            max_calls=1,
        )

        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("batch_query_report_entities", plan["tool_calls"][0]["name"])
        self.assertEqual(["U1", "HQ11112042009"], plan["tool_calls"][0]["args"]["queries"])

    def test_quality_gate_evidence_goal_repair_preserves_seeded_args(self):
        gate = build_final_answer_quality_gate(
            answer="信息不足。",
            citations=[],
            evidence_nodes=[],
            evidence_goal_contract={
                "status": "missing",
                "missing_evidence_types": ["compare_diff", "datasheet_chunk"],
                "repair_actions": [{
                    "type": "tool_call",
                    "tool": "batch_search_datasheet_chunks",
                    "args": {"queries": ["HQ11112042009"], "limit_per_query": 8},
                    "source": "missing_evidence_goal",
                    "priority": 12,
                }],
            },
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_search_datasheet_chunks"},
            max_calls=1,
        )

        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("batch_search_datasheet_chunks", plan["tool_calls"][0]["name"])
        self.assertEqual(["HQ11112042009"], plan["tool_calls"][0]["args"]["queries"])

    def test_evidence_goal_contract_tracks_specific_seed_targets(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U46", "HQ11112042009", "absolute"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[{
                "id": "ev-u46",
                "type": "component",
                "title": "U46 查询结果",
                "summary": "U46 位号已命中。",
            }],
        )

        self.assertEqual("partial", contract["target_status"])
        self.assertEqual(["U46"], [item["value"] for item in contract["covered_targets"]])
        self.assertEqual(["HQ11112042009"], [item["value"] for item in contract["missing_targets"]])
        self.assertEqual("batch_query_report_entities", contract["repair_actions"][0]["tool"])
        self.assertEqual(["HQ11112042009"], contract["repair_actions"][0]["args"]["queries"])

    def test_evidence_goal_contract_matches_targets_as_tokens(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U1"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[{
                "id": "ev-pu10",
                "type": "component",
                "title": "PU10 查询结果",
                "summary": "PU10 位号已命中。",
            }],
        )

        self.assertEqual("missing", contract["target_status"])
        self.assertEqual(["U1"], [item["value"] for item in contract["missing_targets"]])

    def test_quality_gate_repairs_missing_specific_target_coverage(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U46", "HQ11112042009"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[{"id": "ev-u46", "type": "component", "title": "U46 查询结果"}],
        )
        gate = build_final_answer_quality_gate(
            answer="已检查 U46。",
            citations=[{"id": "ev-u46", "valid": True}],
            evidence_nodes=[{"id": "ev-u46", "type": "component", "title": "U46 查询结果"}],
            evidence_goal_contract=contract,
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_query_report_entities"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("missing_target_coverage", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("batch_query_report_entities", plan["tool_calls"][0]["name"])
        self.assertEqual(["HQ11112042009"], plan["tool_calls"][0]["args"]["queries"])

    def test_quality_gate_repairs_missing_target_even_if_answer_mentions_it(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U46", "HQ11112042009"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[{"id": "ev-u46", "type": "component", "title": "U46 查询结果"}],
        )
        gate = build_final_answer_quality_gate(
            answer="已检查 U46 和 HQ11112042009。",
            citations=[{"id": "ev-u46", "valid": True}],
            evidence_nodes=[{"id": "ev-u46", "type": "component", "title": "U46 查询结果"}],
            evidence_goal_contract=contract,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("missing_target_coverage", reason_ids)

    def test_quality_gate_repairs_partial_evidence_goal(self):
        gate = build_final_answer_quality_gate(
            answer="已有 compare diff，可以判断规格书差异。",
            citations=[{"id": "ev-diff", "valid": True}],
            evidence_nodes=[{"id": "ev-diff", "type": "compare_diff"}],
            evidence_goal_contract={
                "status": "partial",
                "missing_evidence_types": ["datasheet_chunk"],
                "repair_actions": [{
                    "source": "missing_evidence_goal",
                    "tool": "batch_search_datasheet_chunks",
                    "args": {"queries": ["HQ100"]},
                }],
            },
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_search_datasheet_chunks"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("missing_evidence_goal", reason_ids)
        self.assertEqual("batch_search_datasheet_chunks", plan["tool_calls"][0]["name"])

    def test_quality_gate_repairs_incomplete_datasheet_connection_phase(self):
        plan = build_playbook_plan(
            question="请根据 datasheet 反查 U1 的 I2C 连接是否有问题。",
            capability_profiles=["connection_datasheet_review"],
            allowed_tools=[
                "batch_query_llm_topology_netlist",
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
            ],
            playbooks=REPORT_AGENT_PLAYBOOKS,
        ).to_dict()
        evidence_nodes = [{"id": "ev-edge", "type": "llm_topology_edge", "title": "U1 I2C edge"}]
        contract = build_evidence_goal_contract(playbook_plan=plan, evidence_nodes=evidence_nodes)
        gate = build_final_answer_quality_gate(
            answer="U1 I2C 连接看起来没有明显问题。",
            citations=[{"id": "ev-edge", "valid": True}],
            evidence_nodes=evidence_nodes,
            evidence_goal_contract=contract,
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={
                "batch_get_component_identity_cards",
                "batch_match_component_datasheets",
                "search_datasheet_parameters",
                "batch_search_datasheet_chunks",
            },
            max_calls=2,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("missing_connection_review_phase", reason_ids)
        selected_tools = [item["name"] for item in plan["tool_calls"]]
        self.assertIn("batch_get_component_identity_cards", selected_tools)
        self.assertIn("batch_match_component_datasheets", selected_tools)

    def test_quality_gate_allows_complete_datasheet_connection_phase(self):
        contract = {
            "status": "partial",
            "missing_evidence_types": ["source_trace"],
            "connection_review_phase_status": "satisfied",
            "missing_connection_review_phases": [],
            "connection_review_repair_actions": [],
        }
        gate = build_final_answer_quality_gate(
            answer="U1 I2C 连接需要人工复核电平余量。",
            citations=[
                {"id": "ev-edge", "valid": True},
                {"id": "ev-id", "valid": True},
                {"id": "ev-detail", "valid": True},
            ],
            evidence_nodes=[
                {"id": "ev-edge", "type": "llm_topology_edge", "title": "U1 I2C edge"},
                {"id": "ev-id", "type": "component_identity", "title": "U1 identity"},
                {"id": "ev-detail", "type": "datasheet_chunk", "title": "U1 VIH/VIL detail", "source": {"tool": "get_datasheet_chunk"}},
            ],
            evidence_goal_contract=contract,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertNotIn("missing_connection_review_phase", reason_ids)

    def test_quality_gate_repairs_answer_that_omits_covered_target(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U46", "U47"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[
                {"id": "ev-u46", "type": "component", "title": "U46 查询结果"},
                {"id": "ev-u47", "type": "component", "title": "U47 查询结果"},
            ],
        )
        gate = build_final_answer_quality_gate(
            answer="U46 已检查，未发现明显异常。",
            citations=[{"id": "ev-u46", "valid": True}],
            evidence_nodes=[
                {"id": "ev-u46", "type": "component", "title": "U46 查询结果"},
                {"id": "ev-u47", "type": "component", "title": "U47 查询结果"},
            ],
            evidence_goal_contract=contract,
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_query_report_entities"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("answer_missing_target_coverage", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("batch_query_report_entities", plan["tool_calls"][0]["name"])
        self.assertEqual(["U47"], plan["tool_calls"][0]["args"]["queries"])

    def test_quality_gate_answer_target_coverage_uses_token_boundaries(self):
        contract = {
            "status": "satisfied",
            "target_status": "satisfied",
            "covered_targets": [
                {"value": "U1", "normalized": "U1", "repair_tool": "batch_query_report_entities", "repair_arg_key": "queries"},
                {"value": "U2", "normalized": "U2", "repair_tool": "batch_query_report_entities", "repair_arg_key": "queries"},
            ],
            "missing_targets": [],
        }
        gate = build_final_answer_quality_gate(
            answer="PU10 和 U2 均已检查。",
            citations=[
                {"id": "ev-pu10", "valid": True},
                {"id": "ev-u2", "valid": True},
            ],
            evidence_nodes=[
                {"id": "ev-pu10", "type": "component", "title": "PU10 查询结果"},
                {"id": "ev-u2", "type": "component", "title": "U2 查询结果"},
            ],
            evidence_goal_contract=contract,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("answer_missing_target_coverage", reason_ids)

    def test_quality_gate_repairs_answer_that_mentions_target_without_citation(self):
        contract = build_evidence_goal_contract(
            playbook_plan={
                "selected_playbooks": [{
                    "id": "report_entity_batch_lookup",
                    "title": "报告实体批量查询",
                    "required_evidence": ["component"],
                }],
                "seeded_tool_calls": [{
                    "name": "batch_query_report_entities",
                    "args": {"queries": ["U46", "U47"], "limit_per_query": 10},
                }],
            },
            evidence_nodes=[
                {"id": "ev-u46", "type": "component", "title": "U46 查询结果"},
                {"id": "ev-u47", "type": "component", "title": "U47 查询结果"},
            ],
        )
        gate = build_final_answer_quality_gate(
            answer="U46 和 U47 均已检查。",
            citations=[{"id": "ev-u46", "valid": True}],
            evidence_nodes=[
                {"id": "ev-u46", "type": "component", "title": "U46 查询结果"},
                {"id": "ev-u47", "type": "component", "title": "U47 查询结果"},
            ],
            evidence_goal_contract=contract,
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"batch_query_report_entities"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("answer_target_citation_missing", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("batch_query_report_entities", plan["tool_calls"][0]["name"])
        self.assertEqual(["U47"], plan["tool_calls"][0]["args"]["queries"])

    def test_quality_gate_requires_detail_for_cited_datasheet_search_hit(self):
        gate = build_final_answer_quality_gate(
            answer="HQ100 推荐工作条件是 3.3V。",
            citations=[{"id": "ev-ds-hit", "valid": True}],
            evidence_nodes=[{
                "id": "ev-ds-hit",
                "type": "datasheet_chunk",
                "title": "HQ100 datasheet 搜索命中",
                "source": {"tool": "search_datasheet_chunks"},
                "detail_tool": {"name": "get_datasheet_chunk", "args": {"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000}},
            }],
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"get_datasheet_chunk"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("citation_detail_required", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("get_datasheet_chunk", plan["tool_calls"][0]["name"])
        self.assertEqual({"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000}, plan["tool_calls"][0]["args"])

    def test_quality_gate_requires_detail_for_memory_recalled_datasheet_hit(self):
        gate = build_final_answer_quality_gate(
            answer="HQ100 推荐工作条件是 3.3V。",
            citations=[{"id": "ev-memory-ds-hit", "valid": True}],
            evidence_nodes=[{
                "id": "ev-memory-ds-hit",
                "type": "datasheet_chunk",
                "title": "项目记忆中的 HQ100 datasheet 搜索命中",
                "source": {"tool": "get_project_memory_evidence"},
                "locator": {"original_type": "datasheet_chunk"},
                "detail_tool": {"name": "get_datasheet_chunk", "args": {"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000}},
            }],
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"get_datasheet_chunk"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("citation_detail_required", reason_ids)
        self.assertIn("quantitative_claim_detail_required", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("get_datasheet_chunk", plan["tool_calls"][0]["name"])

    def test_quality_gate_requires_detail_for_quantitative_spec_claim(self):
        gate = build_final_answer_quality_gate(
            answer="U1 推荐工作电压为 3.3V。",
            citations=[{"id": "ev-u1", "valid": True}],
            evidence_nodes=[
                {"id": "ev-u1", "type": "component", "title": "U1 身份卡"},
                {
                    "id": "ev-ds-hit",
                    "type": "datasheet_chunk",
                    "title": "HQ100 datasheet 搜索命中",
                    "source": {"tool": "search_datasheet_chunks"},
                    "detail_tool": {"name": "get_datasheet_chunk", "args": {"doc_id": 1, "chunk_id": "p1-c1", "max_chars": 4000}},
                },
            ],
        )
        plan = build_quality_repair_tool_calls(
            gate,
            allowed_tools={"get_datasheet_chunk"},
            max_calls=1,
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertIn("quantitative_claim_detail_required", reason_ids)
        self.assertEqual(1, plan["selected_tool_call_count"])
        self.assertEqual("get_datasheet_chunk", plan["tool_calls"][0]["name"])

    def test_quality_gate_allows_quantitative_claim_with_detail_datasheet_citation(self):
        gate = build_final_answer_quality_gate(
            answer="U1 推荐工作电压为 3.3V。",
            citations=[{"id": "ev-ds-detail", "valid": True}],
            evidence_nodes=[{
                "id": "ev-ds-detail",
                "type": "datasheet_chunk",
                "title": "HQ100 datasheet 原文 chunk",
                "source": {"tool": "get_datasheet_chunk"},
            }],
        )

        reason_ids = {item["id"] for item in gate["reasons"]}
        self.assertNotIn("quantitative_claim_detail_required", reason_ids)
        self.assertNotIn("citation_detail_required", reason_ids)

    def test_runtime_session_state_compacts_memory_and_pending_questions(self):
        runtime_state = build_runtime_state(
            goal="做 DFMEA 准备",
            capability_plan=[{"id": "dfmea_prep", "title": "DFMEA 准备"}],
            observations=[{"tool": "identity", "summary": "U1 缺规格", "evidence_node_ids": ["ev-u1"]}],
            project_context={
                "answers": [{"question_id": "q1", "answer": "U1 是 FPGA"}],
                "pending_questions": [{"question_id": "q2", "question": "请补 U2 规格", "missing_fields": ["spec"]}],
            },
        )
        session = build_agent_session_state(
            agent_run_id="run-1",
            goal="做 DFMEA 准备",
            runtime_state=runtime_state,
            project_context={
                "answers": [{"question_id": "q1", "answer": "U1 是 FPGA"}],
                "pending_questions": [{"question_id": "q2", "question": "请补 U2 规格", "missing_fields": ["spec"]}],
            },
            observations=[{"evidence_node_ids": ["ev-u1", "ev-u2"]}],
        )

        self.assertEqual("pstx-agent-runtime/v1", session["protocol_version"])
        self.assertEqual("run-1", session["agent_run_id"])
        self.assertEqual("agent-task-ledger/v1", session["task_ledger"]["version"])
        self.assertIn("ev-u1", session["recent_evidence_ids"])
        self.assertEqual("q2", session["pending_questions"][0]["question_id"])
        self.assertEqual(1, session["context_answer_count"])

    def test_guidance_loader_reads_agents_and_ignores_archive(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".git").mkdir()
            (root / "AGENTS.md").write_text(
                "# AGENTS\n\n## 硬边界\n\n- Harness 工具必须只读。\n\n## 运行入口\n\n- python pstx_web.py\n",
                encoding="utf-8",
            )
            (root / "docs" / "archive").mkdir(parents=True)
            (root / "docs" / "archive" / "AGENTS.md").write_text("old", encoding="utf-8")

            guidance = load_project_guidance(root)

            self.assertEqual(1, guidance["source_count"])
            self.assertIn("Harness 工具必须只读", "\n".join(guidance["hard_boundaries"]))
            self.assertIn("python pstx_web.py", "\n".join(guidance["quick_start"]))

    def test_skill_registry_selects_matching_skill(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".git").mkdir()
            skill_dir = root / "harness_skills" / "demo"
            skill_dir.mkdir(parents=True)
            (skill_dir / "SKILL.md").write_text(
                "---\n"
                "name: demo\n"
                "description: Demo topology skill\n"
                "triggers: [拓扑, topology]\n"
                "capability_profiles: [chip_topology]\n"
                "playbooks: [chip_level_topology]\n"
                "allowed_tools: [summarize_llm_topology_netlist]\n"
                "---\n"
                "Use topology tools first.\n",
                encoding="utf-8",
            )

            skills = load_harness_skills(root)
            selected = select_harness_skills(
                question="看一下 U46 拓扑",
                capability_profiles=[],
                playbook_plan={},
                root=root,
            )

            self.assertEqual(1, len(skills))
            self.assertEqual("demo", selected["selected_skills"][0]["id"])
            self.assertIn("Use topology", selected["selected_skills"][0]["body"])

    def test_builtin_datasheet_key_info_skill_is_selectable(self):
        repo_root = Path(__file__).resolve().parents[1]

        skills = load_harness_skills(repo_root)
        selected = select_harness_skills(
            question="请用 MinerU 读取 64144 datasheet 的关键电源和时序参数",
            capability_profiles=["datasheet_qa"],
            playbook_plan={},
            root=repo_root,
        )
        skill_ids = {skill.id for skill in skills}
        selected_ids = {skill["id"] for skill in selected["selected_skills"]}
        selected_text = "\n".join(str(skill.get("body") or "") for skill in selected["selected_skills"])
        skill_body = next(skill.body for skill in skills if skill.id == "datasheet-key-info")

        self.assertIn("datasheet-key-info", skill_ids)
        self.assertIn("datasheet-key-info", selected_ids)
        self.assertIn("MinerU", selected_text)
        self.assertIn("已确认的 datasheet 事实", selected_text)
        self.assertIn("已确认的 datasheet 事实", skill_body)
        self.assertIn("get_datasheet_parameter", skill_body)

    def test_effort_policy_wraps_premature_refusal_retry(self):
        state = build_effort_policy_state(
            step_type="final_answer",
            answer="无法回答，信息不足。",
            tool_call_count=0,
            max_tool_calls=4,
            playbook_plan={"recommended_first_tools": ["list_report_tables"]},
            tool_result_contracts=[],
            task_ledger={"progress": {"open": 1}, "next_actions": [{"type": "tool_call", "tool": "list_report_tables"}]},
            evidence_node_count=0,
            citation_count=0,
            retry_count=0,
        )

        self.assertTrue(state["retry_available"])
        self.assertIn("list_report_tables", state["recommended_tools"])
        self.assertIn("不要", state["retry_note"])

    def test_markdown_task_memory_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_env = os.environ.get("PSTX_AGENT_MEMORY_DIR")
            os.environ["PSTX_AGENT_MEMORY_DIR"] = str(Path(tmp) / "memory")
            try:
                written = write_task_memory("run-1", {
                    "agent_run_id": "agent-1",
                    "profile": "dfmea_prep",
                    "status": "waiting_for_user",
                    "answer": "需要补充 U1 规格。",
                    "guidance_summary": {"source_count": 1, "hard_boundaries": ["只读"]},
                    "selected_skills": {"selected_skills": [{"id": "dfmea-prep", "description": "DFMEA"}]},
                    "needs_user_input": {"questions": [{"question_id": "q1", "question": "U1 规格是什么？"}]},
                    "citations": [{"id": "ev-u1"}],
                    "trace_summary": {"tool_call_count": 1, "evidence_node_count": 1},
                    "runtime_state": {"goal": "做 DFMEA", "task_ledger": {"next_actions": [{"tool": "get_component_identity_card", "args": {"refdes": "U1"}}]}},
                })
                loaded = read_task_memory("run-1")
            finally:
                if old_env is None:
                    os.environ.pop("PSTX_AGENT_MEMORY_DIR", None)
                else:
                    os.environ["PSTX_AGENT_MEMORY_DIR"] = old_env

            self.assertTrue(written["found"])
            self.assertTrue(loaded["found"])
            self.assertIn("ev-u1", loaded["summary"])
            self.assertIn("U1 规格是什么", loaded["summary"])

    def test_agentic_envelope_combines_guidance_skills_and_memory(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".git").mkdir()
            (root / "AGENTS.md").write_text("## 硬边界\n- 只读工具。\n", encoding="utf-8")
            skill_dir = root / "harness_skills" / "feishu"
            skill_dir.mkdir(parents=True)
            (skill_dir / "SKILL.md").write_text(
                "---\nname: feishu\ndescription: Feishu skill\ntriggers: [飞书]\nallowed_tools: [search_feishu_cache_rows]\n---\nUse cache evidence.\n",
                encoding="utf-8",
            )

            envelope = build_agentic_envelope(
                run_id="run-2",
                question="查询飞书 HQ 料号",
                capability_profiles=[],
                playbook_plan={},
                root=root,
            )

            self.assertEqual("pstx-agent-kernel/v2", envelope["version"])
            self.assertEqual(1, envelope["guidance_summary"]["source_count"])
            self.assertEqual("feishu", envelope["selected_skills"]["selected_skills"][0]["id"])
            self.assertFalse(envelope["task_memory_summary"]["found"])


if __name__ == "__main__":
    unittest.main()
