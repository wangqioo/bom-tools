# -*- coding: utf-8 -*-
"""Planning and prompt construction for the report harness agent."""

from __future__ import annotations

import json
from typing import Callable, List, Optional, Sequence

from pstx_agent_runtime import (
    build_agent_protocol_brief,
    build_capability_plan_items,
    build_playbook_plan,
    dedupe_profile_ids as runtime_dedupe_profile_ids,
    infer_capability_profiles,
    allowed_tool_names as runtime_allowed_tool_names,
    REPORT_AGENT_PLAYBOOKS,
)
from pstx_harness.report_agent_config import (
    HARNESS_AGENT_CAPABILITY_RULES,
    HARNESS_AGENT_PROFILES,
    HARNESS_SKILL_TOOL_NAMES,
    PROJECT_MEMORY_TOOL_NAMES,
    HarnessAgentRequest,
    profile_config,
)
from pstx_harness.report_tools import HarnessToolRegistry


def _preview_text(value, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[:limit] + "..."


def question_text(request: HarnessAgentRequest) -> str:
    return f"{request.question or ''} {profile_config(request.profile).get('default_question') or ''}".lower()


def dedupe_profile_ids(profile_ids: Sequence[str]) -> List[str]:
    return runtime_dedupe_profile_ids(profile_ids, HARNESS_AGENT_PROFILES)


def memory_planning_context(project_context: Optional[dict]) -> str:
    if not isinstance(project_context, dict):
        return ""
    parts: List[str] = []

    def _extend_from(value, *, limit: int = 6):
        if isinstance(value, str):
            if value.strip():
                parts.append(value.strip())
            return
        if isinstance(value, list):
            for item in value[:limit]:
                if isinstance(item, dict):
                    text = " ".join(
                        str(item.get(key) or "")
                        for key in ("title", "question", "summary", "name", "next_intent", "goal")
                    ).strip()
                    if text:
                        parts.append(text)
                elif str(item).strip():
                    parts.append(str(item).strip())

    memory = project_context.get("session_memory_summary")
    if isinstance(memory, dict):
        for key in ("goal", "facts", "open_questions", "open_items", "next_actions"):
            _extend_from(memory.get(key))
    _extend_from(project_context.get("evidence_memory_cards") or [], limit=8)
    for pack_key in ("active_continuation_pack", "latest_continuation_pack"):
        pack = project_context.get(pack_key)
        if not isinstance(pack, dict):
            continue
        for key in ("goal", "next_intent", "continuation_brief", "pending_questions", "open_ledger_items", "suggested_tool_calls"):
            _extend_from(pack.get(key))
    return _preview_text(" ".join(parts), 1600)


def infer_harness_capability_profiles(request: HarnessAgentRequest, planning_context: str = "") -> List[str]:
    return infer_capability_profiles(
        requested_profile=request.profile,
        question=f"{request.question or ''} {planning_context or ''}",
        profiles=HARNESS_AGENT_PROFILES,
        default_profile="quick_scan",
        rules=HARNESS_AGENT_CAPABILITY_RULES,
        quick_profile="quick_scan",
    )


def selected_profile_ids(request: HarnessAgentRequest, planning_context: str = "") -> List[str]:
    if request.profile == "auto":
        return infer_harness_capability_profiles(request, planning_context=planning_context)
    return dedupe_profile_ids([request.profile])


def capability_plan(request: HarnessAgentRequest, planning_context: str = "") -> List[dict]:
    return build_capability_plan_items(selected_profile_ids(request, planning_context=planning_context), HARNESS_AGENT_PROFILES)


def allowed_tool_names(request: HarnessAgentRequest, registry: HarnessToolRegistry, planning_context: str = "") -> List[str]:
    names = runtime_allowed_tool_names(
        profile_ids=selected_profile_ids(request, planning_context=planning_context),
        profiles=HARNESS_AGENT_PROFILES,
        registry_tools=registry.list_tools(),
    )
    available = {str(tool.get("name") or "") for tool in registry.list_tools()}
    for tool_name in (*PROJECT_MEMORY_TOOL_NAMES, *HARNESS_SKILL_TOOL_NAMES):
        if tool_name in available and tool_name not in names:
            names.append(tool_name)
    return names


def filtered_tool_list(request: HarnessAgentRequest, registry: HarnessToolRegistry, planning_context: str = "") -> List[dict]:
    allowed = set(allowed_tool_names(request, registry, planning_context=planning_context))
    return [dict(tool) for tool in registry.list_tools() if tool.get("name") in allowed]


def playbook_plan(request: HarnessAgentRequest, registry: HarnessToolRegistry, planning_context: str = "") -> dict:
    return build_playbook_plan(
        question=f"{request.question or ''} {planning_context or ''}",
        capability_profiles=selected_profile_ids(request, planning_context=planning_context),
        allowed_tools=allowed_tool_names(request, registry, planning_context=planning_context),
        playbooks=REPORT_AGENT_PLAYBOOKS,
    ).to_dict()


def build_agent_prompt(request: HarnessAgentRequest,
                       report: dict,
                       registry: HarnessToolRegistry,
                       observations: List[dict],
                       *,
                       playbook_plan_payload: Optional[dict] = None,
                       project_context: Optional[dict] = None,
                       planning_context: str = "",
                       context_budget: Optional[dict] = None,
                       runtime_state: Optional[dict] = None,
                       session_state: Optional[dict] = None,
                       agentic_context: Optional[dict] = None,
                       retry_note: str = "",
                       compact_project_context: Optional[Callable[[dict], dict]] = None) -> str:
    compact_context = compact_project_context(project_context or {}) if compact_project_context else project_context or {}
    payload = {
        "project_name": report.get("project_name"),
        "profile": {
            "id": request.profile,
            "title": profile_config(request.profile).get("title"),
            "description": profile_config(request.profile).get("description"),
        },
        "continue_agent_run_id": request.continue_agent_run_id,
        "planning_context": planning_context,
        "capability_plan": capability_plan(request, planning_context=planning_context),
        "question": request.question,
        "limits": {
            "max_steps": request.max_steps,
            "max_tool_calls": request.max_tool_calls,
            "max_rows_per_table": request.max_rows_per_table,
        },
        "tools": filtered_tool_list(request, registry, planning_context=planning_context),
        "playbook_plan": playbook_plan_payload or {},
        "observations": observations,
        "context_budget": context_budget or {},
        "observation_bundle": (context_budget or {}).get("observation_bundle", {}),
        "runtime_state": runtime_state or {},
        "session_state": session_state or {},
        "agentic_context": agentic_context or {},
        "guidance_summary": (agentic_context or {}).get("guidance_summary", {}),
        "selected_skills": (agentic_context or {}).get("selected_skills", {}),
        "task_memory_summary": (agentic_context or {}).get("task_memory_summary", {}),
        "effort_policy": (agentic_context or {}).get("effort_policy", {}),
        "task_ledger": (runtime_state or {}).get("task_ledger", {}),
        "runtime_protocol": build_agent_protocol_brief(allow_batch_tools=True, allow_task_dispatch=True),
        "project_context": compact_context,
    }
    retry = f"\n上一次输出无效：{retry_note}\n请严格改为合法 JSON。" if retry_note else ""
    return (
        "你是 PSTX 原理图审查的受控 agent。\n"
        f"{build_agent_protocol_brief(allow_batch_tools=True, allow_task_dispatch=True)}\n"
        "你不能直接执行工具，只能请求本地 harness 执行白名单只读工具。\n"
        "每轮只能输出一个 JSON 对象，五选一：\n"
        '1. {"tool_call":{"name":"工具名","args":{...},"reason":"为什么需要这个工具"}}\n'
        '2. {"tool_batch_call":[{"name":"工具名","args":{...},"reason":"为什么需要这个工具"}]}，一轮最多 4 个工具。\n'
        '3. {"needs_user_input":{"reason":"为什么必须由用户补充","questions":[{"question_id":"q-1","question":"要问用户的问题","applies_to":{"refdes":"U1","field":"spec"},"missing_fields":["spec"]}],"missing_fields":[...],"related_evidence_ids":["ev-..."]}}\n'
        '4. {"dispatch_tasks":[{"task_id":"task-1","title":"子任务标题","profile":"auto|datasheet_qa|cadence_pages|...","question":"给 child run 的完整问题","reason":"为什么可独立后台执行","depends_on":[],"expected_outputs":[]}],"reason":"为什么这些分支适合分发"}。\n'
        '5. {"final_answer":"最终回答","confidence":"high|medium|low","citations":[{"id":"ev-...","note":"引用理由"}],"proposed_actions":[...],"scratch_files":[{"filename":"临时笔记.md","content":"临时中间产物","content_type":"text/markdown"}]}。\n'
        "只有当多个长耗时分支互相独立、适合后台 durable child runs 时才使用 dispatch_tasks；"
        "不要为了普通单轮查询或一个 detail 工具调用而分发。"
        "dispatch_tasks 只是任务声明，不授权新工具，不替代 evidence；子任务仍必须通过本地只读白名单工具取证。"
        "如果你生成了临时分析笔记、候选清单或中间 JSON，可在 final_answer 同一个 JSON 对象里附带 scratch_files；"
        "本地 runtime 会写入 agent_workspace scratch 临时区。scratch_files 只适合临时文本，不要放原始大表、PDF 全文、CSA/CSV 全文或凭据。"
        "当用户问题包含多个位号、网络、料号、页码、表格或规格书候选时，优先调用 batch_* 工具一次性取证；"
        "只有需要单条详情时再调用 get_* detail 工具。"
        "不要过早拒绝用户问题，也不要在未尝试本地只读工具前说无法回答；"
        "只要仍有白名单工具、推荐路线或 detail/aggregation 工具可用，就应继续取证或缩小问题范围。"
        "如果 observations 中出现工具调用错误，说明本地 harness 已安全拒绝该调用；"
        "不要重复同一个失败工具和参数，应改用其它白名单工具、修正参数，或在确实缺少用户信息时 needs_user_input。"
        "本地 playbook_plan 给出了推荐取证路线、推荐首批工具、证据目标和 anti-pattern；"
        "如果 playbook_plan.recommended_first_tools 与白名单工具可用，应优先沿该路线取证。"
        "如果 playbook_plan.seeded_tool_calls 非空，说明本地 runtime 已从用户问题中提取实体并准备好安全参数；"
        "runtime 可能已在第一轮模型调用前执行少量安全预取，请先检查 observations，"
        "不要重复同一工具参数；若仍缺证据，再使用剩余带参 tool_call/tool_batch_call。"
        "runtime_state.task_ledger 是本地任务账本，记录 completed/in_progress/pending/blocked 项、证据绑定和 next_actions；"
        "请像 Codex/Claude Code 一样持续推进任务账本：优先处理 in_progress 和 next_actions，"
        "如果仍有安全可用的推荐工具，不要过早 final_answer。"
        "runtime_state.evidence_goal_contract 是本地证据目标契约；"
        "如果它的 status 是 missing 或 partial，说明当前 playbook 还缺关键 evidence 类型，应优先调用其中推荐工具补齐。"
        "agentic_context.guidance_summary 是从 AGENTS/Agent/CLAUDE 项目说明中提取的边界和阅读导航；"
        "必须遵守其中硬边界，但不能把文档说明当作项目事实证据。"
        "agentic_context.selected_skills 是本轮自动命中的 Harness Skill；"
        "它们描述当前任务的推荐工具、输出约束和业务打法，优先级低于白名单和 evidence，但高于自由猜测。"
        "如果初始 selected_skills 不足，且白名单工具中有 list_harness_skills、select_harness_skills 或 get_harness_skill，"
        "可以主动调用这些 skill 工具读取更多打法说明；skill 仍只是指导，不会授权新工具，也不能替代项目 evidence。"
        "agentic_context.task_memory_summary 是当前 run_id 的 Markdown 任务摘要；"
        "它用于恢复多轮任务目标、开放问题和 evidence id，不替代本轮工具证据。"
        "agentic_context.effort_policy 描述本地抗放弃策略；如果 retry_available 或 has_safe_next_step 为真，"
        "不要直接用低努力话术结束，应沿推荐工具继续或结构化追问。"
        "project_context.active_continuation_pack 是 continue_agent_run_id 对应上一轮的交接包；如果存在，"
        "必须优先依据它的 next_intent、pending_questions、open_ledger_items、suggested_tool_calls 和 evidence_ids 继续推进。"
        "project_context.latest_continuation_pack 只作为最近上下文参考；当 active_continuation_pack 与 latest 不一致时，以 active 为准。"
        "continuation_pack 是压缩交接摘要，不替代原始 evidence；高风险或定量结论仍需通过 detail_tool/aggregation_tool 或白名单工具回拉证据。"
        "project_context.session_memory_summary 是当前 run_id 的滚动任务记忆，包含多轮事实、未解决问题、开放任务、下一步动作和证据 id；"
        "它用于保持连续性，但同样不能替代 evidence，本轮下结论前仍需确认相关证据是否足够。"
        "project_context.evidence_memory_cards 是当前项目会话内沉淀的证据卡片索引；"
        "如果上一轮已找到相关 evidence，应优先用 list_project_memory_evidence/get_project_memory_evidence 复核卡片来源、locator 和 detail_tool，再决定是否重新查询。"
        "planning_context 是本地从项目记忆/continuation pack 压缩出的能力路由提示；它只帮助选择工具范围，不能作为最终事实依据。"
        "每个 observation 的 tool_result_contract 会声明 completeness、recommended_next_tools、detail_tool、aggregation_tool；"
        "当 completeness 是 preview、partial 或 truncated 时，不得把当前 preview 当完整事实下最终统计结论。"
        "每个 observation 还包含 evidence_layers：summary_layer 只用于快速规划，evidence_card_layer 提供 evidence id、来源、定位和 detail_tool，"
        "raw_layer 表示完整工具结果保留在本地 trace/store 且默认不进入模型上下文。"
        "当用户要求追溯到原始文件、底层证据、PSTX/Cadence 文本，或需要核对报告结论来源时，"
        "优先使用 trace_project_source，并用位号、网络、页码或 table_id+row_index 定位 line-number excerpt；"
        "只有需要更大上下文窗口时再沿 detail_tool 调用 read_project_text。"
        "对高风险、不确定、需要定量确认或会影响结论的项，如果 evidence card 或 tool_result_contract 提供 detail_tool/recommended_next_tools，"
        "必须先继续读取原始详情或聚合证据，不能只基于摘要层下结论。"
        "当问题是统计表格某列唯一值、页码总数、覆盖范围、top values，或 observations 里表格 count 很大且 result_preview 被截断时，"
        "优先调用 summarize_table_column_values，不要尝试用 get_table_rows 读取完整大表。"
        "但当用户问“原理图一共有多少页/总页数/最后一页”时，必须优先调用 summarize_schematic_page_count；"
        "page_rows 只表示有记录或有元件的页码，不能代表原理图总页数，尤其不能覆盖尾部空白页。"
        "最终回答必须优先引用 observations 中的 evidence_nodes.id；如果缺失信息会影响判断，优先输出 needs_user_input，不要硬猜。"
        "如果 context_budget.truncated 为 true，说明 observations 只包含压缩摘要/preview，不能当作完整数据；"
        "需要更多细节时继续调用白名单工具按表格、实体、row_id 或 evidence id 读取。"
        "不要输出 Markdown，不要输出代码块。"
        f"{retry}\n输入：\n{json.dumps(payload, ensure_ascii=False, indent=2)}"
    )
