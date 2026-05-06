# -*- coding: utf-8 -*-
"""Planning and prompt construction for the compare agent."""

from __future__ import annotations

import json
from typing import List, Optional, Sequence

from pstx_agent_runtime import (
    COMPARE_AGENT_PLAYBOOKS,
    build_agent_protocol_brief,
    build_capability_plan_items,
    build_playbook_plan,
    dedupe_profile_ids as runtime_dedupe_profile_ids,
    infer_capability_profiles,
    allowed_tool_names as runtime_allowed_tool_names,
)
from pstx_harness.compare_agent_config import (
    COMPARE_AGENT_CAPABILITY_RULES,
    COMPARE_AGENT_PROFILES,
    CompareAgentRequest,
    profile_config,
)
from pstx_harness.report_tools import HarnessToolRegistry
from pstx_harness.skill_tools import HARNESS_SKILL_TOOL_NAMES


def question_text(request: CompareAgentRequest) -> str:
    return f"{request.question or ''} {profile_config(request.profile).get('default_question') or ''}".lower()


def dedupe_profile_ids(profile_ids: Sequence[str]) -> List[str]:
    return runtime_dedupe_profile_ids(profile_ids, COMPARE_AGENT_PROFILES)


def infer_compare_capability_profiles(request: CompareAgentRequest) -> List[str]:
    return infer_capability_profiles(
        requested_profile=request.profile,
        question=request.question,
        profiles=COMPARE_AGENT_PROFILES,
        default_profile="compare_quick_scan",
        rules=COMPARE_AGENT_CAPABILITY_RULES,
        quick_profile="compare_quick_scan",
    )


def selected_profile_ids(request: CompareAgentRequest) -> List[str]:
    if request.profile == "auto":
        return infer_compare_capability_profiles(request)
    return dedupe_profile_ids([request.profile])


def capability_plan(request: CompareAgentRequest) -> List[dict]:
    return build_capability_plan_items(selected_profile_ids(request), COMPARE_AGENT_PROFILES)


def allowed_tool_names(request: CompareAgentRequest, registry: HarnessToolRegistry) -> List[str]:
    names = runtime_allowed_tool_names(
        profile_ids=selected_profile_ids(request),
        profiles=COMPARE_AGENT_PROFILES,
        registry_tools=registry.list_tools(),
    )
    available = {str(tool.get("name") or "") for tool in registry.list_tools()}
    for tool_name in HARNESS_SKILL_TOOL_NAMES:
        if tool_name in available and tool_name not in names:
            names.append(tool_name)
    return names


def filtered_tool_list(request: CompareAgentRequest, registry: HarnessToolRegistry) -> List[dict]:
    allowed = set(allowed_tool_names(request, registry))
    return [dict(tool) for tool in registry.list_tools() if tool.get("name") in allowed]


def playbook_plan(request: CompareAgentRequest, registry: HarnessToolRegistry) -> dict:
    return build_playbook_plan(
        question=request.question,
        capability_profiles=selected_profile_ids(request),
        allowed_tools=allowed_tool_names(request, registry),
        playbooks=COMPARE_AGENT_PLAYBOOKS,
    ).to_dict()


def build_agent_prompt(request: CompareAgentRequest,
                       compare_payload: dict,
                       registry: HarnessToolRegistry,
                       observations: List[dict],
                       *,
                       playbook_plan: Optional[dict] = None,
                       context_budget: Optional[dict] = None,
                       runtime_state: Optional[dict] = None,
                       session_state: Optional[dict] = None,
                       agentic_context: Optional[dict] = None,
                       retry_note: str = "") -> str:
    payload = {
        "left": compare_payload.get("left", {}),
        "right": compare_payload.get("right", {}),
        "diff_totals": compare_payload.get("diff_totals", {}),
        "profile": {
            "id": request.profile,
            "title": profile_config(request.profile).get("title"),
            "description": profile_config(request.profile).get("description"),
        },
        "capability_plan": capability_plan(request),
        "question": request.question,
        "limits": {
            "max_steps": request.max_steps,
            "max_tool_calls": request.max_tool_calls,
            "detail_limit": request.detail_limit,
        },
        "tools": filtered_tool_list(request, registry),
        "playbook_plan": playbook_plan or {},
        "observations": observations,
        "context_budget": context_budget or {},
        "runtime_state": runtime_state or {},
        "session_state": session_state or {},
        "agentic_context": agentic_context or {},
        "guidance_summary": (agentic_context or {}).get("guidance_summary", {}),
        "selected_skills": (agentic_context or {}).get("selected_skills", {}),
        "task_memory_summary": (agentic_context or {}).get("task_memory_summary", {}),
        "effort_policy": (agentic_context or {}).get("effort_policy", {}),
        "task_ledger": (runtime_state or {}).get("task_ledger", {}),
        "runtime_protocol": build_agent_protocol_brief(allow_batch_tools=True, allow_task_dispatch=True),
    }
    retry = f"\n上一次输出无效：{retry_note}\n请严格改为合法 JSON。" if retry_note else ""
    return (
        "你是 PSTX 项目对比审查的受控 agent。\n"
        f"{build_agent_protocol_brief(allow_batch_tools=True, allow_task_dispatch=True)}\n"
        "你不能直接执行工具，只能请求本地 compare harness 执行白名单只读工具。\n"
        "每轮只能输出一个 JSON 对象，四选一：\n"
        '1. {"tool_call":{"name":"工具名","args":{...},"reason":"为什么需要这个工具"}}\n'
        '2. {"tool_batch_call":[{"name":"工具名","args":{...},"reason":"为什么需要这个工具"}]}，一轮最多 4 个工具。\n'
        '3. {"dispatch_tasks":[{"task_id":"task-1","title":"子任务标题","profile":"auto|compare_datasheet_qa|compare_cadence_pages|...","question":"给 child run 的完整问题","reason":"为什么可独立后台执行","depends_on":[],"expected_outputs":[]}],"reason":"为什么这些 A/B 分支适合分发"}。\n'
        '4. {"final_answer":"最终回答","confidence":"high|medium|low","citations":[{"id":"ev-...","note":"引用理由"}],"proposed_actions":[...],"scratch_files":[{"filename":"compare-notes.md","content":"临时中间产物","content_type":"text/markdown"}]}。\n'
        "只有当多个 A/B 长耗时分支互相独立、适合后台 durable child runs 时才使用 dispatch_tasks；"
        "普通单轮 compare 查询或一个 detail 工具调用不要分发。"
        "dispatch_tasks 只是任务声明，不授权新工具，不替代 A/B evidence；子任务仍必须通过本地只读白名单工具取证。"
        "如果你生成了临时对比笔记、候选清单或中间 JSON，可在 final_answer 同一个 JSON 对象里附带 scratch_files；"
        "本地 runtime 会写入 agent_workspace scratch 临时区。scratch_files 只适合临时文本，不要放原始大表、PDF 全文、CSA/CSV 全文或凭据。"
        "当用户问题包含多个位号、网络、HQ 料号、PI、页码或 Cadence 对象时，优先使用 batch_* compare 工具一次性取证；"
        "只有需要单条 diff row 或对象详情时再调用 get_* detail 工具。"
        "不要过早拒绝用户问题，也不要在未尝试本地只读工具前说无法回答；"
        "只要仍有白名单工具、推荐路线或 detail/aggregation 工具可用，就应继续取证或缩小问题范围。"
        "如果 observations 中出现工具调用错误，说明本地 compare harness 已安全拒绝该调用；"
        "不要重复同一个失败工具和参数，应改用其它白名单 compare 工具、修正参数，或缩小查询范围。"
        "本地 playbook_plan 给出了推荐取证路线、推荐首批工具、证据目标和 anti-pattern；"
        "如果 playbook_plan.recommended_first_tools 与白名单工具可用，应优先沿该路线取证。"
        "如果 playbook_plan.seeded_tool_calls 非空，说明本地 runtime 已从用户问题中提取对比关键词并准备好安全参数；"
        "runtime 可能已在第一轮模型调用前执行少量安全预取，请先检查 observations，"
        "不要重复同一工具参数；若仍缺证据，再使用剩余带参 compare tool_call/tool_batch_call。"
        "runtime_state.task_ledger 是本地任务账本，记录 completed/in_progress/pending/blocked 项、证据绑定和 next_actions；"
        "请像 Codex/Claude Code 一样持续推进任务账本：如果 next_actions 里有 compare detail、batch 或 aggregation 工具，应优先继续取证。"
        "runtime_state.evidence_goal_contract 是本地证据目标契约；"
        "如果它的 status 是 missing 或 partial，说明当前 compare playbook 还缺关键 evidence 类型，应优先调用其中推荐工具补齐。"
        "agentic_context.guidance_summary 是从 AGENTS/Agent/CLAUDE 项目说明中提取的边界和阅读导航；"
        "必须遵守其中硬边界，但不能把文档说明当作 A/B 对比事实证据。"
        "agentic_context.selected_skills 是本轮自动命中的 Harness Skill；"
        "它们描述当前任务的推荐工具、输出约束和业务打法，优先级低于白名单和 compare evidence。"
        "如果初始 selected_skills 不足，且白名单工具中有 list_harness_skills、select_harness_skills 或 get_harness_skill，"
        "可以主动调用这些 skill 工具读取更多 A/B 取证打法；skill 仍只是指导，不会授权新工具，也不能替代 A/B evidence。"
        "agentic_context.task_memory_summary 是当前对比 run 的 Markdown 任务摘要；"
        "它用于恢复多轮任务目标、开放问题和 evidence id，不替代本轮工具证据。"
        "agentic_context.effort_policy 描述本地抗放弃策略；如果 retry_available 或 has_safe_next_step 为真，"
        "不要直接用低努力话术结束，应沿推荐工具继续或结构化缩小问题。"
        "每个 observation 的 tool_result_contract 会声明 completeness、recommended_next_tools、detail_tool、aggregation_tool；"
        "当 completeness 是 preview、partial 或 truncated 时，不得把当前 preview 当完整 A/B 差异事实。"
        "每个 observation 还包含 evidence_layers：summary_layer 用于规划，evidence_card_layer 保留 evidence id、来源、定位、字段和 detail_tool，"
        "raw_layer 表示完整工具结果留在本地 trace/store。"
        "如果差异属于高风险、不确定或需要精确定位，必须沿 evidence card 的 detail_tool 或 recommended_next_tools 读取原始差异行、对象或文件片段，"
        "不要只凭摘要层给最终结论。"
        "最终回答必须引用当前 observations 中的 evidence_nodes.id；历史经验不能替代当前 A/B 证据。"
        "如果 context_budget.truncated 为 true，说明 observations 只是压缩摘要/preview，"
        "需要更多细节时继续调用 get_compare_section_rows、get_compare_row 或文件 excerpt 工具。"
        "不要输出 Markdown，不要输出代码块。"
        f"{retry}\n输入：\n{json.dumps(payload, ensure_ascii=False, indent=2, default=str)}"
    )
