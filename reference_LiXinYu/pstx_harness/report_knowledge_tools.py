# -*- coding: utf-8 -*-
"""Knowledge-source tools used by the report harness agent."""

from __future__ import annotations

from typing import Optional

from pstx_harness.tool_core import HarnessToolContext, HarnessToolError
from pstx_harness.report_tool_utils import (
    _as_int,
    _batch_limit,
    _batch_summary,
    _compact_mapping,
    _feishu_row_summary,
    _safe_text,
    _sanitize_feishu_row,
)
from pstx_knowledge.business_dictionary import business_dictionary_summary
from pstx_knowledge.component_identity import (
    USER_VISIBLE_REAL_PAGE_LABEL,
    build_component_identity_cards,
    filter_component_identity_cards,
    summarize_dfmea_readiness,
)
from pstx_knowledge.datasheets import (
    batch_search_datasheet_chunks,
    build_datasheet_status,
    get_datasheet_chunk,
    get_datasheet_excerpt,
    get_datasheet_page_excerpt,
    get_datasheet_parameter,
    list_datasheet_documents,
    match_component_datasheets,
    search_datasheet_chunks,
    search_datasheet_parameters,
    search_datasheets,
    summarize_datasheet_coverage,
)
from pstx_knowledge.datasheet_review_templates import (
    get_datasheet_review_template,
    list_datasheet_review_templates,
)
from pstx_knowledge.document_search import (
    batch_search_documents,
    build_document_search_status,
    get_document_excerpt,
    search_documents,
)
from pstx_knowledge.feishu_cache import (
    build_feishu_database_overview,
    get_feishu_cache_row,
    get_feishu_cache_rows,
)
from pstx_knowledge.reference_library import (
    build_agent_ref_status,
    build_review_checklist_status,
    get_agent_ref_excerpt,
    get_review_checklist_excerpt,
    search_agent_ref,
    search_review_checklists,
)
from pstx_knowledge.topology import (
    batch_expand_topology_review_tasks,
    batch_query_llm_topology_netlist,
    batch_query_chip_topology,
    build_chip_topology,
    build_llm_topology_netlist,
    get_chip_topology_edge,
    get_llm_topology_edge,
    get_llm_topology_node,
    get_topology_review_task,
    query_llm_topology_netlist,
    query_chip_topology,
    summarize_topology_review_tasks,
)


def _list_feishu_cache_libraries_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    args = args or {}
    include_sheets = bool(args.get("include_sheets", True))
    overview = build_feishu_database_overview()
    libraries = []
    for library in overview.get("libraries", []) or []:
        if not isinstance(library, dict):
            continue
        item = {
            "lib_id": _safe_text(library.get("lib_id", ""), 120),
            "lib_name": _safe_text(library.get("lib_name", ""), 160),
            "cache_count": _as_int(library.get("cache_count"), 0),
            "last_synced_at": _safe_text(library.get("last_synced_at", ""), 120),
            "sheet_config_count": _as_int(library.get("sheet_config_count"), 0),
            "enabled_sheet_count": _as_int(library.get("enabled_sheet_count"), 0),
        }
        if include_sheets:
            item["sheet_stats"] = [
                {
                    "sheet_name": _safe_text(sheet.get("sheet_name", ""), 160),
                    "count": _as_int(sheet.get("count"), 0),
                    "last_synced_at": _safe_text(sheet.get("last_synced_at", ""), 120),
                }
                for sheet in list(library.get("sheet_stats") or [])[:24]
                if isinstance(sheet, dict)
            ]
            item["configured_sheets"] = [
                {
                    "sheet_id": _safe_text(sheet.get("sheet_id", ""), 120),
                    "title": _safe_text(sheet.get("title", ""), 160),
                    "header_row": _as_int(sheet.get("header_row"), 1),
                    "hq_code_col": _safe_text(sheet.get("hq_code_col", ""), 120),
                    "spec_model_col": _safe_text(sheet.get("spec_model_col", ""), 120),
                    "pi_col": _safe_text(sheet.get("pi_col", ""), 120),
                    "selection_order_col": _safe_text(sheet.get("selection_order_col", ""), 120),
                }
                for sheet in list(library.get("configured_sheets") or [])[:24]
                if isinstance(sheet, dict)
            ]
        libraries.append(item)

    status = "可用" if overview.get("available") else "不可用"
    return {
        "id": "list_feishu_cache_libraries",
        "title": "飞书缓存库清单",
        "target": "bom",
        "summary": (
            f"飞书本地缓存{status}，库数量 {len(libraries)}，"
            f"缓存行 {overview.get('cache_count', 0)}。"
        ),
        "ok": bool(overview.get("ok", True)),
        "available": bool(overview.get("available")),
        "configured": bool(overview.get("configured")),
        "cache_count": _as_int(overview.get("cache_count"), 0),
        "library_count": len(libraries),
        "libraries": libraries,
        "saved_field_order": [
            _safe_text(item, 120)
            for item in list(overview.get("saved_field_order") or [])[:40]
        ],
        "readonly": True,
    }


def _list_business_dictionary_tool(context: HarnessToolContext, args: Optional[dict] = None) -> dict:
    dictionary = business_dictionary_summary()
    interface_count = int(dictionary.get("interface_count") or 0)
    return {
        "id": "list_business_dictionary",
        "title": "项目业务词典/缩写表",
        "target": "topology",
        "summary": (
            f"当前使用 {dictionary.get('source') or 'builtin'} 业务词典，"
            f"接口缩写组 {interface_count} 个；用于拓扑接口识别、查询同义词和 review focus。"
        ),
        "ok": True,
        "readonly": True,
        "dictionary": dictionary,
    }


def _search_feishu_cache_rows_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_feishu_cache_rows 需要 query。")
    limit = _as_int(args.get("limit", 20), 20)
    offset = _as_int(args.get("offset", 0), 0)
    result = get_feishu_cache_rows(
        lib_id=str(args.get("lib_id") or "").strip(),
        sheet_name=str(args.get("sheet_name") or "").strip(),
        query=query,
        limit=limit,
        offset=offset,
    )
    rows = [_sanitize_feishu_row(row) for row in result.get("rows", []) or []]
    total = _as_int(result.get("total"), 0)
    ok = bool(result.get("ok", False))
    if not ok:
        summary = result.get("error") or "飞书缓存搜索失败。"
    elif total:
        summary = f"飞书缓存搜索 `{query}` 命中 {total} 条，返回 {len(rows)} 条。"
    else:
        summary = f"飞书缓存搜索 `{query}` 无命中；建议补充 HQ 料号、规格型号、PI 或选型顺序关键词。"
    return {
        "id": "search_feishu_cache_rows",
        "title": "搜索飞书缓存物料",
        "target": "bom",
        "summary": summary,
        "ok": ok,
        "query": _safe_text(query, 220),
        "lib_id": _safe_text(args.get("lib_id", ""), 120),
        "sheet_name": _safe_text(args.get("sheet_name", ""), 160),
        "total_rows": total,
        "limit": _as_int(result.get("limit"), limit),
        "offset": _as_int(result.get("offset"), offset),
        "rows": rows,
        "readonly": True,
    }


def _get_feishu_cache_row_tool(context: HarnessToolContext, args: dict) -> dict:
    row_id = _as_int(args.get("row_id"), 0)
    if row_id <= 0:
        raise HarnessToolError("get_feishu_cache_row 需要正整数 row_id。")
    result = get_feishu_cache_row(row_id)
    row = _sanitize_feishu_row(result.get("row") or {}) if result.get("row") else None
    ok = bool(result.get("ok", False))
    summary = _feishu_row_summary(row) if row else (result.get("error") or f"未找到缓存行 id={row_id}。")
    return {
        "id": "get_feishu_cache_row",
        "title": f"飞书缓存行 {row_id}",
        "target": "bom",
        "summary": summary,
        "ok": ok,
        "row_id": row_id,
        "row": row,
        "readonly": True,
    }


def _identity_card_preview(card: dict) -> dict:
    return {
        "refdes": _safe_text(card.get("refdes", ""), 80),
        "category": _safe_text(card.get("category", ""), 80),
        "candidate_chip_type": _safe_text(card.get("candidate_chip_type", ""), 120),
        "hq_no": _safe_text(card.get("hq_no", ""), 140),
        "spec": _safe_text(card.get("spec", ""), 260),
        "pi": _safe_text(card.get("pi", ""), 120),
        "selection_order": _safe_text(card.get("selection_order", ""), 80),
        "package": _safe_text(card.get("package", ""), 100),
        "value": _safe_text(card.get("value", ""), 120),
        "bom_option": _safe_text(card.get("bom_option", ""), 120),
        USER_VISIBLE_REAL_PAGE_LABEL: _safe_text(card.get(USER_VISIBLE_REAL_PAGE_LABEL) or card.get("user_visible_page", ""), 80),
        "user_visible_page": _safe_text(card.get("user_visible_page", ""), 80),
        "pin_net_summary": list(card.get("pin_net_summary") or [])[:8],
        "power_nets": list(card.get("power_nets") or [])[:8],
        "interface_nets": list(card.get("interface_nets") or [])[:8],
        "feishu_match": dict(card.get("feishu_match") or {}),
        "datasheet_match": dict(card.get("datasheet_match") or {}),
        "datasheet_missing_reason": _safe_text(card.get("datasheet_missing_reason", ""), 220),
        "missing_fields": list(card.get("missing_fields") or []),
        "confidence": _safe_text(card.get("confidence", ""), 80),
    }


def _list_component_identity_cards_tool(context: HarnessToolContext, args: dict) -> dict:
    cards = build_component_identity_cards(context.report, context.bundle)
    cards = filter_component_identity_cards(
        cards,
        category=str(args.get("category") or "").strip(),
        refdes_prefix=str(args.get("refdes_prefix") or "").strip(),
        hq_no=str(args.get("hq_no") or "").strip(),
        feishu_status=str(args.get("feishu_status") or "").strip(),
    )
    limit = _as_int(args.get("limit", 20), 20)
    offset = _as_int(args.get("offset", 0), 0)
    selected = cards[offset:offset + limit]
    return {
        "id": "list_component_identity_cards",
        "title": "元件身份卡清单",
        "target": "dfmea",
        "summary": f"当前项目筛选出 {len(cards)} 张元件身份卡，返回 {len(selected)} 张。",
        "total_cards": len(cards),
        "limit": limit,
        "offset": offset,
        "cards": [_identity_card_preview(card) for card in selected],
        "readonly": True,
    }


def _get_component_identity_card_tool(context: HarnessToolContext, args: dict) -> dict:
    refdes = str(args.get("refdes") or "").strip()
    if not refdes:
        raise HarnessToolError("get_component_identity_card 需要 refdes。")
    cards = build_component_identity_cards(context.report, context.bundle)
    for card in cards:
        if str(card.get("refdes") or "").upper() == refdes.upper():
            return {
                "id": "get_component_identity_card",
                "title": f"元件身份卡 {card.get('refdes')}",
                "target": "dfmea",
                "summary": (
                    f"{card.get('refdes')} 分类为 {card.get('category')}，"
                    f"HQ={card.get('hq_no') or '缺失'}，规格={card.get('spec') or '缺失'}。"
                ),
                "card": card,
                "readonly": True,
            }
    raise HarnessToolError(f"未找到元件身份卡：{refdes}")


def _search_component_identity_cards_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_component_identity_cards 需要 query。")
    cards = build_component_identity_cards(context.report, context.bundle)
    cards = filter_component_identity_cards(
        cards,
        category=str(args.get("category") or "").strip(),
        query=query,
    )
    limit = _as_int(args.get("limit", 20), 20)
    offset = _as_int(args.get("offset", 0), 0)
    selected = cards[offset:offset + limit]
    return {
        "id": "search_component_identity_cards",
        "title": f"搜索元件身份卡：{query}",
        "target": "dfmea",
        "summary": f"搜索 `{query}` 命中 {len(cards)} 张元件身份卡，返回 {len(selected)} 张。",
        "query": _safe_text(query, 200),
        "total_cards": len(cards),
        "limit": limit,
        "offset": offset,
        "cards": [_identity_card_preview(card) for card in selected],
        "readonly": True,
    }


def _summarize_dfmea_readiness_tool(context: HarnessToolContext, args: dict) -> dict:
    cards = build_component_identity_cards(context.report, context.bundle)
    summary = summarize_dfmea_readiness(cards)
    return {
        "id": "summarize_dfmea_readiness",
        "title": "DFMEA 准备度摘要",
        "target": "dfmea",
        "summary": (
            f"当前项目 {summary.get('total_components', 0)} 个元件中，"
            f"{summary.get('ready_count', 0)} 个关键器件具备第一阶段 DFMEA 输入条件，"
            f"{summary.get('needs_context_count', 0)} 个关键器件仍需补充上下文。"
        ),
        **summary,
        "ready_cards": [_identity_card_preview(card) for card in summary.get("ready_cards", [])],
        "needs_context_cards": [_identity_card_preview(card) for card in summary.get("needs_context_cards", [])],
        "readonly": True,
    }


def _summarize_chip_topology_tool(context: HarnessToolContext, args: dict) -> dict:
    result = build_llm_topology_netlist(
        context.report,
        context.bundle,
        focus_refdes=str(args.get("focus_refdes") or ""),
        role_filter=str(args.get("role_filter") or ""),
        include_connectors=bool(args.get("include_connectors", False)),
        limit=_as_int(args.get("limit", 30), 30),
        view=str(args.get("view") or "summary"),
        supply_mode=str(args.get("supply_mode") or "grouped"),
        supply_limit=_as_int(args.get("supply_limit", 12), 12),
    )
    return {
        "id": "summarize_chip_topology",
        "title": "芯片级连接拓扑摘要",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _summarize_llm_topology_netlist_tool(context: HarnessToolContext, args: dict) -> dict:
    result = build_llm_topology_netlist(
        context.report,
        context.bundle,
        focus_refdes=str(args.get("focus_refdes") or ""),
        role_filter=str(args.get("role_filter") or ""),
        include_connectors=bool(args.get("include_connectors", False)),
        limit=_as_int(args.get("limit", 30), 30),
        view=str(args.get("view") or "summary"),
        supply_mode=str(args.get("supply_mode") or "grouped"),
        supply_limit=_as_int(args.get("supply_limit", 12), 12),
    )
    return {
        "id": "summarize_llm_topology_netlist",
        "title": "LLM 拓扑网表摘要",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _query_chip_topology_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("query_chip_topology 需要 query。")
    result = query_chip_topology(
        context.report,
        context.bundle,
        query,
        include_connectors=bool(args.get("include_connectors", False)),
        limit=_as_int(args.get("limit", 30), 30),
    )
    return {
        "id": "query_chip_topology",
        "title": f"查询芯片级拓扑：{_safe_text(query, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _query_llm_topology_netlist_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("query_llm_topology_netlist 需要 query。")
    result = query_llm_topology_netlist(
        context.report,
        context.bundle,
        query,
        include_connectors=bool(args.get("include_connectors", False)),
        limit=_as_int(args.get("limit", 30), 30),
    )
    return {
        "id": "query_llm_topology_netlist",
        "title": f"查询 LLM 拓扑网表：{_safe_text(query, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _get_llm_topology_node_tool(context: HarnessToolContext, args: dict) -> dict:
    refdes = str(args.get("refdes") or "").strip()
    if not refdes:
        raise HarnessToolError("get_llm_topology_node 需要 refdes。")
    result = get_llm_topology_node(
        context.report,
        context.bundle,
        refdes,
        include_connectors=bool(args.get("include_connectors", False)),
        max_pin_nets=_as_int(args.get("max_pin_nets", 240), 240),
    )
    if not result.get("ok", True):
        raise HarnessToolError(result.get("summary") or f"未找到 LLM 拓扑节点：{refdes}")
    return {
        "id": "get_llm_topology_node",
        "title": f"读取 LLM 拓扑节点：{_safe_text(refdes, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _get_chip_topology_edge_tool(context: HarnessToolContext, args: dict) -> dict:
    edge_id = str(args.get("edge_id") or "").strip()
    if not edge_id:
        raise HarnessToolError("get_chip_topology_edge 需要 edge_id。")
    result = get_chip_topology_edge(
        context.report,
        context.bundle,
        edge_id,
        include_connectors=bool(args.get("include_connectors", False)),
    )
    if not result.get("ok", True):
        raise HarnessToolError(result.get("summary") or f"未找到芯片级拓扑连接：{edge_id}")
    return {
        "id": "get_chip_topology_edge",
        "title": f"读取芯片级拓扑连接：{_safe_text(edge_id, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _get_llm_topology_edge_tool(context: HarnessToolContext, args: dict) -> dict:
    edge_id = str(args.get("edge_id") or "").strip()
    if not edge_id:
        raise HarnessToolError("get_llm_topology_edge 需要 edge_id。")
    result = get_llm_topology_edge(
        context.report,
        context.bundle,
        edge_id,
        include_connectors=bool(args.get("include_connectors", False)),
    )
    if not result.get("ok", True):
        raise HarnessToolError(result.get("summary") or f"未找到 LLM 拓扑连接：{edge_id}")
    return {
        "id": "get_llm_topology_edge",
        "title": f"读取 LLM 拓扑连接：{_safe_text(edge_id, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _batch_query_chip_topology_tool(context: HarnessToolContext, args: dict) -> dict:
    result = batch_query_chip_topology(
        context.report,
        context.bundle,
        args.get("queries") or [],
        include_connectors=bool(args.get("include_connectors", False)),
        limit_per_query=_as_int(args.get("limit_per_query", args.get("limit", 8)), 8),
    )
    return {
        "id": "batch_query_chip_topology",
        "title": "批量查询芯片级拓扑",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _batch_query_llm_topology_netlist_tool(context: HarnessToolContext, args: dict) -> dict:
    result = batch_query_llm_topology_netlist(
        context.report,
        context.bundle,
        args.get("queries") or [],
        include_connectors=bool(args.get("include_connectors", False)),
        limit_per_query=_as_int(args.get("limit_per_query", args.get("limit", 8)), 8),
    )
    return {
        "id": "batch_query_llm_topology_netlist",
        "title": "批量查询 LLM 拓扑网表",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _summarize_topology_review_tasks_tool(context: HarnessToolContext, args: dict) -> dict:
    result = summarize_topology_review_tasks(
        context.report,
        context.bundle,
        include_connectors=bool(args.get("include_connectors", False)),
        focus_refdes=str(args.get("focus_refdes") or ""),
        interface_group=str(args.get("interface_group") or ""),
        priority=str(args.get("priority") or ""),
        limit=_as_int(args.get("limit", 30), 30),
    )
    return {
        "id": "summarize_topology_review_tasks",
        "title": "拓扑 review 任务队列",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _get_topology_review_task_tool(context: HarnessToolContext, args: dict) -> dict:
    task_id = str(args.get("task_id") or "").strip()
    if not task_id:
        raise HarnessToolError("get_topology_review_task 需要 task_id。")
    result = get_topology_review_task(
        context.report,
        context.bundle,
        task_id,
        include_connectors=bool(args.get("include_connectors", False)),
    )
    if not result.get("ok", True):
        raise HarnessToolError(result.get("summary") or f"未找到拓扑 review task：{task_id}")
    return {
        "id": "get_topology_review_task",
        "title": f"拓扑 review task：{_safe_text(task_id, 80)}",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _batch_expand_topology_review_tasks_tool(context: HarnessToolContext, args: dict) -> dict:
    result = batch_expand_topology_review_tasks(
        context.report,
        context.bundle,
        args.get("task_ids") or [],
        include_connectors=bool(args.get("include_connectors", False)),
    )
    return {
        "id": "batch_expand_topology_review_tasks",
        "title": "批量展开拓扑 review task",
        "target": "topology",
        **result,
        "readonly": True,
    }


def _list_document_search_sources_tool(context: HarnessToolContext, args: dict) -> dict:
    status = build_document_search_status()
    return {
        "id": "list_document_search_sources",
        "title": "本地文档搜索状态",
        "target": "document_search",
        **status,
        "readonly": True,
    }


def _search_documents_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_documents 需要 query。")
    result = search_documents(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        max_files=_as_int(args.get("max_files", 500), 500),
    )
    return {
        "id": "search_documents",
        "title": f"搜索本地文档：{_safe_text(query, 80)}",
        "target": "document_search",
        **result,
        "readonly": True,
    }


def _get_document_excerpt_tool(context: HarnessToolContext, args: dict) -> dict:
    doc_id = str(args.get("doc_id") or "").strip()
    if not doc_id:
        raise HarnessToolError("get_document_excerpt 需要 doc_id。")
    result = get_document_excerpt(
        doc_id,
        char_start=_as_int(args.get("char_start", 0), 0),
        before_chars=_as_int(args.get("before_chars", 800), 800),
        after_chars=_as_int(args.get("after_chars", 1600), 1600),
        max_chars=_as_int(args.get("max_chars", 5000), 5000),
    )
    if not result.get("ok", True):
        raise HarnessToolError(result.get("summary") or f"读取文档片段失败：{doc_id}")
    return {
        "id": "get_document_excerpt",
        "title": result.get("title") or f"文档片段 {doc_id}",
        "target": "document_search",
        **result,
        "readonly": True,
    }


def _batch_search_documents_tool(context: HarnessToolContext, args: dict) -> dict:
    result = batch_search_documents(
        args.get("queries") or [],
        limit_per_query=_as_int(args.get("limit_per_query", args.get("limit", 8)), 8),
    )
    return {
        "id": "batch_search_documents",
        "title": "批量搜索本地文档",
        "target": "document_search",
        **result,
        "readonly": True,
    }


def _datasheet_match_preview(match: dict) -> dict:
    return {
        "doc_id": _as_int(match.get("doc_id"), 0),
        "title": _safe_text(match.get("title", ""), 220),
        "page": _as_int(match.get("page"), 1),
        "chunk_id": _safe_text(match.get("chunk_id", ""), 80),
        "section_title": _safe_text(match.get("section_title", ""), 180),
        "score": _as_int(match.get("score"), 0),
        "matched_terms": list(match.get("matched_terms") or [])[:8],
        "keywords": _safe_text(match.get("keywords", ""), 240),
        "char_range": list(match.get("char_range") or [])[:2],
        "snippet": _safe_text(match.get("snippet", ""), 420),
    }


def _datasheet_parameter_preview(parameter: dict) -> dict:
    return {
        "parameter_id": _as_int(parameter.get("parameter_id"), 0),
        "evidence_id": _safe_text(parameter.get("evidence_id", ""), 120),
        "doc_id": _as_int(parameter.get("doc_id"), 0),
        "title": _safe_text(parameter.get("title", ""), 220),
        "parameter_key": _safe_text(parameter.get("parameter_key", ""), 120),
        "parameter_name": _safe_text(parameter.get("parameter_name", ""), 220),
        "value_text": _safe_text(parameter.get("value_text", ""), 220),
        "value_min": parameter.get("value_min"),
        "value_typ": parameter.get("value_typ"),
        "value_max": parameter.get("value_max"),
        "unit": _safe_text(parameter.get("unit", ""), 60),
        "condition": _safe_text(parameter.get("condition", ""), 260),
        "page": _as_int(parameter.get("page"), 1),
        "chunk_id": _safe_text(parameter.get("chunk_id", ""), 100),
        "confidence": _safe_text(parameter.get("confidence", ""), 80),
        "extraction_method": _safe_text(parameter.get("extraction_method", ""), 120),
        "source_text": _safe_text(parameter.get("source_text", ""), 420),
        "detail_locator": dict(parameter.get("detail_locator") or {}),
    }


def _list_datasheet_review_templates_tool(context: HarnessToolContext, args: dict) -> dict:
    result = list_datasheet_review_templates(
        str(args.get("category") or ""),
        include_questions=bool(args.get("include_questions", True)),
    )
    return {
        "id": "list_datasheet_review_templates",
        "title": "Datasheet 审查模板清单",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _get_datasheet_review_template_tool(context: HarnessToolContext, args: dict) -> dict:
    template_id = str(args.get("template_id") or "").strip()
    if not template_id:
        raise HarnessToolError("get_datasheet_review_template 需要 template_id。")
    result = get_datasheet_review_template(template_id)
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取 datasheet 审查模板失败。"))
    return {
        "id": "get_datasheet_review_template",
        "title": result.get("template", {}).get("title") or "Datasheet 审查模板",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _list_datasheet_sources_tool(context: HarnessToolContext, args: dict) -> dict:
    status = build_datasheet_status()
    return {
        "id": "list_datasheet_sources",
        "title": "本地规格书索引状态",
        "target": "dfmea",
        "summary": status.get("summary", ""),
        **status,
        "readonly": True,
    }


def _search_datasheet_parameters_tool(context: HarnessToolContext, args: dict) -> dict:
    result = search_datasheet_parameters(
        str(args.get("query") or ""),
        parameter_key=str(args.get("parameter_key") or ""),
        doc_id=_as_int(args.get("doc_id"), 0) or None,
        limit=_as_int(args.get("limit", 30), 30),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "规格书参数卡检索失败。"))
    return {
        "id": "search_datasheet_parameters",
        "title": "搜索规格书参数卡",
        "target": "dfmea",
        "summary": result.get("summary", ""),
        "query": _safe_text(result.get("query", ""), 220),
        "parameter_key": _safe_text(result.get("parameter_key", ""), 120),
        "doc_id": result.get("doc_id"),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 30),
        "offset": result.get("offset", 0),
        "parameters": [_datasheet_parameter_preview(item) for item in result.get("parameters", [])],
        "readonly": True,
    }


def _get_datasheet_parameter_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_datasheet_parameter(
        _as_int(args.get("parameter_id"), 0),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书参数卡失败。"))
    return {
        "id": "get_datasheet_parameter",
        "title": result.get("parameter_name") or "规格书参数卡",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _list_datasheet_documents_tool(context: HarnessToolContext, args: dict) -> dict:
    result = list_datasheet_documents(
        limit=_as_int(args.get("limit", 200), 200),
        offset=_as_int(args.get("offset", 0), 0),
    )
    return {
        "id": "list_datasheet_documents",
        "title": "本地规格书文档清单",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _search_datasheet_chunks_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_datasheet_chunks 需要 query。")
    result = search_datasheet_chunks(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "规格书 chunk 检索失败。"))
    return {
        "id": "search_datasheet_chunks",
        "title": f"搜索规格书片段：{query}",
        "target": "dfmea",
        "summary": result.get("summary", ""),
        "query": _safe_text(query, 220),
        "terms": result.get("terms", []),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 20),
        "offset": result.get("offset", 0),
        "matches": [_datasheet_match_preview(match) for match in result.get("matches", [])],
        "readonly": True,
    }


def _get_datasheet_chunk_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_datasheet_chunk(
        _as_int(args.get("doc_id"), 0),
        str(args.get("chunk_id") or ""),
        max_chars=_as_int(args.get("max_chars", 4000), 4000),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书 chunk 失败。"))
    return {
        "id": "get_datasheet_chunk",
        "title": result.get("title") or "规格书 chunk",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _get_datasheet_page_excerpt_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_datasheet_page_excerpt(
        _as_int(args.get("doc_id"), 0),
        _as_int(args.get("page"), 1),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书页片段失败。"))
    return {
        "id": "get_datasheet_page_excerpt",
        "title": result.get("title") or "规格书页片段",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _batch_search_datasheet_chunks_tool(context: HarnessToolContext, args: dict) -> dict:
    result = batch_search_datasheet_chunks(
        args.get("queries") or [],
        limit_per_query=_as_int(args.get("limit_per_query", args.get("limit", 8)), 8),
    )
    items = []
    for item in result.get("items", []) or []:
        if not isinstance(item, dict):
            continue
        compact = dict(item)
        compact["query"] = _safe_text(item.get("query", ""), 220)
        compact["matches"] = [_datasheet_match_preview(match) for match in item.get("matches", []) or []]
        compact["missing_reason"] = _safe_text(item.get("missing_reason", ""), 260)
        if item.get("error"):
            compact["error"] = _safe_text(item.get("error", ""), 260)
        items.append(compact)
    return {
        "id": "batch_search_datasheet_chunks",
        "title": "批量搜索规格书片段",
        "target": "dfmea",
        "summary": result.get("summary", ""),
        "query_count": result.get("query_count", 0),
        "limit_per_query": result.get("limit_per_query", 8),
        "truncated": bool(result.get("truncated")),
        "items": items,
        "readonly": True,
    }


def _search_datasheets_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_datasheets 需要 query。")
    result = search_datasheets(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "规格书检索失败。"))
    return {
        "id": "search_datasheets",
        "title": f"搜索规格书：{query}",
        "target": "dfmea",
        "summary": result.get("summary", ""),
        "query": _safe_text(query, 220),
        "terms": result.get("terms", []),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 20),
        "offset": result.get("offset", 0),
        "matches": [_datasheet_match_preview(match) for match in result.get("matches", [])],
        "readonly": True,
    }


def _get_datasheet_excerpt_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_datasheet_excerpt(
        _as_int(args.get("doc_id"), 0),
        _as_int(args.get("page"), 1),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取规格书片段失败。"))
    return {
        "id": "get_datasheet_excerpt",
        "title": result.get("title") or "规格书片段",
        "target": "dfmea",
        **result,
        "readonly": True,
    }


def _list_agent_ref_sources_tool(context: HarnessToolContext, args: dict) -> dict:
    status = build_agent_ref_status()
    return {
        "id": "list_agent_ref_sources",
        "title": "Agent Lab ref PDF 索引状态",
        "target": "agent_ref",
        "summary": status.get("summary", ""),
        **status,
        "readonly": True,
    }


def _search_agent_ref_pdfs_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_agent_ref_pdfs 需要 query。")
    result = search_agent_ref(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "ref PDF 检索失败。"))
    return {
        "id": "search_agent_ref_pdfs",
        "title": f"搜索 ref PDF：{query}",
        "target": "agent_ref",
        "summary": result.get("summary", ""),
        "query": _safe_text(query, 220),
        "terms": result.get("terms", []),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 20),
        "offset": result.get("offset", 0),
        "matches": [_datasheet_match_preview(match) | {"rel_path": _safe_text(match.get("rel_path", ""), 240)} for match in result.get("matches", [])],
        "readonly": True,
    }


def _get_agent_ref_pdf_excerpt_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_agent_ref_excerpt(
        _as_int(args.get("doc_id"), 0),
        _as_int(args.get("page"), 1),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取 ref PDF 片段失败。"))
    return {
        "id": "get_agent_ref_pdf_excerpt",
        "title": result.get("title") or "ref PDF 片段",
        "target": "agent_ref",
        **result,
        "readonly": True,
    }


def _list_review_checklist_sources_tool(context: HarnessToolContext, args: dict) -> dict:
    status = build_review_checklist_status()
    return {
        "id": "list_review_checklist_sources",
        "title": "Review checklist 索引状态",
        "target": "review_checklist",
        "summary": status.get("summary", ""),
        **status,
        "readonly": True,
    }


def _search_review_checklists_tool(context: HarnessToolContext, args: dict) -> dict:
    query = str(args.get("query") or "").strip()
    if not query:
        raise HarnessToolError("search_review_checklists 需要 query。")
    result = search_review_checklists(
        query,
        limit=_as_int(args.get("limit", 20), 20),
        offset=_as_int(args.get("offset", 0), 0),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "review checklist 检索失败。"))
    return {
        "id": "search_review_checklists",
        "title": f"搜索 review checklist：{query}",
        "target": "review_checklist",
        "summary": result.get("summary", ""),
        "query": _safe_text(query, 220),
        "terms": result.get("terms", []),
        "total_matches": result.get("total_matches", 0),
        "limit": result.get("limit", 20),
        "offset": result.get("offset", 0),
        "matches": [_datasheet_match_preview(match) | {"rel_path": _safe_text(match.get("rel_path", ""), 240)} for match in result.get("matches", [])],
        "readonly": True,
    }


def _get_review_checklist_excerpt_tool(context: HarnessToolContext, args: dict) -> dict:
    result = get_review_checklist_excerpt(
        _as_int(args.get("doc_id"), 0),
        _as_int(args.get("page"), 1),
        max_chars=_as_int(args.get("max_chars", 2400), 2400),
    )
    if not result.get("ok", True):
        raise HarnessToolError(str(result.get("error") or "读取 review checklist 片段失败。"))
    return {
        "id": "get_review_checklist_excerpt",
        "title": result.get("title") or "review checklist 片段",
        "target": "review_checklist",
        **result,
        "readonly": True,
    }


def _match_component_datasheets_tool(context: HarnessToolContext, args: dict) -> dict:
    refdes = str(args.get("refdes") or "").strip()
    if not refdes:
        raise HarnessToolError("match_component_datasheets 需要 refdes。")
    cards = build_component_identity_cards(context.report, context.bundle)
    for card in cards:
        if str(card.get("refdes") or "").upper() == refdes.upper():
            result = match_component_datasheets(card, limit=_as_int(args.get("limit", 5), 5))
            matches = [_datasheet_match_preview(match) for match in result.get("matches", [])]
            return {
                "id": "match_component_datasheets",
                "title": f"{refdes} 规格书候选",
                "target": "dfmea",
                "summary": f"{refdes} 命中 {len(matches)} 个规格书页级候选。" if matches else f"{refdes} 暂未命中规格书候选。",
                "refdes": refdes,
                "card": _identity_card_preview(card),
                "query": _safe_text(result.get("query", ""), 320),
                "matches": matches,
                "missing_reason": _safe_text(result.get("missing_reason", ""), 260),
                "readonly": True,
            }
    raise HarnessToolError(f"未找到元件身份卡：{refdes}")


def _summarize_dfmea_datasheet_coverage_tool(context: HarnessToolContext, args: dict) -> dict:
    cards = build_component_identity_cards(context.report, context.bundle)
    coverage = summarize_datasheet_coverage(cards, limit=_as_int(args.get("limit", 12), 12))
    return {
        "id": "summarize_dfmea_datasheet_coverage",
        "title": "DFMEA 规格书覆盖摘要",
        "target": "dfmea",
        "summary": coverage.get("summary", ""),
        **coverage,
        "readonly": True,
    }
