# -*- coding: utf-8 -*-
"""Flask app factory and route dependency wiring."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from pstx_queries.project_query import query_project_data
from pstx_core.cadence.csa_connectivity_overlay import build_csa_connectivity_overlay
from pstx_core.cadence.csa_geometry import build_csa_geometry_payload
from pstx_core.cadence.page_model import build_cadence_page_payload
from pstx_core.cadence.semantic_index import build_cadence_index_payload
from pstx_core.schematic_pdf_annotation import build_schematic_pdf_annotation_payload
from pstx_rules.project_analysis import analyze_project_contents
from pstx_rules.result_meta import DRC_ISSUE_KEYS as _DRC_ISSUE_KEYS
from pstx_harness.eval import AgentEvalError, build_agent_eval_status, run_agent_eval
from pstx_harness.compare_agent import (
    CompareAgentRequest,
    CompareMockModelProvider,
    list_compare_agent_profiles,
    run_compare_agent,
)
from pstx_harness.model import AsterHarnessModelProvider, MockHarnessModelProvider
from pstx_harness.report_agent import HarnessAgentRequest, list_harness_agent_profiles, run_harness_agent
from pstx_harness.report_tools import build_default_harness_registry
from pstx_harness.review import HarnessError, HarnessRunRequest, build_harness_status, run_harness_review
from pstx_integrations.aster.service import (
    ask_aster_model,
    aster_error_payload,
    build_aster_status,
    build_aster_summary,
    clear_aster_runtime_config,
    set_aster_runtime_config,
)
from pstx_integrations.diagnostics import (
    build_diagnostics_status,
    diagnostics_export_bytes,
    format_exception,
    new_diagnostic_request_id,
    summarize_mapping,
    tail_diagnostics,
    write_diagnostic_event,
)
from pstx_integrations.feishu.gateway import (
    FeishuBomError,
    build_feishu_bom_status,
    build_feishu_database_overview,
    build_feishu_mapping_from_headers,
    create_feishu_cache_row,
    delete_feishu_cache_library,
    delete_feishu_cache_row,
    fetch_feishu_sheet_list,
    get_feishu_cache_rows,
    get_saved_feishu_field_order,
    match_rows_with_feishu_cache,
    preview_feishu_sheet,
    suggest_feishu_mapping_from_preview,
    sync_feishu_library,
    update_feishu_cache_row,
)
from pstx_knowledge.datasheets import build_datasheet_status, reindex_datasheets
from pstx_knowledge.topology import build_llm_topology_netlist
from pstx_knowledge.reference_library import (
    build_agent_ref_status,
    build_review_checklist_status,
    reindex_agent_ref,
    reindex_review_checklists,
)
from pstx_webapp.agent_context import (
    agent_context_public,
    append_agent_context_answers,
    get_agent_context,
    new_agent_context,
    update_agent_context_after_run,
)
from pstx_webapp.compare_payload import build_compare_payload
from pstx_webapp.compare_view import coerce_compare_detail_limit
from pstx_webapp.form_parsing import parse_checkbox_flag, parse_voltage_map_text
from pstx_webapp.pages import render_home_page, render_named_page, render_report_page
from pstx_webapp.project_io import discover_project_files_with_snapshot, read_local_text_file
from pstx_webapp.report_feishu import FEISHU_BOM_ROW_SOURCES, MAX_FEISHU_PREVIEW_ROWS
from pstx_webapp.report_view import build_report_payload
from pstx_webapp.routes.agent_lab import register_agent_lab_routes
from pstx_webapp.routes.compare import register_compare_routes
from pstx_webapp.routes.diagnostics import register_diagnostics_routes
from pstx_webapp.routes.dfmea import register_dfmea_routes
from pstx_webapp.routes.feishu import register_feishu_routes
from pstx_webapp.routes.harness import register_harness_routes
from pstx_webapp.routes.pages import register_page_routes
from pstx_webapp.routes.projects import register_project_routes
from pstx_webapp.routes.reports import register_report_routes
from pstx_webapp.routes.schematic_pdf import register_schematic_pdf_routes
from pstx_webapp.routes.system import register_system_routes
from pstx_webapp.run_store import get_run, list_project_summaries, remember_run
from pstx_webapp.server import DEFAULT_HOST, DEFAULT_PORT
from pstx_webapp.state import (
    AGENT_BACKGROUND_RUNNER,
    AGENT_CONTEXT_CACHE,
    AGENT_DURABLE_STORE,
    AGENT_RUN_CACHE,
    RUN_CACHE,
)

BASE_DIR = Path(__file__).resolve().parent.parent
WEB_DIR = BASE_DIR / 'web'

Flask = abort = jsonify = render_template = request = send_file = url_for = None


def _export_to_excel(data: dict, out_path: str) -> str:
    from pstx_exports.excel import export_to_excel
    return export_to_excel(data, out_path)


def _export_bom_to_excel(data: dict, out_path: str, *, mode: str = "all") -> str:
    from pstx_exports.excel import export_bom_to_excel
    return export_bom_to_excel(data, out_path, mode=mode)


def _ensure_flask():
    try:
        from flask import (  # type: ignore
            Flask,
            abort,
            jsonify,
            render_template,
            request,
            send_file,
            url_for,
        )
        return Flask, abort, jsonify, render_template, request, send_file, url_for
    except Exception:
        print("未检测到可用的 Flask 环境，正在自动修复本地 Web 依赖...")
        subprocess.check_call([
            sys.executable,
            '-m',
            'pip',
            'install',
            '--upgrade',
            'Flask>=3.1,<4',
            'Jinja2>=3.1.6,<4',
            'Werkzeug>=3.1,<4',
            'MarkupSafe>=2.1,<4',
            'itsdangerous>=2.2,<3',
            'click>=8.1,<9',
            'blinker>=1.9,<2',
        ])
        from flask import (  # type: ignore
            Flask,
            abort,
            jsonify,
            render_template,
            request,
            send_file,
            url_for,
        )
        return Flask, abort, jsonify, render_template, request, send_file, url_for


def _ensure_flask_loaded():
    global Flask, abort, jsonify, render_template, request, send_file, url_for
    if Flask is None:
        (
            Flask,
            abort,
            jsonify,
            render_template,
            request,
            send_file,
            url_for,
        ) = _ensure_flask()
    return Flask, abort, jsonify, render_template, request, send_file, url_for


def _remember_agent_run(agent_run_id: str, payload: dict) -> None:
    AGENT_RUN_CACHE.remember(payload, agent_run_id=agent_run_id)


def _get_run(run_id: str) -> dict:
    payload = get_run(run_id)
    if not payload:
        abort(404)
    return payload


def create_app() -> "Flask":
    _ensure_flask_loaded()
    app = Flask(
        __name__,
        template_folder=str(WEB_DIR / 'templates'),
        static_folder=str(WEB_DIR / 'static'),
    )

    register_diagnostics_routes(
        app,
        request=request,
        jsonify=jsonify,
        send_file=send_file,
        build_diagnostics_status=build_diagnostics_status,
        diagnostics_export_bytes=diagnostics_export_bytes,
        format_exception=format_exception,
        new_diagnostic_request_id=new_diagnostic_request_id,
        summarize_mapping=summarize_mapping,
        tail_diagnostics=tail_diagnostics,
        write_diagnostic_event=write_diagnostic_event,
    )

    register_page_routes(
        app,
        render_template=render_template,
        request=request,
        abort=abort,
        render_home_page=render_home_page,
        render_named_page=render_named_page,
        render_report_page=render_report_page,
        default_host=DEFAULT_HOST,
        default_port=DEFAULT_PORT,
        get_run=_get_run,
    )

    register_agent_lab_routes(
        app,
        request=request,
        jsonify=jsonify,
        build_agent_ref_status=build_agent_ref_status,
        build_aster_status=build_aster_status,
        build_review_checklist_status=build_review_checklist_status,
        list_harness_agent_profiles=list_harness_agent_profiles,
        reindex_agent_ref=reindex_agent_ref,
        reindex_review_checklists=reindex_review_checklists,
        HarnessAgentRequest=HarnessAgentRequest,
        HarnessError=HarnessError,
        AsterHarnessModelProvider=AsterHarnessModelProvider,
        MockHarnessModelProvider=MockHarnessModelProvider,
        new_agent_context=new_agent_context,
        run_harness_agent=run_harness_agent,
    )

    register_system_routes(
        app,
        request=request,
        jsonify=jsonify,
        build_aster_status=build_aster_status,
        set_aster_runtime_config=set_aster_runtime_config,
        clear_aster_runtime_config=clear_aster_runtime_config,
        aster_error_payload=aster_error_payload,
        build_harness_status=build_harness_status,
        build_default_harness_registry=build_default_harness_registry,
        list_harness_agent_profiles=list_harness_agent_profiles,
        build_datasheet_status=build_datasheet_status,
        reindex_datasheets=reindex_datasheets,
        build_agent_eval_status=build_agent_eval_status,
        run_agent_eval=run_agent_eval,
        AgentEvalError=AgentEvalError,
    )

    register_project_routes(
        app,
        request=request,
        jsonify=jsonify,
        url_for=url_for,
        discover_project_files=discover_project_files_with_snapshot,
        read_local_text_file=read_local_text_file,
        parse_voltage_map_text=parse_voltage_map_text,
        parse_checkbox_flag=parse_checkbox_flag,
        analyze_project_contents=analyze_project_contents,
        remember_run=remember_run,
        list_project_summaries=list_project_summaries,
        build_report_payload=build_report_payload,
        drc_issue_keys=_DRC_ISSUE_KEYS,
    )

    register_report_routes(
        app,
        request=request,
        jsonify=jsonify,
        send_file=send_file,
        get_run=_get_run,
        build_aster_summary=build_aster_summary,
        aster_error_payload=aster_error_payload,
        match_rows_with_feishu_cache=match_rows_with_feishu_cache,
        query_project_data=query_project_data,
        export_to_excel=_export_to_excel,
        export_bom_to_excel=_export_bom_to_excel,
        build_llm_topology_netlist=build_llm_topology_netlist,
        build_cadence_page_payload=build_cadence_page_payload,
        build_cadence_index_payload=build_cadence_index_payload,
        build_csa_geometry_payload=build_csa_geometry_payload,
        build_csa_connectivity_overlay=build_csa_connectivity_overlay,
        feishu_bom_row_sources=FEISHU_BOM_ROW_SOURCES,
        max_feishu_preview_rows=MAX_FEISHU_PREVIEW_ROWS,
    )

    register_schematic_pdf_routes(
        app,
        request=request,
        jsonify=jsonify,
        get_run=_get_run,
        build_schematic_pdf_annotation_payload=build_schematic_pdf_annotation_payload,
    )

    register_dfmea_routes(
        app,
        request=request,
        jsonify=jsonify,
        send_file=send_file,
        get_run=_get_run,
    )

    register_harness_routes(
        app,
        request=request,
        jsonify=jsonify,
        run_cache=RUN_CACHE,
        agent_context_cache=AGENT_CONTEXT_CACHE,
        agent_run_cache=AGENT_RUN_CACHE,
        durable_store=AGENT_DURABLE_STORE,
        background_runner=AGENT_BACKGROUND_RUNNER,
        get_agent_context=get_agent_context,
        new_agent_context=new_agent_context,
        agent_context_public=agent_context_public,
        append_agent_context_answers=append_agent_context_answers,
        update_agent_context_after_run=update_agent_context_after_run,
        build_aster_status=build_aster_status,
        remember_agent_run=_remember_agent_run,
        HarnessRunRequest=HarnessRunRequest,
        HarnessAgentRequest=HarnessAgentRequest,
        HarnessError=HarnessError,
        AsterHarnessModelProvider=AsterHarnessModelProvider,
        MockHarnessModelProvider=MockHarnessModelProvider,
        run_harness_review=run_harness_review,
        run_harness_agent=run_harness_agent,
        build_compare_payload=lambda left_run_id, right_run_id, *, detail_limit: build_compare_payload(
            left_run_id,
            right_run_id,
            get_run_payload=_get_run,
            detail_limit=detail_limit,
        ),
        CompareAgentRequest=CompareAgentRequest,
        CompareMockModelProvider=CompareMockModelProvider,
        run_compare_agent=run_compare_agent,
    )

    register_compare_routes(
        app,
        request=request,
        jsonify=jsonify,
        run_cache=RUN_CACHE,
        durable_store=AGENT_DURABLE_STORE,
        background_runner=AGENT_BACKGROUND_RUNNER,
        build_compare_payload=lambda left_run_id, right_run_id, *, detail_limit: build_compare_payload(
            left_run_id,
            right_run_id,
            get_run_payload=_get_run,
            detail_limit=detail_limit,
        ),
        coerce_compare_detail_limit=coerce_compare_detail_limit,
        list_compare_agent_profiles=list_compare_agent_profiles,
        CompareAgentRequest=CompareAgentRequest,
        CompareMockModelProvider=CompareMockModelProvider,
        HarnessError=HarnessError,
        AsterHarnessModelProvider=AsterHarnessModelProvider,
        build_aster_status=build_aster_status,
        run_compare_agent=run_compare_agent,
        remember_agent_run=_remember_agent_run,
    )

    register_feishu_routes(
        app,
        request=request,
        jsonify=jsonify,
        FeishuBomError=FeishuBomError,
        build_feishu_bom_status=build_feishu_bom_status,
        build_feishu_database_overview=build_feishu_database_overview,
        get_feishu_cache_rows=get_feishu_cache_rows,
        create_feishu_cache_row=create_feishu_cache_row,
        update_feishu_cache_row=update_feishu_cache_row,
        delete_feishu_cache_library=delete_feishu_cache_library,
        delete_feishu_cache_row=delete_feishu_cache_row,
        fetch_feishu_sheet_list=lambda **kwargs: fetch_feishu_sheet_list(**kwargs),
        preview_feishu_sheet=lambda **kwargs: preview_feishu_sheet(**kwargs),
        get_saved_feishu_field_order=get_saved_feishu_field_order,
        suggest_feishu_mapping_from_preview=suggest_feishu_mapping_from_preview,
        build_feishu_mapping_from_headers=build_feishu_mapping_from_headers,
        build_aster_status=build_aster_status,
        ask_aster_model=lambda *args, **kwargs: ask_aster_model(*args, **kwargs),
        AsterHarnessModelProvider=AsterHarnessModelProvider,
        sync_feishu_library=lambda **kwargs: sync_feishu_library(**kwargs),
    )

    return app


__all__ = ["create_app"]
