# -*- coding: utf-8 -*-
"""CLI-friendly JSON interfaces for PSTX analysis capabilities."""

from __future__ import annotations

import argparse
from collections import Counter
import json
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from pstx_apps.offline_migration import (
    DEFAULT_MINERU_MODEL_SOURCE,
    DEFAULT_MINERU_MODEL_TYPE,
    DEFAULT_MINERU_WHEEL_SPEC,
    build_python_download_url,
    prepare_offline_bundle,
    verify_offline_bundle,
)
from pstx_core import pages as page_logic
from pstx_core.cadence import csa_geometry
from pstx_core.cadence.csa_connectivity_overlay import build_csa_connectivity_overlay
from pstx_core.cadence.page_model import build_cadence_page_payload
from pstx_core.cadence.semantic_index import build_cadence_index_payload
from pstx_core.schematic_pdf_annotation import (
    build_schematic_pdf_annotation_payload,
    load_json_mapping_or_sequence,
    load_targets_json,
)
from pstx_agent_runtime import (
    AgentDurableRunStore,
    COMPARE_AGENT_PLAYBOOKS,
    REPORT_AGENT_PLAYBOOKS,
    load_harness_skills,
    select_harness_skills,
)
from pstx_core.page_resolution import component_user_visible_page
from pstx_queries.project_query import query_project_data
from pstx_knowledge.business_dictionary import business_dictionary_summary
from pstx_knowledge.datasheet_review_templates import (
    get_datasheet_review_template,
    list_datasheet_review_templates,
)
from pstx_knowledge.datasheets import (
    build_datasheet_status,
    list_datasheet_documents,
    search_datasheet_chunks,
    search_datasheet_parameters,
)
from pstx_knowledge.topology import build_llm_topology_netlist
from pstx_rules.module_scope import filter_module_review
from pstx_rules.project_analysis import analyze_project_contents, append_analysis_timing
from pstx_webapp.compare_payload import build_compare_payload
from pstx_webapp.compare_view import coerce_compare_detail_limit
from pstx_webapp.form_parsing import parse_voltage_map_text
from pstx_webapp.project_io import _is_supported_archive, discover_project_files_with_snapshot, read_local_text_file
from pstx_webapp.report_view import build_report_payload


CLI_VERSION = "1"
SCHEMA_VERSION = "pstx-cli.v1"


class CliArgumentError(ValueError):
    """Argument parsing error that should be returned as a JSON envelope."""


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:  # pragma: no cover - routed through main.
        raise CliArgumentError(message)


CLI_COMMAND_SCHEMAS: Dict[str, Dict[str, Any]] = {
    "capabilities": {
        "purpose": "列出当前 CLI 可用能力。",
        "inputs": [],
        "outputs": ["capabilities", "notes"],
        "cache": "none",
    },
    "schema": {
        "purpose": "输出机器可读 CLI 命令协议说明。",
        "inputs": ["command?"],
        "outputs": ["commands", "schema"],
        "cache": "none",
    },
    "analyze": {
        "purpose": "分析一个 PSTX 项目，并可写出 bundle/report/excel/cache。",
        "inputs": ["project_root|project_container|archive", "--include-depop", "--include-total-bom", "--ratio-limit", "--custom-volt-map"],
        "outputs": ["summary", "module_scope", "written", "bundle?", "report?", "module_review?"],
        "cache": "writes --bundle-cache-out",
    },
    "inspect": {
        "purpose": "快速检查项目目录、关键文件和建议下一步 CLI 调用。",
        "inputs": ["project_root|project_container|archive"],
        "outputs": ["project", "files", "page_sources", "suggested_workflow"],
        "cache": "none",
    },
    "query": {
        "purpose": "查询单个位号或网络。",
        "inputs": ["project_root? or --bundle-cache-in", "--mode 位号|网络", "--keyword"],
        "outputs": ["summary", "module_scope", "query"],
        "cache": "reads --bundle-cache-in",
    },
    "batch-query": {
        "purpose": "一次查询多个位号、网络、HQ 料号或页码。",
        "inputs": ["project_root? or --bundle-cache-in", "--mode 位号|网络|HQ料号|页码", "--items or --items-file"],
        "outputs": ["summary", "module_scope", "requested_count", "found_count", "missing_count", "results"],
        "cache": "reads --bundle-cache-in",
    },
    "module-review": {
        "purpose": "输出 module_order 主模块/子模块视角。",
        "inputs": ["project_root? or --bundle-cache-in", "--module-id?", "--module-name?", "--module-type?"],
        "outputs": ["summary", "module_review"],
        "cache": "reads --bundle-cache-in",
    },
    "report-table": {
        "purpose": "列出报告表格，或分页读取一个 table_id 的行。",
        "inputs": ["project_root? or --bundle-cache-in", "--table-id?", "--offset", "--limit", "--module-*"],
        "outputs": ["summary", "module_scope", "tables", "table?"],
        "cache": "reads --bundle-cache-in",
    },
    "report-aggregate": {
        "purpose": "对报告表格某列做本地确定性聚合。",
        "inputs": ["project_root? or --bundle-cache-in", "--table-id", "--column", "--operation top|count|unique"],
        "outputs": ["summary", "module_scope", "table", "aggregation"],
        "cache": "reads --bundle-cache-in",
    },
    "evidence-pack": {
        "purpose": "按位号、网络、HQ 料号、页码和报告表格一次性打包模型可读证据。",
        "inputs": ["project_root? or --bundle-cache-in", "--refdes?", "--net?", "--hq?", "--page?", "--table-id?"],
        "outputs": ["summary", "module_scope", "evidence_pack"],
        "cache": "reads --bundle-cache-in",
    },
    "net-catalog": {
        "purpose": "列出/筛选项目网标目录，帮助外部 Agent 先发现可查询的网络，再按需拉详情证据。",
        "inputs": ["project_root? or --bundle-cache-in", "--query?", "--kind?", "--min-nodes?", "--limit", "--offset", "--include-nodes"],
        "outputs": ["summary", "module_scope", "net_catalog"],
        "cache": "reads --bundle-cache-in",
    },
    "topology-netlist": {
        "purpose": "导出芯片级 LLM 语义拓扑网表，供外部 Agent/LLM 快速理解大芯片连接关系。",
        "inputs": ["project_root? or --bundle-cache-in", "--focus-refdes?", "--include-connectors", "--limit", "--view summary|full", "--supply-mode grouped|details|hidden", "--supply-limit", "--out?"],
        "outputs": ["summary", "module_scope", "topology_netlist", "written"],
        "cache": "reads --bundle-cache-in",
    },
    "cadence-page": {
        "purpose": "读取单页 Cadence pageX.csv|csa 连接语义模型摘要、对象列表或对象详情。",
        "inputs": ["project_root? or --bundle-cache-in", "--page", "--stdout summary|objects|full", "--object-id?", "--limit"],
        "outputs": ["summary", "cadence_page"],
        "cache": "reads --bundle-cache-in",
    },
    "cadence-index": {
        "purpose": "汇总项目级 Cadence 页图语义索引，按网络标签、端口、跨页连接、Bus、No Connect 和 unbound 语义取证。",
        "inputs": ["project_root? or --bundle-cache-in", "--stdout summary|nets|ports|links|full", "--query?", "--kind all|net|port|offpage|bus|no_connect|unbound", "--page?", "--limit"],
        "outputs": ["summary", "cadence_index"],
        "cache": "reads --bundle-cache-in",
    },
    "csa-geometry": {
        "purpose": "扫描 Cadence DE HDL pageX.csa 几何对象，输出 DOT 四向十字、CIRCLE/可选 ARC 画圈对象和礼包同名 CSV/JSON。",
        "inputs": [
            "project_root|sch_1|pageX.csa? or --bundle-cache-in",
            "--recursive",
            "--workers?",
            "--executor thread|process|serial",
            "--include-arcs",
            "--circle-two-point-mode center_radius|bbox",
            "--check-missing",
            "--include-connectivity",
            "--page?",
            "--out-dir?",
            "--json",
            "--html",
            "--fail-on-findings",
            "--fail-on-circles",
            "--stdout summary|hits|details|full",
            "--demo",
        ],
        "outputs": ["summary", "csa_geometry", "written", "demo?"],
        "cache": "reads --bundle-cache-in",
    },
    "schematic-pdf-annotate": {
        "purpose": "将位号、网络、页码、坐标等 review target 定位到原理图 PDF，输出可绘制 overlay JSON。",
        "inputs": [
            "pdf",
            "project_root? or --bundle-cache-in",
            "--targets-json?",
            "--target-json?",
            "--refdes?",
            "--net?",
            "--page?",
            "--pdf-page-map-json?",
            "--calibrations-json?",
            "--allow-page-number-fallback",
            "--stdout summary|annotations|full",
            "--limit",
        ],
        "outputs": ["summary", "schematic_pdf_annotation"],
        "cache": "reads --bundle-cache-in",
    },
    "business-dictionary": {
        "purpose": "输出项目内置业务词典、接口别名、角色别名和 review focus，供外部 Agent 统一口径。",
        "inputs": ["--json-out?"],
        "outputs": ["business_dictionary", "usage"],
        "cache": "none",
    },
    "harness-skills": {
        "purpose": "输出 Harness Agent 可读取的 skill 卡，可按问题/profile 选中并返回完整 body，供 Trae/外部 Agent 共用取证路线。",
        "inputs": ["skill_id?", "--query?", "--capability-profile?", "--playbook?", "--tool?", "--include-body", "--max-body-chars?", "--json-out?"],
        "outputs": ["harness_skills"],
        "cache": "none",
    },
    "datasheet-status": {
        "purpose": "查看本地 datasheet PDF 索引、MinerU/pypdf 后端、参数卡和 chunk 统计。",
        "inputs": ["--limit?", "--offset?", "--json-out?"],
        "outputs": ["datasheet_status", "documents?"],
        "cache": "reads datasheet SQLite index",
    },
    "datasheet-search": {
        "purpose": "按关键词搜索已索引 datasheet chunk，返回可引用片段摘要。",
        "inputs": ["--query", "--limit?", "--offset?", "--json-out?"],
        "outputs": ["datasheet_search"],
        "cache": "reads datasheet SQLite index",
    },
    "datasheet-parameters": {
        "purpose": "搜索确定性抽取的 datasheet 参数卡，适合电压、电流、热、时序等参数级证据。",
        "inputs": ["--query?", "--parameter-key?", "--doc-id?", "--limit?", "--offset?", "--json-out?"],
        "outputs": ["datasheet_parameters"],
        "cache": "reads datasheet SQLite index",
    },
    "datasheet-template": {
        "purpose": "输出 LLM 可读的 datasheet 审查模板，用于把 PDF 证据映射到原理图 review 检查点。",
        "inputs": ["template_id?", "--category?", "--without-questions?", "--json-out?"],
        "outputs": ["datasheet_template or datasheet_templates"],
        "cache": "none",
    },
    "compare": {
        "purpose": "分析并对比两个 PSTX 项目。",
        "inputs": [
            "left_project_root|left_project_container|left_archive",
            "right_project_root|right_project_container|right_archive",
            "--detail-limit",
            "--include-depop",
            "--include-total-bom",
        ],
        "outputs": ["compare"],
        "cache": "none",
    },
    "agent-run-status": {
        "purpose": "读取 agent_workspace 中后台 Agent run 的 durable 状态。",
        "inputs": ["agent_run_id"],
        "outputs": ["agent_run_status"],
        "cache": "reads agent_workspace",
    },
    "agent-run-artifacts": {
        "purpose": "列出 agent_workspace 中某个 Agent run 的草稿/报告 artifact。",
        "inputs": ["agent_run_id"],
        "outputs": ["agent_run_artifacts"],
        "cache": "reads agent_workspace",
    },
    "agent-run-trace": {
        "purpose": "读取 agent_workspace 中某个 Agent run 的 partial/final trace。",
        "inputs": ["agent_run_id"],
        "outputs": ["agent_run_trace"],
        "cache": "reads agent_workspace",
    },
    "offline-migration": {
        "purpose": "电脑 A 准备离线迁移包，电脑 B 离线校验 Python/MinerU/依赖/项目文件完整性。",
        "inputs": [
            "build-python-url|prepare|verify",
            "--out-dir?",
            "--name?",
            "--target-platform?",
            "--target-profile?",
            "--python-url?",
            "--python-version?",
            "--python-mirror official|tuna|npmmirror",
            "--python-mirror-base?",
            "--python-filename?",
            "--python-archive?",
            "--python-dir?",
            "--no-extract-python",
            "--allow-system-python-on-b",
            "--mineru-venv?",
            "--mineru-model-dir?",
            "--mineru-config?",
            "--download-mineru-models",
            "--mineru-model-source huggingface|modelscope",
            "--mineru-model-type pipeline|vlm|all",
            "--huggingface-endpoint?",
            "--mineru-model-downloader?",
            "--download-wheels",
            "--pip-index-url?",
            "--pip-extra-index-url?",
            "--include-mineru-wheels",
            "--mineru-wheel-spec?",
            "--strict-mineru-wheels",
            "--asset-cache-dir?",
            "--no-reuse-assets",
            "--include-datasheet-source",
            "--skip-datasheet-data",
            "--no-zip",
            "--skip-runtime-probe",
        ],
        "outputs": ["offline_migration", "written?", "verification?"],
        "cache": "writes/reads offline package and _asset_cache",
    },
}


def _json_default(value: Any):
    if isinstance(value, Path):
        return str(value)
    return str(value)


def _write_json_file(path: Optional[str], payload: Dict[str, Any], *, pretty: bool) -> Optional[str]:
    if not path:
        return None
    target = Path(path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2 if pretty else None, default=_json_default),
        encoding="utf-8",
    )
    return str(target)


def _read_json_file(path: str) -> Dict[str, Any]:
    target = Path(path).expanduser()
    if not target.is_file():
        raise FileNotFoundError(f"JSON 文件不存在：{target}")
    data = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"JSON 文件必须是对象：{target}")
    return data


def _read_bundle_cache(path: str) -> Dict[str, Any]:
    data = _read_json_file(path)
    bundle = data.get("bundle") if isinstance(data.get("bundle"), dict) else data
    if not isinstance(bundle, dict) or "components" not in bundle or "nets" not in bundle:
        raise ValueError(f"不是有效的 PSTX bundle 缓存：{Path(path).expanduser()}")
    loaded = dict(bundle)
    loaded["_cli_bundle_cache"] = {
        "loaded": True,
        "path": str(Path(path).expanduser()),
    }
    return loaded


def _emit(payload: Dict[str, Any], *, pretty: bool) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2 if pretty else None, default=_json_default))


def _read_project_inputs(project_root: str) -> Dict[str, Any]:
    root, prt_path, net_path, ref_path, snapshot_meta = discover_project_files_with_snapshot(project_root)
    prt_content, prt_info = read_local_text_file(prt_path, "pstxprt.dat", required=True)
    net_content, net_info = read_local_text_file(net_path, "pstxnet.dat", required=True)
    ref_content, ref_info = read_local_text_file(ref_path, "pstxref.dat", required=False) if ref_path else (None, {
        "label": "pstxref.dat",
        "filename": "",
        "size": "0",
        "encoding": "",
    })
    return {
        "project_root": root,
        "project_name": root.name,
        "prt_content": prt_content or "",
        "net_content": net_content or "",
        "ref_content": ref_content or "",
        "files": [prt_info, net_info, ref_info],
        "snapshot": snapshot_meta,
    }


def _file_status(path: Optional[Path], *, label: str, required: bool = False) -> Dict[str, Any]:
    exists = bool(path and Path(path).is_file())
    return {
        "label": label,
        "path": str(path) if path else "",
        "exists": exists,
        "required": required,
        "size": Path(path).stat().st_size if exists else 0,
    }


def _inspect_project_root(project_root: str) -> Dict[str, Any]:
    raw = str(project_root or "").strip().strip('"')
    if not raw:
        raise ValueError("请输入项目根路径")
    root = Path(raw).expanduser()
    snapshot_meta: Dict[str, Any] = {}
    try:
        resolved_root, _prt_path, _net_path, _ref_path, snapshot_meta = discover_project_files_with_snapshot(raw)
        root = resolved_root
    except Exception:
        if root.exists() and (
            root.is_file()
            or (root.is_dir() and list(root.glob("*.cpm")))
        ):
            raise
        if root.name.lower() == "packaged":
            root = root.parent
    packaged_dir = root / "packaged"
    prt_path = packaged_dir / "pstxprt.dat"
    net_path = packaged_dir / "pstxnet.dat"
    ref_path = packaged_dir / "pstxref.dat"
    module_order_candidates = [
        root / "module_order",
        root / "module_order.dat",
        root / "sch_1" / "module_order.dat",
    ]
    module_order_files = [path for path in module_order_candidates if path.is_file()]
    sch_dir = root / "sch_1"
    page_csv_count = len(list(sch_dir.glob("page*.csv"))) if sch_dir.is_dir() else 0
    page_csa_count = len(list(sch_dir.glob("page*.csa"))) if sch_dir.is_dir() else 0
    cache_example = "out/bundle-cache.json"
    return {
        "project": {
            "root": str(root),
            "name": root.name,
            "exists": root.exists(),
            "is_directory": root.is_dir(),
            "packaged_exists": packaged_dir.is_dir(),
            "input": raw,
            "snapshot": snapshot_meta,
        },
        "files": [
            _file_status(prt_path, label="pstxprt.dat", required=True),
            _file_status(net_path, label="pstxnet.dat", required=True),
            _file_status(ref_path, label="pstxref.dat", required=False),
        ],
        "page_sources": {
            "sch_1": str(sch_dir),
            "sch_1_exists": sch_dir.is_dir(),
            "page_csv_count": page_csv_count,
            "page_csa_count": page_csa_count,
            "module_order_files": [str(path) for path in module_order_files],
            "module_order_available": bool(module_order_files),
        },
        "suggested_workflow": [
            f"python pstx_cli.py analyze {json.dumps(str(root), ensure_ascii=False)} --bundle-cache-out {cache_example}",
            f"python pstx_cli.py evidence-pack --bundle-cache-in {cache_example} --refdes U1,U2 --table-id chip_pin_rows",
            f"python pstx_cli.py report-table --bundle-cache-in {cache_example} --table-id chip_pin_rows --offset 0 --limit 200",
            f"python pstx_cli.py report-aggregate --bundle-cache-in {cache_example} --table-id page_rows --column 页码",
        ],
    }


def _analysis_source_args_for_command(args: argparse.Namespace) -> str:
    cache_path = str(getattr(args, "bundle_cache_in", "") or "").strip()
    if cache_path:
        return f"--bundle-cache-in {json.dumps(cache_path, ensure_ascii=False)}"
    project_root = str(getattr(args, "project_root", "") or "").strip()
    if project_root:
        return json.dumps(project_root, ensure_ascii=False)
    return "--bundle-cache-in <bundle-cache.json>"


def _analysis_summary(bundle: Dict[str, Any]) -> Dict[str, Any]:
    module_summary = (bundle.get("module_review") or {}).get("summary", {})
    return {
        "project_name": bundle.get("project_name", ""),
        "project_root": bundle.get("project_root", ""),
        "component_count": len(bundle.get("components", {}) or {}),
        "net_count": len(bundle.get("nets", {}) or {}),
        "all_component_count": len(bundle.get("all_components", {}) or {}),
        "all_net_count": len(bundle.get("all_nets", {}) or {}),
        "module_summary": module_summary,
        "include_depop": bool(bundle.get("include_depop", False)),
        "include_total_bom": bool(bundle.get("include_total_bom", False)),
        "bundle_cache": dict(bundle.get("_cli_bundle_cache", {}) or {}),
        "analysis_timings": dict(bundle.get("analysis_timings", {}) or {}),
        "project_input_snapshot": dict(bundle.get("project_input_snapshot", {}) or {}),
        "warnings": list(bundle.get("page_warnings", []) or []),
    }


def _analyze_project_from_args(args: argparse.Namespace) -> Dict[str, Any]:
    if getattr(args, "bundle_cache_in", ""):
        return _read_bundle_cache(args.bundle_cache_in)
    project_root = str(getattr(args, "project_root", "") or "").strip()
    if not project_root:
        raise ValueError("缺少 project_root；使用缓存时请提供 --bundle-cache-in。")
    inputs = _read_project_inputs(project_root)
    volt_map = None
    voltage_warnings = []
    if getattr(args, "custom_volt_map", ""):
        volt_map, voltage_warnings = parse_voltage_map_text(args.custom_volt_map)
    bundle = analyze_project_contents(
        inputs["prt_content"],
        inputs["net_content"],
        project_name=args.project_name or inputs["project_name"],
        project_root=str(inputs["project_root"]),
        ratio_limit=float(getattr(args, "ratio_limit", 70.0)),
        custom_volt_map=volt_map,
        include_depop=bool(getattr(args, "include_depop", False)),
        include_total_bom=bool(getattr(args, "include_total_bom", False)),
    )
    if voltage_warnings:
        bundle.setdefault("page_warnings", []).extend(voltage_warnings)
    bundle["input_files"] = inputs["files"]
    bundle["project_input_snapshot"] = inputs.get("snapshot", {})
    snapshot_warnings = (inputs.get("snapshot", {}) or {}).get("warnings", [])
    if snapshot_warnings:
        bundle.setdefault("page_warnings", []).extend(snapshot_warnings)
    return bundle


def _report_payload_for_bundle(bundle: Dict[str, Any], run_id: str = "cli-run") -> Dict[str, Any]:
    return build_report_payload(run_id, bundle)


def _module_review_for_args(bundle: Dict[str, Any], args: argparse.Namespace) -> Dict[str, Any]:
    return filter_module_review(
        bundle.get("module_review", {}) or {},
        module_id=getattr(args, "module_id", ""),
        module_name=getattr(args, "module_name", ""),
        module_type=getattr(args, "module_type", "all"),
    )


def _module_filter_active(args: argparse.Namespace) -> bool:
    return bool(
        getattr(args, "module_id", "")
        or getattr(args, "module_name", "")
        or getattr(args, "module_type", "all") != "all"
    )


def _selected_module_refdes(module_review: Dict[str, Any]) -> Set[str]:
    return {
        str(row.get("位号", "")).strip()
        for row in (module_review or {}).get("component_rows", [])
        if isinstance(row, dict) and str(row.get("位号", "")).strip()
    }


def _row_refdes_values(row: Dict[str, Any]) -> Set[str]:
    refdes_keys = {
        "位号",
        "refdes",
        "RefDes",
        "REFDES",
        "元件",
        "元件位号",
        "位号列表",
        "器件位号",
        "芯片位号",
        "关联芯片",
        "起点位号",
        "终点位号",
    }
    values: Set[str] = set()
    for key, value in (row or {}).items():
        if str(key) not in refdes_keys:
            continue
        text = str(value or "")
        for chunk in text.replace("，", ",").replace(";", ",").split(","):
            token = chunk.strip()
            if token:
                values.add(token)
    return values


def _filter_rows_by_refdes(rows: List[Dict[str, Any]], selected_refdes: Set[str]) -> List[Dict[str, Any]]:
    if not selected_refdes:
        return []
    selected_upper = {item.upper() for item in selected_refdes}
    filtered: List[Dict[str, Any]] = []
    for row in rows:
        row_refdes = _row_refdes_values(row)
        if not row_refdes:
            filtered.append(row)
            continue
        if any(refdes.upper() in selected_upper for refdes in row_refdes):
            filtered.append(row)
    return filtered


def _iter_report_tables(report: Dict[str, Any]) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []
    for section in report.get("sections", []) or []:
        for table in section.get("tables", []) or []:
            item = dict(table)
            item["section_id"] = section.get("id", "")
            item["section_title"] = section.get("title", "")
            tables.append(item)
    return tables


def _report_table_summary(table: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "section_id": table.get("section_id", ""),
        "section_title": table.get("section_title", ""),
        "table_id": table.get("id", ""),
        "title": table.get("title", ""),
        "count": int(table.get("count", 0) or 0),
        "columns": list(table.get("columns", []) or []),
        "kind_counts": dict(table.get("kind_counts", {}) or {}),
    }


def _report_table_for_args(bundle: Dict[str, Any],
                           args: argparse.Namespace,
                           table_id: str) -> tuple[Dict[str, Any], Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]]]:
    report = _report_payload_for_bundle(bundle)
    module_review = _module_review_for_args(bundle, args)
    selected_refdes = _selected_module_refdes(module_review) if _module_filter_active(args) else set()
    tables = _iter_report_tables(report)
    table = next((item for item in tables if item.get("id") == table_id), None)
    if table is None:
        raise ValueError(f"unknown table_id: {table_id}")
    all_rows = list(table.get("rows", []) or [])
    if _module_filter_active(args):
        if table_id == "module_scope_rows":
            all_rows = list(module_review.get("module_rows", []) or [])
        elif table_id == "module_component_rows":
            all_rows = list(module_review.get("component_rows", []) or [])
        else:
            all_rows = _filter_rows_by_refdes(all_rows, selected_refdes)
    return module_review, table, all_rows, tables


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple, dict)):
        return json.dumps(value, ensure_ascii=False, default=_json_default)
    return str(value).strip()


def _aggregate_rows(rows: List[Dict[str, Any]],
                    *,
                    column: str,
                    operation: str,
                    limit: int,
                    include_empty: bool) -> Dict[str, Any]:
    if rows and column not in rows[0]:
        available = sorted({key for row in rows for key in row.keys()})
        raise ValueError(f"unknown column: {column}; available columns: {', '.join(available[:30])}")
    values: List[str] = []
    for row in rows:
        value = _cell_text(row.get(column, ""))
        if value or include_empty:
            values.append(value)
    counts = Counter(values)
    if operation == "top":
        ordered = sorted(counts.items(), key=lambda item: (-item[1], item[0]))
    else:
        ordered = sorted(counts.items(), key=lambda item: item[0])
    items = [{"value": value, "count": count} for value, count in ordered[:limit]]
    return {
        "column": column,
        "operation": operation,
        "row_count": len(rows),
        "non_empty_count": sum(1 for value in values if value),
        "empty_count": sum(1 for value in values if not value),
        "unique_count": len(counts),
        "limit": limit,
        "truncated": len(ordered) > limit,
        "items": items,
    }


def _analysis_scope_for_args(bundle: Dict[str, Any], args: argparse.Namespace) -> tuple[Dict[str, Any], Dict[str, Any], Dict[str, Any]]:
    module_review = _module_review_for_args(bundle, args)
    components = dict(bundle.get("components", {}) or {})
    nets = dict(bundle.get("nets", {}) or {})
    if _module_filter_active(args):
        selected_refdes = _selected_module_refdes(module_review)
        components = {refdes: comp for refdes, comp in components.items() if refdes in selected_refdes}
        nets = {
            net_name: [
                node for node in nodes
                if str(node.get("refdes", "")) in selected_refdes
            ]
            for net_name, nodes in nets.items()
        }
        nets = {net_name: nodes for net_name, nodes in nets.items() if nodes}
    return module_review, components, nets


def _parse_batch_items(args: argparse.Namespace) -> List[str]:
    chunks: List[str] = []
    raw_items = str(getattr(args, "items", "") or "")
    if raw_items:
        chunks.append(raw_items)
    items_file = str(getattr(args, "items_file", "") or "").strip()
    if items_file:
        text = Path(items_file).expanduser().read_text(encoding="utf-8")
        stripped = text.strip()
        if stripped.startswith("["):
            loaded = json.loads(stripped)
            if not isinstance(loaded, list):
                raise ValueError("--items-file JSON must be an array")
            chunks.extend(str(item) for item in loaded)
        else:
            chunks.append(text)
    items: List[str] = []
    seen: Set[str] = set()
    for chunk in chunks:
        for part in str(chunk).replace("\n", ",").replace("，", ",").split(","):
            item = part.strip()
            if not item:
                continue
            key = item.upper()
            if key in seen:
                continue
            seen.add(key)
            items.append(item)
    if not items:
        raise ValueError("batch-query requires --items or --items-file")
    max_items = max(1, min(500, int(getattr(args, "max_items", 100) or 100)))
    return items[:max_items]


def _split_cli_values(raw_values: Any, *, max_items: int = 100) -> List[str]:
    values = raw_values if isinstance(raw_values, list) else [raw_values]
    output: List[str] = []
    seen: Set[str] = set()
    for value in values:
        if value is None:
            continue
        for part in str(value).replace("\n", ",").replace("，", ",").split(","):
            item = part.strip()
            if not item:
                continue
            key = item.upper()
            if key in seen:
                continue
            seen.add(key)
            output.append(item)
            if len(output) >= max_items:
                return output
    return output


def _trim_query_result(result: Dict[str, Any], limit: int) -> Dict[str, Any]:
    trimmed = dict(result)
    items = list(trimmed.get("items", []) or [])
    trimmed["items"] = items[:limit]
    trimmed["items_truncated"] = len(items) > limit
    cards = []
    for card in trimmed.get("cards", []) or []:
        card_item = dict(card)
        card_items = list(card_item.get("items", []) or [])
        card_item["items"] = card_items[:limit]
        card_item["total_count"] = len(card_items)
        card_item["truncated"] = len(card_items) > limit
        cards.append(card_item)
    trimmed["cards"] = cards
    return trimmed


def _query_result_count(result: Dict[str, Any]) -> int:
    if result.get("match_type") == "missing":
        return 0
    if result.get("entity_type") == "component" and result.get("match_type") == "exact":
        return 1
    if result.get("entity_type") == "network" and result.get("match_type") == "exact":
        for item in (result.get("summary", {}) or {}).get("meta", []) or []:
            if item.get("label") == "节点数":
                try:
                    return int(item.get("value", 0) or 0)
                except (TypeError, ValueError):
                    return 1
        return 1
    return len(result.get("items", []) or [])


def _component_summary_row(refdes: str, comp: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "位号": refdes,
        "类型": str(comp.get("comp_type", "") or ""),
        "HQ料号": str(comp.get("hq_code", "") or ""),
        "VALUE": str(comp.get("value", "") or comp.get("part_name", "") or ""),
        "封装": str(comp.get("package", "") or ""),
        "页码": component_user_visible_page(comp),
        "BOM_OPTION": str(comp.get("bom_option", "") or ""),
    }


def _query_by_hq(components: Dict[str, Any], query: str, limit: int) -> Dict[str, Any]:
    needle = str(query or "").strip().upper()
    rows = [
        _component_summary_row(refdes, comp)
        for refdes, comp in sorted(components.items(), key=lambda item: item[0])
        if needle and needle in str(comp.get("hq_code", "") or "").upper()
    ]
    return {
        "mode": "HQ料号",
        "query": query,
        "status": "found" if rows else "missing",
        "result_count": len(rows),
        "truncated": len(rows) > limit,
        "items": rows[:limit],
    }


def _query_by_page(components: Dict[str, Any], query: str, limit: int) -> Dict[str, Any]:
    raw = str(query or "").strip()
    target = page_logic.normalize_page_label(raw if raw.upper().startswith("PAGE") else f"PAGE{raw}")
    rows = [
        _component_summary_row(refdes, comp)
        for refdes, comp in sorted(components.items(), key=lambda item: item[0])
        if component_user_visible_page(comp).upper() == target.upper()
    ]
    return {
        "mode": "页码",
        "query": query,
        "normalized_query": target,
        "status": "found" if rows else "missing",
        "result_count": len(rows),
        "truncated": len(rows) > limit,
        "items": rows[:limit],
    }


def _query_alias_terms(query: str) -> List[str]:
    raw = str(query or "").strip()
    if not raw:
        return []
    terms = [raw]
    upper = raw.upper()
    dictionary = business_dictionary_summary()
    for interface_id, aliases in (dictionary.get("interface_aliases") or {}).items():
        alias_values = [str(alias) for alias in aliases or []]
        if upper == str(interface_id).upper() or any(upper == alias.upper() for alias in alias_values):
            terms.extend(alias_values)
            terms.append(str(interface_id))
            continue
        if any(upper in alias.upper() or alias.upper() in upper for alias in alias_values):
            terms.extend(alias_values)
            terms.append(str(interface_id))
    output: List[str] = []
    seen: Set[str] = set()
    for term in terms:
        cleaned = str(term or "").strip()
        key = cleaned.upper()
        if cleaned and key not in seen:
            seen.add(key)
            output.append(cleaned)
    return output


def _is_ground_net(net_name: str) -> bool:
    upper = str(net_name or "").upper()
    return bool(re.search(r"(^|[_\-.+])(GND|AGND|DGND|PGND|SGND|VSS|0V)([_\-.+]|$)", upper))


def _is_power_net(net_name: str) -> bool:
    upper = str(net_name or "").upper()
    if _is_ground_net(upper):
        return False
    if re.search(r"(^|[_\-.+])(VDD|VCC|VEE|VTT|VIN|VOUT|VBAT|AVDD|DVDD|PVDD)([_\-.+]|$)", upper):
        return True
    if re.search(r"(^|[_\-.+])P\d+V\d*([A-Z0-9_+\-.]*)([_\-.+]|$)", upper):
        return True
    if re.search(r"^[+-]?\d+(\.\d+)?V([A-Z0-9_+\-.]*)$", upper):
        return True
    return False


def _is_differential_net(net_name: str) -> bool:
    upper = str(net_name or "").upper()
    if re.search(r"(^|[_\-.])(TX|RX|CLK|REFCLK|D|DATA|LANE)\d*([PN])($|[_\-.])", upper):
        return True
    if re.search(r"(^|[_\-.])(DP|DN|DM|DP\d+|DN\d+)($|[_\-.])", upper):
        return True
    high_speed_markers = ("PCIE", "PCE", "P5E", "P4E", "P3E", "P2E", "P1E", "USB", "MIPI", "LVDS", "HDMI", "SATA", "SGMII")
    if any(marker in upper for marker in high_speed_markers) and re.search(r"([_+\-.][PN]|\d[PN])($|[_+\-.])", upper):
        return True
    return False


def _is_unnamed_net(net_name: str) -> bool:
    upper = str(net_name or "").upper()
    return bool(
        re.search(r"(^|[_\-.])(NET|N|UNNAMED|NO_NAME|NONAME)\d*($|[_\-.])", upper)
        or re.search(r"^N\d{3,}$", upper)
    )


def _net_kind(net_name: str) -> str:
    if _is_ground_net(net_name):
        return "ground"
    if _is_power_net(net_name):
        return "power"
    if _is_differential_net(net_name):
        return "differential"
    if _is_unnamed_net(net_name):
        return "unnamed"
    return "signal"


def _component_display_type(comp: Dict[str, Any]) -> str:
    return str(
        comp.get("comp_type")
        or comp.get("CDS_PART_NAME")
        or comp.get("part_name")
        or comp.get("value")
        or comp.get("VALUE")
        or ""
    )


def _net_node_summary(node: Dict[str, Any], components: Dict[str, Any]) -> Dict[str, Any]:
    refdes = str(node.get("refdes", "") or "").strip()
    comp = components.get(refdes, {}) if refdes else {}
    return {
        "refdes": refdes,
        "pin": str(node.get("pin", "") or node.get("pin_number", "") or "").strip(),
        "pin_name": str(node.get("pin_name", "") or node.get("name", "") or "").strip(),
        "page": component_user_visible_page(comp) if comp else "",
        "component_type": _component_display_type(comp) if comp else "",
    }


def _net_catalog_item(net_name: str,
                      nodes: List[Dict[str, Any]],
                      components: Dict[str, Any],
                      args: argparse.Namespace,
                      *,
                      include_nodes: bool) -> Dict[str, Any]:
    kind = _net_kind(net_name)
    node_summaries = [_net_node_summary(node, components) for node in nodes if isinstance(node, dict)]
    refdes_sample = list(dict.fromkeys(item["refdes"] for item in node_summaries if item["refdes"]))[:10]
    pin_sample = list(dict.fromkeys(item["pin"] for item in node_summaries if item["pin"]))[:10]
    page_sample = list(dict.fromkeys(item["page"] for item in node_summaries if item["page"]))[:10]
    component_type_sample = list(dict.fromkeys(item["component_type"] for item in node_summaries if item["component_type"]))[:8]
    item: Dict[str, Any] = {
        "net_name": net_name,
        "kind": kind,
        "node_count": len(nodes),
        "refdes_sample": refdes_sample,
        "pin_sample": pin_sample,
        "page_sample": page_sample,
        "component_type_sample": component_type_sample,
        "is_power": kind == "power",
        "is_ground": kind == "ground",
        "is_differential": kind == "differential",
        "is_unnamed": kind == "unnamed",
        "detail_command": (
            f"python3 pstx_cli.py evidence-pack {_analysis_source_args_for_command(args)} "
            f"--net {json.dumps(net_name, ensure_ascii=False)} --table-id chip_pin_rows --pretty"
        ),
    }
    if include_nodes:
        item["nodes"] = node_summaries[:50]
        item["nodes_truncated"] = len(node_summaries) > 50
    return item


def _build_net_catalog(bundle: Dict[str, Any], args: argparse.Namespace) -> Dict[str, Any]:
    _module_review, components, nets = _analysis_scope_for_args(bundle, args)
    query = str(getattr(args, "query", "") or "").strip()
    query_terms = _query_alias_terms(query)
    query_terms_upper = [term.upper() for term in query_terms]
    kind_filter = str(getattr(args, "kind", "all") or "all")
    min_nodes = max(1, int(getattr(args, "min_nodes", 1) or 1))
    limit = max(1, min(5000, int(getattr(args, "limit", 100) or 100)))
    offset = max(0, int(getattr(args, "offset", 0) or 0))
    include_nodes = bool(getattr(args, "include_nodes", False))
    all_items: List[Dict[str, Any]] = []
    for net_name, raw_nodes in sorted((nets or {}).items(), key=lambda item: str(item[0]).upper()):
        nodes = [node for node in (raw_nodes or []) if isinstance(node, dict)]
        if len(nodes) < min_nodes:
            continue
        kind = _net_kind(str(net_name))
        if kind_filter != "all" and kind != kind_filter:
            continue
        if query_terms_upper and not any(term in str(net_name).upper() for term in query_terms_upper):
            continue
        all_items.append(_net_catalog_item(str(net_name), nodes, components, args, include_nodes=include_nodes))
    kind_counts = Counter(_net_kind(str(net_name)) for net_name in (nets or {}).keys())
    paged_items = all_items[offset:offset + limit]
    return {
        "schema_version": "pstx-net-catalog.v1",
        "total_net_count": len(nets or {}),
        "matched_count": len(all_items),
        "returned_count": len(paged_items),
        "offset": offset,
        "limit": limit,
        "truncated": offset + len(paged_items) < len(all_items),
        "filters": {
            "query": query,
            "expanded_query_terms": query_terms,
            "kind": kind_filter,
            "min_nodes": min_nodes,
            "include_nodes": include_nodes,
        },
        "kind_counts": dict(sorted(kind_counts.items())),
        "items": paged_items,
        "recommended_next_commands": [
            "Use evidence-pack --net <net_name> for full network evidence and report-table snippets.",
            "Use topology-netlist when the question is chip-level relationship rather than raw net listing.",
            "Use business-dictionary before interpreting project abbreviations such as PCE/P5E as PCIe aliases.",
            "If truncated=true, page with --offset/--limit instead of assuming the returned list is complete.",
        ],
    }


def _envelope(command: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "ok": True,
        "interface": "pstx-cli",
        "interface_version": CLI_VERSION,
        "schema_version": SCHEMA_VERSION,
        "command": command,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        **payload,
    }


def _error_code_for_exception(exc: Exception) -> str:
    if isinstance(exc, FileNotFoundError):
        return "file_not_found"
    if isinstance(exc, (ValueError, argparse.ArgumentError)):
        return "invalid_request"
    return "internal_error"


def _error_envelope(command: str, exc: Exception) -> Dict[str, Any]:
    code = _error_code_for_exception(exc)
    message = str(exc)
    return {
        "ok": False,
        "interface": "pstx-cli",
        "interface_version": CLI_VERSION,
        "schema_version": SCHEMA_VERSION,
        "command": command,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "error_code": code,
        "error_message": message,
        "error": {
            "code": code,
            "message": message,
            "type": exc.__class__.__name__,
        },
    }


def _add_common_analysis_args(parser: argparse.ArgumentParser,
                              *,
                              project_required: bool = True,
                              allow_cache_in: bool = False) -> None:
    if project_required:
        parser.add_argument("project_root", help="PSTX project root, packaged directory, CPM container, or supported archive")
    else:
        parser.add_argument("project_root", nargs="?", default="", help="PSTX project root/container/archive; optional with --bundle-cache-in")
    parser.add_argument("--project-name", default="", help="override project name")
    parser.add_argument("--ratio-limit", type=float, default=70.0, help="capacitor derating ratio limit, default 70")
    parser.add_argument("--custom-volt-map", default="", help="custom voltage map text, e.g. 'P3V3=3.3\\nVDD=1.8'")
    parser.add_argument("--include-depop", action="store_true", help="include DEPOP/DNP components in rule analysis")
    parser.add_argument("--include-total-bom", action="store_true", help="include total BOM summary")
    if allow_cache_in:
        parser.add_argument("--bundle-cache-in", default="", help="read an existing analyze bundle JSON instead of analyzing project_root")


def _add_module_filter_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--module-id", default="", help="filter module review by exact module id")
    parser.add_argument("--module-name", default="", help="filter module review by module name keyword")
    parser.add_argument("--module-type", choices=["all", "主模块", "子模块"], default="all", help="filter module review by type")


def _add_pretty_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--pretty", action="store_true", help="pretty-print JSON output")


def cmd_capabilities(args: argparse.Namespace) -> int:
    capabilities = [
        {
            "id": command_id,
            "description": schema.get("purpose", ""),
            "outputs": list(schema.get("outputs", []) or []),
            "cache": schema.get("cache", "none"),
        }
        for command_id, schema in CLI_COMMAND_SCHEMAS.items()
    ]
    skills = load_harness_skills()
    harness_runtime = {
        "skills": [
            {
                "id": skill.id,
                "title": skill.title,
                "description": skill.description,
                "capability_profiles": list(skill.capability_profiles),
                "playbooks": list(skill.playbooks),
                "allowed_tools": list(skill.allowed_tools),
            }
            for skill in skills
        ],
        "playbooks": {
            "report": [item.to_dict() for item in REPORT_AGENT_PLAYBOOKS],
            "compare": [item.to_dict() for item in COMPARE_AGENT_PLAYBOOKS],
        },
        "tool_contract_fields": [
            "completeness",
            "evidence_layers",
            "detail_tool",
            "aggregation_tool",
            "recommended_next_tools",
            "scope_summary",
        ],
        "task_memory": {
            "directory": "agent_workspace/<scope_id>/TASK.md",
            "stores": ["goal", "answer_summary", "evidence_ids", "open_questions", "next_actions"],
            "raw_content_policy": "raw tables/PDF/CSA stay in trace/store, not in Markdown task/workspace summaries",
        },
        "durable_runs": {
            "directory": "agent_workspace/<scope_id>/runs/<agent_run_id>.json",
            "statuses": ["queued", "running", "waiting_for_user", "completed", "failed", "cancelled", "incomplete"],
            "cli": ["agent-run-status", "agent-run-artifacts", "agent-run-trace"],
            "status_fields": ["current_phase", "heartbeat_at", "progress", "can_continue", "can_cancel", "partial_trace", "next_actions"],
        },
        "scratch_files": {
            "directory": "agent_workspace/<scope_id>/scratch/<agent_run_id>/",
            "schema": "pstx-agent-scratch-files/v1",
            "declared_by": "final_answer.scratch_files",
            "policy": "temporary text artifacts only; no raw large tables, PDF full text, CSA/CSV full text, credentials, or persistent business state",
            "cli": ["agent-run-artifacts"],
        },
    }
    payload = _envelope("capabilities", {
        "capabilities": capabilities,
        "harness_runtime": harness_runtime,
        "notes": [
            "CLI commands are read-only except explicitly requested output files.",
            f"All JSON envelopes include schema_version={SCHEMA_VERSION} and ok/error_code for machine parsing.",
            "All machine outputs are UTF-8 JSON or xlsx files.",
            "Web/UI and CLI share the same analysis, report and compare builders.",
            "Harness Agent v2 exposes skill cards, playbooks and tool-result contract fields for external planners.",
            "Harness Agent v3 stores background run state, artifacts, drafts and temporary scratch files under agent_workspace/.",
        ],
    })
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_agent_run_status(args: argparse.Namespace) -> int:
    store = AgentDurableRunStore()
    status = store.public_status(args.agent_run_id)
    if not status.get("ok"):
        raise ValueError(status.get("error") or f"unknown agent_run_id: {args.agent_run_id}")
    _emit(_envelope("agent-run-status", {"agent_run_status": status}), pretty=args.pretty)
    return 0


def cmd_agent_run_artifacts(args: argparse.Namespace) -> int:
    store = AgentDurableRunStore()
    artifacts = store.list_artifacts(args.agent_run_id)
    if not artifacts:
        raise ValueError(f"unknown agent_run_id: {args.agent_run_id}")
    for item in artifacts.get("artifacts", []) or []:
        path = Path(str(item.get("path") or ""))
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".json", ".jsonl", ".md", ".txt"}:
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        item["content_preview"] = text[:4000]
        item["content_truncated"] = len(text) > 4000
    _emit(_envelope("agent-run-artifacts", {"agent_run_artifacts": artifacts}), pretty=args.pretty)
    return 0


def cmd_agent_run_trace(args: argparse.Namespace) -> int:
    store = AgentDurableRunStore()
    status = store.public_status(args.agent_run_id)
    if not status.get("ok"):
        raise ValueError(status.get("error") or f"unknown agent_run_id: {args.agent_run_id}")
    record = store.read_record(args.agent_run_id)
    trace = {
        "agent_run_id": args.agent_run_id,
        "status": status.get("status", ""),
        "current_phase": status.get("current_phase", ""),
        "progress": status.get("progress") or {},
        "partial_trace": status.get("partial_trace") or {},
        "final_trace": (status.get("agent_run") or {}).get("trace_summary") if isinstance(status.get("agent_run"), dict) else {},
        "steps": record.get("steps") or [],
        "tool_calls": record.get("tool_calls") or [],
        "evidence_ids": record.get("evidence_ids") or [],
        "continuation_pack": record.get("continuation_pack") or {},
    }
    _emit(_envelope("agent-run-trace", {"agent_run_trace": trace}), pretty=args.pretty)
    return 0


def cmd_schema(args: argparse.Namespace) -> int:
    command = str(getattr(args, "schema_command", "") or "").strip()
    if command:
        schema = CLI_COMMAND_SCHEMAS.get(command)
        if not schema:
            raise ValueError(f"unknown CLI command schema: {command}")
        payload = _envelope("schema", {
            "commands": [command],
            "schema": {command: schema},
            "error_codes": ["invalid_request", "file_not_found", "internal_error"],
            "envelope_fields": ["ok", "interface", "interface_version", "schema_version", "command", "generated_at"],
        })
    else:
        payload = _envelope("schema", {
            "commands": list(CLI_COMMAND_SCHEMAS.keys()),
            "schema": CLI_COMMAND_SCHEMAS,
            "error_codes": ["invalid_request", "file_not_found", "internal_error"],
            "envelope_fields": ["ok", "interface", "interface_version", "schema_version", "command", "generated_at"],
        })
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_analyze(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    report_started = time.perf_counter()
    report = _report_payload_for_bundle(bundle)
    append_analysis_timing(bundle, "report_payload", time.perf_counter() - report_started)
    report["analysis_timings"] = bundle.get("analysis_timings", {})
    module_review = _module_review_for_args(bundle, args)
    bundle_path = _write_json_file(args.json_out, bundle, pretty=args.pretty)
    bundle_cache_path = _write_json_file(getattr(args, "bundle_cache_out", ""), bundle, pretty=args.pretty)
    report_path = _write_json_file(args.report_json_out, report, pretty=args.pretty)
    excel_path = None
    if args.excel_out:
        from pstx_exports.excel import export_to_excel

        excel_path = export_to_excel(bundle, args.excel_out)

    stdout_mode = args.stdout
    payload: Dict[str, Any] = {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "written": {
            "bundle_json": bundle_path,
            "bundle_cache": bundle_cache_path,
            "report_json": report_path,
            "excel": excel_path,
        },
    }
    if stdout_mode == "bundle":
        payload["bundle"] = bundle
    elif stdout_mode == "report":
        payload["report"] = report
    elif stdout_mode == "module-review":
        payload["module_review"] = module_review
    _emit(_envelope("analyze", payload), pretty=args.pretty)
    return 0


def cmd_inspect(args: argparse.Namespace) -> int:
    inspection = _inspect_project_root(args.project_root)
    payload = _envelope("inspect", inspection)
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_query(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review, components, nets = _analysis_scope_for_args(bundle, args)
    query = query_project_data(
        components,
        nets,
        args.mode,
        args.keyword,
    )
    payload = _envelope("query", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "query": query,
    })
    _write_json_file(args.json_out, payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_evidence_pack(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review, components, nets = _analysis_scope_for_args(bundle, args)
    limit = max(1, min(200, int(getattr(args, "limit_per_target", 20) or 20)))
    max_targets = max(1, min(200, int(getattr(args, "max_targets", 50) or 50)))
    refdes_items = _split_cli_values(getattr(args, "refdes", []), max_items=max_targets)
    net_items = _split_cli_values(getattr(args, "net", []), max_items=max_targets)
    hq_items = _split_cli_values(getattr(args, "hq", []), max_items=max_targets)
    page_items = _split_cli_values(getattr(args, "page", []), max_items=max_targets)
    table_ids = _split_cli_values(getattr(args, "table_id", []), max_items=max_targets)

    evidence_items: List[Dict[str, Any]] = []
    for item in refdes_items:
        result = _trim_query_result(query_project_data(components, nets, "位号", item), limit)
        evidence_items.append({
            "kind": "refdes",
            "query": item,
            "status": "found" if result.get("match_type") != "missing" else "missing",
            "result_count": _query_result_count(result),
            "result": result,
        })
    for item in net_items:
        result = _trim_query_result(query_project_data(components, nets, "网络", item), limit)
        evidence_items.append({
            "kind": "net",
            "query": item,
            "status": "found" if result.get("match_type") != "missing" else "missing",
            "result_count": _query_result_count(result),
            "result": result,
        })
    for item in hq_items:
        result = _query_by_hq(components, item, limit)
        evidence_items.append({
            "kind": "hq",
            "query": item,
            "status": result.get("status", "missing"),
            "result_count": result.get("result_count", 0),
            "result": result,
        })
    for item in page_items:
        result = _query_by_page(components, item, limit)
        evidence_items.append({
            "kind": "page",
            "query": item,
            "status": result.get("status", "missing"),
            "result_count": result.get("result_count", 0),
            "result": result,
        })

    table_evidence: List[Dict[str, Any]] = []
    for table_id in table_ids:
        table_module_review, table, rows, _tables = _report_table_for_args(bundle, args, table_id)
        row_limit = max(1, min(500, int(getattr(args, "table_limit", 50) or 50)))
        table_evidence.append({
            **_report_table_summary({**table, "count": len(rows)}),
            "returned_count": min(row_limit, len(rows)),
            "total_count": len(rows),
            "truncated": len(rows) > row_limit,
            "rows": rows[:row_limit],
            "module_scope": table_module_review.get("summary", {}),
            "detail_command": (
                f"python pstx_cli.py report-table {_analysis_source_args_for_command(args)} "
                f"--table-id {table_id} --offset {row_limit} --limit {row_limit}"
            ) if len(rows) > row_limit else "",
        })

    if not evidence_items and not table_evidence:
        raise ValueError("evidence-pack requires at least one of --refdes/--net/--hq/--page/--table-id")

    found_count = sum(1 for item in evidence_items if item.get("status") == "found")
    payload = _envelope("evidence-pack", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "evidence_pack": {
            "target_summary": {
                "refdes_count": len(refdes_items),
                "net_count": len(net_items),
                "hq_count": len(hq_items),
                "page_count": len(page_items),
                "table_count": len(table_evidence),
                "found_count": found_count,
                "missing_count": len(evidence_items) - found_count,
            },
            "targets": {
                "refdes": refdes_items,
                "nets": net_items,
                "hq": hq_items,
                "pages": page_items,
                "tables": table_ids,
            },
            "items": evidence_items,
            "tables": table_evidence,
            "recommended_next_commands": [
                "Use report-table with --offset/--limit for truncated table evidence.",
                "Use report-aggregate for column counts and unique values; do not infer counts from evidence-pack previews.",
                "Use batch-query for many homogeneous refdes/net/HQ/page targets.",
            ],
        },
    })
    _write_json_file(args.json_out, payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_net_catalog(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review = _module_review_for_args(bundle, args)
    catalog = _build_net_catalog(bundle, args)
    payload = _envelope("net-catalog", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "net_catalog": catalog,
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_topology_netlist(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review, components, nets = _analysis_scope_for_args(bundle, args)
    scoped_bundle = dict(bundle)
    scoped_bundle["components"] = components
    scoped_bundle["nets"] = nets
    stdout_mode = str(getattr(args, "stdout", "summary") or "summary")
    view = str(getattr(args, "view", "") or ("full" if stdout_mode == "full" or getattr(args, "out", "") else "summary"))
    supply_mode = str(getattr(args, "supply_mode", "") or ("details" if view == "full" else "grouped"))
    topology = build_llm_topology_netlist(
        {
            "project_name": bundle.get("project_name", ""),
        },
        scoped_bundle,
        focus_refdes=str(getattr(args, "focus_refdes", "") or ""),
        role_filter=str(getattr(args, "role_filter", "") or ""),
        include_connectors=bool(getattr(args, "include_connectors", False)),
        limit=max(1, min(100, int(getattr(args, "limit", 30) or 30))),
        return_all_edges=view == "full",
        view=view,
        supply_mode=supply_mode,
        supply_limit=max(0, min(250, int(getattr(args, "supply_limit", 12) or 0))),
    )
    payload = _envelope("topology-netlist", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "topology_netlist": topology,
        "topology_summary": topology.get("summary_layer", {}),
        "topology_business_view": topology.get("business_view", {}),
        "written": {
            "topology_json": None,
        },
    })
    output_path = _write_json_file(getattr(args, "out", ""), topology, pretty=args.pretty)
    if output_path:
        payload["written"]["topology_json"] = output_path
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload if stdout_mode == "full" else _envelope("topology-netlist", {
        "summary": payload["summary"],
        "module_scope": payload["module_scope"],
        "topology_summary": topology.get("summary_layer", {}),
        "topology_business_view": topology.get("business_view", {}),
        "written": payload["written"],
    }), pretty=args.pretty)
    return 0


def _cadence_project_root_from_args(args: argparse.Namespace) -> tuple[str, Dict[str, Any]]:
    if getattr(args, "bundle_cache_in", ""):
        bundle = _read_bundle_cache(args.bundle_cache_in)
        root = str(bundle.get("project_root", "") or "").strip()
        if not root:
            raise ValueError("bundle 缓存中缺少 project_root，无法读取 Cadence 页文件。")
        return root, _analysis_summary(bundle)
    raw = str(getattr(args, "project_root", "") or "").strip()
    if not raw:
        raise ValueError("cadence-page 需要 project_root；使用缓存时请提供 --bundle-cache-in。")
    root, _prt_path, _net_path, _ref_path, snapshot_meta = discover_project_files_with_snapshot(raw)
    return str(root), {
        "project_name": root.name,
        "project_root": str(root),
        "project_input_snapshot": dict(snapshot_meta or {}),
    }


def cmd_cadence_page(args: argparse.Namespace) -> int:
    project_root, summary = _cadence_project_root_from_args(args)
    cadence_page = build_cadence_page_payload(
        project_root,
        int(getattr(args, "page", 0) or 0),
        stdout=str(getattr(args, "stdout", "summary") or "summary"),
        object_id=str(getattr(args, "object_id", "") or "").strip(),
        limit=max(1, min(5000, int(getattr(args, "limit", 200) or 200))),
    )
    payload = _envelope("cadence-page", {
        "summary": summary,
        "cadence_page": cadence_page,
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def _cadence_index_source_from_args(args: argparse.Namespace) -> tuple[str, Dict[str, Any], Dict[str, Any]]:
    if getattr(args, "bundle_cache_in", ""):
        bundle = _read_bundle_cache(args.bundle_cache_in)
        root = str(bundle.get("project_root", "") or "").strip()
        if not root:
            raise ValueError("bundle 缓存中缺少 project_root，无法读取 Cadence 语义索引。")
        return root, _analysis_summary(bundle), bundle
    raw = str(getattr(args, "project_root", "") or "").strip()
    if not raw:
        raise ValueError("cadence-index 需要 project_root；使用缓存时请提供 --bundle-cache-in。")
    root, _prt_path, _net_path, _ref_path, snapshot_meta = discover_project_files_with_snapshot(raw)
    return str(root), {
        "project_name": root.name,
        "project_root": str(root),
        "project_input_snapshot": dict(snapshot_meta or {}),
    }, {}


def cmd_cadence_index(args: argparse.Namespace) -> int:
    project_root, summary, bundle = _cadence_index_source_from_args(args)
    cadence_index = build_cadence_index_payload(
        project_root,
        pstx_nets=bundle.get("nets", {}) if isinstance(bundle, dict) else {},
        stdout=str(getattr(args, "stdout", "summary") or "summary"),
        query=str(getattr(args, "query", "") or ""),
        kind=str(getattr(args, "kind", "all") or "all"),
        page=int(getattr(args, "page", 0) or 0),
        limit=max(1, min(5000, int(getattr(args, "limit", 200) or 200))),
    )
    payload = _envelope("cadence-index", {
        "summary": summary,
        "cadence_index": cadence_index,
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def _csa_geometry_source_from_args(args: argparse.Namespace) -> tuple[str, Dict[str, Any]]:
    if getattr(args, "demo", False):
        demo_dir = Path(__file__).resolve().parents[1] / "pstx_core" / "cadence" / "demo_pages"
        return str(demo_dir), {
            "project_name": "dehdl-csa-demo",
            "project_root": str(demo_dir),
            "source": "builtin-demo",
        }
    if getattr(args, "bundle_cache_in", ""):
        bundle = _read_bundle_cache(args.bundle_cache_in)
        root = str(bundle.get("project_root", "") or "").strip()
        if not root:
            raise ValueError("bundle 缓存中缺少 project_root，无法扫描 CSA 几何。")
        return root, _analysis_summary(bundle)
    raw = str(getattr(args, "project_root", "") or "").strip()
    if not raw:
        raise ValueError("csa-geometry 需要 project_root/sch_1/pageX.csa；使用缓存时请提供 --bundle-cache-in。")
    root = Path(raw).expanduser()
    if root.is_file() and not _is_supported_archive(root):
        return str(root), {
            "project_name": root.stem,
            "project_root": str(root),
            "source": "path",
        }
    if (
        root.is_dir()
        and (
            root.name.lower() == "sch_1"
            or (bool(getattr(args, "recursive", False)) and root.name.lower() == "worklib")
            or (not (root / "packaged").is_dir() and any(root.glob("page*.csa")))
        )
    ):
        return str(root), {
            "project_name": root.name,
            "project_root": str(root),
            "source": "path",
        }
    resolved_root, _prt_path, _net_path, _ref_path, snapshot_meta = discover_project_files_with_snapshot(raw)
    root = resolved_root
    return str(root), {
        "project_name": root.stem if root.is_file() else root.name,
        "project_root": str(root),
        "source": "path",
        "project_input_snapshot": dict(snapshot_meta or {}),
    }


def _csa_geometry_written(args: argparse.Namespace,
                          results: List[csa_geometry.PageResult]) -> Dict[str, Optional[str]]:
    out_dir = str(getattr(args, "out_dir", "") or "").strip()
    json_report = bool(getattr(args, "json", False))
    html_report = bool(getattr(args, "html", False))
    if not out_dir and not json_report and not html_report:
        return {
            "summary_csv": None,
            "cross_detail_csv": None,
            "circle_detail_csv": None,
            "json_report": None,
            "html_report": None,
        }
    return csa_geometry.write_csa_geometry_reports(
        results,
        out_dir or ".",
        summary_name=str(getattr(args, "summary_name", "") or "cross_circle_summary.csv"),
        cross_detail_name=str(getattr(args, "cross_detail_name", "") or "dot_cross_detail.csv"),
        circle_detail_name=str(getattr(args, "circle_detail_name", "") or "circle_detail.csv"),
        json_report=json_report,
        json_name=str(getattr(args, "json_name", "") or "cross_circle_report.json"),
        html_report=html_report,
        html_name=str(getattr(args, "html_name", "") or "cross_circle_report.html"),
    )


def cmd_csa_geometry(args: argparse.Namespace) -> int:
    source, summary = _csa_geometry_source_from_args(args)
    include_arcs = bool(getattr(args, "include_arcs", False))
    if getattr(args, "demo", False):
        include_arcs = True
    results, geometry = csa_geometry.scan_csa_geometry(
        source,
        recursive=bool(getattr(args, "recursive", False)),
        workers=getattr(args, "workers", None),
        executor_kind=str(getattr(args, "executor", "thread") or "thread"),
        circle_two_point_mode=str(getattr(args, "circle_two_point_mode", "center_radius") or "center_radius"),
        include_arcs=include_arcs,
        check_missing=bool(getattr(args, "check_missing", False)),
        strict=True,
        page=getattr(args, "page", None),
    )
    written = _csa_geometry_written(args, results)
    row_limit = max(1, min(5000, int(getattr(args, "limit", 200) or 200)))
    stdout_mode = str(getattr(args, "stdout", "summary") or "summary")
    semantic_overlay = None
    if bool(getattr(args, "include_connectivity", False)):
        semantic_overlay = build_csa_connectivity_overlay(
            geometry,
            source_root=source,
            page=getattr(args, "page", None),
            stdout=stdout_mode,
            limit=row_limit,
        )
    csa_payload = csa_geometry.build_csa_geometry_payload(
        geometry,
        stdout=stdout_mode,
        limit=row_limit,
        page=getattr(args, "page", None),
        semantic_overlay=semantic_overlay,
    )
    csa_payload["written"] = written
    demo = None
    if getattr(args, "demo", False):
        demo_ok = geometry.get("cross_count") == 2 and geometry.get("circle_count") == 3
        demo = {
            "ok": demo_ok,
            "expected_cross": 2,
            "actual_cross": geometry.get("cross_count", 0),
            "expected_circles": 3,
            "actual_circles": geometry.get("circle_count", 0),
            "status": "DEMO_OK" if demo_ok else "DEMO_FAIL",
        }
    payload = _envelope("csa-geometry", {
        "summary": {
            **summary,
            "scan_root": geometry.get("root", ""),
            "recursive": bool(getattr(args, "recursive", False)),
            "executor": str(getattr(args, "executor", "thread") or "thread"),
            "include_arcs": include_arcs,
            "circle_two_point_mode": str(getattr(args, "circle_two_point_mode", "center_radius") or "center_radius"),
            "include_connectivity": bool(getattr(args, "include_connectivity", False)),
            "page": int(getattr(args, "page", 0) or 0),
        },
        "csa_geometry": csa_payload,
        "written": written,
    })
    if demo is not None:
        payload["demo"] = demo
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    if demo is not None and not demo.get("ok"):
        return 2
    if bool(getattr(args, "fail_on_findings", False)) and geometry.get("cross_count", 0):
        return 1
    if bool(getattr(args, "fail_on_circles", False)) and geometry.get("circle_count", 0):
        return 1
    if geometry.get("error_count", 0):
        return 3
    return 0


def _load_target_json_argument(raw: str) -> List[Dict[str, Any]]:
    text = str(raw or "").strip()
    if not text:
        return []
    parsed = json.loads(text)
    if isinstance(parsed, dict) and "targets" not in parsed:
        return [dict(parsed)]
    return load_targets_json(parsed)


def _schematic_pdf_targets_from_args(args: argparse.Namespace) -> List[Dict[str, Any]]:
    targets: List[Dict[str, Any]] = []
    targets_json = str(getattr(args, "targets_json", "") or "").strip()
    if targets_json:
        targets.extend(load_targets_json(targets_json))
    for raw in getattr(args, "target_json", []) or []:
        targets.extend(_load_target_json_argument(raw))
    for refdes in _split_cli_values(getattr(args, "refdes", []), max_items=500):
        targets.append({"kind": "refdes", "refdes": refdes, "source": "cli.refdes"})
    for net_name in _split_cli_values(getattr(args, "net", []), max_items=200):
        targets.append({"kind": "net", "net": net_name, "source": "cli.net"})
    for page in _split_cli_values(getattr(args, "page", []), max_items=200):
        targets.append({"kind": "page", "page": page, "source": "cli.page"})
    if not targets:
        raise ValueError("schematic-pdf-annotate 需要 --targets-json、--target-json、--refdes、--net 或 --page。")
    return targets


def cmd_schematic_pdf_annotate(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review = _module_review_for_args(bundle, args)
    targets = _schematic_pdf_targets_from_args(args)
    pdf_page_map = load_json_mapping_or_sequence(getattr(args, "pdf_page_map_json", ""), default={})
    page_calibrations = load_json_mapping_or_sequence(getattr(args, "calibrations_json", ""), default=[])
    if isinstance(page_calibrations, dict):
        page_calibrations = list(page_calibrations.get("page_calibrations", page_calibrations.get("calibrations", [])) or [])
    if not isinstance(pdf_page_map, dict):
        raise ValueError("--pdf-page-map-json 必须是 JSON object 或 JSON 文件路径。")
    if not isinstance(page_calibrations, list):
        raise ValueError("--calibrations-json 必须是 JSON array 或包含 page_calibrations/calibrations 的对象。")
    stdout_mode = str(getattr(args, "stdout", "summary") or "summary")
    annotation = build_schematic_pdf_annotation_payload(
        str(getattr(args, "pdf", "") or ""),
        bundle,
        targets,
        pdf_page_map=pdf_page_map,
        page_calibrations=page_calibrations,
        stdout=stdout_mode,
        limit=max(1, min(5000, int(getattr(args, "limit", 200) or 200))),
        allow_page_number_fallback=bool(getattr(args, "allow_page_number_fallback", False)),
    )
    payload = _envelope("schematic-pdf-annotate", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "schematic_pdf_annotation": annotation,
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    if stdout_mode == "summary":
        _emit(_envelope("schematic-pdf-annotate", {
            "summary": payload["summary"],
            "module_scope": payload["module_scope"],
            "schematic_pdf_annotation": {
                "schema_version": annotation.get("schema_version", ""),
                "digest": annotation.get("digest", {}),
                "pdf": {
                    "filename": (annotation.get("pdf", {}) or {}).get("filename", ""),
                    "page_count": (annotation.get("pdf", {}) or {}).get("page_count", 0),
                },
                "summary": annotation.get("summary", {}),
                "warnings": annotation.get("warnings", []),
                "truncated": annotation.get("truncated", False),
            },
        }), pretty=args.pretty)
    else:
        _emit(payload, pretty=args.pretty)
    return 0


def cmd_business_dictionary(args: argparse.Namespace) -> int:
    dictionary = business_dictionary_summary()
    payload = _envelope("business-dictionary", {
        "business_dictionary": dictionary,
        "usage": {
            "purpose": "让外部 Agent 统一理解接口别名、角色别名和业务 review focus。",
            "env_override": "PSTX_BUSINESS_DICTIONARY_FILE",
            "recommended_next_commands": [
                "Use topology-netlist after reading the dictionary to inspect chip-level business topology.",
                "Use evidence-pack or batch-query for refdes/net details referenced by dictionary aliases.",
            ],
        },
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_harness_skills(args: argparse.Namespace) -> int:
    skill_id = str(getattr(args, "skill_id", "") or "").strip()
    include_body = bool(getattr(args, "include_body", False))
    max_body_chars = max(200, min(20000, int(getattr(args, "max_body_chars", 4000) or 4000)))
    query = str(getattr(args, "query", "") or "").strip()
    capability_profiles = [str(item).strip() for item in (getattr(args, "capability_profile", []) or []) if str(item).strip()]
    playbooks = [str(item).strip() for item in (getattr(args, "playbook", []) or []) if str(item).strip()]
    tools = [str(item).strip() for item in (getattr(args, "tool", []) or []) if str(item).strip()]
    all_skills = load_harness_skills()

    mode = "list"
    cards: List[Dict[str, Any]]
    if skill_id:
        mode = "single"
        matched = [skill for skill in all_skills if skill.id == skill_id]
        if not matched:
            raise ValueError(f"unknown harness skill: {skill_id}")
        cards = [matched[0].card(include_body=True, max_body_chars=max_body_chars)]
    elif query or capability_profiles or playbooks or tools:
        mode = "select"
        selected = select_harness_skills(
            question=query,
            capability_profiles=capability_profiles,
            playbook_plan={
                "selected_playbooks": [{"id": item} for item in playbooks],
                "recommended_first_tools": tools,
            },
            root=Path(__file__).resolve().parents[1],
            max_selected=max(1, min(24, int(getattr(args, "limit", 24) or 24))),
            include_body=include_body,
            max_body_chars=max_body_chars,
        )
        cards = list(selected.get("selected_skills") or [])
    else:
        limit = max(1, min(200, int(getattr(args, "limit", 200) or 200)))
        cards = [skill.card(include_body=include_body, max_body_chars=max_body_chars) for skill in all_skills[:limit]]

    payload = _envelope("harness-skills", {
        "harness_skills": {
            "schema_version": "pstx-harness-skills.v1",
            "mode": mode,
            "available_count": len(all_skills),
            "returned_count": len(cards),
            "skills": cards,
        },
        "usage": {
            "purpose": "让 Trae/外部 Agent 与 Web Harness Agent 读取同一份技能卡和取证路线。",
            "recommended": [
                "Use harness-skills datasheet-key-info --include-body before broad datasheet review.",
                "Skill cards are guidance only; tool execution is still controlled by CLI/Bridge/Harness profiles.",
            ],
        },
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_datasheet_status(args: argparse.Namespace) -> int:
    status = build_datasheet_status()
    payload_data: Dict[str, Any] = {
        "datasheet_status": status,
        "recommended_next_commands": [
            "Set PSTX_DATASHEET_DIR to one or more PDF folders if configured=false.",
            "Use the Web/API reindex action before expecting datasheet-search to find new PDFs.",
            "Use datasheet-template to plan what the LLM should extract before asking parameter questions.",
        ],
    }
    if bool(getattr(args, "include_documents", False)):
        payload_data["documents"] = list_datasheet_documents(
            limit=max(1, min(1000, int(getattr(args, "limit", 200) or 200))),
            offset=max(0, int(getattr(args, "offset", 0) or 0)),
        )
    payload = _envelope("datasheet-status", payload_data)
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_datasheet_search(args: argparse.Namespace) -> int:
    query = str(getattr(args, "query", "") or "").strip()
    if not query:
        raise ValueError("datasheet-search requires --query")
    result = search_datasheet_chunks(
        query,
        limit=max(1, min(100, int(getattr(args, "limit", 20) or 20))),
        offset=max(0, int(getattr(args, "offset", 0) or 0)),
    )
    payload = _envelope("datasheet-search", {
        "datasheet_search": result,
        "recommended_next_commands": [
            "If a match is relevant, use the harness get_datasheet_chunk tool or Web Agent detail citation to read full evidence.",
            "For exact voltage/current/thermal/timing facts, run datasheet-parameters as a structured follow-up.",
        ],
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_datasheet_parameters(args: argparse.Namespace) -> int:
    result = search_datasheet_parameters(
        str(getattr(args, "query", "") or ""),
        parameter_key=str(getattr(args, "parameter_key", "") or ""),
        doc_id=int(getattr(args, "doc_id", 0) or 0) or None,
        limit=max(1, min(200, int(getattr(args, "limit", 50) or 50))),
        offset=max(0, int(getattr(args, "offset", 0) or 0)),
    )
    payload = _envelope("datasheet-parameters", {
        "datasheet_parameters": result,
        "recommended_next_commands": [
            "Use datasheet-template to decide which parameter categories still need evidence.",
            "High-risk numeric conclusions should cite parameter_id and then read the full detail through the harness get_datasheet_parameter tool.",
        ],
    })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_datasheet_template(args: argparse.Namespace) -> int:
    template_id = str(getattr(args, "template_id", "") or "").strip()
    if template_id:
        result = get_datasheet_review_template(template_id)
        if not result.get("ok", True):
            raise ValueError(str(result.get("error") or "unknown datasheet template"))
        payload = _envelope("datasheet-template", {
            "datasheet_template": result.get("template", {}),
            "schema": {
                "schema_version": result.get("schema_version", ""),
                "mode": "single",
            },
        })
    else:
        result = list_datasheet_review_templates(
            str(getattr(args, "category", "") or ""),
            include_questions=not bool(getattr(args, "without_questions", False)),
        )
        payload = _envelope("datasheet-template", {
            "datasheet_templates": result,
            "schema": {
                "schema_version": result.get("schema_version", ""),
                "mode": "list",
            },
        })
    _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_offline_migration(args: argparse.Namespace) -> int:
    action = str(getattr(args, "offline_action", "") or "").strip()
    if action == "build-python-url":
        url = build_python_download_url(
            python_version=str(getattr(args, "python_version", "") or ""),
            python_mirror=str(getattr(args, "python_mirror", "official") or "official"),
            python_mirror_base=str(getattr(args, "python_mirror_base", "") or ""),
            python_filename=str(getattr(args, "python_filename", "") or ""),
        )
        payload = _envelope("offline-migration", {
            "offline_migration": {
                "action": action,
                "python_url": url,
            },
        })
        _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
        _emit(payload, pretty=args.pretty)
        return 0
    if action == "prepare":
        result = prepare_offline_bundle(
            project_root=str(getattr(args, "project_root", "") or Path.cwd()),
            out_dir=str(getattr(args, "out_dir", "") or "output/offline_migration"),
            name=str(getattr(args, "name", "") or "pstx-offline"),
            target_platform=str(getattr(args, "target_platform", "") or "windows-amd64"),
            target_profile=str(getattr(args, "target_profile", "") or "windows-rtx4060-cuda"),
            python_dir=str(getattr(args, "python_dir", "") or ""),
            python_archive=str(getattr(args, "python_archive", "") or ""),
            python_url=str(getattr(args, "python_url", "") or ""),
            python_version=str(getattr(args, "python_version", "") or ""),
            python_mirror=str(getattr(args, "python_mirror", "official") or "official"),
            python_mirror_base=str(getattr(args, "python_mirror_base", "") or ""),
            python_filename=str(getattr(args, "python_filename", "") or ""),
            extract_python=not bool(getattr(args, "no_extract_python", False)),
            allow_system_python_on_b=bool(getattr(args, "allow_system_python_on_b", False)),
            mineru_venv=str(getattr(args, "mineru_venv", "") or ""),
            mineru_model_dir=str(getattr(args, "mineru_model_dir", "") or ""),
            mineru_config=str(getattr(args, "mineru_config", "") or ""),
            download_mineru_models=bool(getattr(args, "download_mineru_models", False)),
            mineru_model_source=str(getattr(args, "mineru_model_source", "") or DEFAULT_MINERU_MODEL_SOURCE),
            mineru_model_type=str(getattr(args, "mineru_model_type", "") or DEFAULT_MINERU_MODEL_TYPE),
            huggingface_endpoint=str(getattr(args, "huggingface_endpoint", "") or ""),
            mineru_model_downloader=str(getattr(args, "mineru_model_downloader", "") or ""),
            download_wheels=bool(getattr(args, "download_wheels", False)),
            pip_index_url=str(getattr(args, "pip_index_url", "") or ""),
            pip_extra_index_url=str(getattr(args, "pip_extra_index_url", "") or ""),
            include_mineru_wheels=bool(getattr(args, "include_mineru_wheels", False)),
            strict_mineru_wheels=bool(getattr(args, "strict_mineru_wheels", False)),
            mineru_wheel_spec=str(getattr(args, "mineru_wheel_spec", "") or DEFAULT_MINERU_WHEEL_SPEC),
            asset_cache_dir=str(getattr(args, "asset_cache_dir", "") or ""),
            reuse_assets=not bool(getattr(args, "no_reuse_assets", False)),
            include_datasheet_data=not bool(getattr(args, "skip_datasheet_data", False)),
            include_datasheet_source=bool(getattr(args, "include_datasheet_source", False)),
            make_zip=not bool(getattr(args, "no_zip", False)),
        )
        payload = _envelope("offline-migration", {
            "offline_migration": {
                "action": action,
                **result,
            },
            "written": {
                "bundle_root": result.get("bundle_root", ""),
                "zip_path": result.get("zip_path", ""),
                "manifest_path": result.get("manifest_path", ""),
            },
            "computer_b_command": [
                "电脑 B 解压后优先运行 RUN_SETUP_B.bat 或 powershell -ExecutionPolicy Bypass -File RUN_SETUP_B.ps1",
                "排错时再运行 RUN_VERIFY_B.bat / RUN_VERIFY_B.ps1 / ./RUN_VERIFY_B.sh 或 RUN_INSTALL_WHEELHOUSE_B.*",
            ],
        })
        _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
        _emit(payload, pretty=args.pretty)
        return 0
    if action == "verify":
        package_root = str(getattr(args, "package_root", "") or "").strip()
        if not package_root:
            raise ValueError("offline-migration verify requires package_root")
        result = verify_offline_bundle(
            package_root,
            probe_runtime=not bool(getattr(args, "skip_runtime_probe", False)),
        )
        payload = _envelope("offline-migration", {
            "offline_migration": {
                "action": action,
                **result,
            },
            "verification": result,
        })
        _write_json_file(getattr(args, "json_out", ""), payload, pretty=args.pretty)
        _emit(payload, pretty=args.pretty)
        return 0 if result.get("ok") else 1
    raise ValueError(f"unsupported offline-migration action: {action}")


def cmd_batch_query(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review, components, nets = _analysis_scope_for_args(bundle, args)
    mode = str(args.mode)
    limit = max(1, min(200, int(getattr(args, "limit_per_item", 20) or 20)))
    items = _parse_batch_items(args)
    results: List[Dict[str, Any]] = []
    for item in items:
        if mode == "HQ料号":
            result = _query_by_hq(components, item, limit)
        elif mode == "页码":
            result = _query_by_page(components, item, limit)
        else:
            query_result = _trim_query_result(query_project_data(components, nets, mode, item), limit)
            result = {
                "mode": mode,
                "query": item,
                "status": "found" if query_result.get("match_type") != "missing" else "missing",
                "result_count": _query_result_count(query_result),
                "truncated": bool(query_result.get("items_truncated")) or any(
                    bool(card.get("truncated")) for card in query_result.get("cards", []) or []
                ),
                "result": query_result,
            }
        results.append(result)
    found_count = sum(1 for result in results if result.get("status") == "found")
    payload = _envelope("batch-query", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "mode": mode,
        "requested_count": len(items),
        "found_count": found_count,
        "missing_count": len(items) - found_count,
        "limit_per_item": limit,
        "results": results,
    })
    _write_json_file(args.json_out, payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_module_review(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    module_review = _module_review_for_args(bundle, args)
    payload = _envelope("module-review", {
        "summary": _analysis_summary(bundle),
        "module_review": module_review,
    })
    _write_json_file(args.json_out, payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def cmd_report_table(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    report = _report_payload_for_bundle(bundle)
    tables = _iter_report_tables(report)
    module_review = _module_review_for_args(bundle, args)
    summaries = [_report_table_summary(table) for table in tables]
    target_id = str(getattr(args, "table_id", "") or "").strip()
    payload: Dict[str, Any] = {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "tables": summaries,
    }
    if target_id:
        module_review, table, all_rows, _tables = _report_table_for_args(bundle, args, target_id)
        payload["module_scope"] = module_review.get("summary", {})
        offset = max(0, int(getattr(args, "offset", 0) or 0))
        limit = max(1, min(5000, int(getattr(args, "limit", 200) or 200)))
        rows = all_rows[offset:offset + limit]
        payload["table"] = {
            **_report_table_summary({**table, "count": len(all_rows)}),
            "offset": offset,
            "limit": limit,
            "returned_count": len(rows),
            "total_count": len(all_rows),
            "truncated": offset + len(rows) < len(all_rows),
            "rows": rows,
        }
    envelope = _envelope("report-table", payload)
    _write_json_file(getattr(args, "json_out", ""), envelope, pretty=args.pretty)
    _emit(envelope, pretty=args.pretty)
    return 0


def cmd_report_aggregate(args: argparse.Namespace) -> int:
    bundle = _analyze_project_from_args(args)
    table_id = str(getattr(args, "table_id", "") or "").strip()
    if not table_id:
        raise ValueError("report-aggregate requires --table-id")
    column = str(getattr(args, "column", "") or "").strip()
    if not column:
        raise ValueError("report-aggregate requires --column")
    module_review, table, rows, _tables = _report_table_for_args(bundle, args, table_id)
    limit = max(1, min(5000, int(getattr(args, "limit", 100) or 100)))
    aggregation = _aggregate_rows(
        rows,
        column=column,
        operation=str(getattr(args, "operation", "top") or "top"),
        limit=limit,
        include_empty=bool(getattr(args, "include_empty", False)),
    )
    payload = _envelope("report-aggregate", {
        "summary": _analysis_summary(bundle),
        "module_scope": module_review.get("summary", {}),
        "table": _report_table_summary({**table, "count": len(rows)}),
        "aggregation": aggregation,
    })
    _write_json_file(args.json_out, payload, pretty=args.pretty)
    _emit(payload, pretty=args.pretty)
    return 0


def _payload_for_compare_side(project_root: str,
                              project_name: str,
                              args: argparse.Namespace,
                              run_id: str) -> Dict[str, Any]:
    side_args = argparse.Namespace(
        project_root=project_root,
        project_name=project_name,
        ratio_limit=args.ratio_limit,
        custom_volt_map=args.custom_volt_map,
        include_depop=args.include_depop,
        include_total_bom=args.include_total_bom,
    )
    bundle = _analyze_project_from_args(side_args)
    return {
        "bundle": bundle,
        "report": _report_payload_for_bundle(bundle, run_id=run_id),
    }


def cmd_compare(args: argparse.Namespace) -> int:
    left_payload = _payload_for_compare_side(args.left_project_root, args.left_name, args, "left-cli")
    right_payload = _payload_for_compare_side(args.right_project_root, args.right_name, args, "right-cli")
    payloads = {
        "left-cli": left_payload,
        "right-cli": right_payload,
    }
    detail_limit = coerce_compare_detail_limit(args.detail_limit)
    compare_payload = build_compare_payload(
        "left-cli",
        "right-cli",
        get_run_payload=lambda run_id: payloads[run_id],
        detail_limit=detail_limit,
    )
    envelope = _envelope("compare", {
        "compare": compare_payload,
    })
    _write_json_file(args.json_out, envelope, pretty=args.pretty)
    if args.stdout == "summary":
        summary = {
            "ok": True,
            "interface": "pstx-cli",
            "interface_version": CLI_VERSION,
            "schema_version": SCHEMA_VERSION,
            "command": "compare",
            "generated_at": envelope["generated_at"],
            "diff_totals": compare_payload.get("diff_totals", {}),
            "left": compare_payload.get("left", {}),
            "right": compare_payload.get("right", {}),
            "written": {"compare_json": str(Path(args.json_out).expanduser()) if args.json_out else None},
        }
        _emit(summary, pretty=args.pretty)
    else:
        _emit(envelope, pretty=args.pretty)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(
        prog="pstx_cli",
        description="PSTX CLI-friendly interfaces for external processes",
    )
    parser.add_argument("--pretty", action="store_true", help="pretty-print JSON output")
    subparsers = parser.add_subparsers(dest="command", required=True, parser_class=JsonArgumentParser)

    capabilities = subparsers.add_parser("capabilities", help="list CLI-accessible capabilities")
    _add_pretty_arg(capabilities)
    capabilities.set_defaults(func=cmd_capabilities)

    schema = subparsers.add_parser("schema", help="describe CLI JSON schemas")
    _add_pretty_arg(schema)
    schema.add_argument("schema_command", nargs="?", default="", help="optional command name to describe")
    schema.set_defaults(func=cmd_schema)

    analyze = subparsers.add_parser("analyze", help="analyze one PSTX project")
    _add_pretty_arg(analyze)
    _add_common_analysis_args(analyze)
    analyze.add_argument("--json-out", default="", help="write full analysis bundle JSON")
    analyze.add_argument("--bundle-cache-out", default="", help="write reusable analysis bundle cache JSON")
    analyze.add_argument("--report-json-out", default="", help="write UI report payload JSON")
    analyze.add_argument("--excel-out", default="", help="write Excel report")
    analyze.add_argument(
        "--stdout",
        choices=["summary", "bundle", "report", "module-review"],
        default="summary",
        help="payload printed to stdout, default summary",
    )
    _add_module_filter_args(analyze)
    analyze.set_defaults(func=cmd_analyze)

    inspect = subparsers.add_parser("inspect", help="inspect PSTX project files and suggested workflow")
    _add_pretty_arg(inspect)
    inspect.add_argument("project_root", help="PSTX project root, packaged directory, CPM container, or supported archive")
    inspect.add_argument("--json-out", default="", help="write inspect envelope JSON")
    inspect.set_defaults(func=cmd_inspect)

    query = subparsers.add_parser("query", help="query a refdes or net after analysis")
    _add_pretty_arg(query)
    _add_common_analysis_args(query, project_required=False, allow_cache_in=True)
    query.add_argument("--mode", choices=["位号", "网络"], default="位号", help="query mode")
    query.add_argument("--keyword", required=True, help="refdes or net keyword")
    query.add_argument("--json-out", default="", help="write query envelope JSON")
    _add_module_filter_args(query)
    query.set_defaults(func=cmd_query)

    batch_query = subparsers.add_parser("batch-query", help="query multiple refdes, nets, HQ numbers or pages")
    _add_pretty_arg(batch_query)
    _add_common_analysis_args(batch_query, project_required=False, allow_cache_in=True)
    _add_module_filter_args(batch_query)
    batch_query.add_argument("--mode", choices=["位号", "网络", "HQ料号", "页码"], default="位号", help="batch query mode")
    batch_query.add_argument("--items", default="", help="comma/newline separated query items")
    batch_query.add_argument("--items-file", default="", help="UTF-8 text lines or JSON array of query items")
    batch_query.add_argument("--max-items", type=int, default=100, help="maximum input items, max 500")
    batch_query.add_argument("--limit-per-item", type=int, default=20, help="maximum rows returned per item, max 200")
    batch_query.add_argument("--json-out", default="", help="write batch-query envelope JSON")
    batch_query.set_defaults(func=cmd_batch_query)

    module_review = subparsers.add_parser("module-review", help="output module_order based module scope")
    _add_pretty_arg(module_review)
    _add_common_analysis_args(module_review, project_required=False, allow_cache_in=True)
    _add_module_filter_args(module_review)
    module_review.add_argument("--json-out", default="", help="write module review envelope JSON")
    module_review.set_defaults(func=cmd_module_review)

    report_table = subparsers.add_parser("report-table", help="list report tables or page one table")
    _add_pretty_arg(report_table)
    _add_common_analysis_args(report_table, project_required=False, allow_cache_in=True)
    _add_module_filter_args(report_table)
    report_table.add_argument("--table-id", default="", help="report table id to page, omit to list catalog only")
    report_table.add_argument("--offset", type=int, default=0, help="row offset for --table-id")
    report_table.add_argument("--limit", type=int, default=200, help="row limit for --table-id, max 5000")
    report_table.add_argument("--json-out", default="", help="write report-table envelope JSON")
    report_table.set_defaults(func=cmd_report_table)

    report_aggregate = subparsers.add_parser("report-aggregate", help="aggregate one report table column")
    _add_pretty_arg(report_aggregate)
    _add_common_analysis_args(report_aggregate, project_required=False, allow_cache_in=True)
    _add_module_filter_args(report_aggregate)
    report_aggregate.add_argument("--table-id", required=True, help="report table id to aggregate")
    report_aggregate.add_argument("--column", required=True, help="column name to aggregate")
    report_aggregate.add_argument("--operation", choices=["top", "count", "unique"], default="top", help="aggregation ordering")
    report_aggregate.add_argument("--limit", type=int, default=100, help="aggregation item limit, max 5000")
    report_aggregate.add_argument("--include-empty", action="store_true", help="include empty values in aggregation")
    report_aggregate.add_argument("--json-out", default="", help="write report-aggregate envelope JSON")
    report_aggregate.set_defaults(func=cmd_report_aggregate)

    evidence_pack = subparsers.add_parser("evidence-pack", help="collect mixed target evidence for external agents")
    _add_pretty_arg(evidence_pack)
    _add_common_analysis_args(evidence_pack, project_required=False, allow_cache_in=True)
    _add_module_filter_args(evidence_pack)
    evidence_pack.add_argument("--refdes", action="append", default=[], help="comma/newline separated refdes targets; repeatable")
    evidence_pack.add_argument("--net", action="append", default=[], help="comma/newline separated net targets; repeatable")
    evidence_pack.add_argument("--hq", action="append", default=[], help="comma/newline separated HQ material numbers; repeatable")
    evidence_pack.add_argument("--page", action="append", default=[], help="comma/newline separated page numbers or PAGE labels; repeatable")
    evidence_pack.add_argument("--table-id", action="append", default=[], help="report table ids to include; repeatable")
    evidence_pack.add_argument("--max-targets", type=int, default=50, help="maximum targets per kind, max 200")
    evidence_pack.add_argument("--limit-per-target", type=int, default=20, help="maximum rows per non-table target, max 200")
    evidence_pack.add_argument("--table-limit", type=int, default=50, help="maximum preview rows per table, max 500")
    evidence_pack.add_argument("--json-out", default="", help="write evidence-pack envelope JSON")
    evidence_pack.set_defaults(func=cmd_evidence_pack)

    net_catalog = subparsers.add_parser("net-catalog", help="list and filter net labels for external agents")
    _add_pretty_arg(net_catalog)
    _add_common_analysis_args(net_catalog, project_required=False, allow_cache_in=True)
    _add_module_filter_args(net_catalog)
    net_catalog.add_argument("--query", default="", help="net keyword or business alias, e.g. PCE/PCIE/I2C/P3V3")
    net_catalog.add_argument(
        "--kind",
        choices=["all", "power", "ground", "signal", "differential", "unnamed"],
        default="all",
        help="net kind filter",
    )
    net_catalog.add_argument("--min-nodes", type=int, default=1, help="minimum node count, default 1")
    net_catalog.add_argument("--offset", type=int, default=0, help="net list offset")
    net_catalog.add_argument("--limit", type=int, default=100, help="net list limit, max 5000")
    net_catalog.add_argument("--include-nodes", action="store_true", help="include up to 50 node summaries per returned net")
    net_catalog.add_argument("--json-out", default="", help="write net-catalog envelope JSON")
    net_catalog.set_defaults(func=cmd_net_catalog)

    topology_netlist = subparsers.add_parser("topology-netlist", help="export LLM semantic topology netlist")
    _add_pretty_arg(topology_netlist)
    _add_common_analysis_args(topology_netlist, project_required=False, allow_cache_in=True)
    _add_module_filter_args(topology_netlist)
    topology_netlist.add_argument("--focus-refdes", default="", help="optional refdes to focus topology edges")
    topology_netlist.add_argument("--role-filter", default="", help="optional role substring filter, e.g. level_shifter")
    topology_netlist.add_argument("--include-connectors", action="store_true", help="include connector nodes in topology")
    topology_netlist.add_argument("--limit", type=int, default=30, help="node/edge return limit, max 100")
    topology_netlist.add_argument("--view", choices=["summary", "full"], default="", help="topology payload view, default follows stdout/out")
    topology_netlist.add_argument("--supply-mode", choices=["grouped", "details", "hidden"], default="", help="supply edge display mode")
    topology_netlist.add_argument("--supply-limit", type=int, default=12, help="supply samples/groups return limit, max 250")
    topology_netlist.add_argument("--out", default="", help="write raw topology netlist JSON artifact")
    topology_netlist.add_argument("--json-out", default="", help="write topology-netlist envelope JSON")
    topology_netlist.add_argument("--stdout", choices=["summary", "full"], default="summary", help="stdout payload")
    topology_netlist.set_defaults(func=cmd_topology_netlist)

    cadence_page = subparsers.add_parser("cadence-page", help="read one Cadence page connectivity semantic model")
    _add_pretty_arg(cadence_page)
    _add_common_analysis_args(cadence_page, project_required=False, allow_cache_in=True)
    cadence_page.add_argument("--page", type=int, required=True, help="page number for sch_1/pageX.csv|csa")
    cadence_page.add_argument("--stdout", choices=["summary", "objects", "full"], default="summary", help="Cadence page payload detail")
    cadence_page.add_argument("--object-id", default="", help="optional Cadence object_id or connectivity component id")
    cadence_page.add_argument("--limit", type=int, default=200, help="object/connectivity return limit, max 5000")
    cadence_page.add_argument("--json-out", default="", help="write cadence-page envelope JSON")
    cadence_page.set_defaults(func=cmd_cadence_page)

    cadence_index = subparsers.add_parser("cadence-index", help="read project-level Cadence semantic index")
    _add_pretty_arg(cadence_index)
    _add_common_analysis_args(cadence_index, project_required=False, allow_cache_in=True)
    cadence_index.add_argument("--stdout", choices=["summary", "nets", "ports", "links", "full"], default="summary", help="Cadence index payload detail")
    cadence_index.add_argument("--query", default="", help="case-insensitive substring filter for semantic names")
    cadence_index.add_argument("--kind", choices=["all", "net", "port", "offpage", "bus", "no_connect", "unbound"], default="all", help="semantic row kind filter")
    cadence_index.add_argument("--page", type=int, default=0, help="only return rows that appear on one page number")
    cadence_index.add_argument("--limit", type=int, default=200, help="row return limit per row set, max 5000")
    cadence_index.add_argument("--json-out", default="", help="write cadence-index envelope JSON")
    cadence_index.set_defaults(func=cmd_cadence_index)

    csa_geometry_parser = subparsers.add_parser("csa-geometry", help="scan Cadence DE HDL CSA geometry checks")
    _add_pretty_arg(csa_geometry_parser)
    csa_geometry_parser.add_argument("project_root", nargs="?", default="", help="project root, sch_1 directory, or one pageX.csa file")
    csa_geometry_parser.add_argument("--bundle-cache-in", default="", help="read project_root from an existing analyze bundle JSON")
    csa_geometry_parser.add_argument("--recursive", action="store_true", help="recursively scan pageX.csa below the input root")
    csa_geometry_parser.add_argument("--workers", type=int, default=None, help="parallel worker count; default min(CPU cores, page count)")
    csa_geometry_parser.add_argument("--executor", choices=["thread", "process", "serial"], default="thread", help="CSA scan backend")
    csa_geometry_parser.add_argument("--include-arcs", action="store_true", help="parse ARC objects as fitted/guessed circles")
    csa_geometry_parser.add_argument("--circle-two-point-mode", choices=["center_radius", "bbox"], default="center_radius", help="two-point CIRCLE interpretation")
    csa_geometry_parser.add_argument("--check-missing", action="store_true", help="warn if page numbers are missing between min and max")
    csa_geometry_parser.add_argument("--include-connectivity", action="store_true", help="overlay Cadence page connectivity semantics on CSA geometry findings")
    csa_geometry_parser.add_argument("--page", type=int, default=0, help="only scan/return one page number")
    csa_geometry_parser.add_argument("--out-dir", default="", help="write package-style CSA CSV reports to this directory")
    csa_geometry_parser.add_argument("--summary-name", default="cross_circle_summary.csv", help="summary CSV filename")
    csa_geometry_parser.add_argument("--cross-detail-name", default="dot_cross_detail.csv", help="DOT cross detail CSV filename")
    csa_geometry_parser.add_argument("--circle-detail-name", default="circle_detail.csv", help="circle detail CSV filename")
    csa_geometry_parser.add_argument("--json", action="store_true", help="also write package-style cross_circle_report.json")
    csa_geometry_parser.add_argument("--json-name", default="cross_circle_report.json", help="package-style JSON report filename")
    csa_geometry_parser.add_argument("--html", action="store_true", help="also write self-contained CSA geometry HTML report")
    csa_geometry_parser.add_argument("--html-name", default="cross_circle_report.html", help="HTML report filename")
    csa_geometry_parser.add_argument("--json-out", default="", help="write csa-geometry CLI envelope JSON")
    csa_geometry_parser.add_argument("--fail-on-findings", action="store_true", help="return code 1 if DOT four-way crosses are found")
    csa_geometry_parser.add_argument("--fail-on-circles", action="store_true", help="return code 1 if circle marks are found")
    csa_geometry_parser.add_argument("--stdout", choices=["summary", "hits", "details", "full"], default="summary", help="stdout payload detail")
    csa_geometry_parser.add_argument("--limit", type=int, default=200, help="row return limit, max 5000")
    csa_geometry_parser.add_argument("--demo", action="store_true", help="run built-in demo pages and verify expected results")
    csa_geometry_parser.set_defaults(func=cmd_csa_geometry)

    schematic_pdf = subparsers.add_parser("schematic-pdf-annotate", help="locate review targets on a schematic PDF and return overlay JSON")
    _add_pretty_arg(schematic_pdf)
    schematic_pdf.add_argument("pdf", help="schematic PDF path on the analysis machine")
    schematic_pdf.add_argument("project_root", nargs="?", default="", help="PSTX project root/container/archive; optional with --bundle-cache-in")
    schematic_pdf.add_argument("--project-name", default="", help="override project name when analyzing project_root")
    schematic_pdf.add_argument("--ratio-limit", type=float, default=70.0, help="capacitor derating ratio limit when analyzing project_root")
    schematic_pdf.add_argument("--custom-volt-map", default="", help="custom voltage map text when analyzing project_root")
    schematic_pdf.add_argument("--include-depop", action="store_true", help="include DEPOP/DNP components when analyzing project_root")
    schematic_pdf.add_argument("--include-total-bom", action="store_true", help="include total BOM summary when analyzing project_root")
    schematic_pdf.add_argument("--bundle-cache-in", default="", help="read an existing analyze bundle JSON instead of analyzing project_root")
    _add_module_filter_args(schematic_pdf)
    schematic_pdf.add_argument("--targets-json", default="", help="JSON file/string containing a target array or {'targets': [...]}")
    schematic_pdf.add_argument("--target-json", action="append", default=[], help="single target JSON object or target array; repeatable")
    schematic_pdf.add_argument("--refdes", action="append", default=[], help="comma/newline separated refdes targets; repeatable")
    schematic_pdf.add_argument("--net", action="append", default=[], help="comma/newline separated net targets; repeatable")
    schematic_pdf.add_argument("--page", action="append", default=[], help="comma/newline separated project page targets; repeatable")
    schematic_pdf.add_argument("--pdf-page-map-json", default="", help="JSON/file mapping project PAGE labels to 1-based PDF pages")
    schematic_pdf.add_argument("--calibrations-json", default="", help="JSON/file page calibration array for schematic XY -> PDF coordinate mapping")
    schematic_pdf.add_argument("--allow-page-number-fallback", action="store_true", help="allow weak PAGE<N> -> PDF page N fallback when no explicit/text page map exists")
    schematic_pdf.add_argument("--stdout", choices=["summary", "annotations", "full"], default="summary", help="stdout payload detail")
    schematic_pdf.add_argument("--limit", type=int, default=200, help="annotation return limit, max 5000")
    schematic_pdf.add_argument("--json-out", default="", help="write schematic PDF annotation envelope JSON")
    schematic_pdf.set_defaults(func=cmd_schematic_pdf_annotate)

    business_dictionary = subparsers.add_parser("business-dictionary", help="output business aliases and review focus dictionary")
    _add_pretty_arg(business_dictionary)
    business_dictionary.add_argument("--json-out", default="", help="write business dictionary envelope JSON")
    business_dictionary.set_defaults(func=cmd_business_dictionary)

    harness_skills = subparsers.add_parser("harness-skills", help="list or read Harness Agent skill cards")
    _add_pretty_arg(harness_skills)
    harness_skills.add_argument("skill_id", nargs="?", default="", help="optional skill id, e.g. datasheet-key-info")
    harness_skills.add_argument("--query", default="", help="optional question text used to select matching skills")
    harness_skills.add_argument("--capability-profile", action="append", default=[], help="capability profile filter; repeatable")
    harness_skills.add_argument("--playbook", action="append", default=[], help="playbook id filter; repeatable")
    harness_skills.add_argument("--tool", action="append", default=[], help="tool name filter; repeatable")
    harness_skills.add_argument("--include-body", action="store_true", help="include skill markdown body")
    harness_skills.add_argument("--max-body-chars", type=int, default=4000, help="body preview limit when included")
    harness_skills.add_argument("--limit", type=int, default=24, help="max returned skills when listing/selecting")
    harness_skills.add_argument("--json-out", default="", help="write harness skills envelope JSON")
    harness_skills.set_defaults(func=cmd_harness_skills)

    datasheet_status = subparsers.add_parser("datasheet-status", help="show local datasheet PDF index status")
    _add_pretty_arg(datasheet_status)
    datasheet_status.add_argument("--include-documents", action="store_true", help="include indexed document list")
    datasheet_status.add_argument("--limit", type=int, default=200, help="document list limit, max 1000")
    datasheet_status.add_argument("--offset", type=int, default=0, help="document list offset")
    datasheet_status.add_argument("--json-out", default="", help="write datasheet status envelope JSON")
    datasheet_status.set_defaults(func=cmd_datasheet_status)

    datasheet_search = subparsers.add_parser("datasheet-search", help="search indexed datasheet chunks")
    _add_pretty_arg(datasheet_search)
    datasheet_search.add_argument("--query", required=True, help="datasheet search keyword/query")
    datasheet_search.add_argument("--limit", type=int, default=20, help="match limit, max 100")
    datasheet_search.add_argument("--offset", type=int, default=0, help="match offset")
    datasheet_search.add_argument("--json-out", default="", help="write datasheet search envelope JSON")
    datasheet_search.set_defaults(func=cmd_datasheet_search)

    datasheet_parameters = subparsers.add_parser("datasheet-parameters", help="search deterministic datasheet parameter cards")
    _add_pretty_arg(datasheet_parameters)
    datasheet_parameters.add_argument("--query", default="", help="free-text parameter query")
    datasheet_parameters.add_argument("--parameter-key", default="", help="parameter key filter")
    datasheet_parameters.add_argument("--doc-id", type=int, default=0, help="optional datasheet doc_id filter")
    datasheet_parameters.add_argument("--limit", type=int, default=50, help="parameter card limit, max 200")
    datasheet_parameters.add_argument("--offset", type=int, default=0, help="parameter card offset")
    datasheet_parameters.add_argument("--json-out", default="", help="write datasheet parameter envelope JSON")
    datasheet_parameters.set_defaults(func=cmd_datasheet_parameters)

    datasheet_template = subparsers.add_parser("datasheet-template", help="output LLM-readable datasheet review templates")
    _add_pretty_arg(datasheet_template)
    datasheet_template.add_argument("template_id", nargs="?", default="", help="optional template id, e.g. complex_chip")
    datasheet_template.add_argument("--category", default="", help="template category filter when template_id is omitted")
    datasheet_template.add_argument("--without-questions", action="store_true", help="omit long review questions/playbook for compact output")
    datasheet_template.add_argument("--json-out", default="", help="write datasheet template envelope JSON")
    datasheet_template.set_defaults(func=cmd_datasheet_template)

    compare = subparsers.add_parser("compare", help="compare two PSTX projects")
    _add_pretty_arg(compare)
    compare.add_argument("left_project_root", help="left PSTX project root, packaged directory, CPM container, worklib, or supported archive")
    compare.add_argument("right_project_root", help="right PSTX project root, packaged directory, CPM container, worklib, or supported archive")
    compare.add_argument("--left-name", default="", help="override left project name")
    compare.add_argument("--right-name", default="", help="override right project name")
    compare.add_argument("--ratio-limit", type=float, default=70.0, help="capacitor derating ratio limit")
    compare.add_argument("--custom-volt-map", default="", help="custom voltage map text")
    compare.add_argument("--include-depop", action="store_true", help="include DEPOP/DNP components in rule analysis")
    compare.add_argument("--include-total-bom", action="store_true", help="include total BOM summary")
    compare.add_argument("--detail-limit", default=500, help="compare detail limit, max follows Web API")
    compare.add_argument("--json-out", default="", help="write compare envelope JSON")
    compare.add_argument("--stdout", choices=["summary", "full"], default="summary", help="stdout payload")
    compare.set_defaults(func=cmd_compare)

    agent_run_status = subparsers.add_parser("agent-run-status", help="show durable background Agent run status")
    _add_pretty_arg(agent_run_status)
    agent_run_status.add_argument("agent_run_id", help="agent run id returned by async Web API")
    agent_run_status.set_defaults(func=cmd_agent_run_status)

    agent_run_artifacts = subparsers.add_parser("agent-run-artifacts", help="list durable Agent run artifacts")
    _add_pretty_arg(agent_run_artifacts)
    agent_run_artifacts.add_argument("agent_run_id", help="agent run id returned by async Web API")
    agent_run_artifacts.set_defaults(func=cmd_agent_run_artifacts)

    agent_run_trace = subparsers.add_parser("agent-run-trace", help="show partial/final durable Agent run trace")
    _add_pretty_arg(agent_run_trace)
    agent_run_trace.add_argument("agent_run_id", help="agent run id returned by async Web API")
    agent_run_trace.set_defaults(func=cmd_agent_run_trace)

    offline = subparsers.add_parser("offline-migration", help="prepare or verify offline migration bundles")
    _add_pretty_arg(offline)
    offline_sub = offline.add_subparsers(dest="offline_action", required=True)

    offline_url = offline_sub.add_parser("build-python-url", help="build a mirror URL for a portable Python archive")
    _add_pretty_arg(offline_url)
    offline_url.add_argument("--python-version", required=True, help="Python version, e.g. 3.10.11")
    offline_url.add_argument("--python-mirror", choices=["official", "tuna", "npmmirror"], default="official", help="Python mirror alias")
    offline_url.add_argument("--python-mirror-base", default="", help="custom mirror base, e.g. https://mirrors.example/python")
    offline_url.add_argument("--python-filename", default="", help="archive filename; default python-<version>-embed-amd64.zip")
    offline_url.add_argument("--json-out", default="", help="write offline migration envelope JSON")
    offline_url.set_defaults(func=cmd_offline_migration)

    offline_prepare = offline_sub.add_parser("prepare", help="prepare an offline migration bundle on computer A")
    _add_pretty_arg(offline_prepare)
    offline_prepare.add_argument("--project-root", default=str(Path.cwd()), help="source project root; default current working directory")
    offline_prepare.add_argument("--out-dir", default="output/offline_migration", help="output parent directory")
    offline_prepare.add_argument("--name", default="pstx-offline", help="bundle folder/zip name")
    offline_prepare.add_argument("--target-platform", default="windows-amd64", help="target platform label stored in manifest")
    offline_prepare.add_argument("--target-profile", default="windows-rtx4060-cuda", help="runtime profile label, e.g. windows-rtx4060-cuda")
    offline_prepare.add_argument("--python-dir", default="", help="copy an already extracted portable Python directory")
    offline_prepare.add_argument("--python-archive", default="", help="copy a local portable Python archive")
    offline_prepare.add_argument("--python-url", default="", help="download portable Python archive from this URL")
    offline_prepare.add_argument("--python-version", default="", help="build Python mirror URL from this version when --python-url is omitted")
    offline_prepare.add_argument("--python-mirror", choices=["official", "tuna", "npmmirror"], default="official", help="Python mirror alias")
    offline_prepare.add_argument("--python-mirror-base", default="", help="custom Python mirror base")
    offline_prepare.add_argument("--python-filename", default="", help="portable Python archive filename")
    offline_prepare.add_argument("--no-extract-python", action="store_true", help="keep Python archive only; do not extract into runtime/python")
    offline_prepare.add_argument("--allow-system-python-on-b", action="store_true", help="allow packages without an extracted portable Python runtime")
    offline_prepare.add_argument("--mineru-venv", default="", help="copy a tested MinerU virtualenv/portable env into runtime/mineru_venv")
    offline_prepare.add_argument("--mineru-model-dir", default="", help="copy local MinerU model directory into runtime/mineru_models")
    offline_prepare.add_argument("--mineru-config", default="", help="copy mineru.json as a bundle-local config template")
    offline_prepare.add_argument("--download-mineru-models", action="store_true", help="download MinerU models on computer A before packaging")
    offline_prepare.add_argument(
        "--mineru-model-source",
        choices=["huggingface", "modelscope"],
        default=DEFAULT_MINERU_MODEL_SOURCE,
        help=f"MinerU model download source; default {DEFAULT_MINERU_MODEL_SOURCE}",
    )
    offline_prepare.add_argument(
        "--mineru-model-type",
        choices=["pipeline", "vlm", "all"],
        default=DEFAULT_MINERU_MODEL_TYPE,
        help=f"MinerU model type to download; default {DEFAULT_MINERU_MODEL_TYPE}",
    )
    offline_prepare.add_argument("--huggingface-endpoint", default="", help="HF_ENDPOINT used when downloading MinerU models from a Hugging Face mirror")
    offline_prepare.add_argument("--mineru-model-downloader", default="", help="path to mineru-models-download; default resolves from --mineru-venv or PATH")
    offline_prepare.add_argument("--download-wheels", action="store_true", help="download project dependency wheels into wheelhouse")
    offline_prepare.add_argument("--pip-index-url", default="", help="pip mirror index URL used by pip download")
    offline_prepare.add_argument("--pip-extra-index-url", default="", help="extra pip index URL used by pip download")
    offline_prepare.add_argument("--include-mineru-wheels", action="store_true", help="also download optional MinerU runtime wheels")
    offline_prepare.add_argument(
        "--mineru-wheel-spec",
        default=DEFAULT_MINERU_WHEEL_SPEC,
        help=f"MinerU pip requirement for optional runtime wheels; default {DEFAULT_MINERU_WHEEL_SPEC}",
    )
    offline_prepare.add_argument("--strict-mineru-wheels", action="store_true", help="fail prepare when optional MinerU runtime wheel download fails")
    offline_prepare.add_argument("--asset-cache-dir", default="", help="local cache for reusable heavy migration assets; default <out-dir>/_asset_cache")
    offline_prepare.add_argument("--no-reuse-assets", action="store_true", help="disable reuse of cached Python archives, MinerU models and wheelhouse")
    offline_prepare.add_argument("--skip-datasheet-data", action="store_true", help="do not copy local datasheet_data index")
    offline_prepare.add_argument("--include-datasheet-source", action="store_true", help="copy PDF folders from PSTX_DATASHEET_DIR into data/datasheets")
    offline_prepare.add_argument("--no-zip", action="store_true", help="do not create a zip archive")
    offline_prepare.add_argument("--json-out", default="", help="write offline migration envelope JSON")
    offline_prepare.set_defaults(func=cmd_offline_migration)

    offline_verify = offline_sub.add_parser("verify", help="verify an offline migration bundle on computer B")
    _add_pretty_arg(offline_verify)
    offline_verify.add_argument("package_root", help="extracted bundle directory or zip path")
    offline_verify.add_argument("--skip-runtime-probe", action="store_true", help="skip portable Python dependency import probe")
    offline_verify.add_argument("--json-out", default="", help="write offline verification envelope JSON")
    offline_verify.set_defaults(func=cmd_offline_migration)

    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    argv = list(argv) if argv is not None else sys.argv[1:]
    try:
        args = parser.parse_args(argv)
        return args.func(args)
    except Exception as exc:  # pragma: no cover - exercised by CLI users.
        command = ""
        for item in argv:
            if item and not item.startswith("-"):
                command = item
                break
        pretty = "--pretty" in argv
        if "args" in locals():
            command = getattr(args, "command", command)
            pretty = bool(getattr(args, "pretty", pretty))
        _emit(_error_envelope(command, exc), pretty=pretty)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
