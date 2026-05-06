# -*- coding: utf-8 -*-
"""Top-level PSTX project analysis orchestration.

This module wires the core parser and individual rule modules together.
It intentionally contains no Web route code and no Excel export code.
"""

import time
from typing import Dict, Optional

from pstx_core.analysis_cache import get_or_compute_cached_result
from pstx_core.cadence import csa_geometry as pstx_csa_geometry
from pstx_core.cadence.page_model import build_cadence_connectivity_summary
from pstx_core.page_resolution import apply_component_pages, prepare_page_resolution
from pstx_core.pstx_parser import parse_all
from pstx_rules.bom import build_bom, build_total_bom
from pstx_rules.bom_option_circle import check_bom_option_circle_coverage
from pstx_rules.common import _build_analysis_scope
from pstx_rules.derating import analyze_derating
from pstx_rules.drc import check_drc
from pstx_rules.module_scope import build_module_review
from pstx_rules.network import analyze_networks
from pstx_rules.resistor_bias import analyze_resistors


ANALYSIS_TIMINGS_SCHEMA_VERSION = "pstx-analysis-timings.v1"


def _new_analysis_timings() -> dict:
    return {
        "schema_version": ANALYSIS_TIMINGS_SCHEMA_VERSION,
        "stages": [],
        "cache": {},
        "total_stage_elapsed_ms": 0.0,
        "analysis_elapsed_ms": 0.0,
    }


def _append_timing_stage(timings: dict,
                         stage: str,
                         started_at: float,
                         **extra) -> dict:
    elapsed_ms = round((time.perf_counter() - started_at) * 1000.0, 3)
    row = {
        "stage": stage,
        "elapsed_ms": elapsed_ms,
    }
    row.update({key: value for key, value in extra.items() if value is not None})
    timings.setdefault("stages", []).append(row)
    timings["total_stage_elapsed_ms"] = round(
        sum(float(item.get("elapsed_ms", 0.0) or 0.0) for item in timings.get("stages", [])),
        3,
    )
    return row


def append_analysis_timing(bundle: dict,
                           stage: str,
                           elapsed_seconds: float,
                           **extra) -> dict:
    timings = bundle.setdefault("analysis_timings", _new_analysis_timings())
    row = {
        "stage": stage,
        "elapsed_ms": round(float(elapsed_seconds or 0.0) * 1000.0, 3),
    }
    row.update({key: value for key, value in extra.items() if value is not None})
    timings.setdefault("stages", []).append(row)
    timings["total_stage_elapsed_ms"] = round(
        sum(float(item.get("elapsed_ms", 0.0) or 0.0) for item in timings.get("stages", [])),
        3,
    )
    return row


def analyze_project_contents(prt_content: str,
                             net_content: str,
                             *,
                             project_name: str = '',
                             project_root: str = '',
                             ratio_limit: float = 70.0,
                             custom_volt_map: Optional[Dict[str, float]] = None,
                             include_depop: bool = False,
                             include_total_bom: bool = False) -> dict:
    analysis_started = time.perf_counter()
    timings = _new_analysis_timings()

    stage_started = time.perf_counter()
    components, nets, comp_nets = parse_all(prt_content, net_content)
    _append_timing_stage(timings, "parse_pstx", stage_started)

    stage_started = time.perf_counter()
    page_context = prepare_page_resolution(project_root)
    page_mapping = page_context.get('page_mapping', {})
    module_order_index = page_context.get('module_order_index', {})
    page_warnings = list(page_context.get('warnings', []))
    apply_component_pages(components, page_context)
    module_review = build_module_review(
        components,
        module_order_index,
        project_name=project_name,
    )
    _append_timing_stage(timings, "page_resolution", stage_started)

    stage_started = time.perf_counter()
    bom_normal_detail, bom_depop_detail, bom_normal_merged, bom_depop_merged = build_bom(components)
    bom_total_detail, bom_total_merged = build_total_bom(bom_normal_detail, bom_depop_detail)
    _append_timing_stage(timings, "bom", stage_started)

    stage_started = time.perf_counter()
    analysis_components, analysis_nets, depop_refdes, excluded_depop_refdes = _build_analysis_scope(
        components,
        nets,
        include_depop=include_depop,
    )
    analysis_comp_nets = {refdes: dict(comp.get('nets', {})) for refdes, comp in analysis_components.items()}
    _append_timing_stage(timings, "analysis_scope", stage_started)

    stage_started = time.perf_counter()
    net_analysis = analyze_networks(
        analysis_nets,
        analysis_components,
        single_node_topology_nets=nets,
    )
    _append_timing_stage(timings, "network_rules", stage_started)

    stage_started = time.perf_counter()
    drc = check_drc(
        analysis_components,
        analysis_nets,
        option_components_source=components,
        single_pin_components=components,
        single_pin_nets=nets,
    )
    _append_timing_stage(timings, "drc_rules", stage_started)

    stage_started = time.perf_counter()
    derating = analyze_derating(analysis_components, analysis_nets, ratio_limit, custom_volt_map)
    _append_timing_stage(timings, "derating_rules", stage_started)

    stage_started = time.perf_counter()
    resistor_analysis = analyze_resistors(analysis_components, analysis_nets, exclude_depop=not include_depop)
    _append_timing_stage(timings, "resistor_rules", stage_started)

    if project_root:
        stage_started = time.perf_counter()
        csa_geometry, csa_cache = get_or_compute_cached_result(
            "csa_geometry",
            project_root,
            params={"include_arcs": True, "circle_two_point_mode": "center_radius"},
            extensions=(".csa",),
            compute=lambda: pstx_csa_geometry.analyze_csa_geometry(project_root),
        )
        timings["cache"]["csa_geometry"] = csa_cache
        _append_timing_stage(timings, "csa_geometry", stage_started, cache_status=csa_cache.get("status"))
    else:
        csa_geometry = {
            'enabled': False,
            'root': '',
            'page_count': 0,
            'cross_count': 0,
            'circle_count': 0,
            'error_count': 0,
            'summary_rows': [],
            'dot_cross_rows': [],
            'circle_rows': [],
            'warnings': [],
        }

    if project_root:
        stage_started = time.perf_counter()
        cadence_page_semantics, cadence_cache = get_or_compute_cached_result(
            "cadence_page_semantics",
            project_root,
            params={"summary_only": True, "include_raw_unknown": True, "collect_junctions": False},
            extensions=(".csa", ".csv"),
            compute=lambda: build_cadence_connectivity_summary(project_root),
        )
        timings["cache"]["cadence_page_semantics"] = cadence_cache
        _append_timing_stage(
            timings,
            "cadence_page_semantics",
            stage_started,
            cache_status=cadence_cache.get("status"),
        )
    else:
        cadence_page_semantics = {
            'enabled': False,
            'schema_version': 'pstx-cadence-page.v1',
            'root': '',
            'page_count': 0,
            'rows': [],
            'warnings': [],
        }

    stage_started = time.perf_counter()
    bom_circle_coverage = check_bom_option_circle_coverage(components, csa_geometry, project_root=project_root)
    _append_timing_stage(timings, "bom_option_circle_coverage", stage_started)
    drc['bom_option_circle_coverage'] = bom_circle_coverage['coverage_rows']
    drc['bom_option_circle_issues'] = bom_circle_coverage['issue_rows']
    page_warnings.extend(csa_geometry.get('warnings', []))
    if depop_refdes:
        if include_depop:
            page_warnings.append(
                f'DEPOP 排查开关已开启：共有 {len(depop_refdes)} 个 BOM_OPTION=DEPOP/DNP 元件继续参与后续分析。'
            )
        else:
            preview = ', '.join(depop_refdes[:8])
            suffix = ' ...' if len(depop_refdes) > 8 else ''
            page_warnings.append(
                f'DEPOP 排查开关默认关闭：已在后续分析中忽略 {len(depop_refdes)} 个 BOM_OPTION=DEPOP/DNP 元件'
                f'（{preview}{suffix}）。'
            )
    timings["analysis_elapsed_ms"] = round((time.perf_counter() - analysis_started) * 1000.0, 3)
    return {
        'project_name': project_name,
        'project_root': project_root,
        'components': analysis_components,
        'nets': analysis_nets,
        'comp_nets': analysis_comp_nets,
        'all_components': components,
        'all_nets': nets,
        'all_comp_nets': comp_nets,
        'bom_normal_detail': bom_normal_detail,
        'bom_depop_detail': bom_depop_detail,
        'bom_total_detail': bom_total_detail,
        'bom_normal_merged': bom_normal_merged,
        'bom_depop_merged': bom_depop_merged,
        'bom_total_merged': bom_total_merged,
        'net_analysis': net_analysis,
        'page_mapping_rows': page_mapping.get('rows', []),
        'module_review': module_review,
        'drc': drc,
        'derating': derating,
        'resistor_analysis': resistor_analysis,
        'csa_geometry': csa_geometry,
        'cadence_page_semantics': cadence_page_semantics,
        'ratio_limit': ratio_limit,
        'custom_volt_map': custom_volt_map or None,
        'include_depop': include_depop,
        'include_total_bom': include_total_bom,
        'analysis_timings': timings,
        'depop_refdes': depop_refdes,
        'excluded_depop_refdes': excluded_depop_refdes,
        'page_warnings': page_warnings,
    }
