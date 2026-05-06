# -*- coding: utf-8 -*-
"""BOM_OPTION circle coverage checks, including submodule circle lookup."""

import math
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pstx_core.cadence import csa_geometry as pstx_csa_geometry
from pstx_core.page_resolution import (
    USER_VISIBLE_REAL_PAGE_LABEL,
    _component_display_page,
    _component_logical_page,
    _component_page_fields,
    _component_submodule_mapped_page,
    _normalize_page_label,
    component_user_visible_page,
)
from pstx_rules.common import (
    COMP_TYPE_CN,
    _NUMBER_VALUE_RE,
    _format_measure,
    _format_percent,
    _format_ratio,
    _is_depop_option,
    _natural_sort_key,
    _normalize_bom_option,
    _parse_xy_point,
)
from pstx_rules.result_meta import meta_fields as _meta_fields, with_meta as _with_meta

def _component_center_point(comp: Dict) -> Optional[Tuple[float, float]]:
    x_raw = comp.get('xy_x', '')
    y_raw = comp.get('xy_y', '')
    if x_raw != '' and y_raw != '':
        try:
            return float(x_raw), float(y_raw)
        except (TypeError, ValueError):
            pass
    return _parse_xy_point(comp.get('xy', ''))


def _bom_option_candidate_pages(comp: Dict) -> List[Tuple[str, str]]:
    candidates = []
    mapped_page = _normalize_page_label(_component_submodule_mapped_page(comp))
    real_page = _normalize_page_label(_component_display_page(comp))
    page_sources = (
        [(USER_VISIBLE_REAL_PAGE_LABEL, mapped_page)]
        if mapped_page else
        [(USER_VISIBLE_REAL_PAGE_LABEL, real_page)]
    )
    for source, page_label in page_sources:
        if page_label and page_label not in {label for _, label in candidates}:
            candidates.append((source, page_label))
    return candidates


def _page_label_number(page_label: str) -> str:
    match = re.search(r'(\d+)', str(page_label or ''))
    return str(int(match.group(1))) if match else ''


def _extract_module_name_from_order_key(module_order_key: str) -> str:
    text = str(module_order_key or '').strip()
    if not text:
        return ''
    matches = list(re.finditer(r'@(?P<head>[^@]+?)\(\s*SCH_1\s*\)', text, re.IGNORECASE))
    if len(matches) < 2:
        return ''
    head = matches[-1].group('head').strip()
    _, _, cell = head.rpartition('.')
    return (cell or head).strip()


def _find_sibling_module_root(project_root: str, module_name: str) -> str:
    if not project_root or not module_name:
        return ''
    try:
        main_root = Path(project_root).expanduser()
        worklib_root = main_root.parent
        direct = worklib_root / module_name.lower()
        if direct.is_dir():
            return str(direct)
        direct = worklib_root / module_name
        if direct.is_dir():
            return str(direct)
        if worklib_root.is_dir():
            module_lower = module_name.lower()
            for child in worklib_root.iterdir():
                if child.is_dir() and child.name.lower() == module_lower:
                    return str(child)
    except OSError:
        return ''
    return ''


def _submodule_circle_target(comp: Dict,
                             project_root: str,
                             geometry_cache: Dict[str, Dict]) -> Optional[Dict[str, object]]:
    local_page = _normalize_page_label(
        comp.get('page_submodule_real', '')
        or comp.get('module_order_local_page', '')
    )
    if not local_page:
        return None
    module_name = _extract_module_name_from_order_key(comp.get('module_order_key', ''))
    module_root = _find_sibling_module_root(project_root, module_name)
    if not module_root:
        return None
    if module_root not in geometry_cache:
        geometry_cache[module_root] = pstx_csa_geometry.analyze_csa_geometry(module_root)
    geometry = geometry_cache[module_root]
    if not geometry.get('enabled'):
        return None
    circles_by_page = _index_csa_circles(geometry)
    local_circles = circles_by_page.get(local_page, [])
    user_visible_page = component_user_visible_page(comp)
    display_page = _normalize_page_label(user_visible_page) or _normalize_page_label(_component_display_page(comp))
    translated_circles = []
    for circle in local_circles:
        circle_copy = dict(circle)
        row_copy = dict(circle.get('row', {}))
        row_copy['子模块名'] = module_name
        row_copy['子模块CSA页'] = local_page
        row_copy['子模块路径'] = module_root
        row_copy[USER_VISIBLE_REAL_PAGE_LABEL] = display_page
        if display_page:
            row_copy['页面'] = display_page
            circle_copy['page'] = display_page
        circle_copy['row'] = row_copy
        translated_circles.append(circle_copy)
    source = f'子模块CSA:{module_name}:{local_page}' if module_name else f'子模块CSA:{local_page}'
    return {
        'source': source,
        'display_page': display_page,
        'circle_page': local_page,
        'circles': translated_circles,
        'csa_enabled': True,
        'csa_file': str(Path(module_root) / 'sch_1' / f'page{_page_label_number(local_page)}.csa'),
    }


def _bom_option_circle_targets(comp: Dict,
                               root_circles_by_page: Dict[str, List[Dict[str, object]]],
                               project_root: str,
                               geometry_cache: Dict[str, Dict],
                               root_csa_enabled: bool) -> List[Dict[str, object]]:
    submodule_target = _submodule_circle_target(comp, project_root, geometry_cache) if project_root else None
    if submodule_target:
        return [submodule_target]
    targets: List[Dict[str, object]] = []
    for source, page_label in _bom_option_candidate_pages(comp):
        targets.append({
            'source': source,
            'display_page': page_label,
            'circle_page': page_label,
            'circles': root_circles_by_page.get(page_label, []),
            'csa_enabled': root_csa_enabled,
            'csa_file': '',
        })
    return targets


def _bom_option_coverage_units(comp: Dict) -> List[Dict[str, object]]:
    sections = comp.get('sections', [])
    units: List[Dict[str, object]] = []
    if isinstance(sections, list) and sections:
        for section in sections:
            if not isinstance(section, dict):
                continue
            section_comp = dict(comp)
            for aggregate_key in ('page_user_visible_pages', 'page_logical_pages', 'page_real_pages'):
                section_comp.pop(aggregate_key, None)
            section_comp.update(section)
            section_comp['bom_option'] = comp.get('bom_option', '')
            section_comp['comp_type'] = comp.get('comp_type', '')
            units.append({
                'section_number': str(section.get('section_number') or ''),
                'component': section_comp,
            })
    if units:
        return units
    return [{'section_number': '', 'component': comp}]


def _parse_csa_circle_row(row: Dict) -> Optional[Dict[str, object]]:
    center = _parse_xy_point(row.get('圆心', ''))
    radius_match = _NUMBER_VALUE_RE.search(str(row.get('半径', '') or ''))
    if not center or not radius_match:
        return None
    radius = abs(float(radius_match.group(0)))
    if radius <= 0:
        return None
    page_label = _normalize_page_label(row.get('页面', ''))
    if not page_label:
        return None
    return {
        'page': page_label,
        'center': center,
        'radius': radius,
        'row': row,
    }


def _index_csa_circles(csa_geometry: Dict) -> Dict[str, List[Dict[str, object]]]:
    circles_by_page: Dict[str, List[Dict[str, object]]] = defaultdict(list)
    for row in csa_geometry.get('circle_rows', []) or []:
        parsed = _parse_csa_circle_row(row)
        if parsed:
            circles_by_page[str(parsed['page'])].append(parsed)
    return circles_by_page


def _circle_boundary_tolerance(radius: float) -> float:
    return min(max(radius * 0.10, 20.0), 100.0)


def _build_bom_circle_base_row(refdes: str, comp: Dict, unit_comp: Optional[Dict] = None) -> Dict[str, str]:
    display_comp = unit_comp or comp
    ctype = comp.get('comp_type', '')
    row = {
        '位号': refdes,
        '类型': COMP_TYPE_CN.get(ctype, ctype),
        'BOM_OPTION值': _normalize_bom_option(comp.get('bom_option', '')),
        '是否DEPOP': '是' if _is_depop_option(comp.get('bom_option', '')) else '否',
        'SECTION_NUMBER': str(display_comp.get('section_number', '') or ''),
        '页面': component_user_visible_page(display_comp),
        'XY': display_comp.get('xy', ''),
    }
    row.update(_component_page_fields(display_comp))
    return row


def check_bom_option_circle_coverage(components: Dict,
                                     csa_geometry: Dict,
                                     project_root: str = '') -> Dict[str, List[dict]]:
    circles_by_page = _index_csa_circles(csa_geometry or {})
    csa_enabled = bool((csa_geometry or {}).get('enabled'))
    submodule_geometry_cache: Dict[str, Dict] = {}
    coverage_rows: List[dict] = []
    issue_rows: List[dict] = []

    for refdes, comp in sorted(components.items(), key=lambda item: _natural_sort_key(item[0])):
        ctype = comp.get('comp_type', '')
        if ctype == 'TESTPOINT':
            continue
        bom_option = _normalize_bom_option(comp.get('bom_option', ''))
        if not bom_option:
            continue

        for unit in _bom_option_coverage_units(comp):
            unit_comp = unit.get('component') if isinstance(unit, dict) else comp
            if not isinstance(unit_comp, dict):
                unit_comp = comp
            center = _component_center_point(unit_comp)
            circle_targets = _bom_option_circle_targets(
                unit_comp,
                circles_by_page,
                project_root,
                submodule_geometry_cache,
                csa_enabled,
            )
            candidate_page_labels = [
                str(target.get('display_page', ''))
                for target in circle_targets
                if str(target.get('display_page', ''))
            ]
            candidate_pages = [(str(target.get('source', '')), str(target.get('display_page', ''))) for target in circle_targets]
            has_any_csa = csa_enabled or any(bool(target.get('csa_enabled')) for target in circle_targets)
            base = _build_bom_circle_base_row(refdes, comp, unit_comp)
            base.update({
                '候选检查页': ', '.join(candidate_page_labels),
                '检查CSA页': '',
                '检查CSA文件': '',
                '覆盖状态': '',
                '中心重合度': '',
                '距离/半径': '',
                '中心距离': '',
                '边界余量': '',
                '最近画圈页': '',
                '最近画圈对象': '',
                '最近画圈行号': '',
                '最近画圈圆心': '',
                '最近画圈半径': '',
                '匹配来源': '',
                '说明': '',
            })

            if not has_any_csa:
                base['覆盖状态'] = '无法判断'
                base['说明'] = '未找到主模块或子模块 sch_1/page*.csa，无法检查 BOM_OPTION 打圈覆盖。'
                coverage_rows.append(base)
                issue_rows.append(_with_meta(base.copy(), 'indeterminate', 'medium', 'low', 'bom_option_circle_unknown_no_csa'))
                continue
            if center is None:
                base['覆盖状态'] = '无法判断'
                base['说明'] = '元件缺少可解析 XY 坐标，无法用中心点比对画圈范围。'
                coverage_rows.append(base)
                issue_rows.append(_with_meta(base.copy(), 'indeterminate', 'medium', 'low', 'bom_option_circle_unknown_no_xy'))
                continue
            if not candidate_pages:
                base['覆盖状态'] = '无法判断'
                base['说明'] = '元件缺少页码，无法定位对应 CSA 页面。'
                coverage_rows.append(base)
                issue_rows.append(_with_meta(base.copy(), 'indeterminate', 'medium', 'low', 'bom_option_circle_unknown_no_page'))
                continue

            best: Optional[Dict[str, object]] = None
            for target in circle_targets:
                source = str(target.get('source', ''))
                circles = target.get('circles', [])
                for circle in circles:
                    cx, cy = circle['center']  # type: ignore[misc]
                    radius = float(circle['radius'])
                    distance = math.hypot(center[0] - float(cx), center[1] - float(cy))
                    ratio = distance / radius if radius else float('inf')
                    margin = radius - distance
                    score = max(0.0, (1.0 - ratio) * 100.0)
                    candidate = {
                        'source': source,
                        'circle': circle,
                        'distance': distance,
                        'ratio': ratio,
                        'margin': margin,
                        'score': score,
                    }
                    candidate.update({
                        'circle_page': str(target.get('circle_page', '')),
                        'display_page': str(target.get('display_page', '')),
                        'csa_file': str(target.get('csa_file', '')),
                    })
                    if best is None or ratio < float(best['ratio']):
                        best = candidate

            if best is None:
                base['覆盖状态'] = '未打圈'
                base['说明'] = '该元件候选检查页未发现可用于覆盖判断的 CSA 画圈对象。'
                coverage_rows.append(base)
                issue_rows.append(_with_meta(base.copy(), 'candidate', 'medium', 'medium', 'bom_option_circle_missing_no_circle_on_page'))
                continue

            circle = best['circle']  # type: ignore[assignment]
            circle_row = circle['row']  # type: ignore[index]
            radius = float(circle['radius'])  # type: ignore[index]
            ratio = float(best['ratio'])
            distance = float(best['distance'])
            margin = float(best['margin'])
            object_type = str(circle_row.get('对象类型', ''))
            tolerance = _circle_boundary_tolerance(radius)
            is_arc_candidate = object_type.upper().startswith('ARC')
            if ratio <= 1.0 and not is_arc_candidate:
                status = '已打圈'
                note = '元件 XY 中心落在 CIRCLE 画圈范围内。'
            elif distance <= radius + tolerance:
                status = '疑似打圈'
                note = '元件 XY 中心接近画圈边界或命中 ARC 候选，建议人工快速复核。'
            else:
                status = '未打圈'
                note = '最近画圈对象未覆盖该元件 XY 中心。'

            base.update({
                '覆盖状态': status,
                '中心重合度': _format_percent(float(best['score'])),
                '距离/半径': _format_ratio(ratio),
                '中心距离': _format_measure(distance),
                '边界余量': _format_measure(margin),
                '最近画圈页': str(circle['page']),  # type: ignore[index]
                '检查CSA页': str(best.get('circle_page', '')),
                '检查CSA文件': str(best.get('csa_file', '')),
                '最近画圈对象': object_type,
                '最近画圈行号': circle_row.get('行号', ''),
                '最近画圈圆心': circle_row.get('圆心', ''),
                '最近画圈半径': _format_measure(radius),
                '匹配来源': str(best['source']),
                '说明': note,
            })
            coverage_rows.append(base)
            if status == '未打圈':
                issue_rows.append(_with_meta(base.copy(), 'candidate', 'medium', 'medium', 'bom_option_circle_missing'))

    return {
        'coverage_rows': sorted(
            coverage_rows,
            key=lambda row: (_natural_sort_key(row.get('位号', '')), _natural_sort_key(row.get('页面', ''))),
        ),
        'issue_rows': sorted(
            issue_rows,
            key=lambda row: (_natural_sort_key(row.get('位号', '')), _natural_sort_key(row.get('页面', ''))),
        ),
    }


# ══════════════════════════════════════════════════════════
# 三、网络分析
# ══════════════════════════════════════════════════════════
