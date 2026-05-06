# -*- coding: utf-8 -*-
"""Feishu HQ inline view helpers for report and compare payloads."""

from __future__ import annotations

from typing import Dict, List, Optional

from pstx_integrations.feishu.gateway import match_rows_with_feishu_cache

MAX_FEISHU_PREVIEW_ROWS = 500
MAX_FEISHU_REPORT_LINK_ROWS = 5000

FEISHU_BOM_ROW_SOURCES = {
    'bom_normal_detail': '贴装 BOM 明细',
    'bom_depop_detail': 'DEPOP BOM 明细',
    'bom_total_detail': '总 BOM 明细',
    'bom_normal_merged': '贴装 BOM 汇总',
    'bom_depop_merged': 'DEPOP BOM 汇总',
    'bom_total_merged': '总 BOM 汇总',
}


def build_feishu_hq_link_rows(bundle: dict) -> List[dict]:
    project_rows: List[dict] = []
    for source_key, bom_state in [
        ('bom_normal_detail', '贴装'),
        ('bom_depop_detail', 'DEPOP'),
    ]:
        for row in bundle.get(source_key, []) or []:
            enriched = dict(row)
            enriched['BOM状态'] = bom_state
            project_rows.append(enriched)
    if not project_rows:
        return []

    try:
        match = match_rows_with_feishu_cache(
            project_rows,
            '料号',
            limit=MAX_FEISHU_REPORT_LINK_ROWS,
            match_mode='hq_no',
        )
    except Exception:
        return []
    if not match.get('ok'):
        return []

    def review_status(row: dict) -> tuple[str, str]:
        status = str(row.get('匹配状态') or '')
        if status.startswith('跳过'):
            return '待补充：Cadence 料号为空', '项目HQ料号'
        if status == '未匹配':
            return '未匹配：飞书缓存库没有该 HQ 料号', '飞书HQ料号, 飞书规格型号, PI, 选型顺序'
        if status != '已匹配':
            return status or '未知状态', ''
        missing = []
        if not str(row.get('飞书规格型号') or row.get('HQ规格型号') or '').strip():
            missing.append('飞书规格型号')
        if not str(row.get('PI') or '').strip():
            missing.append('PI')
        if not str(row.get('选型顺序') or '').strip():
            missing.append('选型顺序')
        if int(row.get('匹配数量') or 0) > 1:
            prefix = f"需复核：同一 HQ 料号命中 {row.get('匹配数量')} 条"
            if missing:
                return f"{prefix}，且缺少 {', '.join(missing)}", ', '.join(missing)
            return prefix, ''
        if missing:
            return f"已匹配，字段待补充：{', '.join(missing)}", ', '.join(missing)
        return '通过：HQ 料号命中，飞书规格/PI/选型顺序已补齐', ''

    ordered_rows: List[dict] = []
    for row in match.get('rows', []) or []:
        conclusion, missing_fields = review_status(row)
        ordered_rows.append({
            '序号': row.get('序号', ''),
            '校对结论': conclusion,
            '缺失字段': missing_fields,
            'BOM状态': row.get('BOM状态', ''),
            '位号': row.get('位号', ''),
            '项目HQ料号': row.get('项目HQ料号', ''),
            '项目规格型号': row.get('项目规格型号', ''),
            '项目值': row.get('项目值', ''),
            '项目封装': row.get('项目封装', ''),
            '项目类型': row.get('项目类型', ''),
            '匹配状态': row.get('匹配状态', ''),
            '匹配数量': row.get('匹配数量', 0),
            '飞书HQ料号': row.get('飞书HQ料号', row.get('HQ料号', '')),
            '飞书规格型号': row.get('飞书规格型号', row.get('HQ规格型号', '')),
            'PI': row.get('PI', ''),
            '选型顺序': row.get('选型顺序', ''),
            '来源库': row.get('来源库', ''),
            '来源Sheet': row.get('来源Sheet', ''),
            'HQ制造商': row.get('HQ制造商', ''),
            'HQ描述': row.get('HQ描述', ''),
            '缓存行ID': row.get('缓存行ID', ''),
            '匹配方式': row.get('匹配方式', ''),
            '全部匹配': row.get('全部匹配', []),
        })
    return ordered_rows


def feishu_links_by_refdes(rows: List[dict]) -> Dict[str, dict]:
    return {
        str(row.get('位号', '')).strip().upper(): row
        for row in rows or []
        if str(row.get('位号', '')).strip()
    }


def feishu_links_by_hq_code(rows: List[dict]) -> Dict[str, dict]:
    links: Dict[str, dict] = {}
    for row in rows or []:
        hq_code = str(row.get('项目HQ料号', '') or row.get('飞书HQ料号', '')).strip().upper()
        if hq_code and hq_code not in links:
            links[hq_code] = row
    return links


def feishu_links_from_report(report: dict) -> Dict[str, dict]:
    if isinstance(report.get('feishu_hq_links'), list):
        return feishu_links_by_refdes(report.get('feishu_hq_links') or [])
    for section in report.get('sections', []) or []:
        for table in section.get('tables', []) or []:
            if table.get('id') == 'feishu_hq_links':
                return feishu_links_by_refdes(table.get('rows', []) or [])
    return {}


def feishu_link_value(link: Optional[dict]) -> dict:
    if not link:
        return {
            '项目HQ料号': '',
            '飞书HQ料号': '',
            '飞书规格型号': '',
            'PI': '',
            '选型顺序': '',
            '飞书校对结论': '',
            '来源库': '',
            '来源Sheet': '',
        }
    return {
        '项目HQ料号': link.get('项目HQ料号', ''),
        '飞书HQ料号': link.get('飞书HQ料号', ''),
        '飞书规格型号': link.get('飞书规格型号', ''),
        'PI': link.get('PI', ''),
        '选型顺序': link.get('选型顺序', ''),
        '飞书校对结论': link.get('校对结论', ''),
        '来源库': link.get('来源库', ''),
        '来源Sheet': link.get('来源Sheet', ''),
    }


def empty_feishu_link_for_hq(hq_code: object) -> dict:
    hq_text = str(hq_code or '').strip()
    if not hq_text:
        conclusion = '待补充：Cadence 料号为空'
    else:
        conclusion = '未匹配：飞书缓存库没有该 HQ 料号'
    return {
        '项目HQ料号': hq_text,
        '飞书HQ料号': '',
        '飞书规格型号': '',
        'PI': '',
        '选型顺序': '',
        '飞书校对结论': conclusion,
        '来源库': '',
        '来源Sheet': '',
    }


def enrich_bom_rows_with_feishu(rows: List[dict], feishu_by_hq_code: Dict[str, dict]) -> List[dict]:
    enriched_rows = []
    for row in rows or []:
        hq_code = str(row.get('料号', '') or row.get('HQ料号', '')).strip()
        feishu = feishu_link_value(feishu_by_hq_code.get(hq_code.upper())) if hq_code else empty_feishu_link_for_hq('')
        if hq_code and not feishu['飞书校对结论']:
            feishu = empty_feishu_link_for_hq(hq_code)
        enriched = dict(row)
        enriched.update({
            '飞书HQ料号': feishu['飞书HQ料号'],
            '飞书规格型号': feishu['飞书规格型号'],
            'PI': feishu['PI'],
            '选型顺序': feishu['选型顺序'],
            '飞书校对结论': feishu['飞书校对结论'],
            '飞书来源库': feishu['来源库'],
            '飞书来源Sheet': feishu['来源Sheet'],
        })
        enriched_rows.append(enriched)
    return enriched_rows


def enrich_chip_pin_rows_with_feishu(rows: List[dict], feishu_by_refdes: Dict[str, dict]) -> List[dict]:
    enriched_rows = []
    for row in rows or []:
        refdes = str(row.get('芯片位号', '')).strip().upper()
        feishu = feishu_link_value(feishu_by_refdes.get(refdes))
        enriched = dict(row)
        enriched.update({
            '芯片项目HQ料号': feishu['项目HQ料号'],
            '芯片飞书HQ料号': feishu['飞书HQ料号'],
            '芯片飞书规格型号': feishu['飞书规格型号'],
            '芯片PI': feishu['PI'],
            '芯片选型顺序': feishu['选型顺序'],
            '芯片飞书校对结论': feishu['飞书校对结论'],
            '芯片飞书来源库': feishu['来源库'],
            '芯片飞书来源Sheet': feishu['来源Sheet'],
        })
        enriched_rows.append(enriched)
    return enriched_rows
