# -*- coding: utf-8 -*-
"""Report payload and report-oriented view model builders."""

from __future__ import annotations

from pstx_rules.result_meta import (
    DRC_ISSUE_KEYS as _DRC_ISSUE_KEYS,
    count_result_kinds as _count_result_kinds,
    iter_list_rows as _iter_list_rows,
)
from pstx_webapp.report_feishu import (
    build_feishu_hq_link_rows,
    enrich_bom_rows_with_feishu,
    enrich_chip_pin_rows_with_feishu,
    feishu_links_by_hq_code,
    feishu_links_by_refdes,
)
from pstx_webapp.report_tables import (
    METRIC_TARGETS,
    SECTION_LAYOUT,
    build_review_plan,
    build_report_table,
    build_section_cards,
    build_table_display_policy,
    build_top_insights,
)


def build_report_payload(run_id: str, bundle: dict) -> dict:
    na = bundle.get('net_analysis', {})
    drc = bundle.get('drc', {})
    drt = bundle.get('derating', [])
    res = bundle.get('resistor_analysis', {})
    csa = bundle.get('csa_geometry', {})
    cadence = bundle.get('cadence_page_semantics', {})
    module_review = bundle.get('module_review', {})
    mn = bundle.get('bom_normal_merged', [])
    md = bundle.get('bom_depop_merged', [])
    mt = bundle.get('bom_total_merged', [])

    resistor_rows = _iter_list_rows(res, ['divider_risks', 'dup_pullups', 'dup_pulldowns'])
    resistor_kind_counts = _count_result_kinds(resistor_rows)
    drc_total = sum(len(drc.get(key, [])) for key in _DRC_ISSUE_KEYS)
    derating_fail = sum(1 for row in drt if str(row.get('状态', '')).startswith('❌'))
    csa_candidate_total = int(csa.get('cross_count', 0) or 0)
    bom_circle_issue_count = len(drc.get('bom_option_circle_issues', []) or [])
    include_depop = bool(bundle.get('include_depop', False))
    include_total_bom = bool(bundle.get('include_total_bom', False))
    depop_refdes = list(bundle.get('depop_refdes', []) or [])
    excluded_depop_refdes = list(bundle.get('excluded_depop_refdes', []) or [])
    feishu_hq_links = build_feishu_hq_link_rows(bundle)
    feishu_by_refdes = feishu_links_by_refdes(feishu_hq_links)
    feishu_by_hq_code = feishu_links_by_hq_code(feishu_hq_links)
    bom_normal_merged = enrich_bom_rows_with_feishu(mn, feishu_by_hq_code)
    bom_depop_merged = enrich_bom_rows_with_feishu(md, feishu_by_hq_code)
    bom_total_merged = enrich_bom_rows_with_feishu(mt, feishu_by_hq_code)
    chip_pin_rows = enrich_chip_pin_rows_with_feishu(res.get('chip_pin_rows', []), feishu_by_refdes)
    module_summary = module_review.get('summary', {}) if isinstance(module_review, dict) else {}

    metrics = [
        {'label': '贴装种类', 'value': len(mn), 'tone': 'neutral', 'target': METRIC_TARGETS['贴装种类'], 'caption': 'BOM 视图'},
        {'label': '贴装总数', 'value': sum(row.get('数量', 0) for row in mn), 'tone': 'neutral', 'target': METRIC_TARGETS['贴装总数'], 'caption': '贴装器件总量'},
        {'label': 'DEPOP 总数', 'value': sum(row.get('数量', 0) for row in md), 'tone': 'muted', 'target': METRIC_TARGETS['DEPOP 总数'], 'caption': '去装配器件'},
        {'label': 'BOM圈问题', 'value': bom_circle_issue_count, 'tone': 'warning' if bom_circle_issue_count else 'ok', 'target': METRIC_TARGETS['BOM圈问题'], 'caption': 'BOM_OPTION 画圈覆盖'},
        {'label': '子模块数', 'value': module_summary.get('submodule_count', 0), 'tone': 'neutral', 'target': METRIC_TARGETS['子模块数'], 'caption': 'module_order 视角'},
        {'label': '网络总数', 'value': na.get('total', 0), 'tone': 'neutral', 'target': METRIC_TARGETS['网络总数'], 'caption': '网络总览'},
        {'label': 'DRC 总数', 'value': drc_total, 'tone': 'warning' if drc_total else 'ok', 'target': METRIC_TARGETS['DRC 总数'], 'caption': '设计检查结果'},
        {'label': '降额不合格', 'value': derating_fail, 'tone': 'warning' if derating_fail else 'ok', 'target': METRIC_TARGETS['降额不合格'], 'caption': '优先核查电容'},
        {'label': '电阻候选', 'value': resistor_kind_counts.get('候选判断', 0), 'tone': 'neutral', 'target': METRIC_TARGETS['电阻候选'], 'caption': '电阻规则候选项'},
        {'label': '电阻无法判断', 'value': resistor_kind_counts.get('无法判断', 0), 'tone': 'muted', 'target': METRIC_TARGETS['电阻无法判断'], 'caption': '待人工复核'},
        {'label': '规范候选', 'value': csa_candidate_total, 'tone': 'warning' if csa_candidate_total else 'ok', 'target': METRIC_TARGETS['规范候选'], 'caption': 'CSA 几何对象'},
    ]
    if include_total_bom:
        metrics.insert(3, {
            'label': '总BOM 总数',
            'value': sum(row.get('数量', 0) for row in mt),
            'tone': 'neutral',
            'target': METRIC_TARGETS['总BOM 总数'],
            'caption': '贴装 + DEPOP',
        })

    dataset_map = {
        'bom_total_merged': build_report_table(
            'bom_total_merged',
            '总 BOM',
            bom_total_merged,
            default_hidden_columns=['飞书来源库', '飞书来源Sheet'],
        ),
        'bom_normal_merged': build_report_table(
            'bom_normal_merged',
            '贴装 BOM',
            bom_normal_merged,
            default_hidden_columns=['飞书来源库', '飞书来源Sheet'],
        ),
        'bom_depop_merged': build_report_table(
            'bom_depop_merged',
            'DEPOP BOM',
            bom_depop_merged,
            default_hidden_columns=['飞书来源库', '飞书来源Sheet'],
        ),
        'bom_option_components': build_report_table('bom_option_components', 'BOM_OPTION 元件', drc.get('bom_option_components', [])),
        'power_net_rows': build_report_table('power_net_rows', '候选电源网络', na.get('power_net_rows', [])),
        'gnd_net_rows': build_report_table('gnd_net_rows', '候选 GND 网络', na.get('gnd_net_rows', [])),
        'diff_pair_rows': build_report_table('diff_pair_rows', '候选差分对', na.get('diff_pair_rows', [])),
        'single_node_rows': build_report_table('single_node_rows', '单节点网络概览', na.get('single_node_rows', [])),
        'page_rows': build_report_table('page_rows', '页码元件分布', na.get('page_rows', [])),
        'page_mapping_rows': build_report_table(
            'page_mapping_rows',
            '主模块页/页码映射检查',
            bundle.get('page_mapping_rows', []),
            default_hidden_columns=['涉及模块', '映射文件'],
        ),
        'module_scope_rows': build_report_table(
            'module_scope_rows',
            '模块范围汇总',
            module_review.get('module_rows', []) if isinstance(module_review, dict) else [],
            default_hidden_columns=['模块ID', 'module_order路径', 'module_order来源'],
        ),
        'module_component_rows': build_report_table(
            'module_component_rows',
            '模块元件索引',
            module_review.get('component_rows', []) if isinstance(module_review, dict) else [],
            default_hidden_columns=['模块ID'],
            sort_profiles=[
                {'id': 'column', 'label': '字段排序'},
                {'id': 'submodule', 'label': '子模块优先'},
            ],
            default_sort_mode='submodule',
        ),
        'missing_hq_code': build_report_table('missing_hq_code', '缺少料号', drc.get('missing_hq_code', [])),
        'missing_value': build_report_table('missing_value', '缺少 VALUE', drc.get('missing_value', [])),
        'missing_package': build_report_table('missing_package', '缺少封装', drc.get('missing_package', [])),
        'tbd_attrs': build_report_table('tbd_attrs', 'TBD 待确认属性', drc.get('tbd_attrs', [])),
        'single_pin_nets': build_report_table('single_pin_nets', '单端候选网络', drc.get('single_pin_nets', [])),
        'unnamed_nets': build_report_table('unnamed_nets', '未命名网络', drc.get('unnamed_nets', [])),
        'bom_option_typos': build_report_table('bom_option_typos', 'BOM_OPTION 候选拼写', drc.get('bom_option_typos', [])),
        'bom_option_circle_issues': build_report_table(
            'bom_option_circle_issues',
            'BOM_OPTION 打圈覆盖问题',
            drc.get('bom_option_circle_issues', []),
            default_hidden_columns=['候选检查页', '最近画圈行号', '最近画圈圆心', '最近画圈半径'],
        ),
        'bom_option_circle_coverage': build_report_table(
            'bom_option_circle_coverage',
            'BOM_OPTION 打圈覆盖明细',
            drc.get('bom_option_circle_coverage', []),
            default_hidden_columns=['候选检查页', '最近画圈行号', '最近画圈圆心', '最近画圈半径', '说明'],
        ),
        'csa_summary_rows': build_report_table('csa_summary_rows', 'CSA 页级汇总', csa.get('summary_rows', [])),
        'cadence_connectivity_rows': build_report_table(
            'cadence_connectivity_rows',
            'Cadence 连接语义页摘要',
            cadence.get('rows', []),
        ),
        'csa_dot_cross_rows': build_report_table(
            'csa_dot_cross_rows',
            'CSA DOT四向十字交叉',
            csa.get('dot_cross_rows', []),
            default_hidden_columns=['文件', '全部WIRE行号'],
        ),
        'divider_risks': build_report_table('divider_risks', '串阻分压候选风险', res.get('divider_risks', [])),
        'dup_pullups': build_report_table('dup_pullups', '重复上拉候选', res.get('dup_pullups', [])),
        'dup_pulldowns': build_report_table('dup_pulldowns', '重复下拉候选', res.get('dup_pulldowns', [])),
        'chip_pin_rows': build_report_table(
            'chip_pin_rows',
            '芯片 Pin 电阻状态',
            chip_pin_rows,
            default_hidden_columns=['后缀组', '子模块路径', '芯片飞书来源库', '芯片飞书来源Sheet'],
            sort_profiles=[
                {'id': 'column', 'label': '字段排序'},
                {'id': 'suffix_group', 'label': '后缀组优先'},
                {'id': 'submodule', 'label': '子模块优先'},
            ],
            default_sort_mode='submodule',
        ),
        'derating': build_report_table('derating', '电容降额结果', drt),
    }

    sections = []
    for section in SECTION_LAYOUT:
        table_entries = list(section['tables'])
        if section['id'] == 'bom' and include_total_bom:
            table_entries = [('总 BOM', 'bom_total_merged')] + table_entries
        tables = [dataset_map[key] for _, key in table_entries]
        sections.append({
            'id': section['id'],
            'title': section['title'],
            'lead': section['lead'],
            'tables': tables,
            'total_rows': sum(table['count'] for table in tables),
        })
    section_cards = build_section_cards(sections)

    summary_lines = [
        (
            f'DEPOP 排查：开启，{len(depop_refdes)} 个 DEPOP/DNP 元件继续参与分析'
            if include_depop else
            f'DEPOP 排查：关闭，后续分析已忽略 {len(excluded_depop_refdes)} 个 DEPOP/DNP 元件'
        ),
        (
            f'总 BOM：开启，显示贴装 + DEPOP 汇总（{sum(row.get("数量", 0) for row in mt)} 个器件）'
            if include_total_bom else
            '总 BOM：关闭，仅显示贴装 BOM 与 DEPOP BOM'
        ),
        (
            f'模块视角：识别到 {module_summary.get("submodule_count", 0)} 个子模块实例，可按主模块/子模块拆分复核'
            if module_summary else
            '模块视角：未生成 module_order 模块索引'
        ),
        '结果计数已汇总在顶部指标卡和重点提示中，下面分区只保留可展开的明细表。',
    ]
    top_insights = build_top_insights(
        drc_total=drc_total,
        derating_fail=derating_fail,
        resistor_kind_counts=resistor_kind_counts,
        csa_candidate_total=csa_candidate_total,
        warnings=bundle.get('warnings', []),
        section_cards=section_cards,
    )
    review_plan = build_review_plan(sections)
    table_display_policy = build_table_display_policy(sections)

    return {
        'run_id': run_id,
        'project_name': bundle.get('project_name') or '未命名项目',
        'generated_at': bundle.get('generated_at', ''),
        'ratio_limit': bundle.get('ratio_limit', 70.0),
        'include_depop': include_depop,
        'include_total_bom': include_total_bom,
        'depop_count': len(depop_refdes),
        'excluded_depop_count': len(excluded_depop_refdes),
        'custom_volt_map': bundle.get('custom_volt_map') or {},
        'warnings': bundle.get('warnings', []),
        'input_files': bundle.get('input_files', []),
        'analysis_timings': bundle.get('analysis_timings', {}),
        'metrics': metrics,
        'top_insights': top_insights,
        'section_cards': section_cards,
        'review_plan': review_plan,
        'table_display_policy': table_display_policy,
        'summary_lines': summary_lines,
        'feishu_hq_links': feishu_hq_links,
        'sections': sections,
    }
