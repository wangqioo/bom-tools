# -*- coding: utf-8 -*-
"""Project compare payload orchestration."""

from __future__ import annotations

import time

from pstx_rules.result_meta import DRC_ISSUE_KEYS as _DRC_ISSUE_KEYS
from pstx_webapp.compare_diff import (
    compare_component_inventory,
    compare_component_pin_nets,
    compare_report_tables,
    component_compare_value_with_feishu,
    build_net_view_diff,
    diff_named_maps,
    net_compare_value,
)
from pstx_webapp.compare_sections import build_compare_sections
from pstx_webapp.compare_view import (
    DEFAULT_COMPARE_DETAIL_LIMIT,
    build_compare_scalar_metrics,
    is_key_refdes,
    is_passive_refdes,
)
from pstx_webapp.report_feishu import (
    build_feishu_hq_link_rows,
    feishu_links_by_refdes,
    feishu_links_from_report,
)
from pstx_webapp.run_store import build_project_summary


def _project_summary(run_id: str, payload: dict) -> dict:
    return build_project_summary(run_id, payload, drc_issue_keys=_DRC_ISSUE_KEYS)


def build_compare_payload(left_run_id: str, right_run_id: str, *, get_run_payload, detail_limit: int = DEFAULT_COMPARE_DETAIL_LIMIT) -> dict:
    left_payload = get_run_payload(left_run_id)
    right_payload = get_run_payload(right_run_id)
    left_bundle = left_payload.get('bundle', {})
    right_bundle = right_payload.get('bundle', {})
    left_feishu = feishu_links_from_report(left_payload.get('report', {}))
    right_feishu = feishu_links_from_report(right_payload.get('report', {}))
    if not left_feishu:
        left_feishu = feishu_links_by_refdes(build_feishu_hq_link_rows(left_bundle))
    if not right_feishu:
        right_feishu = feishu_links_by_refdes(build_feishu_hq_link_rows(right_bundle))

    left_summary = _project_summary(left_run_id, left_payload)
    right_summary = _project_summary(right_run_id, right_payload)
    left_components_for_compare = {
        refdes: component_compare_value_with_feishu(comp, left_feishu.get(refdes.upper()))
        for refdes, comp in (left_bundle.get('components', {}) or {}).items()
    }
    right_components_for_compare = {
        refdes: component_compare_value_with_feishu(comp, right_feishu.get(refdes.upper()))
        for refdes, comp in (right_bundle.get('components', {}) or {}).items()
    }
    component_diff = diff_named_maps(
        left_components_for_compare,
        right_components_for_compare,
        title='元件差异',
        key_label='位号',
        detail_limit=detail_limit,
    )
    key_component_diff = compare_component_inventory(
        left_bundle.get('components', {}) or {},
        right_bundle.get('components', {}) or {},
        title='关键器件增删',
        predicate=is_key_refdes,
        detail_limit=detail_limit,
        left_feishu=left_feishu,
        right_feishu=right_feishu,
    )
    key_pin_net_diff = compare_component_pin_nets(
        left_bundle.get('components', {}) or {},
        right_bundle.get('components', {}) or {},
        left_bundle.get('nets', {}) or {},
        right_bundle.get('nets', {}) or {},
        title='关键器件 Pin/Net 连接差异',
        predicate=is_key_refdes,
        detail_limit=detail_limit,
        include_all_rows=True,
        left_feishu=left_feishu,
        right_feishu=right_feishu,
    )
    passive_pin_net_diff = compare_component_pin_nets(
        left_bundle.get('components', {}) or {},
        right_bundle.get('components', {}) or {},
        left_bundle.get('nets', {}) or {},
        right_bundle.get('nets', {}) or {},
        title='R/C/L Pin/Net 连接差异',
        predicate=is_passive_refdes,
        detail_limit=detail_limit,
        include_all_rows=True,
        left_feishu=left_feishu,
        right_feishu=right_feishu,
    )
    net_diff = diff_named_maps(
        left_bundle.get('nets', {}) or {},
        right_bundle.get('nets', {}) or {},
        title='网络差异',
        key_label='网络名',
        value_builder=net_compare_value,
        detail_limit=detail_limit,
        include_all_rows=True,
    )
    net_view_diff = build_net_view_diff(
        left_bundle.get('nets', {}) or {},
        right_bundle.get('nets', {}) or {},
        key_pin_net_diff,
        passive_pin_net_diff,
        net_diff,
        detail_limit=detail_limit,
    )
    for diff in (key_pin_net_diff, passive_pin_net_diff, net_diff):
        diff.pop('_all_rows', None)
    table_diffs = compare_report_tables(
        left_payload.get('report', {}),
        right_payload.get('report', {}),
        detail_limit=detail_limit,
    )
    overview = build_compare_scalar_metrics(left_summary, right_summary)
    payload = {
        'ok': True,
        'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'detail_limit': detail_limit,
        'left': left_summary,
        'right': right_summary,
        'overview': overview,
        'key_component_diff': key_component_diff,
        'key_pin_net_diff': key_pin_net_diff,
        'passive_pin_net_diff': passive_pin_net_diff,
        'net_view_diff': net_view_diff,
        'component_diff': component_diff,
        'net_diff': net_diff,
        'table_diffs': table_diffs,
        'diff_totals': {
            'overview': len(overview),
            'key_components': key_component_diff['added_count'] + key_component_diff['removed_count'] + key_component_diff['changed_count'],
            'net_view': net_view_diff['added_count'] + net_view_diff['removed_count'] + net_view_diff['changed_count'],
            'key_pin_nets': key_pin_net_diff['added_count'] + key_pin_net_diff['removed_count'] + key_pin_net_diff['changed_count'],
            'passive_pin_nets': passive_pin_net_diff['added_count'] + passive_pin_net_diff['removed_count'] + passive_pin_net_diff['changed_count'],
            'components': component_diff['added_count'] + component_diff['removed_count'] + component_diff['changed_count'],
            'nets': net_diff['added_count'] + net_diff['removed_count'] + net_diff['changed_count'],
            'tables': sum(item['added_count'] + item['removed_count'] + item['changed_count'] for item in table_diffs),
        },
    }
    payload['compare_sections'] = build_compare_sections(payload)
    return payload
