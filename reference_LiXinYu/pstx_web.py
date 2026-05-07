# -*- coding: utf-8 -*-
"""
PSTX localhost Web UI

Run:
    python pstx_web.py
"""

from __future__ import annotations

import argparse
import io
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from collections import OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pstx_analyzer import (
    _DRC_ISSUE_KEYS,
    _count_result_kinds,
    _iter_list_rows,
    analyze_project_contents,
    export_to_excel,
    query_project_data,
)
from pstx_aster_service import (
    aster_error_payload,
    build_aster_status,
    build_aster_summary,
    clear_aster_runtime_config,
    set_aster_runtime_config,
)


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


Flask = abort = jsonify = render_template = request = send_file = url_for = None


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


BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / 'web'
DEFAULT_HOST = '127.0.0.1'
DEFAULT_PORT = 44441
MAX_RUNS = 12
MAX_COMPARE_DETAIL_ROWS = 200
RUN_CACHE: "OrderedDict[str, dict]" = OrderedDict()
METRIC_TARGETS = {
    '贴装种类': 'bom',
    '贴装总数': 'bom',
    'DEPOP 总数': 'bom',
    '网络总数': 'network',
    'DRC 总数': 'drc',
    '降额不合格': 'derating',
    '电阻候选': 'resistor',
    '电阻无法判断': 'resistor',
    '规范候选': 'csa',
}


SECTION_LAYOUT = [
    {
        'id': 'bom',
        'title': 'BOM 视图',
        'lead': '展示贴装、去装配与变体配置，支持快速复核物料范围与装配差异。',
        'tables': [
            ('贴装 BOM', 'bom_normal_merged'),
            ('DEPOP BOM', 'bom_depop_merged'),
            ('BOM_OPTION 元件', 'bom_option_components'),
        ],
    },
    {
        'id': 'network',
        'title': '网络分析',
        'lead': '按网络视角汇总候选电源、接地、差分对、单节点网络与页面分布。',
        'tables': [
            ('候选电源网络', 'power_net_rows'),
            ('候选 GND 网络', 'gnd_net_rows'),
            ('候选差分对', 'diff_pair_rows'),
            ('单节点候选网络', 'single_node_rows'),
            ('各页面元件数（真实页）', 'page_rows'),
            ('逻辑页/真实页映射检查', 'page_mapping_rows'),
        ],
    },
    {
        'id': 'drc',
        'title': '设计检查',
        'lead': '集中展示属性缺失、命名异常和 BOM_OPTION 相关检查项。',
        'tables': [
            ('缺少料号', 'missing_hq_code'),
            ('缺少 VALUE', 'missing_value'),
            ('缺少封装', 'missing_package'),
            ('TBD 待确认属性', 'tbd_attrs'),
            ('单端候选网络', 'single_pin_nets'),
            ('未命名网络', 'unnamed_nets'),
            ('BOM_OPTION 候选拼写', 'bom_option_typos'),
            ('BOM_OPTION 元件', 'bom_option_components'),
        ],
    },
    {
        'id': 'csa',
        'title': '规范检查',
        'lead': '扫描 sch_1/page*.csa 几何对象，复核带 DOT 四向十字交叉与画圈标注。',
        'tables': [
            ('CSA 页级汇总', 'csa_summary_rows'),
            ('CSA DOT四向十字交叉', 'csa_dot_cross_rows'),
            ('CSA 画圈对象', 'csa_circle_rows'),
        ],
    },
    {
        'id': 'resistor',
        'title': '电阻检查',
        'lead': '面向偏置、串阻、OD/OC 以及芯片引脚关联的规则检查结果。',
        'tables': [
            ('串阻分压候选风险', 'divider_risks'),
            ('重复上拉候选', 'dup_pullups'),
            ('重复下拉候选', 'dup_pulldowns'),
            ('OD/OC 候选缺上拉', 'od_missing'),
            ('芯片 Pin 电阻状态', 'chip_pin_rows'),
        ],
    },
    {
        'id': 'derating',
        'title': '电容降额',
        'lead': '汇总工作电压推断、降额比值、原因代码与无法判断项。',
        'tables': [
            ('电容降额结果', 'derating'),
        ],
    },
]


def _parse_voltage_map_text(text: str) -> Tuple[Optional[Dict[str, float]], List[str]]:
    mapping: Dict[str, float] = {}
    warnings: List[str] = []
    for idx, raw_line in enumerate((text or '').splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith('#'):
            continue
        if '=' not in line:
            warnings.append(f'电压映射第 {idx} 行缺少 "="：{raw_line.strip()}')
            continue
        key, _, value = line.partition('=')
        key = key.strip()
        value = value.strip()
        if not key:
            warnings.append(f'电压映射第 {idx} 行前缀为空：{raw_line.strip()}')
            continue
        try:
            mapping[key] = float(value)
        except ValueError:
            warnings.append(f'电压映射第 {idx} 行电压不是有效数字：{raw_line.strip()}')
    return mapping or None, warnings


def _parse_checkbox_flag(value: object) -> bool:
    return str(value or '').strip().lower() in {'1', 'true', 'yes', 'on'}


def _port_is_available(port: int, host: str = DEFAULT_HOST) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def _resolve_port(preferred_port: int, host: str = DEFAULT_HOST, max_attempts: int = 20) -> int:
    for offset in range(max_attempts + 1):
        candidate = preferred_port + offset
        if _port_is_available(candidate, host):
            return candidate
    raise RuntimeError(
        f'Unable to find a free localhost port in range {preferred_port}-{preferred_port + max_attempts}.'
    )


TEXT_DECODE_ENCODINGS = (
    'utf-8-sig',
    'utf-8',
    'utf-16',
    'utf-16-le',
    'utf-16-be',
    'gb18030',
    'cp936',
)
TEXT_DECODE_MARKERS = (
    'PART_NAME',
    'NET_NAME',
    'NODE_NAME',
    'SECTION_NUMBER',
    'PAGE_NUMBER',
    'BOM_OPTION',
    'C_PATH',
    'P_PATH',
)


def _score_decoded_text(text: str) -> int:
    upper_text = str(text or '').upper()
    marker_score = sum(upper_text.count(marker) for marker in TEXT_DECODE_MARKERS) * 1000
    control_penalty = sum(
        1
        for char in text
        if ord(char) < 32 and char not in {'\r', '\n', '\t'}
    ) * 50
    ascii_score = sum(1 for char in text if 32 <= ord(char) < 127)
    return marker_score + ascii_score - control_penalty


def _decode_text_bytes(data: bytes) -> Tuple[str, str]:
    candidates = []
    for order, encoding in enumerate(TEXT_DECODE_ENCODINGS):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        candidates.append((_score_decoded_text(text), -order, text, encoding))
    if candidates:
        _, _, text, encoding = max(candidates, key=lambda item: (item[0], item[1]))
        return text, encoding
    return data.decode('utf-8', errors='replace'), 'utf-8-replace'


def _read_local_text_file(path: Path, label: str, required: bool) -> Tuple[Optional[str], Dict[str, str]]:
    if not path.exists():
        if required:
            raise ValueError(f'缺少必需文件：{path}')
        return None, {'label': label, 'filename': str(path), 'size': '0', 'encoding': ''}
    data = path.read_bytes()
    text, encoding = _decode_text_bytes(data)
    return text, {
        'label': label,
        'filename': str(path),
        'size': str(len(data)),
        'encoding': encoding,
    }


def _resolve_project_root(root_text: str) -> Path:
    raw = (root_text or '').strip().strip('"')
    if not raw:
        raise ValueError('请输入项目根路径')
    root = Path(raw).expanduser()
    if root.name.lower() == 'packaged':
        root = root.parent
    if not root.exists():
        raise ValueError(f'项目根路径不存在：{root}')
    if not root.is_dir():
        raise ValueError(f'项目根路径不是文件夹：{root}')
    return root


def _discover_project_files(root_text: str) -> Tuple[Path, Path, Path, Optional[Path]]:
    project_root = _resolve_project_root(root_text)
    packaged_dir = project_root / 'packaged'
    if not packaged_dir.is_dir():
        raise ValueError(f'项目根路径下缺少 packaged 文件夹：{packaged_dir}')

    prt_path = packaged_dir / 'pstxprt.dat'
    net_path = packaged_dir / 'pstxnet.dat'
    ref_path = packaged_dir / 'pstxref.dat'
    if not prt_path.is_file():
        raise ValueError(f'未找到输入文件：{prt_path}')
    if not net_path.is_file():
        raise ValueError(f'未找到输入文件：{net_path}')
    return project_root, prt_path, net_path, (ref_path if ref_path.is_file() else None)


def _decode_upload(file_storage, label: str, required: bool) -> Tuple[Optional[str], Dict[str, str]]:
    if not file_storage or not getattr(file_storage, 'filename', ''):
        if required:
            raise ValueError(f'请上传 {label}')
        return None, {'label': label, 'filename': '', 'size': '0', 'encoding': ''}
    data = file_storage.read()
    text, encoding = _decode_text_bytes(data)
    return text, {
        'label': label,
        'filename': file_storage.filename,
        'size': str(len(data)),
        'encoding': encoding,
    }


def _report_table(table_id: str,
                  title: str,
                  rows: List[dict],
                  *,
                  default_hidden_columns: Optional[List[str]] = None,
                  sort_profiles: Optional[List[dict]] = None,
                  default_sort_mode: str = 'column') -> dict:
    columns = list(rows[0].keys()) if rows else []
    return {
        'id': table_id,
        'title': title,
        'count': len(rows),
        'columns': columns,
        'rows': rows,
        'kind_counts': dict(_count_result_kinds(rows)),
        'default_hidden_columns': default_hidden_columns or [],
        'sort_profiles': sort_profiles or [{'id': 'column', 'label': '字段排序'}],
        'default_sort_mode': default_sort_mode,
    }


def _build_top_insights(
    *,
    drc_total: int,
    derating_fail: int,
    resistor_kind_counts,
    csa_candidate_total: int,
    warnings: List[str],
    section_cards: List[dict],
) -> List[dict]:
    insights = []
    if drc_total:
        insights.append({
            'title': '优先处理 DRC 检查项',
            'body': f'当前共有 {drc_total} 项 DRC 结果，建议先进入设计检查分区复核。',
            'tone': 'warning',
            'target': 'drc',
        })
    if derating_fail:
        insights.append({
            'title': '存在降额不合格项',
            'body': f'当前识别到 {derating_fail} 项降额不满足阈值，建议优先确认电容工作电压依据。',
            'tone': 'danger',
            'target': 'derating',
        })
    resistor_candidates = resistor_kind_counts.get('候选判断', 0)
    if resistor_candidates:
        insights.append({
            'title': '电阻相关候选项较多',
            'body': f'当前有 {resistor_candidates} 项电阻候选结果，建议结合偏置和串阻关系优先筛查。',
            'tone': 'neutral',
            'target': 'resistor',
        })
    if csa_candidate_total:
        insights.append({
            'title': '发现 CSA 几何规范候选项',
            'body': f'当前有 {csa_candidate_total} 项 CSA 几何候选结果，建议进入规范检查分区核对页面坐标和原始行号。',
            'tone': 'warning',
            'target': 'csa',
        })
    if warnings:
        insights.append({
            'title': '存在补充说明',
            'body': warnings[0],
            'tone': 'neutral',
            'target': 'summary',
        })
    if not insights:
        top_section = max(section_cards, key=lambda item: item.get('rows', 0), default=None)
        section_name = top_section['title'] if top_section else '本次报告'
        insights.append({
            'title': '结果已生成，可按分区复核',
            'body': f'当前未发现需要优先弹出的高风险项，建议从 {section_name} 开始浏览。',
            'tone': 'ok',
            'target': top_section['id'] if top_section else 'summary',
        })
    return insights[:4]


def _build_section_cards(sections: List[dict]) -> List[dict]:
    cards = []
    for section in sections:
        non_empty_tables = [table for table in section['tables'] if table['count'] > 0]
        top_table = max(section['tables'], key=lambda table: table['count'], default=None)
        top_label = top_table['title'] if top_table and top_table['count'] else '暂无重点表'
        top_value = top_table['count'] if top_table else 0
        tone = 'ok' if section['total_rows'] == 0 else ('warning' if section['id'] in {'drc', 'derating', 'resistor', 'csa'} else 'neutral')
        cards.append({
            'id': section['id'],
            'title': section['title'],
            'rows': section['total_rows'],
            'active_tables': len(non_empty_tables),
            'top_label': top_label,
            'top_value': top_value,
            'tone': tone,
            'lead': section['lead'],
        })
    return cards


def _build_report_payload(run_id: str, bundle: dict) -> dict:
    na = bundle.get('net_analysis', {})
    drc = bundle.get('drc', {})
    drt = bundle.get('derating', [])
    res = bundle.get('resistor_analysis', {})
    csa = bundle.get('csa_geometry', {})
    mn = bundle.get('bom_normal_merged', [])
    md = bundle.get('bom_depop_merged', [])

    network_rows = _iter_list_rows(na, ['power_net_rows', 'gnd_net_rows', 'diff_pair_rows', 'single_node_rows'])
    drc_rows = _iter_list_rows(drc, _DRC_ISSUE_KEYS)
    resistor_rows = _iter_list_rows(res, ['divider_risks', 'dup_pullups', 'dup_pulldowns', 'od_missing'])
    net_kind_counts = _count_result_kinds(network_rows)
    drc_kind_counts = _count_result_kinds(drc_rows)
    drt_kind_counts = _count_result_kinds(drt)
    resistor_kind_counts = _count_result_kinds(resistor_rows)
    drc_total = sum(len(drc.get(key, [])) for key in _DRC_ISSUE_KEYS)
    derating_fail = sum(1 for row in drt if str(row.get('状态', '')).startswith('❌'))
    csa_candidate_total = int(csa.get('cross_count', 0) or 0) + int(csa.get('circle_count', 0) or 0)
    include_depop = bool(bundle.get('include_depop', False))
    depop_refdes = list(bundle.get('depop_refdes', []) or [])
    excluded_depop_refdes = list(bundle.get('excluded_depop_refdes', []) or [])

    metrics = [
        {'label': '贴装种类', 'value': len(mn), 'tone': 'neutral', 'target': METRIC_TARGETS['贴装种类'], 'caption': 'BOM 视图'},
        {'label': '贴装总数', 'value': sum(row.get('数量', 0) for row in mn), 'tone': 'neutral', 'target': METRIC_TARGETS['贴装总数'], 'caption': '贴装器件总量'},
        {'label': 'DEPOP 总数', 'value': sum(row.get('数量', 0) for row in md), 'tone': 'muted', 'target': METRIC_TARGETS['DEPOP 总数'], 'caption': '去装配器件'},
        {'label': '网络总数', 'value': na.get('total', 0), 'tone': 'neutral', 'target': METRIC_TARGETS['网络总数'], 'caption': '网络总览'},
        {'label': 'DRC 总数', 'value': drc_total, 'tone': 'warning' if drc_total else 'ok', 'target': METRIC_TARGETS['DRC 总数'], 'caption': '设计检查结果'},
        {'label': '降额不合格', 'value': derating_fail, 'tone': 'warning' if derating_fail else 'ok', 'target': METRIC_TARGETS['降额不合格'], 'caption': '优先核查电容'},
        {'label': '电阻候选', 'value': resistor_kind_counts.get('候选判断', 0), 'tone': 'neutral', 'target': METRIC_TARGETS['电阻候选'], 'caption': '电阻规则候选项'},
        {'label': '电阻无法判断', 'value': resistor_kind_counts.get('无法判断', 0), 'tone': 'muted', 'target': METRIC_TARGETS['电阻无法判断'], 'caption': '待人工复核'},
        {'label': '规范候选', 'value': csa_candidate_total, 'tone': 'warning' if csa_candidate_total else 'ok', 'target': METRIC_TARGETS['规范候选'], 'caption': 'CSA 几何对象'},
    ]

    dataset_map = {
        'bom_normal_merged': _report_table('bom_normal_merged', '贴装 BOM', mn),
        'bom_depop_merged': _report_table('bom_depop_merged', 'DEPOP BOM', md),
        'bom_option_components': _report_table('bom_option_components', 'BOM_OPTION 元件', drc.get('bom_option_components', [])),
        'power_net_rows': _report_table('power_net_rows', '候选电源网络', na.get('power_net_rows', [])),
        'gnd_net_rows': _report_table('gnd_net_rows', '候选 GND 网络', na.get('gnd_net_rows', [])),
        'diff_pair_rows': _report_table('diff_pair_rows', '候选差分对', na.get('diff_pair_rows', [])),
        'single_node_rows': _report_table('single_node_rows', '单节点候选网络', na.get('single_node_rows', [])),
        'page_rows': _report_table('page_rows', '各页面元件数（真实页）', na.get('page_rows', [])),
        'page_mapping_rows': _report_table(
            'page_mapping_rows',
            '逻辑页/真实页映射检查',
            bundle.get('page_mapping_rows', []),
            default_hidden_columns=['涉及模块', '映射文件'],
        ),
        'missing_hq_code': _report_table('missing_hq_code', '缺少料号', drc.get('missing_hq_code', [])),
        'missing_value': _report_table('missing_value', '缺少 VALUE', drc.get('missing_value', [])),
        'missing_package': _report_table('missing_package', '缺少封装', drc.get('missing_package', [])),
        'tbd_attrs': _report_table('tbd_attrs', 'TBD 待确认属性', drc.get('tbd_attrs', [])),
        'single_pin_nets': _report_table('single_pin_nets', '单端候选网络', drc.get('single_pin_nets', [])),
        'unnamed_nets': _report_table('unnamed_nets', '未命名网络', drc.get('unnamed_nets', [])),
        'bom_option_typos': _report_table('bom_option_typos', 'BOM_OPTION 候选拼写', drc.get('bom_option_typos', [])),
        'csa_summary_rows': _report_table('csa_summary_rows', 'CSA 页级汇总', csa.get('summary_rows', [])),
        'csa_dot_cross_rows': _report_table(
            'csa_dot_cross_rows',
            'CSA DOT四向十字交叉',
            csa.get('dot_cross_rows', []),
            default_hidden_columns=['文件', '全部WIRE行号'],
        ),
        'csa_circle_rows': _report_table(
            'csa_circle_rows',
            'CSA 画圈对象',
            csa.get('circle_rows', []),
            default_hidden_columns=['文件', '原始行', '外接框', '解析说明'],
        ),
        'divider_risks': _report_table('divider_risks', '串阻分压候选风险', res.get('divider_risks', [])),
        'dup_pullups': _report_table('dup_pullups', '重复上拉候选', res.get('dup_pullups', [])),
        'dup_pulldowns': _report_table('dup_pulldowns', '重复下拉候选', res.get('dup_pulldowns', [])),
        'od_missing': _report_table('od_missing', 'OD/OC 候选缺上拉', res.get('od_missing', [])),
        'chip_pin_rows': _report_table(
            'chip_pin_rows',
            '芯片 Pin 电阻状态',
            res.get('chip_pin_rows', []),
            default_hidden_columns=['后缀组', '子模块路径'],
            sort_profiles=[
                {'id': 'column', 'label': '字段排序'},
                {'id': 'suffix_group', 'label': '后缀组优先'},
                {'id': 'submodule', 'label': '子模块优先'},
            ],
            default_sort_mode='submodule',
        ),
        'derating': _report_table('derating', '电容降额结果', drt),
    }

    sections = []
    for section in SECTION_LAYOUT:
        tables = [dataset_map[key] for _, key in section['tables']]
        sections.append({
            'id': section['id'],
            'title': section['title'],
            'lead': section['lead'],
            'tables': tables,
            'total_rows': sum(table['count'] for table in tables),
        })
    section_cards = _build_section_cards(sections)

    summary_lines = [
        (
            f'DEPOP 排查：开启，{len(depop_refdes)} 个 DEPOP/DNP 元件继续参与分析'
            if include_depop else
            f'DEPOP 排查：关闭，后续分析已忽略 {len(excluded_depop_refdes)} 个 DEPOP/DNP 元件'
        ),
        f'网络候选判断：{net_kind_counts.get("候选判断", 0)}',
        f'DRC 确定结论：{drc_kind_counts.get("确定结论", 0)}',
        f'DRC 候选判断：{drc_kind_counts.get("候选判断", 0)}',
        f'降额候选判断：{drt_kind_counts.get("候选判断", 0)}',
        f'降额无法判断：{drt_kind_counts.get("无法判断", 0)}',
        f'电阻候选判断：{resistor_kind_counts.get("候选判断", 0)}',
        f'CSA 几何候选：{csa_candidate_total}',
    ]
    top_insights = _build_top_insights(
        drc_total=drc_total,
        derating_fail=derating_fail,
        resistor_kind_counts=resistor_kind_counts,
        csa_candidate_total=csa_candidate_total,
        warnings=bundle.get('warnings', []),
        section_cards=section_cards,
    )

    return {
        'run_id': run_id,
        'project_name': bundle.get('project_name') or '未命名项目',
        'generated_at': bundle.get('generated_at', ''),
        'ratio_limit': bundle.get('ratio_limit', 70.0),
        'include_depop': include_depop,
        'depop_count': len(depop_refdes),
        'excluded_depop_count': len(excluded_depop_refdes),
        'custom_volt_map': bundle.get('custom_volt_map') or {},
        'warnings': bundle.get('warnings', []),
        'input_files': bundle.get('input_files', []),
        'metrics': metrics,
        'top_insights': top_insights,
        'section_cards': section_cards,
        'summary_lines': summary_lines,
        'sections': sections,
    }


def _remember_run(run_id: str, payload: dict) -> None:
    RUN_CACHE[run_id] = payload
    RUN_CACHE.move_to_end(run_id)
    while len(RUN_CACHE) > MAX_RUNS:
        RUN_CACHE.popitem(last=False)


def _get_run(run_id: str) -> dict:
    payload = RUN_CACHE.get(run_id)
    if not payload:
        abort(404)
    return payload


def _json_fingerprint(value) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def _compact_value(value, limit: int = 180) -> str:
    if isinstance(value, (dict, list, tuple)):
        text = _json_fingerprint(value)
    else:
        text = str(value if value is not None else '')
    return text if len(text) <= limit else text[:limit - 1] + '…'


def _project_summary(run_id: str, payload: dict) -> dict:
    bundle = payload.get('bundle', {})
    report = payload.get('report', {})
    drc = bundle.get('drc', {})
    metrics = report.get('metrics', [])
    metric_map = {str(item.get('label', '')): item.get('value') for item in metrics}
    return {
        'run_id': run_id,
        'project_name': report.get('project_name') or bundle.get('project_name') or '未命名项目',
        'project_root': bundle.get('project_root', ''),
        'generated_at': report.get('generated_at') or bundle.get('generated_at', ''),
        'ratio_limit': report.get('ratio_limit', bundle.get('ratio_limit', '')),
        'include_depop': bool(report.get('include_depop', bundle.get('include_depop', False))),
        'component_count': len(bundle.get('components', {}) or {}),
        'net_count': len(bundle.get('nets', {}) or {}),
        'drc_count': sum(len(drc.get(key, [])) for key in _DRC_ISSUE_KEYS),
        'metrics': metrics,
        'metric_map': metric_map,
    }


def _compare_scalar_metrics(left: dict, right: dict) -> List[dict]:
    labels = list(dict.fromkeys(
        list(left.get('metric_map', {}).keys())
        + list(right.get('metric_map', {}).keys())
        + ['component_count', 'net_count', 'drc_count']
    ))
    rows = []
    for label in labels:
        if label == 'component_count':
            left_value, right_value, display_label = left.get('component_count', 0), right.get('component_count', 0), '元件数'
        elif label == 'net_count':
            left_value, right_value, display_label = left.get('net_count', 0), right.get('net_count', 0), '网络数'
        elif label == 'drc_count':
            left_value, right_value, display_label = left.get('drc_count', 0), right.get('drc_count', 0), 'DRC 问题数'
        else:
            left_value = left.get('metric_map', {}).get(label, '')
            right_value = right.get('metric_map', {}).get(label, '')
            display_label = label
        if left_value == right_value:
            continue
        delta = ''
        if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
            diff = right_value - left_value
            delta = f'{diff:+g}'
        rows.append({
            '指标': display_label,
            '左侧': left_value,
            '右侧': right_value,
            '变化': delta or 'changed',
        })
    return rows


def _component_compare_value(comp: dict) -> dict:
    return {
        '类型': comp.get('comp_type', ''),
        '料号': comp.get('hq_code', ''),
        '值': comp.get('value', ''),
        '封装': comp.get('package', ''),
        'BOM_OPTION': comp.get('bom_option', ''),
        '页面': comp.get('page', ''),
        '网络': comp.get('nets', {}),
    }


def _net_compare_value(nodes: List[dict]) -> List[str]:
    return sorted(
        (
            f"{node.get('refdes', '')}:{node.get('pin', '')}:{node.get('pin_name', '')}"
            for node in nodes or []
        ),
        key=str.upper,
    )


def _diff_named_maps(left_map: dict,
                     right_map: dict,
                     *,
                     title: str,
                     key_label: str,
                     value_builder=None) -> dict:
    value_builder = value_builder or (lambda value: value)
    left_keys = set(left_map)
    right_keys = set(right_map)
    added = sorted(right_keys - left_keys, key=str.upper)
    removed = sorted(left_keys - right_keys, key=str.upper)
    shared = sorted(left_keys & right_keys, key=str.upper)
    rows = []
    for key in added:
        rows.append({'类型': '新增', key_label: key, '左侧': '', '右侧': _compact_value(value_builder(right_map[key])), '变化字段': '新增'})
    for key in removed:
        rows.append({'类型': '删除', key_label: key, '左侧': _compact_value(value_builder(left_map[key])), '右侧': '', '变化字段': '删除'})
    changed_count = 0
    for key in shared:
        left_value = value_builder(left_map[key])
        right_value = value_builder(right_map[key])
        if left_value == right_value:
            continue
        changed_count += 1
        changed_fields = []
        if isinstance(left_value, dict) and isinstance(right_value, dict):
            for field in sorted(set(left_value) | set(right_value)):
                if left_value.get(field) != right_value.get(field):
                    changed_fields.append(field)
        rows.append({
            '类型': '变化',
            key_label: key,
            '左侧': _compact_value(left_value),
            '右侧': _compact_value(right_value),
            '变化字段': ', '.join(changed_fields) or '内容变化',
        })
    return {
        'title': title,
        'key_label': key_label,
        'added_count': len(added),
        'removed_count': len(removed),
        'changed_count': changed_count,
        'rows': rows[:MAX_COMPARE_DETAIL_ROWS],
        'total_rows': len(rows),
        'truncated': len(rows) > MAX_COMPARE_DETAIL_ROWS,
    }


def _row_compare_key(row: dict) -> str:
    priority = [
        '位号', '网络名', '芯片位号', '引脚', '基础名', 'P端网络', 'N端网络',
        '使用该值的位号', '料号', '值', '封装', '逻辑页', '真实页', '页面', '原因代码', '状态',
    ]
    fields = [field for field in priority if field in row]
    if fields:
        return ' | '.join(f'{field}={row.get(field, "")}' for field in fields[:4])
    return _json_fingerprint(row)


def _table_rows_by_key(table: dict) -> Dict[str, dict]:
    rows_by_key: Dict[str, dict] = {}
    for index, row in enumerate(table.get('rows', []) or []):
        key = _row_compare_key(row)
        if key in rows_by_key:
            key = f'{key} #{index + 1}'
        rows_by_key[key] = row
    return rows_by_key


def _flatten_report_tables(report: dict) -> Dict[str, dict]:
    tables: Dict[str, dict] = {}
    for section in report.get('sections', []) or []:
        for table in section.get('tables', []) or []:
            tables[table.get('id') or table.get('title')] = table
    return tables


def _compare_report_tables(left_report: dict, right_report: dict) -> List[dict]:
    left_tables = _flatten_report_tables(left_report)
    right_tables = _flatten_report_tables(right_report)
    results = []
    for table_id in sorted(set(left_tables) | set(right_tables)):
        left_table = left_tables.get(table_id, {'title': table_id, 'rows': []})
        right_table = right_tables.get(table_id, {'title': table_id, 'rows': []})
        diff = _diff_named_maps(
            _table_rows_by_key(left_table),
            _table_rows_by_key(right_table),
            title=right_table.get('title') or left_table.get('title') or table_id,
            key_label='对象',
        )
        diff['id'] = table_id
        if diff['added_count'] or diff['removed_count'] or diff['changed_count']:
            results.append(diff)
    return results


def _build_compare_payload(left_run_id: str, right_run_id: str) -> dict:
    left_payload = _get_run(left_run_id)
    right_payload = _get_run(right_run_id)
    left_bundle = left_payload.get('bundle', {})
    right_bundle = right_payload.get('bundle', {})
    left_summary = _project_summary(left_run_id, left_payload)
    right_summary = _project_summary(right_run_id, right_payload)
    component_diff = _diff_named_maps(
        left_bundle.get('components', {}) or {},
        right_bundle.get('components', {}) or {},
        title='元件差异',
        key_label='位号',
        value_builder=_component_compare_value,
    )
    net_diff = _diff_named_maps(
        left_bundle.get('nets', {}) or {},
        right_bundle.get('nets', {}) or {},
        title='网络差异',
        key_label='网络名',
        value_builder=_net_compare_value,
    )
    table_diffs = _compare_report_tables(left_payload.get('report', {}), right_payload.get('report', {}))
    overview = _compare_scalar_metrics(left_summary, right_summary)
    return {
        'ok': True,
        'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'left': left_summary,
        'right': right_summary,
        'overview': overview,
        'component_diff': component_diff,
        'net_diff': net_diff,
        'table_diffs': table_diffs,
        'diff_totals': {
            'overview': len(overview),
            'components': component_diff['added_count'] + component_diff['removed_count'] + component_diff['changed_count'],
            'nets': net_diff['added_count'] + net_diff['removed_count'] + net_diff['changed_count'],
            'tables': sum(item['added_count'] + item['removed_count'] + item['changed_count'] for item in table_diffs),
        },
    }


def create_app() -> "Flask":
    _ensure_flask_loaded()
    app = Flask(
        __name__,
        template_folder=str(WEB_DIR / 'templates'),
        static_folder=str(WEB_DIR / 'static'),
    )

    @app.get('/')
    def home():
        host_text = request.host or f'{DEFAULT_HOST}:{DEFAULT_PORT}'
        listen_port = host_text.rsplit(':', 1)[-1] if ':' in host_text else str(DEFAULT_PORT)
        return render_template('index.html', listen_host=DEFAULT_HOST, listen_port=listen_port)

    @app.get('/debug/effects')
    def debug_effects():
        return render_template('debug_effects.html')

    @app.get('/debug/report-open')
    def debug_report_open():
        return render_template('debug_report_open.html')

    @app.post('/api/analyze')
    def analyze_upload():
        try:
            project_root, prt_path, net_path, ref_path = _discover_project_files(request.form.get('project_root', ''))
            prt_text, prt_meta = _read_local_text_file(prt_path, 'pstxprt.dat', True)
            net_text, net_meta = _read_local_text_file(net_path, 'pstxnet.dat', True)
            ref_text, ref_meta = _read_local_text_file(
                ref_path or (project_root / 'packaged' / 'pstxref.dat'),
                'pstxref.dat',
                False,
            )
            project_name = (request.form.get('project_name') or '').strip()
            ratio_limit = float(request.form.get('ratio_limit') or 70)
            custom_volt_map, map_warnings = _parse_voltage_map_text(request.form.get('custom_volt_map', ''))
            include_depop = _parse_checkbox_flag(request.form.get('include_depop'))
        except ValueError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'参数解析失败：{exc}'}), 400

        run_id = uuid.uuid4().hex[:12]
        bundle = analyze_project_contents(
            prt_text or '',
            net_text or '',
            project_name=project_name or project_root.name,
            project_root=str(project_root),
            ratio_limit=ratio_limit,
            custom_volt_map=custom_volt_map,
            include_depop=include_depop,
        )
        warnings = list(map_warnings) + list(bundle.get('page_warnings', []))
        if ref_text is not None:
            warnings.append('pstxref.dat 已接收，当前版本仅保留文件记录，暂不参与分析结果。')
        bundle.update({
            'project_name': project_name or '未命名项目',
            'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
            'warnings': warnings,
            'input_files': [prt_meta, net_meta, ref_meta],
        })
        bundle['project_name'] = project_name or project_root.name
        bundle['project_root'] = str(project_root)
        payload = {
            'bundle': bundle,
            'report': _build_report_payload(run_id, bundle),
        }
        _remember_run(run_id, payload)
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'redirect_url': url_for('report_page', run_id=run_id),
        })

    @app.get('/report/<run_id>')
    def report_page(run_id: str):
        payload = _get_run(run_id)
        report = payload['report']
        return render_template('report.html', run_id=run_id, report=report)

    @app.get('/api/report/<run_id>')
    def report_data(run_id: str):
        payload = _get_run(run_id)
        return jsonify(payload['report'])

    @app.get('/api/report/<run_id>/aster-summary')
    def aster_summary(run_id: str):
        payload = _get_run(run_id)
        try:
            return jsonify(build_aster_summary(payload['report'], payload['bundle']))
        except Exception as exc:
            error_payload, status = aster_error_payload(exc)
            return jsonify(error_payload), status

    @app.get('/api/aster/status')
    def aster_status():
        return jsonify(build_aster_status())

    @app.post('/api/aster/runtime-config')
    def aster_runtime_config_update():
        data = request.get_json(silent=True) or request.form.to_dict()
        try:
            return jsonify(set_aster_runtime_config(data))
        except Exception as exc:
            error_payload, status = aster_error_payload(exc)
            return jsonify(error_payload), status

    @app.delete('/api/aster/runtime-config')
    def aster_runtime_config_clear():
        return jsonify(clear_aster_runtime_config())

    @app.get('/api/projects')
    def project_list():
        projects = [
            _project_summary(run_id, payload)
            for run_id, payload in reversed(RUN_CACHE.items())
        ]
        return jsonify({
            'ok': True,
            'count': len(projects),
            'projects': projects,
        })

    @app.post('/api/compare')
    def compare_projects():
        data = request.get_json(silent=True) or request.form
        left_run_id = str(data.get('left_run_id') or '').strip()
        right_run_id = str(data.get('right_run_id') or '').strip()
        if not left_run_id or not right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个项目后再对比。'}), 400
        if left_run_id == right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个不同项目进行对比。'}), 400
        return jsonify(_build_compare_payload(left_run_id, right_run_id))

    @app.post('/api/report/<run_id>/query')
    def query_report(run_id: str):
        payload = _get_run(run_id)
        data = request.get_json(silent=True) or {}
        mode = data.get('mode') or '位号'
        keyword = data.get('keyword') or ''
        result = query_project_data(payload['bundle']['components'], payload['bundle']['nets'], mode, keyword)
        return jsonify(result)

    @app.get('/api/report/<run_id>/export')
    def export_report(run_id: str):
        payload = _get_run(run_id)
        fd, target = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        os.unlink(target)
        actual = export_to_excel(payload['bundle'], target)
        try:
            with open(actual, 'rb') as handle:
                data = handle.read()
        finally:
            try:
                os.remove(actual)
            except OSError:
                pass
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=Path(actual).name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    return app


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description='Run PSTX localhost web UI')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='localhost port, default 44441')
    parser.add_argument('--no-browser', action='store_true', help='do not auto-open the browser')
    args = parser.parse_args(argv)

    resolved_port = _resolve_port(args.port, DEFAULT_HOST)
    app = create_app()
    url = f'http://{DEFAULT_HOST}:{resolved_port}/'
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    if resolved_port != args.port:
        print(f'Requested port {args.port} is busy; falling back to localhost port {resolved_port}.')
    print(f'PSTX Web UI is listening on {url}')
    print('This service is bound to 127.0.0.1 only and cannot be accessed from other machines.')
    app.run(host=DEFAULT_HOST, port=resolved_port, debug=False)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
