# -*- coding: utf-8 -*-
"""CSA 几何规范检查工具 — Blueprint（与 csa_checker.py 功能完全一致）"""

import csv
import io
import math
import os
import re
import uuid
from pathlib import Path

from shared import (
    render_template, request, jsonify,
    UPLOAD_DIR,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint, Response

csa_bp = Blueprint('csa_tool', __name__)


# ── 常量 ─────────────────────────────────────────────────────

PAGE_FILE_RE = re.compile(r"^page(\d+)\.csa$", re.IGNORECASE)
WIRE_RE = re.compile(
    r"\bWIRE\s+\S+(?:\s+\S+)?\s+"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)\s*"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)",
    re.IGNORECASE,
)
DOT_RE = re.compile(
    r"(?<![A-Za-z0-9_])DOT\s+\S+(?:\s+\S+)?\s+"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)",
    re.IGNORECASE,
)
SIG_NAME_RE = re.compile(
    r"\bFORCEPROP\s+\S+\s+LAST\s+SIG_NAME\s+(.+?)(?=\s+J\s+\d+\b|;|$)",
    re.IGNORECASE,
)
PAGE_NUMBER_RE = re.compile(
    r"(?:\bSET\s+)?['\"]?\bPAGE_NUMBER\b['\"]?\s*(?:=\s*)?['\"]?([A-Z]*\d+)['\"]?",
    re.IGNORECASE,
)
CIRCLE_LINE_RE = re.compile(r"^\s*CIRCLE\b", re.IGNORECASE)
ARC_LINE_RE = re.compile(r"^\s*ARC\b", re.IGNORECASE)
COORD_RE = re.compile(r"\((-?\d+)\s*,?\s*(-?\d+)\)")
NUMBER_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")


# ── 工具函数 ─────────────────────────────────────────────────

def _natural_sort_key(value: str):
    parts = re.split(r'(\d+)', str(value or '').upper())
    return [int(p) if p.isdigit() else p for p in parts]


def _csa_page_no(path):
    m = PAGE_FILE_RE.match(os.path.basename(str(path)))
    return int(m.group(1)) if m else 10 ** 12


def _csa_page_label(no):
    return f"PAGE{no}" if no < 10 ** 12 else "UNKNOWN"


def _normalize_sig(sig: str) -> str:
    return sig.strip().rstrip(";").strip()


def _extract_coords(raw: str):
    return [(float(x), float(y)) for x, y in COORD_RE.findall(raw)]


def _fmt_num(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.2f}"


# ── 画圈解析 ─────────────────────────────────────────────────

def _circle_from_center_radius(obj_type, line_no, raw, center, radius, note):
    cx, cy = center
    r = abs(radius)
    return {
        'object_type': obj_type, 'line_no': line_no,
        'center_x': cx, 'center_y': cy, 'radius': r, 'diameter': 2 * r,
        'bbox_xmin': cx - r, 'bbox_ymin': cy - r,
        'bbox_xmax': cx + r, 'bbox_ymax': cy + r,
        'width': 2 * r, 'height': 2 * r,
        'raw': raw, 'parse_note': note,
    }


def _circle_from_bbox(obj_type, line_no, raw, p1, p2, note):
    x1, y1 = p1
    x2, y2 = p2
    xmin, xmax = sorted([x1, x2])
    ymin, ymax = sorted([y1, y2])
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2
    radius = max(xmax - xmin, ymax - ymin) / 2
    return {
        'object_type': obj_type, 'line_no': line_no,
        'center_x': cx, 'center_y': cy, 'radius': radius, 'diameter': 2 * radius,
        'bbox_xmin': xmin, 'bbox_ymin': ymin,
        'bbox_xmax': xmax, 'bbox_ymax': ymax,
        'width': xmax - xmin, 'height': ymax - ymin,
        'raw': raw, 'parse_note': note,
    }


def _fit_circle_three_points(p1, p2, p3):
    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    temp = x2 * x2 + y2 * y2
    bc = (x1 * x1 + y1 * y1 - temp) / 2.0
    cd = (temp - x3 * x3 - y3 * y3) / 2.0
    det = (x1 - x2) * (y2 - y3) - (x2 - x3) * (y1 - y2)
    if abs(det) < 1e-9:
        return None
    cx = (bc * (y2 - y3) - cd * (y1 - y2)) / det
    cy = ((x1 - x2) * cd - (x2 - x3) * bc) / det
    return cx, cy, math.hypot(cx - x1, cy - y1)


def parse_circle_line(raw, line_no, mode='center_radius'):
    coords = _extract_coords(raw)
    if len(coords) >= 2:
        if mode == 'bbox':
            return _circle_from_bbox(
                'CIRCLE', line_no, raw, coords[0], coords[1],
                "CIRCLE two-point mode: bbox diagonal points.")
        radius = math.hypot(coords[1][0] - coords[0][0], coords[1][1] - coords[0][1])
        return _circle_from_center_radius(
            'CIRCLE', line_no, raw, coords[0], radius,
            "CIRCLE two-point mode: center + radius point.")
    if len(coords) == 1:
        nums = [float(n) for n in NUMBER_RE.findall(COORD_RE.sub(" ", raw))]
        if nums:
            return _circle_from_center_radius(
                'CIRCLE', line_no, raw, coords[0], nums[-1],
                "CIRCLE parsed as center + numeric radius.")
    return None


def parse_arc_line_as_circle(raw, line_no):
    coords = _extract_coords(raw)
    if len(coords) >= 3:
        fit = _fit_circle_three_points(coords[0], coords[1], coords[2])
        if fit is None:
            return None
        cx, cy, radius = fit
        return _circle_from_center_radius(
            'ARC_FIT', line_no, raw, (cx, cy), radius,
            "ARC parsed by fitting a circle through three points; manually confirm.")
    if len(coords) == 2:
        return _circle_from_bbox(
            'ARC_DIAMETER_GUESS', line_no, raw, coords[0], coords[1],
            "ARC with two points parsed as a weak diameter/bbox guess; manually confirm.")
    return None


# ── DSU (Disjoint Set Union) ─────────────────────────────────

class _DSU:
    def __init__(self, ids):
        self.parent = {i: i for i in ids}

    def find(self, v):
        while self.parent[v] != v:
            self.parent[v] = self.parent[self.parent[v]]
            v = self.parent[v]
        return v

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


# ── 解析引擎 ─────────────────────────────────────────────────

def parse_csa_text(text, page_no, *, circle_mode='center_radius', include_arcs=True):
    wires = []
    dots = []
    circles = []
    page_name = _csa_page_label(page_no)
    last_wire_idx = None

    for line_no, raw_line in enumerate(str(text or '').splitlines(), start=1):
        raw = raw_line.strip()
        if not raw:
            continue

        page_match = PAGE_NUMBER_RE.search(raw)
        if page_match:
            page_name = page_match.group(1).upper()

        events = []
        for m in WIRE_RE.finditer(raw):
            events.append((m.start(), 'WIRE', m))
        for m in DOT_RE.finditer(raw):
            events.append((m.start(), 'DOT', m))
        for m in SIG_NAME_RE.finditer(raw):
            events.append((m.start(), 'SIG', m))
        events.sort(key=lambda x: x[0])

        for _, kind, match in events:
            if kind == 'WIRE':
                x1, y1, x2, y2 = map(int, match.groups())
                wires.append({
                    'wid': len(wires), 'line_no': line_no, 'raw': match.group(0).strip(),
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2, 'sig_name': '',
                })
                last_wire_idx = len(wires) - 1
            elif kind == 'DOT':
                x, y = map(int, match.groups()[-2:])
                dots.append({
                    'line_no': line_no, 'x': x, 'y': y, 'raw': match.group(0).strip(),
                })
            elif kind == 'SIG' and last_wire_idx is not None:
                wires[last_wire_idx]['sig_name'] = _normalize_sig(match.group(1))

        if CIRCLE_LINE_RE.match(raw):
            c = parse_circle_line(raw, line_no, circle_mode)
            if c:
                circles.append(c)
        elif include_arcs and ARC_LINE_RE.match(raw):
            c = parse_arc_line_as_circle(raw, line_no)
            if c:
                circles.append(c)

    return wires, dots, circles, page_name


def _wire_component_labels(wires):
    """使用 DSU 将端点相连的 WIRE 连通，并收集每个 WIRE 所属分量的信号名。"""
    dsu = _DSU(w['wid'] for w in wires)
    endpoint_map = {}
    for w in wires:
        is_h = w['y1'] == w['y2'] and w['x1'] != w['x2']
        is_v = w['x1'] == w['x2'] and w['y1'] != w['y2']
        if not (is_h or is_v):
            continue
        for pt in [(w['x1'], w['y1']), (w['x2'], w['y2'])]:
            endpoint_map.setdefault(pt, []).append(w['wid'])
    for ids in endpoint_map.values():
        if len(ids) >= 2:
            for other in ids[1:]:
                dsu.union(ids[0], other)
    root_labels = {}
    for w in wires:
        root = dsu.find(w['wid'])
        root_labels.setdefault(root, set())
        if w['sig_name']:
            root_labels[root].add(w['sig_name'])
    return {w['wid']: root_labels.get(dsu.find(w['wid']), set()) for w in wires}


def find_dot_four_way_crosses(wires, dots):
    findings = []
    labels_by_wire = _wire_component_labels(wires)
    required = {'left', 'right', 'down', 'up'}

    for dot in dots:
        dx, dy = dot['x'], dot['y']
        touching = [
            w for w in wires
            if ((w['y1'] == w['y2'] == dy and min(w['x1'], w['x2']) <= dx <= max(w['x1'], w['x2']))
                or (w['x1'] == w['x2'] == dx and min(w['y1'], w['y2']) <= dy <= max(w['y1'], w['y2'])))
        ]
        if not touching:
            continue

        dirs = set()
        h_wires = []
        v_wires = []
        for w in touching:
            if w['y1'] == w['y2']:  # horizontal
                h_wires.append(w)
                if min(w['x1'], w['x2']) < dx:
                    dirs.add('left')
                if dx < max(w['x1'], w['x2']):
                    dirs.add('right')
            else:  # vertical
                v_wires.append(w)
                if min(w['y1'], w['y2']) < dy:
                    dirs.add('down')
                if dy < max(w['y1'], w['y2']):
                    dirs.add('up')

        if not required.issubset(dirs):
            continue

        labels = set()
        for w in touching:
            labels.update(labels_by_wire.get(w['wid'], set()))
            if w['sig_name']:
                labels.add(w['sig_name'])

        findings.append({
            'x': dx, 'y': dy, 'dot_line': dot['line_no'],
            'h_wire_lines': ','.join(str(w['line_no']) for w in h_wires),
            'v_wire_lines': ','.join(str(w['line_no']) for w in v_wires),
            'all_wire_lines': ','.join(str(w['line_no']) for w in touching),
            'labels': ','.join(sorted(labels)) if labels else '',
            'detail': 'DOT point has WIREs extending left/right/up/down. T junctions and dotless crosses are ignored.',
        })
    return findings


# ── 单页 / 批量分析 ──────────────────────────────────────────

def analyze_one_page(file_path, *, root='', circle_mode='center_radius', include_arcs=True):
    path = Path(file_path)
    page_no = _csa_page_no(path)
    page_label = _csa_page_label(page_no)
    csa_root = Path(root).expanduser() if root else path.parent
    try:
        relative_file = str(path.relative_to(csa_root))
    except ValueError:
        relative_file = path.name

    result = {
        'page_no': page_no, 'page_label': page_label, 'page_name': page_label,
        'file': str(path), 'relative_file': relative_file,
        'cross_count': 0, 'circle_count': 0,
        'wire_count': 0, 'dot_count': 0,
        'findings': [], 'circles': [], 'error': '',
    }
    try:
        for enc in ('utf-8-sig', 'utf-16', 'gb18030', 'latin-1'):
            try:
                text = path.read_bytes().decode(enc)
                break
            except Exception:
                continue
        else:
            text = path.read_bytes().decode('utf-8', errors='replace')

        wires, dots, circles, page_name = parse_csa_text(
            text, page_no, circle_mode=circle_mode, include_arcs=include_arcs)
        findings = find_dot_four_way_crosses(wires, dots)

        result.update({
            'page_name': page_name,
            'wire_count': len(wires), 'dot_count': len(dots),
            'cross_count': len(findings), 'circle_count': len(circles),
            'findings': findings, 'circles': circles,
        })
    except Exception as exc:
        result['error'] = str(exc)
    return result


def _analyze_csa_dir(project_dir, circle_mode='center_radius', include_arcs=True):
    root = Path(project_dir)
    csa_dir = root if root.name.lower() == 'sch_1' else root / 'sch_1'
    if not csa_dir.is_dir():
        raise FileNotFoundError(f"目录不存在：{csa_dir}")
    files = sorted(
        [f for f in csa_dir.glob('*.csa') if PAGE_FILE_RE.match(f.name)],
        key=lambda f: (_csa_page_no(f), str(f).lower()))
    if not files:
        raise FileNotFoundError(f"在 {csa_dir} 中未找到 page*.csa 文件")

    results = []
    for fp in files:
        r = analyze_one_page(
            str(fp), root=str(csa_dir),
            circle_mode=circle_mode, include_arcs=include_arcs)
        results.append(r)

    total_crosses = sum(r['cross_count'] for r in results)
    total_circles = sum(r['circle_count'] for r in results)
    total_errors = sum(1 for r in results if r['error'])
    return results, total_crosses, total_circles, total_errors


# ── CSV 导出 ─────────────────────────────────────────────────

def _generate_csv(results):
    buf = io.StringIO()
    writer = csv.writer(buf)

    writer.writerow(['=== CSA 几何规范检查 概要 ==='])
    writer.writerow(['页面', 'CSA页名', '文件', 'DOT四向十字数', '画圈对象数', 'WIRE数', 'DOT数', '错误'])
    for r in results:
        writer.writerow([r['page_label'], r['page_name'], r['relative_file'],
                         r['cross_count'], r['circle_count'],
                         r['wire_count'], r['dot_count'], r['error']])

    writer.writerow([])
    writer.writerow(['=== DOT 四向十字详细 ==='])
    writer.writerow(['页面', '序号', '坐标', 'X', 'Y', 'DOT行号',
                     '水平WIRE行号', '垂直WIRE行号', '全部WIRE行号', '关联信号', '说明'])
    for r in results:
        for idx, item in enumerate(r['findings'], start=1):
            writer.writerow([r['page_label'], idx,
                             f'({item["x"]},{item["y"]})', item['x'], item['y'], item['dot_line'],
                             item['h_wire_lines'], item['v_wire_lines'],
                             item['all_wire_lines'], item['labels'], item['detail']])

    writer.writerow([])
    writer.writerow(['=== 画圈对象详细 ==='])
    writer.writerow(['页面', '序号', '对象类型', '行号', '圆心', '半径', '直径',
                     '外接框', '宽', '高', '解析说明'])
    for r in results:
        for idx, item in enumerate(r['circles'], start=1):
            writer.writerow([r['page_label'], idx, item['object_type'], item['line_no'],
                             f'({_fmt_num(item["center_x"])},{_fmt_num(item["center_y"])})',
                             _fmt_num(item['radius']), _fmt_num(item['diameter']),
                             f'({_fmt_num(item["bbox_xmin"])},{_fmt_num(item["bbox_ymin"])})-'
                             f'({_fmt_num(item["bbox_xmax"])},{_fmt_num(item["bbox_ymax"])})',
                             _fmt_num(item['width']), _fmt_num(item['height']),
                             item['parse_note']])

    return buf.getvalue()


# ── 路由 ─────────────────────────────────────────────────────

@csa_bp.route('/csa', methods=['GET', 'POST'])
def tool_csa():
    if request.method == 'POST':
        files_upload = request.files.getlist('files')
        circle_mode = request.form.get('circle_mode', 'center_radius')
        include_arcs = request.form.get('include_arcs', '1') == '1'

        if not files_upload or not any(f.filename for f in files_upload):
            return jsonify({'success': False, 'error': '请上传 page*.csa 文件'})

        uid = str(uuid.uuid4())[:8]
        proj_dir = os.path.join(UPLOAD_DIR, f'csa_proj_{uid}')
        os.makedirs(os.path.join(proj_dir, 'sch_1'), exist_ok=True)

        for f in files_upload:
            if f.filename:
                f.save(os.path.join(proj_dir, 'sch_1', os.path.basename(f.filename)))

        try:
            results, total_crosses, total_circles, total_errors = _analyze_csa_dir(
                proj_dir, circle_mode=circle_mode, include_arcs=include_arcs)
            return jsonify({
                'success': True,
                'results': results,
                'total_crosses': total_crosses,
                'total_circles': total_circles,
                'total_errors': total_errors,
                'pages': len(results),
                'csv_key': uid,
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    return render_template('csa.html', tables=FEISHU_PRESET_TABLES)


@csa_bp.route('/csa_export_csv/<csv_key>')
def csa_export_csv(csv_key):
    proj_dir = os.path.join(UPLOAD_DIR, f'csa_proj_{csv_key}')
    try:
        results, _, _, _ = _analyze_csa_dir(proj_dir)
        csv_content = _generate_csv(results)
        return Response(
            csv_content,
            mimetype='text/csv; charset=utf-8-sig',
            headers={'Content-Disposition': 'attachment; filename=csa_分析结果.csv'},
        )
    except Exception:
        return jsonify({'error': 'CSV 导出失败或数据已过期'}), 404
