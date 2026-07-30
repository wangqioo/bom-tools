# -*- coding: utf-8 -*-
"""BOM 比对工具 — Blueprint"""

import os
import uuid
import json
import re
from decimal import Decimal, InvalidOperation

from flask import Blueprint

from activity import track_tool_activity
from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str,
    _open_workbook, _save_uploaded_excel, _to_int,
)
from manufacturer_alias import lookup_manufacturer
from .customer_hq import preview as customer_hq_preview
from .customer_hq_export import build_report as customer_hq_build_report
from .export_info import write_export_info


bom_compare_bp = Blueprint('bom_compare', __name__)

HQ_STANDARD_HEADER_ROW = 3
HQ_STANDARD_HEADERS = ['\u5e8f\u53f7', '\u6599\u53f7', '\u578b\u53f7', '\u7269\u6599\u63cf\u8ff0', '\u5355\u8017', '\u66ff\u4ee3\u5173\u7cfb', '\u4f4d\u53f7', '\u751f\u4ea7\u5382\u5bb6']
PLM_FULL_HEADER_ROW = 8
PLM_FULL_SHEETS = ['BOM', 'DBG\u4e1a\u52a1BOM', 'DBGBOM\u5236\u63a7\u4fe1\u606f']
PLM_FULL_MERGE_SHEETS = ['BOM', 'DBG\u4e1a\u52a1BOM']
CADENCE_STANDARD_HEADER_ROW = 3
CADENCE_REQUIRED_HEADERS = ['\u5e8f\u53f7', '\u6599\u53f7', '\u578b\u53f7', '\u5355\u8017', '\u4f4d\u53f7']
HQ_FORMAT_ERROR = '\u4e0d\u652f\u6301\u5f53\u524d\u6587\u4ef6\u683c\u5f0f\u3002\u8bf7\u4e0a\u4f20\u7cfb\u7edf\u5bfc\u51fa\u7684\u6807\u51c6 HQ BOM\uff0c\u6216 PLM \u5168\u91cf BOM\uff1a\u4e24\u7c7b\u683c\u5f0f\u90fd\u5fc5\u987b\u5305\u542b\u5e8f\u53f7\u3001\u6599\u53f7\u3001\u578b\u53f7\u3001\u7269\u6599\u63cf\u8ff0\u3001\u5355\u8017\u3001\u66ff\u4ee3\u5173\u7cfb\u3001\u4f4d\u53f7\u3001\u751f\u4ea7\u5382\u5bb6\u7b49\u5217\u3002'

GENERIC_COMPARE_TYPES = {
    'cadence_hq': {
        'title': 'Cadence BOM \u5bf9\u6bd4 HQ BOM \u5dee\u5f02\u603b\u89c8',
        'left_label': 'Cadence BOM',
        'right_label': 'HQ BOM',
        'filename': 'Cadence_BOM\u5bf9\u6bd4HQ_BOM',
    },
}

def _headers(ws, header_row):
    headers = []
    for ci in range(1, ws.max_column + 1):
        value = _cell_str(ws.cell(row=header_row, column=ci).value)
        headers.append(value or f"未命名列{get_column_letter(ci)}")
    return headers



def _save_uploaded_hq_excel(file, prefix, uid):
    return _save_uploaded_excel(file, prefix, uid)


def _missing_required_headers(headers, required):
    normalized = [(_normalize_header(h), h) for h in headers]
    missing = []
    for req in required:
        req_norm = _normalize_header(req)
        if not any(req_norm == norm or req_norm in norm for norm, _ in normalized):
            missing.append(req)
    return missing


def _validate_plm_full_ws(ws):
    if ws.max_row < PLM_FULL_HEADER_ROW:
        raise ValueError(HQ_FORMAT_ERROR)
    headers = _headers(ws, PLM_FULL_HEADER_ROW)
    missing = _missing_required_headers(headers, HQ_STANDARD_HEADERS)
    if missing:
        raise ValueError(HQ_FORMAT_ERROR)
    return headers


def _detect_hq_format_ws(ws):
    try:
        headers = _validate_hq_standard_ws(ws)
        return {'kind': 'standard', 'header_row': HQ_STANDARD_HEADER_ROW, 'headers': headers}
    except ValueError:
        pass
    try:
        headers = _validate_plm_full_ws(ws)
        return {'kind': 'plm_full', 'header_row': PLM_FULL_HEADER_ROW, 'headers': headers}
    except ValueError:
        pass
    raise ValueError(HQ_FORMAT_ERROR)


def _open_hq_workbook_info(path, sheet_name=''):
    wb = _open_workbook(path, data_only=True)
    try:
        selected = _pick_sheet(wb, sheet_name)
        ws = wb[selected]
        fmt = _detect_hq_format_ws(ws)
        return wb, selected, ws, fmt
    except Exception:
        wb.close()
        raise


def _normalize_meta_key(value):
    return _normalize_header(value).replace(':', '').replace('：', '')


def _validate_hq_standard_ws(ws):
    if ws.max_row < HQ_STANDARD_HEADER_ROW:
        raise ValueError(HQ_FORMAT_ERROR)
    meta_keys = {
        _normalize_meta_key(ws.cell(row=ri, column=ci).value)
        for ri in (1, 2)
        for ci in range(1, min(ws.max_column, 8) + 1, 2)
    }
    required_meta = {_normalize_meta_key(v) for v in ('料号', '描述', '项目配置名', '版本', 'BOM名称')}
    if not required_meta.issubset(meta_keys):
        raise ValueError(HQ_FORMAT_ERROR)
    headers = _headers(ws, HQ_STANDARD_HEADER_ROW)
    missing = _missing_required_headers(headers, HQ_STANDARD_HEADERS)
    if missing:
        raise ValueError(HQ_FORMAT_ERROR)
    return headers


def _detect_cadence_format_ws(ws):
    if ws.max_row < CADENCE_STANDARD_HEADER_ROW:
        raise ValueError('不支持当前 Cadence BOM 格式')
    headers = _headers(ws, CADENCE_STANDARD_HEADER_ROW)
    missing = _missing_required_headers(headers, CADENCE_REQUIRED_HEADERS)
    if missing:
        raise ValueError('不支持当前 Cadence BOM 格式')
    return {'kind': 'cadence_standard', 'header_row': CADENCE_STANDARD_HEADER_ROW, 'headers': headers}


def _pick_sheet(wb, sheet_name):
    if sheet_name and sheet_name in wb.sheetnames:
        return sheet_name
    return wb.sheetnames[0] if wb.sheetnames else ''


def _detect_key(headers):
    normalized = [(h or '').lower().replace(' ', '').replace('_', '') for h in headers]
    candidates = ('料号', 'hq料号', 'hqpn', '物料编码', 'partnumber', 'pn')
    for cand in candidates:
        for i, h in enumerate(normalized):
            if cand in h:
                return headers[i]
    return headers[0] if headers else ''





def _safe_sheet_title(value):
    text = str(value or '')[:31] or 'Sheet'
    for ch in r'\/*?:[]':
        text = text.replace(ch, '_')
    return text[:31] or 'Sheet'

def _quote_sheet_name(name):
    return "'" + str(name).replace("'", "''") + "'"


def _set_internal_hyperlink(cell, sheet, target='A1'):
    cell.hyperlink = f"#{_quote_sheet_name(sheet.title)}!{target}"
    cell.style = 'Hyperlink'

def _normalize_header(value):
    return ''.join(str(value or '').lower().split()).replace('_', '').replace('-', '')


def _detect_common_key(left_headers, right_headers, prefer_part_no=False, ignored_left_headers=None):
    ignored_left_headers = set(ignored_left_headers or [])
    left_norm = {_normalize_header(h): h for h in left_headers if h}
    right_norm = {_normalize_header(h): h for h in right_headers if h}
    part_no_pairs = (
        ('\u6599\u53f7', '\u6599\u53f7'), ('hq\u6599\u53f7', '\u6599\u53f7'),
        ('partnumber', '\u6599\u53f7'), ('partno', '\u6599\u53f7'), ('pn', '\u6599\u53f7'),
        ('\u5ba2\u6237\u6599\u53f7', '\u6599\u53f7'), ('\u5ba2\u6237\u6599\u53f7', 'hq\u6599\u53f7'),
    )
    refdes_pairs = (
        ('\u4f4d\u53f7', '\u4f4d\u53f7'), ('reference', '\u4f4d\u53f7'), ('refdes', '\u4f4d\u53f7'),
    )
    other_pairs = (
        ('\u5ba2\u6237\u578b\u53f7', '\u578b\u53f7'),
        ('\u578b\u53f7', '\u578b\u53f7'), ('\u89c4\u683c\u578b\u53f7', '\u578b\u53f7'),
    )
    pairs = part_no_pairs + refdes_pairs + other_pairs if prefer_part_no else refdes_pairs + part_no_pairs + other_pairs
    for left_key, right_key in pairs:
        left = left_norm.get(_normalize_header(left_key))
        right = right_norm.get(_normalize_header(right_key))
        if left and right:
            return left, right
    common = [h for h in left_headers if h and h not in ignored_left_headers and h in right_headers]
    if common:
        key = _detect_key(common)
        return key, key
    return (left_headers[0] if left_headers else ''), (right_headers[0] if right_headers else '')

def _read_generic_headers(path, sheet_name, header_row):
    wb = _open_workbook(path, data_only=True)
    sheet_name = _pick_sheet(wb, sheet_name)
    ws = wb[sheet_name]
    headers = _headers(ws, header_row)
    wb.close()
    return sheet_name, headers


def _preview_rows(path, sheet_name, max_rows=12, max_cols=20):
    wb = _open_workbook(path, data_only=True)
    try:
        sheet_name = _pick_sheet(wb, sheet_name)
        ws = wb[sheet_name]
        rows = []
        row_limit = min(ws.max_row, max_rows)
        col_limit = min(ws.max_column, max_cols)
        for ri in range(1, row_limit + 1):
            rows.append({
                'row_number': ri,
                'values': [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, col_limit + 1)],
            })
        return {
            'sheets': wb.sheetnames,
            'current_sheet': sheet_name,
            'rows': rows,
            'max_row': ws.max_row,
            'max_column': ws.max_column,
            'shown_columns': col_limit,
        }
    finally:
        wb.close()


def _row_is_import_warning(ws, row_idx):
    return _cell_str(ws.cell(row=row_idx, column=1).value).startswith('正式导入前删除')


def _non_empty_headers(ws, header_row, headers, required_headers=None):
    required = set(required_headers or [])
    has_value = {h: False for h in headers if h}
    for ri in range(header_row + 1, ws.max_row + 1):
        row_values = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
        if not any(row_values):
            continue
        if _row_is_import_warning(ws, ri):
            break
        for ci, header in enumerate(headers, 1):
            if header and _cell_str(ws.cell(row=ri, column=ci).value):
                has_value[header] = True
    return [h for h in headers if h in required or has_value.get(h)]


def _comparable_headers(ws, fmt):
    return _non_empty_headers(ws, fmt['header_row'], fmt['headers'])
def _read_cadence_side_headers(path, sheet_name='', fallback_header_row=1):
    wb = _open_workbook(path, data_only=True)
    try:
        sheet_name = _pick_sheet(wb, sheet_name)
        ws = wb[sheet_name]
        try:
            fmt = _detect_cadence_format_ws(ws)
            all_headers = fmt['headers']
            fmt['headers'] = _non_empty_headers(ws, fmt['header_row'], all_headers, CADENCE_REQUIRED_HEADERS)
            fmt['all_headers'] = all_headers
        except ValueError:
            fmt = {'kind': 'generic', 'header_row': fallback_header_row, 'headers': _headers(ws, fallback_header_row)}
        sheets = wb.sheetnames
        return sheet_name, fmt['headers'], fmt, sheets
    finally:
        wb.close()

def _read_hq_side_headers(path, sheet_name=''):
    wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, sheet_name)
    sheets = wb.sheetnames
    headers = fmt['headers']
    bom_sheets = _plm_full_target_sheets(wb, ['DBG业务BOM']) if fmt['kind'] == 'plm_full' else [sheet_name]
    wb.close()
    return sheet_name, headers, fmt, sheets, bom_sheets


def _valid_plm_full_sheets(wb):
    result = []
    for name in wb.sheetnames:
        try:
            _validate_plm_full_ws(wb[name])
            result.append(name)
        except ValueError:
            pass
    return result

def _plm_full_target_sheets(wb, _preferred=None):
    valid = set(_valid_plm_full_sheets(wb))
    return [name for name in PLM_FULL_MERGE_SHEETS if name in valid]

def _load_hq_side_rows(path, sheet_name, key_col, compare_cols, expand_refdes=False, key_cols=None, key_transforms=None):
    rows, duplicates, headers = _load_rows(
        path, sheet_name, None, key_col, compare_cols,
        expand_refdes=expand_refdes, key_cols=key_cols, key_transforms=key_transforms)
    return rows, duplicates, headers, 0


def _is_refdes_header(value):
    norm = _normalize_header(value)
    return '位号' in norm or 'refdes' in norm or 'reference' in norm


def _is_qty_header(value):
    norm = _normalize_header(value)
    return norm in {'单耗', '数量', 'qty', 'quantity'}

_NUMERIC_TEXT_RE = re.compile(r'^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?$|^[+-]?0?\.\d+$')


def _numeric_decimal(value):
    text = _cell_str(value)
    if not text or not _NUMERIC_TEXT_RE.match(text):
        return None
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def _field_value_equal(left_value, right_value):
    left_num = _numeric_decimal(left_value)
    right_num = _numeric_decimal(right_value)
    if left_num is not None and right_num is not None:
        return left_num == right_num
    return _cell_str(left_value) == _cell_str(right_value)

def _map_compare_key_value(value, transform=''):
    text = _cell_str(value)
    if transform == 'manufacturer_alias' and text:
        match = lookup_manufacturer(text)
        if match:
            return _cell_str(match.get('canonical_name'))
    return text


def _normalize_key_config(cols, fallback_col='', transforms=None):
    key_cols = [str(c or '').strip() for c in (cols or []) if str(c or '').strip()]
    if not key_cols and fallback_col:
        key_cols = [str(fallback_col).strip()]
    key_transforms = list(transforms or [])
    if len(key_transforms) < len(key_cols):
        key_transforms += [''] * (len(key_cols) - len(key_transforms))
    elif len(key_transforms) > len(key_cols):
        key_transforms = key_transforms[:len(key_cols)]
    return key_cols, key_transforms


def _expand_compare_values(values, key_col, key):
    expanded = dict(values)
    if key_col in expanded:
        expanded[key_col] = key
    for col in list(expanded):
        if _is_qty_header(col):
            expanded[col] = '1'
    return expanded


def _load_generic_rows(path, sheet_name, header_row, key_col, compare_cols, expand_refdes=False, key_cols=None, key_transforms=None):
    wb = _open_workbook(path, data_only=True)
    sheet_name = _pick_sheet(wb, sheet_name)
    ws = wb[sheet_name]
    headers = _headers(ws, header_row)
    key_cols, key_transforms = _normalize_key_config(key_cols, key_col, key_transforms)
    missing_key_cols = [col for col in key_cols if col not in headers]
    if not key_cols or missing_key_cols:
        wb.close()
        bad_col = missing_key_cols[0] if missing_key_cols else key_col
        raise ValueError(f'\u5339\u914d\u952e\u5217 "{bad_col}" \u4e0d\u5b58\u5728')

    key_indices = [(col, headers.index(col) + 1, key_transforms[i] if i < len(key_transforms) else '') for i, col in enumerate(key_cols)]
    compare_indices = [(col, headers.index(col) + 1) for col in compare_cols if col in headers]
    rows = {}
    duplicates = {}
    skipped_blank_key = 0
    for ri in range(header_row + 1, ws.max_row + 1):
        row_values = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
        if not any(row_values):
            continue
        if _row_is_import_warning(ws, ri):
            break
        key_parts = [
            _map_compare_key_value(ws.cell(row=ri, column=idx).value, transform)
            for _, idx, transform in key_indices
        ]
        if not any(key_parts):
            skipped_blank_key += 1
            continue
        keys = _split_refdes(key_parts[0]) if expand_refdes else ['||'.join(key_parts)]
        values = {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices}
        for item_key in keys:
            if not item_key:
                continue
            item = {
                'row': ri,
                'values': _expand_compare_values(values, key_cols[0], item_key) if expand_refdes else values,
                'raw': row_values,
                'headers': headers,
            }
            if item_key in rows:
                duplicates.setdefault(item_key, [rows[item_key]['row']]).append(ri)
                continue
            rows[item_key] = item
    wb.close()
    return rows, duplicates, headers, skipped_blank_key

def _field_pairs(config_pairs, left_headers, right_headers, ignored_left_headers=None):
    ignored_left_headers = set(ignored_left_headers or [])
    pairs = []
    seen = set()
    for pair in config_pairs or []:
        left = str((pair or {}).get('left') or '').strip()
        right = str((pair or {}).get('right') or '').strip()
        if not left or not right or left in ignored_left_headers or left not in left_headers or right not in right_headers:
            continue
        key = (left, right)
        if key not in seen:
            pairs.append(key)
            seen.add(key)
    if pairs:
        return pairs
    common = [h for h in left_headers if h and h in right_headers]
    return [(col, col) for col in common]


def _generic_field_equal(left_col, right_col, left_value, right_value):
    if _is_refdes_header(left_col) or _is_refdes_header(right_col):
        return set(_split_refdes(left_value)) == set(_split_refdes(right_value))
    return _field_value_equal(left_value, right_value)



def _generic_field_values_for_report(left_col, right_col, left_value, right_value):
    if not (_is_refdes_header(left_col) or _is_refdes_header(right_col)):
        return left_value, right_value
    left_refs = [ref for ref in _split_refdes(left_value) if ref]
    right_refs = [ref for ref in _split_refdes(right_value) if ref]
    right_set = set(right_refs)
    left_set = set(left_refs)
    return (
        '、'.join(ref for ref in left_refs if ref not in right_set),
        '、'.join(ref for ref in right_refs if ref not in left_set),
    )

def _pair_label(left_col, right_col):
    return left_col if left_col == right_col else f'{left_col} <-> {right_col}'


def _generic_diff_items(left_rows, right_rows, field_pairs, labels=None):
    labels = labels or {'left_label': '左侧', 'right_label': '右侧'}
    left_only_type = f"仅{labels['left_label']}存在"
    right_only_type = f"仅{labels['right_label']}存在"
    items = []
    all_keys = sorted(set(left_rows) | set(right_rows))
    for key in all_keys:
        left = left_rows.get(key)
        right = right_rows.get(key)
        changed_fields = []
        if left and not right:
            diff_type = left_only_type
        elif right and not left:
            diff_type = right_only_type
        else:
            for left_col, right_col in field_pairs:
                if not _generic_field_equal(left_col, right_col, left['values'].get(left_col, ''), right['values'].get(right_col, '')):
                    changed_fields.append(f'{left_col} <-> {right_col}' if left_col != right_col else left_col)
            diff_type = '\u5b57\u6bb5\u53d8\u66f4' if changed_fields else '\u4e00\u81f4'
        items.append({
            'key': key,
            'type': diff_type,
            'left': left,
            'right': right,
            'changed_fields': changed_fields,
        })
    return items


def _write_generic_compare_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, labels, left_headers=None, right_headers=None, meta=None):
    wb = Workbook()
    ws = wb.active
    ws.title = '\u5dee\u5f02\u603b\u89c8'

    left_only_type = f"仅{labels['left_label']}存在"
    right_only_type = f"仅{labels['right_label']}存在"
    fills = {
        left_only_type: PatternFill('solid', fgColor='FFEBEE'),
        right_only_type: PatternFill('solid', fgColor='E8F5E9'),
        '\u5b57\u6bb5\u53d8\u66f4': PatternFill('solid', fgColor='FFF9C4'),
        '\u4e00\u81f4': PatternFill('solid', fgColor='F5F5F5'),
    }
    bdr = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    header_fill = PatternFill('solid', fgColor='D9EAF7')
    title_fill = PatternFill('solid', fgColor='1F4E78')
    title_font = Font(bold=True, color='FFFFFF', size=14)

    meta = meta or {}
    tool_name = labels.get('tool_name') or labels.get('title', '').replace(' \u5dee\u5f02\u603b\u89c8', '')
    tool_version_key = labels.get('tool_version_key') or 'cadence-hq-compare'
    summary_start_row = write_export_info(
        ws,
        labels['title'],
        tool_name,
        tool_version_key,
        rows=[
            (f"{labels['left_label']} \u6587\u4ef6", meta.get('left_filename', '')),
            (f"{labels['right_label']} \u6587\u4ef6", meta.get('right_filename', '')),
            ("\u5339\u914d\u952e", labels.get('key_label', '')),
            ("\u6bd4\u5bf9\u5b57\u6bb5", '\uff1b'.join(_pair_label(l, r) for l, r in field_pairs)),
        ],
        note="\u672c\u62a5\u544a\u7531 BOM Tools \u81ea\u52a8\u751f\u6210\uff0c\u7ed3\u679c\u4f9d\u8d56\u4e0a\u4f20\u6587\u4ef6\u5185\u5bb9\u3001\u5339\u914d\u952e\u548c\u6bd4\u5bf9\u5b57\u6bb5\u3002",
        title_fill=title_fill,
        title_font=title_font,
        header_fill=header_fill,
        border=bdr,
        value_alignment=left_align,
    )

    summary_rows = [
        (f"{labels['left_label']} \u552f\u4e00\u9879", stats['left_total']),
        (f"{labels['right_label']} \u552f\u4e00\u9879", stats['right_total']),
        (f"\u4ec5 {labels['left_label']} \u5b58\u5728", stats['left_only']),
        (f"\u4ec5 {labels['right_label']} \u5b58\u5728", stats['right_only']),
        ('\u5b57\u6bb5\u53d8\u66f4', stats['changed']),
        ('\u4e00\u81f4', stats['same']),
        (f"{labels['left_label']} \u91cd\u590d\u952e", stats['left_duplicates']),
        (f"{labels['right_label']} \u91cd\u590d\u952e", stats['right_duplicates']),
        (f"{labels['left_label']} \u7a7a\u952e\u8df3\u8fc7", stats['left_blank_keys']),
        (f"{labels['right_label']} \u7a7a\u952e\u8df3\u8fc7", stats['right_blank_keys']),
    ]
    summary_row_by_name = {}
    for ri, (name, value) in enumerate(summary_rows, summary_start_row):
        summary_row_by_name[name] = ri
        ws.cell(row=ri, column=1, value=name).font = Font(bold=True)
        ws.cell(row=ri, column=2, value=value)
        ws.cell(row=ri, column=1).border = bdr
        ws.cell(row=ri, column=2).border = bdr

    items = _generic_diff_items(left_rows, right_rows, field_pairs, labels)
    detail_headers = ['差异类型', '匹配键', f"{labels['left_label']}行号", f"{labels['right_label']}行号", '差异字段', f"{labels['left_label']}值", f"{labels['right_label']}值"]

    def field_label(left_col, right_col):
        return f'{left_col} <-> {right_col}' if left_col != right_col else left_col

    def item_field_rows(item):
        if item['type'] == '一致':
            return []
        left_item = item['left']
        right_item = item['right']
        rows = []
        for left_col, right_col in field_pairs:
            left_value = left_item['values'].get(left_col, '') if left_item else ''
            right_value = right_item['values'].get(right_col, '') if right_item else ''
            if item['type'] == '字段变更' and _generic_field_equal(left_col, right_col, left_value, right_value):
                continue
            left_value, right_value = _generic_field_values_for_report(left_col, right_col, left_value, right_value)
            rows.append([
                item['type'],
                item['key'],
                left_item['row'] if left_item else '',
                right_item['row'] if right_item else '',
                field_label(left_col, right_col),
                left_value,
                right_value,
            ])
        if not rows:
            rows.append([item['type'], item['key'], left_item['row'] if left_item else '', right_item['row'] if right_item else '', '', '', ''])
        return rows

    def write_table(sheet, table_items):
        change_group_fills = [PatternFill('solid', fgColor='FFF9C4'), PatternFill('solid', fgColor='EAF2F8')]
        group_fill_by_key = {}
        for item in table_items:
            key = item.get('key')
            if item.get('type') == '字段变更' and key not in group_fill_by_key:
                group_fill_by_key[key] = change_group_fills[len(group_fill_by_key) % len(change_group_fills)]
        for ci, header in enumerate(detail_headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        ri = 2
        for item in table_items:
            row_fill = group_fill_by_key.get(item.get('key'), fills[item['type']]) if item['type'] == '字段变更' else fills[item['type']]
            for row_values in item_field_rows(item):
                for ci, value in enumerate(row_values, 1):
                    c = sheet.cell(row=ri, column=ci, value=value)
                    c.alignment = center if ci in (1, 3, 4) else left_align
                    c.border = bdr
                    c.fill = row_fill
                ri += 1
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    def write_original_rows(sheet, table_items, side, headers):
        headers = list(headers or [])
        for ci, header in enumerate(headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        fill = fills[left_only_type if side == 'left' else right_only_type]
        for ri, item in enumerate(table_items, 2):
            source = item['left'] if side == 'left' else item['right']
            raw = list((source or {}).get('raw') or [])
            if len(raw) < len(headers):
                raw.extend([''] * (len(headers) - len(raw)))
            for ci, value in enumerate(raw[:len(headers)], 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = left_align
                c.border = bdr
                c.fill = fill
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    def link_summary_row(name, sheet):
        ri = summary_row_by_name.get(name)
        if not ri or not sheet:
            return
        cell = ws.cell(row=ri, column=2)
        _set_internal_hyperlink(cell, sheet)

    diff_items = [i for i in items if i['type'] != '\u4e00\u81f4']
    detail_sheet = None
    if labels.get('filename') != 'Cadence_BOM\u5bf9\u6bd4HQ_BOM':
        detail_sheet = wb.create_sheet('\u5dee\u5f02\u660e\u7ec6')
        write_table(detail_sheet, diff_items)
    left_only_sheet = wb.create_sheet(_safe_sheet_title(left_only_type))
    write_original_rows(left_only_sheet, [i for i in items if i['type'] == left_only_type], 'left', left_headers)
    right_only_sheet = wb.create_sheet(_safe_sheet_title(right_only_type))
    write_original_rows(right_only_sheet, [i for i in items if i['type'] == right_only_type], 'right', right_headers)
    changed_sheet = wb.create_sheet('\u5b57\u6bb5\u53d8\u66f4')
    write_table(changed_sheet, [i for i in items if i['type'] == '\u5b57\u6bb5\u53d8\u66f4'])

    ws_dup = wb.create_sheet('\u91cd\u590d\u952e')
    ws_dup.append(['\u7c7b\u578b', '\u5339\u914d\u952e\u548c\u884c\u53f7'])
    for cell in ws_dup[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = bdr
    for note in duplicate_notes:
        ws_dup.append(note)
    for row in ws_dup.iter_rows():
        for cell in row:
            cell.border = bdr
            cell.alignment = left_align

    link_summary_row(f"\u4ec5 {labels['left_label']} \u5b58\u5728", left_only_sheet)
    link_summary_row(f"\u4ec5 {labels['right_label']} \u5b58\u5728", right_only_sheet)
    link_summary_row('\u5b57\u6bb5\u53d8\u66f4', changed_sheet)
    link_summary_row(f"{labels['left_label']} \u91cd\u590d\u952e", ws_dup)
    link_summary_row(f"{labels['right_label']} \u91cd\u590d\u952e", ws_dup)
    if detail_sheet:
        ws['D2'] = '\u67e5\u770b\u5dee\u5f02\u660e\u7ec6'
        _set_internal_hyperlink(ws['D2'], detail_sheet)

    for sheet in wb.worksheets:
        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 18 if col not in (2, 5) else 30
    wb.save(out_path)

def _split_refdes(value):
    text = _cell_str(value)
    if not text:
        return ['']
    text = re.sub(r'[,，;；、\s]+', ',', text)
    refs = []
    seen = set()
    for part in text.split(','):
        ref = part.strip()
        if not ref or ref in seen:
            continue
        refs.append(ref)
        seen.add(ref)
    return refs or ['']




def _format_match_key(key):
    text = str(key)
    if '||' not in text:
        return text
    part_no, refdes = text.split('||', 1)
    return f'{part_no} / 位号 {refdes}' if refdes else part_no


def _is_plm_history_header(row_values):
    return row_values[:5] == ['\u0042\u004f\u004d\u7248\u672c', '\u65e5\u671f', '\u4fee\u8ba2\u88c5\u914d\u4ef6', '\u4fee\u8ba2\u7ec4\u4ef6', '\u7ec4\u4ef6\u578b\u53f7']


def _load_rows(path, sheet_name, header_row, key_col, compare_cols, expand_refdes=False, key_cols=None, key_transforms=None):
    wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, sheet_name)
    headers = fmt['headers']
    header_row = fmt['header_row']
    key_cols, key_transforms = _normalize_key_config(key_cols, key_col, key_transforms)
    missing_key_cols = [col for col in key_cols if col not in headers]
    if not key_cols or missing_key_cols:
        wb.close()
        bad_col = missing_key_cols[0] if missing_key_cols else key_col
        raise ValueError(f'\u5339\u914d\u952e\u5217 "{bad_col}" \u4e0d\u5b58\u5728')
    key_indices = [(col, headers.index(col) + 1, key_transforms[i] if i < len(key_transforms) else '') for i, col in enumerate(key_cols)]
    ref_idx = headers.index('\u4f4d\u53f7') + 1 if '\u4f4d\u53f7' in headers else None
    compare_indices = [(col, headers.index(col) + 1) for col in compare_cols if col in headers]

    rows = {}
    duplicates = {}
    for ri in range(header_row + 1, ws.max_row + 1):
        row_values = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, len(headers) + 1)]
        if fmt['kind'] == 'plm_full' and _is_plm_history_header(row_values):
            break
        if not any(row_values):
            if fmt['kind'] == 'plm_full' and ri < ws.max_row:
                next_values = [_cell_str(ws.cell(row=ri + 1, column=ci).value) for ci in range(1, min(len(headers), 16) + 1)]
                if _is_plm_history_header(next_values):
                    break
            continue
        key_parts = [
            _map_compare_key_value(ws.cell(row=ri, column=idx).value, transform)
            for _, idx, transform in key_indices
        ]
        if not any(key_parts):
            continue
        values = {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices}
        all_values = {
            header: value
            for header, value in zip(headers, row_values)
            if header and value
        }
        refdes_list = _split_refdes(ws.cell(row=ri, column=ref_idx).value) if ref_idx else []
        if refdes_list == ['']:
            refdes_list = []
        keys = _split_refdes(key_parts[0]) if expand_refdes else ['||'.join(key_parts)]
        for item_key in keys:
            if not item_key:
                continue
            item = {
                'key': item_key,
                'row': ri,
                'values': _expand_compare_values(values, key_cols[0], item_key) if expand_refdes else values,
                'all_values': all_values,
                'headers': headers,
                'refdes_list': [item_key] if expand_refdes else refdes_list,
                'raw': row_values,
                'sheet': sheet_name,
            }
            if item_key in rows:
                duplicates.setdefault(item_key, [rows[item_key]['row']]).append(ri)
                continue
            rows[item_key] = item
    wb.close()
    return rows, duplicates, headers



def _merge_plm_full_item(key, primary, secondary, compare_cols, headers):
    source = primary or secondary
    values = {}
    all_values = {}
    for col in compare_cols:
        primary_value = (primary.get('values') or {}).get(col, '') if primary else ''
        secondary_value = (secondary.get('values') or {}).get(col, '') if secondary else ''
        values[col] = primary_value if primary_value != '' else secondary_value
    for header in headers:
        primary_value = (primary.get('all_values') or {}).get(header, '') if primary else ''
        secondary_value = (secondary.get('all_values') or {}).get(header, '') if secondary else ''
        value = primary_value if primary_value != '' else secondary_value
        if value:
            all_values[header] = value
    row = source.get('row', '') if source else ''
    if not primary and secondary:
        row = f"{secondary.get('sheet', '')}:{secondary.get('row', '')}"
    return {
        'key': key,
        'row': row,
        'values': values,
        'all_values': all_values,
        'headers': headers,
        'refdes_list': source.get('refdes_list', []) if source else [],
        'raw': source.get('raw', []) if source else [],
        'sheet': '+'.join(PLM_FULL_MERGE_SHEETS),
    }


def _merge_plm_duplicate_rows(target, sheet_name, duplicates):
    for key, rows in duplicates.items():
        target.setdefault(key, []).extend(f'{sheet_name}:{row}' for row in rows)


def _load_plm_full_merged_rows(path, key_col, compare_cols):
    wb = _open_workbook(path, read_only=True, data_only=True)
    target_sheets = _plm_full_target_sheets(wb)
    wb.close()
    if not target_sheets:
        raise ValueError('PLM \u5168\u91cf BOM \u672a\u627e\u5230\u53ef\u63d0\u53d6\u7684 BOM \u6216 DBG\u4e1a\u52a1BOM Sheet')

    rows_by_sheet = {}
    duplicates = {}
    merged_headers = []
    seen_headers = set()
    for sheet in target_sheets:
        rows, sheet_dups, headers = _load_rows(path, sheet, PLM_FULL_HEADER_ROW, key_col, compare_cols)
        rows_by_sheet[sheet] = rows
        _merge_plm_duplicate_rows(duplicates, sheet, sheet_dups)
        for header in headers:
            if header and header not in seen_headers:
                merged_headers.append(header)
                seen_headers.add(header)

    primary_rows = rows_by_sheet.get('BOM', {})
    secondary_rows = rows_by_sheet.get('DBG\u4e1a\u52a1BOM', {})
    all_keys = set()
    for rows in rows_by_sheet.values():
        all_keys.update(rows.keys())
    merged_rows = {
        key: _merge_plm_full_item(key, primary_rows.get(key), secondary_rows.get(key), compare_cols, merged_headers)
        for key in all_keys
    }
    return merged_rows, duplicates, merged_headers, target_sheets



def _load_meta(path, sheet_name):
    wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, sheet_name)
    meta = {}
    meta_rows = (1, 2) if fmt['kind'] == 'standard' else range(1, min(ws.max_row, 7) + 1)
    for ri in meta_rows:
        for ci in range(1, ws.max_column):
            key = _cell_str(ws.cell(row=ri, column=ci).value)
            val = _cell_str(ws.cell(row=ri, column=ci + 1).value)
            if key and val and key not in meta:
                meta[key] = val
    meta['_format'] = fmt['kind']
    meta['_sheet'] = sheet_name
    wb.close()
    return meta


def _refdes_delta(old, new):
    old_refs = old.get('refdes_list', []) if old else []
    new_refs = new.get('refdes_list', []) if new else []
    new_set = set(new_refs)
    old_set = set(old_refs)
    removed = [ref for ref in old_refs if ref not in new_set]
    added = [ref for ref in new_refs if ref not in old_set]
    return removed, added


def _diff_items(old_rows, new_rows, compare_cols):
    items = []
    all_keys = sorted(set(old_rows) | set(new_rows))
    for key in all_keys:
        old = old_rows.get(key)
        new = new_rows.get(key)
        removed_refdes, added_refdes = _refdes_delta(old, new)
        changed_fields = []
        if old and not new:
            diff_type = '删除'
        elif new and not old:
            diff_type = '新增'
        else:
            for col in compare_cols:
                if col == '位号':
                    if removed_refdes or added_refdes:
                        changed_fields.append(col)
                elif not _field_value_equal(old['values'].get(col, ''), new['values'].get(col, '')):
                    changed_fields.append(col)
            if (removed_refdes or added_refdes) and '位号' not in changed_fields:
                changed_fields.append('位号')
            diff_type = '变更' if changed_fields else '未变更'
        items.append({
            'key': key,
            'removed_refdes': removed_refdes,
            'added_refdes': added_refdes,
            'type': diff_type,
            'old': old,
            'new': new,
            'changed_fields': changed_fields,
        })
    return items


def _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols):
    added = sorted(set(new_rows) - set(old_rows))
    removed = sorted(set(old_rows) - set(new_rows))
    common = sorted(set(old_rows) & set(new_rows))
    changed = []
    unchanged = []
    for key in common:
        removed_refdes, added_refdes = _refdes_delta(old_rows[key], new_rows[key])
        has_change = any(
            (removed_refdes or added_refdes) if col == '\u4f4d\u53f7'
            else not _field_value_equal(old_rows[key]['values'].get(col, ''), new_rows[key]['values'].get(col, ''))
            for col in compare_cols
        )
        (changed if has_change else unchanged).append(key)
    return {
        'old_total': len(old_rows),
        'new_total': len(new_rows),
        'added': len(added),
        'removed': len(removed),
        'changed': len(changed),
        'unchanged': len(unchanged),
        'old_duplicates': len(old_dups),
        'new_duplicates': len(new_dups),
    }


def _duplicate_notes(old_dups, new_dups, sheet_label=''):
    prefix = f'{sheet_label} ' if sheet_label else ''
    return (
        [f"{prefix}\u57fa\u51c6\u7248\u672c\u91cd\u590d\u952e {_format_match_key(key)}: \u884c {', '.join(map(str, rows))}" for key, rows in old_dups.items()] +
        [f"{prefix}\u5bf9\u6bd4\u7248\u672c\u91cd\u590d\u952e {_format_match_key(key)}: \u884c {', '.join(map(str, rows))}" for key, rows in new_dups.items()]
    )


def _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta=None, new_meta=None, report_title=None, tool_name=None, tool_version_key=None):
    wb = Workbook()
    ws = wb.active
    ws.title = '差异总览'

    fills = {
        '新增': PatternFill('solid', fgColor='E8F5E9'),
        '删除': PatternFill('solid', fgColor='FFEBEE'),
        '变更': PatternFill('solid', fgColor='FFF9C4'),
        '未变更': PatternFill('solid', fgColor='F5F5F5'),
    }
    change_group_fills = [
        PatternFill('solid', fgColor='FFF9C4'),
        PatternFill('solid', fgColor='EAF2F8'),
    ]
    bdr = Border(left=Side(style='thin'), right=Side(style='thin'),
                 top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    header_fill = PatternFill('solid', fgColor='D9EAF7')
    title_fill = PatternFill('solid', fgColor='1F4E78')
    title_font = Font(bold=True, color='FFFFFF', size=14)
    old_meta = old_meta or {}
    new_meta = new_meta or {}
    items = _diff_items(old_rows, new_rows, compare_cols)

    report_title = report_title or 'HQ BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8'
    tool_name = tool_name or 'HQ BOM \u7248\u672c\u5bf9\u6bd4'
    tool_version_key = tool_version_key or 'hq-version-compare'
    summary_start_row = write_export_info(
        ws,
        report_title,
        tool_name,
        tool_version_key,
        rows=[
            ('\u9879\u76ee\u914d\u7f6e\u540d', new_meta.get('\u9879\u76ee\u914d\u7f6e\u540d') or old_meta.get('\u9879\u76ee\u914d\u7f6e\u540d', '')),
            ('BOM\u540d\u79f0', new_meta.get('BOM\u540d\u79f0') or old_meta.get('BOM\u540d\u79f0', '')),
            ('\u57fa\u51c6\u7248\u672c\u53f7', old_meta.get('\u7248\u672c', '')),
            ('\u5bf9\u6bd4\u7248\u672c\u53f7', new_meta.get('\u7248\u672c', '')),
            ('\u6bd4\u5bf9\u5b57\u6bb5', '\uff1b'.join(compare_cols)),
        ],
        note='\u672c\u62a5\u544a\u7531 BOM Tools \u81ea\u52a8\u751f\u6210\uff0c\u7ed3\u679c\u4f9d\u8d56\u4e0a\u4f20\u6587\u4ef6\u5185\u5bb9\u3001\u7248\u672c\u4fe1\u606f\u548c\u6bd4\u5bf9\u5b57\u6bb5\u3002',
        title_fill=title_fill,
        title_font=title_font,
        header_fill=header_fill,
        border=bdr,
        value_alignment=left,
    )

    summary_rows = [
        ('项目配置名', new_meta.get('项目配置名') or old_meta.get('项目配置名', '')),
        ('BOM名称', new_meta.get('BOM名称') or old_meta.get('BOM名称', '')),
        ('基准版本号', old_meta.get('版本', '')),
        ('对比版本号', new_meta.get('版本', '')),
        ('基准版本唯一料号数', stats['old_total']),
        ('对比版本唯一料号数', stats['new_total']),
        ('新增', stats['added']),
        ('删除', stats['removed']),
        ('变更', stats['changed']),
        ('未变更', stats['unchanged']),
        ('基准版本重复键', stats['old_duplicates']),
        ('对比版本重复键', stats['new_duplicates']),
    ]
    summary_row_by_name = {}
    for ri, (name, value) in enumerate(summary_rows, summary_start_row):
        summary_row_by_name[name] = ri
        ws.cell(row=ri, column=1, value=name).font = Font(bold=True)
        ws.cell(row=ri, column=2, value=value)
        ws.cell(row=ri, column=1).border = bdr
        ws.cell(row=ri, column=2).border = bdr
    legend = [('绿色', '新增'), ('红色', '删除'), ('黄色', '字段变更'), ('灰色', '未变更')]
    for offset, (color, text) in enumerate(legend, 3):
        c = ws.cell(row=offset, column=4, value=f'{color} = {text}')
        c.fill = fills['新增' if color == '绿色' else '删除' if color == '红色' else '变更' if color == '黄色' else '未变更']
        c.border = bdr

    detail_headers = ['差异类型', '料号', '基准版本行号', '对比版本行号', '变更字段', '基准值', '对比值']

    def changed_values(item, field):
        old = item['old']
        new = item['new']
        if field == '位号':
            return '、'.join(item.get('removed_refdes', [])), '、'.join(item.get('added_refdes', []))
        return (
            old['values'].get(field, '') if old else '',
            new['values'].get(field, '') if new else '',
        )

    def expanded_rows(table_items):
        for item in table_items:
            fields = item['changed_fields'] if item['type'] == '变更' and item['changed_fields'] else ['']
            for field in fields:
                old_value, new_value = changed_values(item, field) if field else ('', '')
                yield item, field, old_value, new_value

    def write_table(sheet, table_items):
        group_fill_by_key = {}
        for item in table_items:
            key = item.get('key')
            if key not in group_fill_by_key:
                group_fill_by_key[key] = change_group_fills[len(group_fill_by_key) % len(change_group_fills)]
        for ci, header in enumerate(detail_headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        for ri, (item, field, old_value, new_value) in enumerate(expanded_rows(table_items), 2):
            old = item['old']
            new = item['new']
            row_values = [
                item['type'],
                item['key'],
                old['row'] if old else '',
                new['row'] if new else '',
                field,
                old_value,
                new_value,
            ]
            row_fill = group_fill_by_key.get(item.get('key'), fills[item['type']]) if item['type'] == '\u53d8\u66f4' else fills[item['type']]
            for ci, value in enumerate(row_values, 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = center if ci in (1, 3, 4) else left
                c.border = bdr
                c.fill = row_fill
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    def source_headers(table_items, side):
        result = []
        seen = set()
        for item in table_items:
            source = item.get(side) or {}
            values = source.get('all_values') or {}
            for header in source.get('headers') or values.keys():
                if header in values and header not in seen:
                    result.append(header)
                    seen.add(header)
        return result

    def write_source_table(sheet, table_items, side):
        line_header = '对比版本行号' if side == 'new' else '基准版本行号'
        source_cols = source_headers(table_items, side)
        headers = ['差异类型', '料号', line_header] + source_cols
        for ci, header in enumerate(headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        for ri, item in enumerate(table_items, 2):
            source = item.get(side) or {}
            values = source.get('all_values') or {}
            row_values = [item['type'], item['key'], source.get('row', '')] + [values.get(col, '') for col in source_cols]
            row_fill = group_fill_by_key.get(item.get('key'), fills[item['type']]) if item['type'] == '字段变更' else fills[item['type']]
            for ci, value in enumerate(row_values, 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = center if ci in (1, 3) else left
                c.border = bdr
                c.fill = row_fill
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    added_sheet = wb.create_sheet('\u65b0\u589e\u7269\u6599')
    write_source_table(added_sheet, [i for i in items if i['type'] == '\u65b0\u589e'], 'new')
    removed_sheet = wb.create_sheet('\u5220\u9664\u7269\u6599')
    write_source_table(removed_sheet, [i for i in items if i['type'] == '\u5220\u9664'], 'old')
    changed_sheet = wb.create_sheet('\u53d8\u66f4\u7269\u6599')
    write_table(changed_sheet, [i for i in items if i['type'] == '\u53d8\u66f4'])

    ws_dup = wb.create_sheet('重复料号')
    ws_dup.append(['类型', '料号和行号'])
    for cell in ws_dup[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = bdr
    for note in duplicate_notes:
        kind = '基准版本' if note.startswith('基准版本') else '对比版本'
        ws_dup.append([kind, note])
    for row in ws_dup.iter_rows():
        for cell in row:
            cell.border = bdr
            cell.alignment = left

    for summary_name, target_sheet in [
        ('新增', added_sheet),
        ('删除', removed_sheet),
        ('变更', changed_sheet),
        ('基准版本重复键', ws_dup),
        ('对比版本重复键', ws_dup),
    ]:
        row_idx = summary_row_by_name.get(summary_name)
        if row_idx:
            _set_internal_hyperlink(ws.cell(row=row_idx, column=2), target_sheet)

    for sheet in wb.worksheets:
        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 16 if col not in (2, 5, 6, 7) else 28
    wb.save(out_path)



@bom_compare_bp.route('/api/bom_compare/customer_hq_preview', methods=['POST'])
def api_customer_hq_preview():
    left_file = request.files.get('left_file')
    right_file = request.files.get('right_file')
    if not left_file or not right_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u5ba2\u6237 BOM \u548c HQ BOM \u6587\u4ef6'})
    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config \u53c2\u6570\u683c\u5f0f\u9519\u8bef'})
    left_header_row = _to_int(config.get('left_header_row', 1), 1)
    if left_header_row is None:
        return jsonify({'success': False, 'error': '\u5ba2\u6237 BOM \u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})
    mapping = config.get('mapping') if isinstance(config.get('mapping'), dict) else {}
    match_mode = str(config.get('match_mode') or 'identity')
    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_excel(left_file, 'bomcmp_customer_preview_left', uid)
        right_path = _save_uploaded_hq_excel(right_file, 'bomcmp_customer_preview_right', uid)
        payload = customer_hq_preview(
            left_path=left_path,
            right_path=right_path,
            left_sheet=config.get('left_sheet', ''),
            right_sheet=config.get('right_sheet', ''),
            left_header_row=left_header_row,
            mapping=mapping,
            match_mode=match_mode,
            helpers={
                'pick_sheet': _pick_sheet,
                'headers_fn': _headers,
                'open_hq_info': _open_hq_workbook_info,
                'is_plm_history_header': _is_plm_history_header,
                'normalize_header': _normalize_header,
            },
        )
        return jsonify({'success': True, **payload})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@bom_compare_bp.route('/api/bom_compare/customer_hq_export', methods=['POST'])
@track_tool_activity('客户BOM对比HQ BOM')
def api_customer_hq_export():
    left_file = request.files.get('left_file')
    right_file = request.files.get('right_file')
    if not left_file or not right_file:
        return jsonify({'success': False, 'error': '请上传客户 BOM 和 HQ BOM 文件'})
    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config 参数格式错误'})
    left_header_row = _to_int(config.get('left_header_row', 1), 1)
    if left_header_row is None:
        return jsonify({'success': False, 'error': '客户 BOM 表头行必须是大于等于 1 的数字'})
    mapping = config.get('mapping') if isinstance(config.get('mapping'), dict) else {}
    match_mode = str(config.get('match_mode') or 'identity')
    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_excel(left_file, 'bomcmp_customer_export_left', uid)
        right_path = _save_uploaded_hq_excel(right_file, 'bomcmp_customer_export_right', uid)
        out_name = f"Customer_BOM_vs_HQ_BOM_detail_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        stats = customer_hq_build_report(
            out_path=out_path,
            left_path=left_path,
            right_path=right_path,
            left_sheet=config.get('left_sheet', ''),
            right_sheet=config.get('right_sheet', ''),
            left_header_row=left_header_row,
            mapping=mapping,
            match_mode=match_mode,
            helpers={
                'pick_sheet': _pick_sheet,
                'headers_fn': _headers,
                'open_hq_info': _open_hq_workbook_info,
                'is_plm_history_header': _is_plm_history_header,
                'normalize_header': _normalize_header,
            },
            meta={
                'left_filename': left_file.filename or '',
                'right_filename': right_file.filename or '',
                'left_sheet': config.get('left_sheet', ''),
                'right_sheet': config.get('right_sheet', ''),
                'left_header_row': left_header_row,
                'mapping': mapping,
                'match_mode': match_mode,
            },
        )
        return jsonify({'success': True, 'download': f'/download/{out_name}', **stats})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@bom_compare_bp.route('/api/bom_compare/generic_preview', methods=['POST'])
def api_generic_preview():
    uid = str(uuid.uuid4())[:8]
    result = {'success': True}
    compare_type = str(request.form.get('compare_type') or 'cadence_hq')
    save_left_excel = _save_uploaded_hq_excel if compare_type == 'cadence_hq' else _save_uploaded_excel
    save_right_excel = _save_uploaded_hq_excel if compare_type == 'cadence_hq' else _save_uploaded_excel
    try:
        left_file = request.files.get('left_file')
        right_file = request.files.get('right_file')
        if left_file:
            left_path = save_left_excel(left_file, 'bomcmp_generic_preview_left', uid)
            result['left'] = _preview_rows(left_path, request.form.get('left_sheet', '') or request.form.get('sheet_name', ''))
        if right_file:
            right_path = save_right_excel(right_file, 'bomcmp_generic_preview_right', uid)
            result['right'] = _preview_rows(right_path, request.form.get('right_sheet', ''))
        if not left_file and not right_file:
            return jsonify({'success': False, 'error': '请上传 BOM 文件'})
        return jsonify(result)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@bom_compare_bp.route('/api/bom_compare/generic_sheets', methods=['POST'])
def api_generic_sheets():
    left_file = request.files.get('left_file') or request.files.get('file')
    right_file = request.files.get('right_file')
    if not left_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u5de6\u4fa7 BOM \u6587\u4ef6'})

    header_row = _to_int(request.form.get('header_row', 1), 1)
    left_header_row = _to_int(request.form.get('left_header_row', header_row), 1)
    right_header_row = _to_int(request.form.get('right_header_row', header_row), 1)
    compare_type = str(request.form.get('compare_type') or 'cadence_hq')
    if header_row is None or left_header_row is None or right_header_row is None:
        return jsonify({'success': False, 'error': '\u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})

    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_hq_excel(left_file, 'bomcmp_generic_left', uid) if compare_type == 'cadence_hq' else _save_uploaded_excel(left_file, 'bomcmp_generic_left', uid)
        if compare_type == 'cadence_hq':
            left_sheet, left_headers, left_fmt, left_sheets = _read_cadence_side_headers(left_path, request.form.get('left_sheet', '') or request.form.get('sheet_name', ''), left_header_row)
            ignored_left_headers = [h for h in left_fmt.get('all_headers', left_headers) if h not in left_headers]
        else:
            left_wb = _open_workbook(left_path, read_only=True, data_only=True)
            left_sheets = left_wb.sheetnames
            left_sheet = _pick_sheet(left_wb, request.form.get('left_sheet', '') or request.form.get('sheet_name', ''))
            left_wb.close()
            left_sheet, left_headers = _read_generic_headers(left_path, left_sheet, left_header_row)
            left_fmt = {'kind': 'generic', 'header_row': left_header_row}
            ignored_left_headers = []

        result = {
            'success': True,
            'left_sheets': left_sheets,
            'left_current_sheet': left_sheet,
            'left_headers': left_headers,
            'left_detected_key': _detect_key(left_headers),
            'left_format': left_fmt['kind'],
            'left_header_row': left_fmt['header_row'],
            'left_ignored_headers': ignored_left_headers,
        }
        if right_file:
            right_path = _save_uploaded_hq_excel(right_file, 'bomcmp_generic_right', uid)
            right_sheet, right_headers, right_fmt, right_sheets, right_bom_sheets = _read_hq_side_headers(right_path, request.form.get('right_sheet', ''))
            left_key, right_key = _detect_common_key(left_headers, right_headers, prefer_part_no=(compare_type == 'cadence_hq'), ignored_left_headers=ignored_left_headers)
            result.update({
                'right_sheets': right_sheets,
                'right_current_sheet': right_sheet,
                'right_headers': right_headers,
                'right_detected_key': _detect_key(right_headers),
                'detected_left_key': left_key,
                'detected_right_key': right_key,
                'right_format': right_fmt['kind'],
                'right_header_row': right_fmt['header_row'],
                'right_bom_sheets': right_bom_sheets,
            })
        return jsonify(result)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@bom_compare_bp.route('/api/bom_compare/generic', methods=['POST'])
@track_tool_activity('通用BOM比对')
def api_generic_compare():
    left_file = request.files.get('left_file')
    right_file = request.files.get('right_file')
    if not left_file or not right_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u4e24\u4efd BOM \u6587\u4ef6'})

    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config \u53c2\u6570\u683c\u5f0f\u9519\u8bef'})

    compare_type = str(config.get('compare_type') or 'cadence_hq')
    labels = GENERIC_COMPARE_TYPES.get(compare_type, GENERIC_COMPARE_TYPES['cadence_hq'])
    left_header_row = _to_int(config.get('left_header_row', config.get('header_row', 1)), 1)
    right_header_row = _to_int(config.get('right_header_row', config.get('header_row', 1)), 1)
    if left_header_row is None or right_header_row is None:
        return jsonify({'success': False, 'error': '\u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})

    left_key_col = str(config.get('left_key_col') or '').strip()
    right_key_col = str(config.get('right_key_col') or '').strip()
    left_key_cols, left_key_transforms = _normalize_key_config(
        config.get('left_key_cols'), left_key_col, config.get('left_key_transforms'))
    right_key_cols, right_key_transforms = _normalize_key_config(
        config.get('right_key_cols'), right_key_col, config.get('right_key_transforms'))
    if not left_key_cols or not right_key_cols:
        return jsonify({'success': False, 'error': '\u8bf7\u9009\u62e9\u4e24\u4efd BOM \u7684\u5339\u914d\u952e\u5217'})
    if len(left_key_cols) != len(right_key_cols):
        return jsonify({'success': False, 'error': '\u4e24\u4efd BOM \u7684\u590d\u5408\u5339\u914d\u952e\u6570\u91cf\u5fc5\u987b\u4e00\u81f4'})
    left_key_col = left_key_cols[0]
    right_key_col = right_key_cols[0]

    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_hq_excel(left_file, 'bomcmp_generic_left', uid) if compare_type == 'cadence_hq' else _save_uploaded_excel(left_file, 'bomcmp_generic_left', uid)
        right_path = _save_uploaded_hq_excel(right_file, 'bomcmp_generic_right', uid)

        if compare_type == 'cadence_hq':
            left_sheet, left_headers, left_fmt, _ = _read_cadence_side_headers(left_path, config.get('left_sheet', ''), left_header_row)
            ignored_left_headers = [h for h in left_fmt.get('all_headers', left_headers) if h not in left_headers]
        else:
            left_sheet = config.get('left_sheet', '')
            _, left_headers = _read_generic_headers(left_path, left_sheet, left_header_row)
            left_fmt = {'kind': 'generic', 'header_row': left_header_row}
            ignored_left_headers = []
        right_sheet, right_headers, right_fmt, _, _ = _read_hq_side_headers(right_path, config.get('right_sheet', ''))
        field_pairs = _field_pairs(config.get('field_pairs', []), left_headers, right_headers, ignored_left_headers)
        key_pair_set = set(zip(left_key_cols, right_key_cols))
        field_pairs = [(l, r) for l, r in field_pairs if (l, r) not in key_pair_set]
        if not field_pairs:
            return jsonify({'success': False, 'error': '\u8bf7\u81f3\u5c11\u9009\u62e9\u4e00\u7ec4\u9700\u8981\u6bd4\u5bf9\u7684\u5b57\u6bb5'})

        left_compare_cols = [l for l, _ in field_pairs]
        right_compare_cols = [r for _, r in field_pairs]
        expand_refdes = compare_type == 'cadence_hq' and _is_refdes_header(left_key_col) and _is_refdes_header(right_key_col)
        left_rows, left_dups, _, left_blank = _load_generic_rows(
            left_path, left_sheet, left_fmt['header_row'], left_key_col, left_compare_cols,
            expand_refdes=expand_refdes, key_cols=left_key_cols, key_transforms=left_key_transforms)
        right_rows, right_dups, _, right_blank = _load_hq_side_rows(
            right_path, right_sheet, right_key_col, right_compare_cols,
            expand_refdes=expand_refdes, key_cols=right_key_cols, key_transforms=right_key_transforms)

        left_only = sorted(set(left_rows) - set(right_rows))
        right_only = sorted(set(right_rows) - set(left_rows))
        common = sorted(set(left_rows) & set(right_rows))
        changed = []
        same = []
        for key in common:
            has_change = any(not _generic_field_equal(l, r, left_rows[key]['values'].get(l, ''), right_rows[key]['values'].get(r, '')) for l, r in field_pairs)
            (changed if has_change else same).append(key)

        stats = {
            'left_total': len(left_rows),
            'right_total': len(right_rows),
            'left_only': len(left_only),
            'right_only': len(right_only),
            'changed': len(changed),
            'same': len(same),
            'left_duplicates': len(left_dups),
            'right_duplicates': len(right_dups),
            'left_blank_keys': left_blank,
            'right_blank_keys': right_blank,
        }
        duplicate_notes = (
            [[labels['left_label'], f"\u91cd\u590d\u952e {key}: \u884c {', '.join(map(str, rows))}"] for key, rows in left_dups.items()] +
            [[labels['right_label'], f"\u91cd\u590d\u952e {key}: \u884c {', '.join(map(str, rows))}"] for key, rows in right_dups.items()]
        )
        out_name = f"{labels['filename']}_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        _write_generic_compare_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, labels, left_headers, right_headers, {
            'left_filename': left_file.filename,
            'right_filename': right_file.filename,
        })
        return jsonify({'success': True, 'download': f'/download/{out_name}', 'left_format': left_fmt['kind'], 'left_header_row': left_fmt['header_row'], 'right_format': right_fmt['kind'], 'right_header_row': right_fmt['header_row'], 'expanded_refdes': expand_refdes, 'left_ignored_headers': ignored_left_headers, **stats})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@bom_compare_bp.route('/api/bom_compare/local_sheets', methods=['POST'])
def api_local_sheets():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_hq_excel(file, "bomcmp_pre", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    header_row = _to_int(request.form.get('header_row', 1), 1)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})

    try:
        wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, request.form.get('sheet_name', ''))
        sheets = wb.sheetnames
        headers = _comparable_headers(ws, fmt)
        bom_sheets = _plm_full_target_sheets(wb, ['DBG业务BOM']) if fmt['kind'] == 'plm_full' else [sheet_name]
        wb.close()
        return jsonify({
            'success': True,
            'sheets': sheets,
            'current_sheet': sheet_name,
            'headers': headers,
            'detected_key': _detect_key(headers),
            'format': fmt['kind'],
            'header_row': fmt['header_row'],
            'bom_sheets': bom_sheets,
        })
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@bom_compare_bp.route('/api/bom_compare/machine_local_sheets', methods=['POST'])
def api_machine_local_sheets():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u6587\u4ef6'})
    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_hq_excel(file, "bomcmp_machine_pre", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    try:
        wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, request.form.get('sheet_name', ''))
        sheets = wb.sheetnames
        headers = _comparable_headers(ws, fmt)
        bom_sheets = _plm_full_target_sheets(wb, ['BOM']) if fmt['kind'] == 'plm_full' else [sheet_name]
        wb.close()
        return jsonify({
            'success': True,
            'sheets': sheets,
            'current_sheet': sheet_name,
            'headers': headers,
            'detected_key': _detect_key(headers),
            'format': fmt['kind'],
            'header_row': fmt['header_row'],
            'bom_sheets': bom_sheets,
        })
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@bom_compare_bp.route('/api/bom_compare/machine_hq_version', methods=['POST'])
@track_tool_activity('整机HQ BOM版本对比')
def api_machine_hq_version_compare():
    old_file = request.files.get('old_file')
    new_file = request.files.get('new_file')
    if not old_file or not new_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u57fa\u51c6\u7248\u672c\u548c\u5bf9\u6bd4\u7248\u672c\u6574\u673a HQ BOM'})

    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config \u53c2\u6570\u683c\u5f0f\u9519\u8bef'})

    key_col = str(config.get('key_col') or '').strip()
    compare_cols = [str(c).strip() for c in config.get('compare_cols', []) if str(c).strip()]
    if not key_col:
        return jsonify({'success': False, 'error': '\u8bf7\u9009\u62e9\u5339\u914d\u952e\u5217'})
    if not compare_cols:
        return jsonify({'success': False, 'error': '\u8bf7\u81f3\u5c11\u9009\u62e9\u4e00\u4e2a\u6bd4\u5bf9\u5b57\u6bb5'})

    uid = str(uuid.uuid4())[:8]
    try:
        old_path = _save_uploaded_hq_excel(old_file, "bomcmp_machine_old", uid)
        new_path = _save_uploaded_hq_excel(new_file, "bomcmp_machine_new", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    try:
        old_meta = _load_meta(old_path, config.get('old_sheet', ''))
        new_meta = _load_meta(new_path, config.get('new_sheet', ''))
        old_format = old_meta.get('_format')
        new_format = new_meta.get('_format')
        if old_format != new_format:
            return jsonify({'success': False, 'error': '\u4e24\u4efd BOM \u683c\u5f0f\u4e0d\u540c\uff0c\u8bf7\u4f7f\u7528\u540c\u4e00\u79cd\u6574\u673a BOM \u683c\u5f0f\u8fdb\u884c\u7248\u672c\u5bf9\u6bd4'})

        if old_format == 'plm_full':
            old_rows, old_dups, _, old_sheets = _load_plm_full_merged_rows(old_path, key_col, compare_cols)
            new_rows, new_dups, _, new_sheets = _load_plm_full_merged_rows(new_path, key_col, compare_cols)
            common_sheets = [name for name in old_sheets if name in new_sheets]
            if not common_sheets:
                return jsonify({'success': False, 'error': 'PLM \u5168\u91cf\u6574\u673a BOM \u672a\u627e\u5230\u53ef\u6bd4\u5bf9\u7684 BOM \u6216 DBG\u4e1a\u52a1BOM Sheet'})
            stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
            duplicate_notes = _duplicate_notes(old_dups, new_dups)
            out_name = f"Machine_HQ_BOM_version_diff_{uid}.xlsx"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta, '\u6574\u673a HQ BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8', '\u6574\u673a HQ BOM \u7248\u672c\u5bf9\u6bd4', 'machine-hq-version-compare')
            return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'plm_full', 'sheets': common_sheets, **stats})

        old_rows, old_dups, _ = _load_rows(old_path, config.get('old_sheet', ''), HQ_STANDARD_HEADER_ROW, key_col, compare_cols)
        new_rows, new_dups, _ = _load_rows(new_path, config.get('new_sheet', ''), HQ_STANDARD_HEADER_ROW, key_col, compare_cols)
        stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
        duplicate_notes = _duplicate_notes(old_dups, new_dups)
        out_name = f"Machine_HQ_BOM_version_diff_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta, '\u6574\u673a HQ BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8', '\u6574\u673a HQ BOM \u7248\u672c\u5bf9\u6bd4', 'machine-hq-version-compare')
        return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'standard', **stats})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
@bom_compare_bp.route('/api/bom_compare/hq_version', methods=['POST'])
@track_tool_activity('单板HQ BOM版本对比')
def api_hq_version_compare():
    old_file = request.files.get('old_file')
    new_file = request.files.get('new_file')
    if not old_file or not new_file:
        return jsonify({'success': False, 'error': '请上传基准版本和对比版本 HQ BOM 文件'})

    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config 参数格式错误'})

    header_row = HQ_STANDARD_HEADER_ROW

    key_col = str(config.get('key_col') or '').strip()
    compare_cols = [str(c).strip() for c in config.get('compare_cols', []) if str(c).strip()]
    if not key_col:
        return jsonify({'success': False, 'error': '请选择匹配键列'})
    if not compare_cols:
        return jsonify({'success': False, 'error': '请至少选择一个比对字段'})

    uid = str(uuid.uuid4())[:8]
    try:
        old_path = _save_uploaded_hq_excel(old_file, "bomcmp_old", uid)
        new_path = _save_uploaded_hq_excel(new_file, "bomcmp_new", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    try:
        old_meta = _load_meta(old_path, config.get('old_sheet', ''))
        new_meta = _load_meta(new_path, config.get('new_sheet', ''))
        old_format = old_meta.get('_format')
        new_format = new_meta.get('_format')
        if old_format != new_format:
            return jsonify({'success': False, 'error': '\u4e24\u4efd BOM \u683c\u5f0f\u4e0d\u540c\uff0c\u8bf7\u4f7f\u7528\u540c\u4e00\u79cd\u5bfc\u51fa\u683c\u5f0f\u8fdb\u884c\u7248\u672c\u5bf9\u6bd4'})

        if old_format == 'plm_full':
            old_rows, old_dups, _, old_sheets = _load_plm_full_merged_rows(old_path, key_col, compare_cols)
            new_rows, new_dups, _, new_sheets = _load_plm_full_merged_rows(new_path, key_col, compare_cols)
            common_sheets = [name for name in old_sheets if name in new_sheets]
            if not common_sheets:
                return jsonify({'success': False, 'error': 'PLM \u5168\u91cf BOM \u672a\u627e\u5230\u53ef\u6bd4\u5bf9\u7684 BOM \u6216 DBG\u4e1a\u52a1BOM Sheet'})
            stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
            duplicate_notes = _duplicate_notes(old_dups, new_dups)
            out_name = f"PLM_full_BOM_version_diff_{uid}.xlsx"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta, '\u5355\u677f HQ BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8', '\u5355\u677f HQ BOM \u7248\u672c\u5bf9\u6bd4', 'hq-version-compare')
            return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'plm_full', 'sheets': common_sheets, **stats})

        old_rows, old_dups, _ = _load_rows(old_path, config.get('old_sheet', ''), header_row, key_col, compare_cols)
        new_rows, new_dups, _ = _load_rows(new_path, config.get('new_sheet', ''), header_row, key_col, compare_cols)
        stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
        duplicate_notes = _duplicate_notes(old_dups, new_dups)

        out_name = f"HQ_BOM_version_diff_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta, '\u5355\u677f HQ BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8', '\u5355\u677f HQ BOM \u7248\u672c\u5bf9\u6bd4', 'hq-version-compare')
        return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'standard', **stats})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})



