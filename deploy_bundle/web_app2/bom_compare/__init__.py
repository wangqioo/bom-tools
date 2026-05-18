# -*- coding: utf-8 -*-
"""BOM 比对工具 — Blueprint"""

import os
import uuid
import json
import re
import subprocess

from flask import Blueprint

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str,
    _open_workbook, _save_uploaded_excel, _to_int,
)


bom_compare_bp = Blueprint('bom_compare', __name__)

HQ_STANDARD_HEADER_ROW = 3
HQ_STANDARD_HEADERS = ['\u5e8f\u53f7', '\u6599\u53f7', '\u578b\u53f7', '\u7269\u6599\u63cf\u8ff0', '\u5355\u8017', '\u66ff\u4ee3\u5173\u7cfb', '\u4f4d\u53f7', '\u751f\u4ea7\u5382\u5bb6']
PLM_FULL_HEADER_ROW = 8
PLM_FULL_SHEETS = ['BOM', 'DBG\u4e1a\u52a1BOM', 'DBGBOM\u5236\u63a7\u4fe1\u606f']
CADENCE_STANDARD_HEADER_ROW = 3
CADENCE_REQUIRED_HEADERS = ['\u5e8f\u53f7', '\u6599\u53f7', '\u578b\u53f7', '\u5355\u8017', '\u4f4d\u53f7']
HQ_FORMAT_ERROR = '\u4e0d\u652f\u6301\u5f53\u524d\u6587\u4ef6\u683c\u5f0f\u3002\u8bf7\u4e0a\u4f20\u7cfb\u7edf\u5bfc\u51fa\u7684\u6807\u51c6 HQ BOM\uff0c\u6216 PLM \u5168\u91cf BOM\uff1a\u4e24\u7c7b\u683c\u5f0f\u90fd\u5fc5\u987b\u5305\u542b\u5e8f\u53f7\u3001\u6599\u53f7\u3001\u578b\u53f7\u3001\u7269\u6599\u63cf\u8ff0\u3001\u5355\u8017\u3001\u66ff\u4ee3\u5173\u7cfb\u3001\u4f4d\u53f7\u3001\u751f\u4ea7\u5382\u5bb6\u7b49\u5217\u3002'
HQ_XLS_CONVERT_ERROR = '\u65e0\u6cd5\u76f4\u63a5\u8bfb\u53d6\u8be5 .xls \u6587\u4ef6\u3002\u8bf7\u786e\u8ba4\u670d\u52a1\u5668\u4e3a Windows \u4e14\u5df2\u5b89\u88c5\u53ef\u89e3\u5bc6\u6b64\u6587\u4ef6\u7684 Excel\uff0c\u6216\u5148\u5728 Excel \u4e2d\u53e6\u5b58\u4e3a .xlsx \u540e\u518d\u4e0a\u4f20\u3002'

GENERIC_COMPARE_TYPES = {
    'customer_hq': {
        'title': '\u5ba2\u6237BOM \u5bf9\u6bd4 HQ BOM \u5dee\u5f02\u603b\u89c8',
        'left_label': '\u5ba2\u6237BOM',
        'right_label': 'HQ BOM',
        'filename': '\u5ba2\u6237BOM\u5bf9\u6bd4HQ_BOM',
    },
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



def _ps_single_quote(value):
    return "'" + str(value).replace("'", "''") + "'"

def _convert_xls_with_excel(src_path, uid):
    if os.name != 'nt':
        raise ValueError(HQ_XLS_CONVERT_ERROR)
    out_path = os.path.join(UPLOAD_DIR, f'bomcmp_converted_{uid}.xlsx')
    script = (
        "$ErrorActionPreference='Stop';"
        f"$src={_ps_single_quote(src_path)};"
        f"$dst={_ps_single_quote(out_path)};"
        "$excel=New-Object -ComObject Excel.Application;"
        "$excel.Visible=$false;$excel.DisplayAlerts=$false;"
        "try{$wb=$excel.Workbooks.Open($src);$wb.SaveAs($dst,51);$wb.Close($false)}"
        "finally{$excel.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null}"
    )
    try:
        subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script],
            check=True, capture_output=True, text=True, timeout=90,
        )
    except Exception as exc:
        raise ValueError(HQ_XLS_CONVERT_ERROR) from exc
    if not os.path.exists(out_path):
        raise ValueError(HQ_XLS_CONVERT_ERROR)
    return out_path


def _save_uploaded_hq_excel(file, prefix, uid):
    if not file:
        raise ValueError('\u8bf7\u4e0a\u4f20\u6587\u4ef6')
    filename = file.filename or ''
    lower = filename.lower()
    if lower.endswith('.xls') and not lower.endswith('.xlsx'):
        raw_path = os.path.join(UPLOAD_DIR, f'{prefix}_{uid}.xls')
        file.save(raw_path)
        return _convert_xls_with_excel(raw_path, uid)
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

def _normalize_header(value):
    return ''.join(str(value or '').lower().split()).replace('_', '').replace('-', '')


def _detect_common_key(left_headers, right_headers, prefer_part_no=False):
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
    bom_sheets = [name for name in PLM_FULL_SHEETS if name in sheets] if fmt['kind'] == 'plm_full' else [sheet_name]
    wb.close()
    return sheet_name, headers, fmt, sheets, bom_sheets


def _load_hq_side_rows(path, sheet_name, key_col, compare_cols, expand_refdes=False):
    rows, duplicates, headers = _load_rows(path, sheet_name, None, key_col, compare_cols, expand_refdes=expand_refdes)
    return rows, duplicates, headers, 0


def _is_refdes_header(value):
    norm = _normalize_header(value)
    return '位号' in norm or 'refdes' in norm or 'reference' in norm


def _is_qty_header(value):
    norm = _normalize_header(value)
    return norm in {'单耗', '数量', 'qty', 'quantity'}


def _expand_compare_values(values, key_col, key):
    expanded = dict(values)
    if key_col in expanded:
        expanded[key_col] = key
    for col in list(expanded):
        if _is_qty_header(col):
            expanded[col] = '1'
    return expanded


def _load_generic_rows(path, sheet_name, header_row, key_col, compare_cols, expand_refdes=False):
    wb = _open_workbook(path, data_only=True)
    sheet_name = _pick_sheet(wb, sheet_name)
    ws = wb[sheet_name]
    headers = _headers(ws, header_row)
    if key_col not in headers:
        wb.close()
        raise ValueError(f'匹配键列 "{key_col}" 不存在')

    key_idx = headers.index(key_col) + 1
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
        key = _cell_str(ws.cell(row=ri, column=key_idx).value)
        if not key:
            skipped_blank_key += 1
            continue
        keys = _split_refdes(key) if expand_refdes else [key]
        values = {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices}
        for item_key in keys:
            if not item_key:
                continue
            item = {
                'row': ri,
                'values': _expand_compare_values(values, key_col, item_key) if expand_refdes else values,
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
                if left['values'].get(left_col, '') != right['values'].get(right_col, ''):
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


def _write_generic_compare_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, labels):
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

    ws.merge_cells('A1:D1')
    ws['A1'] = labels['title']
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = center

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
    for ri, (name, value) in enumerate(summary_rows, 3):
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
            if item['type'] == '字段变更' and left_value == right_value:
                continue
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
        for ci, header in enumerate(detail_headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        ri = 2
        for item in table_items:
            row_fill = fills[item['type']]
            for row_values in item_field_rows(item):
                for ci, value in enumerate(row_values, 1):
                    c = sheet.cell(row=ri, column=ci, value=value)
                    c.alignment = center if ci in (1, 3, 4) else left_align
                    c.border = bdr
                    c.fill = row_fill
                ri += 1
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    diff_items = [i for i in items if i['type'] != '一致']
    write_table(wb.create_sheet('差异明细'), diff_items)
    write_table(wb.create_sheet(_safe_sheet_title(left_only_type)), [i for i in items if i['type'] == left_only_type])
    write_table(wb.create_sheet(_safe_sheet_title(right_only_type)), [i for i in items if i['type'] == right_only_type])
    write_table(wb.create_sheet('字段变更'), [i for i in items if i['type'] == '字段变更'])

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


def _load_rows(path, sheet_name, header_row, key_col, compare_cols, expand_refdes=False):
    wb, sheet_name, ws, fmt = _open_hq_workbook_info(path, sheet_name)
    headers = fmt['headers']
    header_row = fmt['header_row']
    if key_col not in headers:
        wb.close()
        raise ValueError(f'\u5339\u914d\u952e\u5217 "{key_col}" \u4e0d\u5b58\u5728')
    key_idx = headers.index(key_col) + 1
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
        key = _cell_str(ws.cell(row=ri, column=key_idx).value)
        if not key:
            continue
        values = {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices}
        refdes_list = _split_refdes(ws.cell(row=ri, column=ref_idx).value) if ref_idx else []
        if refdes_list == ['']:
            refdes_list = []
        keys = _split_refdes(key) if expand_refdes else [key]
        for item_key in keys:
            if not item_key:
                continue
            item = {
                'key': item_key,
                'row': ri,
                'values': _expand_compare_values(values, key_col, item_key) if expand_refdes else values,
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
                elif old['values'].get(col, '') != new['values'].get(col, ''):
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
            else old_rows[key]['values'].get(col, '') != new_rows[key]['values'].get(col, '')
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


def _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta=None, new_meta=None):
    wb = Workbook()
    ws = wb.active
    ws.title = '差异总览'

    fills = {
        '新增': PatternFill('solid', fgColor='E8F5E9'),
        '删除': PatternFill('solid', fgColor='FFEBEE'),
        '变更': PatternFill('solid', fgColor='FFF9C4'),
        '未变更': PatternFill('solid', fgColor='F5F5F5'),
    }
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

    ws.merge_cells('A1:D1')
    ws['A1'] = 'HQ BOM 版本差异总览'
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = center
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
    for ri, (name, value) in enumerate(summary_rows, 3):
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
            row_fill = fills[item['type']]
            for ci, value in enumerate(row_values, 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = center if ci in (1, 3, 4) else left
                c.border = bdr
                c.fill = row_fill
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    write_table(wb.create_sheet('差异明细'), items)
    write_table(wb.create_sheet('新增物料'), [i for i in items if i['type'] == '新增'])
    write_table(wb.create_sheet('删除物料'), [i for i in items if i['type'] == '删除'])
    write_table(wb.create_sheet('变更物料'), [i for i in items if i['type'] == '变更'])

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

    for sheet in wb.worksheets:
        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 16 if col not in (2, 5, 6, 7) else 28
    wb.save(out_path)



def _write_plm_full_diff_report(out_path, sheet_results, compare_cols, old_meta=None, new_meta=None):
    wb = Workbook()
    ws = wb.active
    ws.title = '\u5dee\u5f02\u603b\u89c8'

    fills = {
        '\u65b0\u589e': PatternFill('solid', fgColor='E8F5E9'),
        '\u5220\u9664': PatternFill('solid', fgColor='FFEBEE'),
        '\u53d8\u66f4': PatternFill('solid', fgColor='FFF9C4'),
        '\u672a\u53d8\u66f4': PatternFill('solid', fgColor='F5F5F5'),
    }
    bdr = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    header_fill = PatternFill('solid', fgColor='D9EAF7')
    title_fill = PatternFill('solid', fgColor='1F4E78')
    title_font = Font(bold=True, color='FFFFFF', size=14)
    old_meta = old_meta or {}
    new_meta = new_meta or {}

    ws.merge_cells('A1:I1')
    ws['A1'] = 'PLM \u5168\u91cf BOM \u7248\u672c\u5dee\u5f02\u603b\u89c8'
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = center
    meta_rows = [
        ('\u9879\u76ee\u914d\u7f6e\u540d', new_meta.get('\u9879\u76ee\u914d\u7f6e\u540d') or old_meta.get('\u9879\u76ee\u914d\u7f6e\u540d', '')),
        ('BOM\u540d\u79f0', new_meta.get('BOM\u540d\u79f0') or old_meta.get('BOM\u540d\u79f0', '')),
        ('\u57fa\u51c6\u7248\u672c\u53f7', old_meta.get('\u7248\u672c', '')),
        ('\u5bf9\u6bd4\u7248\u672c\u53f7', new_meta.get('\u7248\u672c', '')),
        ('\u63d0\u4ea4\u65f6\u95f4', new_meta.get('\u63d0\u4ea4\u65f6\u95f4') or old_meta.get('\u63d0\u4ea4\u65f6\u95f4', '')),
        ('\u91cf\u4ea7/\u8bd5\u4ea7', new_meta.get('\u91cf\u4ea7/\u8bd5\u4ea7') or old_meta.get('\u91cf\u4ea7/\u8bd5\u4ea7', '')),
    ]
    for ri, (name, value) in enumerate(meta_rows, 3):
        ws.cell(row=ri, column=1, value=name).font = Font(bold=True)
        ws.cell(row=ri, column=2, value=value)
        ws.cell(row=ri, column=1).border = bdr
        ws.cell(row=ri, column=2).border = bdr

    summary_header_row = 11
    summary_headers = ['Sheet', '\u57fa\u51c6\u7248\u672c\u552f\u4e00\u6599\u53f7\u6570', '\u5bf9\u6bd4\u7248\u672c\u552f\u4e00\u6599\u53f7\u6570', '\u65b0\u589e', '\u5220\u9664', '\u53d8\u66f4', '\u672a\u53d8\u66f4', '\u57fa\u51c6\u91cd\u590d\u952e', '\u5bf9\u6bd4\u91cd\u590d\u952e']
    for ci, header in enumerate(summary_headers, 1):
        c = ws.cell(row=summary_header_row, column=ci, value=header)
        c.font = Font(bold=True)
        c.fill = header_fill
        c.border = bdr
        c.alignment = center
    total = {k: 0 for k in ['old_total', 'new_total', 'added', 'removed', 'changed', 'unchanged', 'old_duplicates', 'new_duplicates']}
    for ri, result in enumerate(sheet_results, summary_header_row + 1):
        stats = result['stats']
        for key in total:
            total[key] += stats.get(key, 0)
        values = [result['sheet'], stats['old_total'], stats['new_total'], stats['added'], stats['removed'], stats['changed'], stats['unchanged'], stats['old_duplicates'], stats['new_duplicates']]
        for ci, value in enumerate(values, 1):
            c = ws.cell(row=ri, column=ci, value=value)
            c.border = bdr
            c.alignment = center if ci != 1 else left
    total_row = summary_header_row + 1 + len(sheet_results)
    values = ['\u5408\u8ba1', total['old_total'], total['new_total'], total['added'], total['removed'], total['changed'], total['unchanged'], total['old_duplicates'], total['new_duplicates']]
    for ci, value in enumerate(values, 1):
        c = ws.cell(row=total_row, column=ci, value=value)
        c.font = Font(bold=True)
        c.border = bdr
        c.alignment = center if ci != 1 else left

    detail_headers = ['Sheet', '\u5dee\u5f02\u7c7b\u578b', '\u6599\u53f7', '\u57fa\u51c6\u7248\u672c\u884c\u53f7', '\u5bf9\u6bd4\u7248\u672c\u884c\u53f7', '\u53d8\u66f4\u5b57\u6bb5', '\u57fa\u51c6\u503c', '\u5bf9\u6bd4\u503c']

    def changed_values(item, field):
        old = item['old']
        new = item['new']
        if field == '\u4f4d\u53f7':
            return '\u3001'.join(item.get('removed_refdes', [])), '\u3001'.join(item.get('added_refdes', []))
        return (
            old['values'].get(field, '') if old else '',
            new['values'].get(field, '') if new else '',
        )

    def expanded_rows(sheet_name, table_items):
        for item in table_items:
            fields = item['changed_fields'] if item['type'] == '\u53d8\u66f4' and item['changed_fields'] else ['']
            for field in fields:
                old_value, new_value = changed_values(item, field) if field else ('', '')
                yield sheet_name, item, field, old_value, new_value

    def write_table(sheet, rows):
        for ci, header in enumerate(detail_headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        ri = 2
        for sheet_name, item, field, old_value, new_value in rows:
            old = item['old']
            new = item['new']
            values = [sheet_name, item['type'], item['key'], old['row'] if old else '', new['row'] if new else '', field, old_value, new_value]
            for ci, value in enumerate(values, 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = center if ci in (1, 2, 4, 5) else left
                c.border = bdr
                c.fill = fills[item['type']]
            ri += 1
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions

    all_rows = []
    added_rows = []
    removed_rows = []
    changed_rows = []
    for result in sheet_results:
        sheet_name = result['sheet']
        items = result['items']
        all_rows.extend(expanded_rows(sheet_name, items))
        added_rows.extend(expanded_rows(sheet_name, [i for i in items if i['type'] == '\u65b0\u589e']))
        removed_rows.extend(expanded_rows(sheet_name, [i for i in items if i['type'] == '\u5220\u9664']))
        changed_rows.extend(expanded_rows(sheet_name, [i for i in items if i['type'] == '\u53d8\u66f4']))
        write_table(wb.create_sheet(_safe_sheet_title(f'{sheet_name}\u5dee\u5f02')), expanded_rows(sheet_name, items))

    write_table(wb.create_sheet('\u5168\u90e8\u5dee\u5f02\u660e\u7ec6'), all_rows)
    write_table(wb.create_sheet('\u5168\u90e8\u65b0\u589e\u7269\u6599'), added_rows)
    write_table(wb.create_sheet('\u5168\u90e8\u5220\u9664\u7269\u6599'), removed_rows)
    write_table(wb.create_sheet('\u5168\u90e8\u53d8\u66f4\u7269\u6599'), changed_rows)

    ws_dup = wb.create_sheet('\u91cd\u590d\u6599\u53f7')
    ws_dup.append(['\u7c7b\u578b', '\u6599\u53f7\u548c\u884c\u53f7'])
    for cell in ws_dup[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = bdr
    for result in sheet_results:
        for note in result['duplicate_notes']:
            kind = '\u57fa\u51c6\u7248\u672c' if '\u57fa\u51c6\u7248\u672c' in note else '\u5bf9\u6bd4\u7248\u672c'
            ws_dup.append([kind, note])
    for row in ws_dup.iter_rows():
        for cell in row:
            cell.border = bdr
            cell.alignment = left

    for sheet in wb.worksheets:
        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 16 if col not in (1, 3, 6, 7, 8) else 28
    wb.save(out_path)





@bom_compare_bp.route('/api/bom_compare/generic_sheets', methods=['POST'])
def api_generic_sheets():
    left_file = request.files.get('left_file') or request.files.get('file')
    right_file = request.files.get('right_file')
    if not left_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u5de6\u4fa7 BOM \u6587\u4ef6'})

    header_row = _to_int(request.form.get('header_row', 1), 1)
    left_header_row = _to_int(request.form.get('left_header_row', header_row), 1)
    right_header_row = _to_int(request.form.get('right_header_row', header_row), 1)
    compare_type = str(request.form.get('compare_type') or 'customer_hq')
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
            left_key, right_key = _detect_common_key(left_headers, right_headers, prefer_part_no=(compare_type == 'cadence_hq'))
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
def api_generic_compare():
    left_file = request.files.get('left_file')
    right_file = request.files.get('right_file')
    if not left_file or not right_file:
        return jsonify({'success': False, 'error': '\u8bf7\u4e0a\u4f20\u4e24\u4efd BOM \u6587\u4ef6'})

    try:
        config = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'success': False, 'error': 'config \u53c2\u6570\u683c\u5f0f\u9519\u8bef'})

    compare_type = str(config.get('compare_type') or 'customer_hq')
    labels = GENERIC_COMPARE_TYPES.get(compare_type, GENERIC_COMPARE_TYPES['customer_hq'])
    left_header_row = _to_int(config.get('left_header_row', config.get('header_row', 1)), 1)
    right_header_row = _to_int(config.get('right_header_row', config.get('header_row', 1)), 1)
    if left_header_row is None or right_header_row is None:
        return jsonify({'success': False, 'error': '\u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})

    left_key_col = str(config.get('left_key_col') or '').strip()
    right_key_col = str(config.get('right_key_col') or '').strip()
    if not left_key_col or not right_key_col:
        return jsonify({'success': False, 'error': '\u8bf7\u9009\u62e9\u4e24\u4efd BOM \u7684\u5339\u914d\u952e\u5217'})

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
        field_pairs = [(l, r) for l, r in field_pairs if l != left_key_col or r != right_key_col]
        if not field_pairs:
            return jsonify({'success': False, 'error': '\u8bf7\u81f3\u5c11\u9009\u62e9\u4e00\u7ec4\u9700\u8981\u6bd4\u5bf9\u7684\u5b57\u6bb5'})

        left_compare_cols = [l for l, _ in field_pairs]
        right_compare_cols = [r for _, r in field_pairs]
        expand_refdes = compare_type == 'cadence_hq' and _is_refdes_header(left_key_col) and _is_refdes_header(right_key_col)
        left_rows, left_dups, _, left_blank = _load_generic_rows(left_path, left_sheet, left_fmt['header_row'], left_key_col, left_compare_cols, expand_refdes=expand_refdes)
        right_rows, right_dups, _, right_blank = _load_hq_side_rows(right_path, right_sheet, right_key_col, right_compare_cols, expand_refdes=expand_refdes)

        left_only = sorted(set(left_rows) - set(right_rows))
        right_only = sorted(set(right_rows) - set(left_rows))
        common = sorted(set(left_rows) & set(right_rows))
        changed = []
        same = []
        for key in common:
            has_change = any(left_rows[key]['values'].get(l, '') != right_rows[key]['values'].get(r, '') for l, r in field_pairs)
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
        _write_generic_compare_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, labels)
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
        headers = fmt['headers']
        bom_sheets = [name for name in PLM_FULL_SHEETS if name in sheets] if fmt['kind'] == 'plm_full' else [sheet_name]
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


@bom_compare_bp.route('/api/bom_compare/hq_version', methods=['POST'])
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
            old_wb = _open_workbook(old_path, read_only=True, data_only=True)
            new_wb = _open_workbook(new_path, read_only=True, data_only=True)
            common_sheets = [name for name in PLM_FULL_SHEETS if name in old_wb.sheetnames and name in new_wb.sheetnames]
            old_wb.close()
            new_wb.close()
            if not common_sheets:
                return jsonify({'success': False, 'error': 'PLM \u5168\u91cf BOM \u672a\u627e\u5230\u53ef\u6bd4\u5bf9\u7684 BOM Sheet'})
            sheet_results = []
            total_stats = {k: 0 for k in ['old_total', 'new_total', 'added', 'removed', 'changed', 'unchanged', 'old_duplicates', 'new_duplicates']}
            for sheet in common_sheets:
                old_rows, old_dups, _ = _load_rows(old_path, sheet, PLM_FULL_HEADER_ROW, key_col, compare_cols)
                new_rows, new_dups, _ = _load_rows(new_path, sheet, PLM_FULL_HEADER_ROW, key_col, compare_cols)
                stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
                for key in total_stats:
                    total_stats[key] += stats.get(key, 0)
                sheet_results.append({
                    'sheet': sheet,
                    'old_rows': old_rows,
                    'new_rows': new_rows,
                    'stats': stats,
                    'items': _diff_items(old_rows, new_rows, compare_cols),
                    'duplicate_notes': _duplicate_notes(old_dups, new_dups, sheet),
                })
            out_name = f"PLM_full_BOM_version_diff_{uid}.xlsx"
            out_path = os.path.join(OUTPUT_DIR, out_name)
            _write_plm_full_diff_report(out_path, sheet_results, compare_cols, old_meta, new_meta)
            return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'plm_full', 'sheets': common_sheets, **total_stats})

        old_rows, old_dups, _ = _load_rows(old_path, config.get('old_sheet', ''), header_row, key_col, compare_cols)
        new_rows, new_dups, _ = _load_rows(new_path, config.get('new_sheet', ''), header_row, key_col, compare_cols)
        stats = _hq_stats(old_rows, old_dups, new_rows, new_dups, compare_cols)
        duplicate_notes = _duplicate_notes(old_dups, new_dups)

        out_name = f"HQ_BOM_version_diff_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta)
        return jsonify({'success': True, 'download': f'/download/{out_name}', 'format': 'standard', **stats})
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
