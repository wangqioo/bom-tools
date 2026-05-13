# -*- coding: utf-8 -*-
"""BOM 比对工具 — Blueprint"""

import os
import uuid
import json
from zipfile import BadZipFile

from flask import Blueprint

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str,
)


bom_compare_bp = Blueprint('bom_compare', __name__)

BAD_EXCEL_ERROR = '无法读取文件，可能原因：① 文件是 .xls 旧格式（请另存为 .xlsx）；② 公司加解密软件未启动导致文件被加密，请检查后重试'
HQ_STANDARD_HEADER_ROW = 3
HQ_STANDARD_HEADERS = ['序号', '料号', '型号', '物料描述', '单耗', '替代关系', '位号', '生产厂家']
HQ_FORMAT_ERROR = '不支持当前文件格式。请上传系统导出的标准 HQ BOM：第 1-2 行为项目信息，第 3 行为表头，且包含序号、料号、型号、物料描述、单耗、替代关系、位号、生产厂家等列。'


def _to_int(value, default=1, min_value=1):
    try:
        result = int(value)
    except (TypeError, ValueError):
        return None
    if min_value is not None and result < min_value:
        return None
    return result


def _headers(ws, header_row):
    headers = []
    for ci in range(1, ws.max_column + 1):
        value = _cell_str(ws.cell(row=header_row, column=ci).value)
        headers.append(value or f"未命名列{get_column_letter(ci)}")
    return headers



def _validate_hq_standard_ws(ws):
    if ws.max_row < HQ_STANDARD_HEADER_ROW:
        raise ValueError(HQ_FORMAT_ERROR)
    meta_keys = {
        _cell_str(ws.cell(row=ri, column=ci).value)
        for ri in (1, 2)
        for ci in range(1, min(ws.max_column, 8) + 1, 2)
    }
    if not {'料号', '描述', '项目配置名', '版本', 'BOM名称'}.issubset(meta_keys):
        raise ValueError(HQ_FORMAT_ERROR)
    headers = _headers(ws, HQ_STANDARD_HEADER_ROW)
    missing = [h for h in HQ_STANDARD_HEADERS if h not in headers]
    if missing:
        raise ValueError(HQ_FORMAT_ERROR)
    return headers


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


def _load_rows(path, sheet_name, header_row, key_col, compare_cols):
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet_name = _pick_sheet(wb, sheet_name)
    ws = wb[sheet_name]
    headers = _validate_hq_standard_ws(ws)
    header_row = HQ_STANDARD_HEADER_ROW
    if key_col not in headers:
        wb.close()
        raise ValueError(f'匹配键列 "{key_col}" 不存在')
    key_idx = headers.index(key_col) + 1
    compare_indices = [(col, headers.index(col) + 1) for col in compare_cols if col in headers]

    rows = {}
    duplicates = {}
    for ri in range(header_row + 1, ws.max_row + 1):
        row_values = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
        if not any(row_values):
            continue
        key = _cell_str(ws.cell(row=ri, column=key_idx).value)
        if not key:
            continue
        item = {
            'row': ri,
            'values': {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices},
            'raw': row_values,
        }
        if key in rows:
            duplicates.setdefault(key, [rows[key]['row']]).append(ri)
            continue
        rows[key] = item
    wb.close()
    return rows, duplicates, headers


def _load_meta(path, sheet_name):
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet_name = _pick_sheet(wb, sheet_name)
    ws = wb[sheet_name]
    meta = {}
    for ri in range(1, min(ws.max_row, 2) + 1):
        for ci in range(1, ws.max_column, 2):
            key = _cell_str(ws.cell(row=ri, column=ci).value)
            val = _cell_str(ws.cell(row=ri, column=ci + 1).value)
            if key:
                meta[key] = val
    wb.close()
    return meta


def _diff_items(old_rows, new_rows, compare_cols):
    items = []
    all_keys = sorted(set(old_rows) | set(new_rows))
    for key in all_keys:
        old = old_rows.get(key)
        new = new_rows.get(key)
        changed_fields = []
        if old and not new:
            diff_type = '删除'
        elif new and not old:
            diff_type = '新增'
        else:
            for col in compare_cols:
                if old['values'].get(col, '') != new['values'].get(col, ''):
                    changed_fields.append(col)
            diff_type = '变更' if changed_fields else '未变更'
        items.append({
            'key': key,
            'type': diff_type,
            'old': old,
            'new': new,
            'changed_fields': changed_fields,
        })
    return items


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
        ('基准版本唯一物料数', stats['old_total']),
        ('对比版本唯一物料数', stats['new_total']),
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

    detail_headers = ['差异类型', '料号', '基准版本行号', '对比版本行号', '变更字段']
    for col in compare_cols:
        detail_headers.extend([f'基准版本{col}', f'对比版本{col}'])

    def write_table(sheet, table_items):
        for ci, header in enumerate(detail_headers, 1):
            c = sheet.cell(row=1, column=ci, value=header)
            c.font = Font(bold=True)
            c.alignment = center
            c.border = bdr
            c.fill = header_fill
        for ri, item in enumerate(table_items, 2):
            old = item['old']
            new = item['new']
            row_values = [
                item['type'],
                item['key'],
                old['row'] if old else '',
                new['row'] if new else '',
                '、'.join(item['changed_fields']),
            ]
            for col in compare_cols:
                row_values.extend([
                    old['values'].get(col, '') if old else '',
                    new['values'].get(col, '') if new else '',
                ])
            row_fill = fills[item['type']]
            for ci, value in enumerate(row_values, 1):
                c = sheet.cell(row=ri, column=ci, value=value)
                c.alignment = center if ci in (1, 3, 4) else left
                c.border = bdr
                c.fill = row_fill
                if item['type'] == '变更' and ci >= 6:
                    field_idx = (ci - 6) // 2
                    if field_idx < len(compare_cols) and compare_cols[field_idx] in item['changed_fields']:
                        c.fill = fills['变更']
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
            sheet.column_dimensions[get_column_letter(col)].width = 16 if col not in (2, 5) else 28
    wb.save(out_path)


@bom_compare_bp.route('/api/bom_compare/local_sheets', methods=['POST'])
def api_local_sheets():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    if file.filename and file.filename.lower().endswith('.xls') and not file.filename.lower().endswith('.xlsx'):
        return jsonify({'success': False, 'error': '不支持 .xls 格式，请在 Excel 中另存为 .xlsx 后重试'})

    uid = str(uuid.uuid4())[:8]
    path = os.path.join(UPLOAD_DIR, f"bomcmp_pre_{uid}.xlsx")
    file.save(path)
    header_row = _to_int(request.form.get('header_row', 1), 1)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        sheet_name = _pick_sheet(wb, request.form.get('sheet_name', ''))
        wb.close()
        wb2 = openpyxl.load_workbook(path, data_only=True)
        ws = wb2[sheet_name]
        headers = _validate_hq_standard_ws(ws)
        wb2.close()
        return jsonify({
            'success': True,
            'sheets': sheets,
            'current_sheet': sheet_name,
            'headers': headers,
            'detected_key': _detect_key(headers),
        })
    except BadZipFile:
        return jsonify({'success': False, 'error': BAD_EXCEL_ERROR})
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
    old_path = os.path.join(UPLOAD_DIR, f"bomcmp_old_{uid}.xlsx")
    new_path = os.path.join(UPLOAD_DIR, f"bomcmp_new_{uid}.xlsx")
    old_file.save(old_path)
    new_file.save(new_path)

    try:
        old_rows, old_dups, _ = _load_rows(old_path, config.get('old_sheet', ''), header_row, key_col, compare_cols)
        new_rows, new_dups, _ = _load_rows(new_path, config.get('new_sheet', ''), header_row, key_col, compare_cols)

        added = sorted(set(new_rows) - set(old_rows))
        removed = sorted(set(old_rows) - set(new_rows))
        common = sorted(set(old_rows) & set(new_rows))
        changed = []
        unchanged = []
        for key in common:
            has_change = any(old_rows[key]['values'].get(col, '') != new_rows[key]['values'].get(col, '') for col in compare_cols)
            (changed if has_change else unchanged).append(key)

        stats = {
            'old_total': len(old_rows),
            'new_total': len(new_rows),
            'added': len(added),
            'removed': len(removed),
            'changed': len(changed),
            'unchanged': len(unchanged),
            'old_duplicates': len(old_dups),
            'new_duplicates': len(new_dups),
        }
        duplicate_notes = (
            [f"基准版本重复键 {key}: 行 {', '.join(map(str, rows))}" for key, rows in old_dups.items()] +
            [f"对比版本重复键 {key}: 行 {', '.join(map(str, rows))}" for key, rows in new_dups.items()]
        )

        out_name = f"HQ_BOM版本差异_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        old_meta = _load_meta(old_path, config.get('old_sheet', ''))
        new_meta = _load_meta(new_path, config.get('new_sheet', ''))
        _write_diff_report(out_path, old_rows, new_rows, compare_cols, stats, duplicate_notes, old_meta, new_meta)
        return jsonify({'success': True, 'download': f'/download/{out_name}', **stats})
    except BadZipFile:
        return jsonify({'success': False, 'error': BAD_EXCEL_ERROR})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
