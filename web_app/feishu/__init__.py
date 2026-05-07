# -*- coding: utf-8 -*-
"""飞书多表格匹配工具 — Blueprint"""

import os, uuid, json

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    render_template, request, jsonify, requests,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint

feishu_bp = Blueprint('feishu_tool', __name__)


def _hq_get_sheets(base_url, origin, user_id, token):
    url = f"{base_url}/fs/sheet/v1/spreadsheetsMetainfo"
    r = requests.get(url, params={
        "origin": origin, "userId": user_id, "spreadsheetToken": token,
    }, timeout=15)
    r.raise_for_status()
    d = r.json()
    if d.get("code") not in (0, 200):
        raise RuntimeError(f"获取Sheet失败：{d.get('msg')}")
    return [s for s in d["data"]["sheets"] if s.get("title")]


def _hq_read_sheet(base_url, origin, user_id, token, sheet_id, row_count=200000):
    """读取单个 Sheet 全部数据，返回二维列表（第0行为表头）"""
    end_col = get_column_letter(100)
    all_rows, start = [], 1
    batch_size = 3000
    while start <= max(row_count, 1):
        end = min(start + batch_size - 1, row_count)
        r = requests.get(f"{base_url}/fs/sheet/v1/getSheetsValue", params={
            "origin": origin, "userId": user_id, "spreadsheetToken": token,
            "range": f"{sheet_id}!A{start}:{end_col}{end}",
        }, timeout=60)
        r.raise_for_status()
        d = r.json()
        if d.get("code") not in (0, 200):
            raise RuntimeError(f"读取失败：{d.get('msg')}")
        batch = d["data"]["valueRange"].get("values") or []
        if all_rows and batch:
            batch = batch[1:]  # 后续分片跳过重复表头
        if not batch:
            break
        all_rows.extend(batch)
        expected = end - start + 1
        skip = 1 if start > 1 else 0
        if len(batch) < expected - skip:
            break
        start = end + 1
    # 移除末尾全空行
    while all_rows and not any(_cell_str(v) for v in all_rows[-1]):
        all_rows.pop()
    return all_rows


def _do_feishu_match(local_ws, header_row, tables, gateway, out_file):
    """tables: [{name, token, local_keys, feishu_keys, fetch_cols}, ...]"""
    base_url, origin, user_id = gateway
    max_local_col = local_ws.max_column
    local_headers = [local_ws.cell(row=header_row, column=ci).value
                     for ci in range(1, max_local_col + 1)]

    prepared = []
    all_fetch_cols = []
    seen_cols = set()

    for t in tables:
        try:
            sheets = _hq_get_sheets(base_url, origin, user_id, t['token'])
            rows = []
            for s in sheets:
                batch = _hq_read_sheet(base_url, origin, user_id, t['token'], s['sheetId'])
                if batch:
                    if not rows:
                        rows = batch
                    else:
                        rows.extend(batch[1:])
            if not rows:
                continue

            headers = [_cell_str(v) for v in rows[0]]
            header_set = set(headers)

            lk_cols, fk_cols = [], []
            for lk, fk in zip(t['local_keys'], t['feishu_keys']):
                if not lk or not fk:
                    continue
                try:
                    lc = next(ci + 1 for ci, h in enumerate(local_headers) if _cell_str(h) == lk)
                except StopIteration:
                    continue
                if fk not in header_set:
                    continue
                fci = headers.index(fk)
                lk_cols.append(lc)
                fk_cols.append(fci)
            if not lk_cols:
                continue

            fetch_idxs = []
            for col_name in t['fetch_cols']:
                if col_name in header_set:
                    fetch_idxs.append(headers.index(col_name))
                else:
                    fetch_idxs.append(-1)

            lookup = {}
            for row in rows[1:]:
                key = tuple(_cell_str(row[fc]) if fc < len(row) else "" for fc in fk_cols)
                if not any(key):
                    continue
                vals = {col_name: (_cell_str(row[idx]) if 0 <= idx < len(row) else "")
                        for col_name, idx in zip(t['fetch_cols'], fetch_idxs)}
                lookup.setdefault(key, []).append(vals)

            for col_name in t['fetch_cols']:
                if col_name not in seen_cols:
                    seen_cols.add(col_name)
                    all_fetch_cols.append(col_name)

            prepared.append({
                'name': t['name'], 'local_key_cols': lk_cols,
                'lookup': lookup, 'fetch_cols': t['fetch_cols'],
            })
        except Exception as e:
            print(f"[{t['name']}] 加载失败：{e}")

    if not prepared:
        raise RuntimeError("没有可用的飞书表格")

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "匹配结果"
    thin = Side('thin')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hq_fill = PatternFill('solid', fgColor='FFFF00')
    src_fill = PatternFill('solid', fgColor='BDD7EE')
    hdr_fill = PatternFill('solid', fgColor='D9D9D9')

    out_hdrs = list(local_headers) + all_fetch_cols + ["来源表格"]
    for ci, h in enumerate(out_hdrs, 1):
        c = ws_out.cell(row=1, column=ci, value=h or '')
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='FFC000') if ci > max_local_col else hdr_fill
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = bdr
    for ci in range(1, max_local_col + 1):
        ws_out.column_dimensions[get_column_letter(ci)].width = 18
    for ci in range(max_local_col + 1, len(out_hdrs) + 1):
        ws_out.column_dimensions[get_column_letter(ci)].width = 22

    dr = 2
    matched = unmatched = 0
    for ri in range(header_row + 1, local_ws.max_row + 1):
        row_vals = [local_ws.cell(row=ri, column=ci).value
                    for ci in range(1, max_local_col + 1)]
        if not any(v is not None and str(v).strip() for v in row_vals):
            continue

        found = False
        for pt in prepared:
            key = tuple(_cell_str(row_vals[lc - 1]) for lc in pt['local_key_cols'])
            hits = pt['lookup'].get(key, [])
            if not hits:
                continue
            first = True
            for mdict in hits:
                for ci, val in enumerate(row_vals, 1):
                    c = ws_out.cell(row=dr, column=ci, value=val if first else None)
                    c.alignment = Alignment(horizontal='left', vertical='center')
                    c.border = bdr
                for j, col_name in enumerate(all_fetch_cols):
                    c = ws_out.cell(
                        row=dr, column=max_local_col + j + 1,
                        value=mdict.get(col_name, ''),
                    )
                    c.fill = hq_fill
                    c.alignment = Alignment(horizontal='left', vertical='center')
                    c.border = bdr
                c = ws_out.cell(
                    row=dr, column=max_local_col + len(all_fetch_cols) + 1,
                    value=pt['name'] if first else '',
                )
                c.fill = src_fill
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = bdr
                first = False
                dr += 1
            found = True
            matched += 1
            break

        if not found:
            for ci, val in enumerate(row_vals, 1):
                c = ws_out.cell(row=dr, column=ci, value=val)
                c.alignment = Alignment(horizontal='left', vertical='center')
                c.border = bdr
            for j in range(len(all_fetch_cols)):
                ws_out.cell(row=dr, column=max_local_col + j + 1).border = bdr
            c = ws_out.cell(
                row=dr, column=max_local_col + len(all_fetch_cols) + 1, value="未匹配")
            c.border = bdr
            unmatched += 1
            dr += 1

    wb_out.save(out_file)
    return dr - 2, matched, unmatched, all_fetch_cols


# ── 路由 ─────────────────────────────────────────────────────

@feishu_bp.route('/feishu', methods=['GET', 'POST'])
def tool_feishu():
    if request.method == 'POST':
        action = request.form.get('action', 'match')

        if action == 'sheets':
            token = request.form.get('token', '')
            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            try:
                sheets = _hq_get_sheets(base_url, origin, user_id, token)
                return jsonify({
                    'success': True,
                    'sheets': [{'sheetId': s['sheetId'], 'title': s['title']} for s in sheets],
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

        elif action == 'load':
            token = request.form.get('token', '')
            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            sheet_ids = json.loads(request.form.get('sheet_ids', '[]'))
            try:
                sheets = _hq_get_sheets(base_url, origin, user_id, token)
                rows = []
                for s in sheets:
                    if s['sheetId'] not in sheet_ids:
                        continue
                    batch = _hq_read_sheet(
                        base_url, origin, user_id, token, s['sheetId'],
                        row_count=s.get('rowCount', 200000),
                    )
                    if batch:
                        if not rows:
                            rows = batch
                        else:
                            rows.extend(batch[1:])
                headers = [_cell_str(v) for v in (rows[0] if rows else [])]
                return jsonify({
                    'success': True, 'rows': rows, 'headers': headers,
                    'row_count': max(len(rows) - 1, 0),
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

        elif action == 'match':
            file = request.files.get('file')
            if not file:
                return "请上传文件", 400
            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            header_row = int(request.form.get('header_row', 1))
            tables_json = request.form.get('tables', '[]')
            try:
                tables = json.loads(tables_json)
            except Exception:
                return "表格配置格式错误", 400

            uid = str(uuid.uuid4())[:8]
            in_path = os.path.join(UPLOAD_DIR, f"feishu_in_{uid}.xlsx")
            out_path = os.path.join(OUTPUT_DIR, f"飞书匹配结果_{uid}.xlsx")
            file.save(in_path)

            wb = openpyxl.load_workbook(in_path, data_only=True)
            ws = wb[wb.sheetnames[0]]

            try:
                total, matched, unmatched, cols = _do_feishu_match(
                    ws, header_row, tables, (base_url, origin, user_id), out_path)
                return jsonify({
                    'success': True, 'total': total, 'matched': matched,
                    'unmatched': unmatched,
                    'download': f'/download/飞书匹配结果_{uid}.xlsx',
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

    return render_template('index.html', tables=FEISHU_PRESET_TABLES)
