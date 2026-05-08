# -*- coding: utf-8 -*-
"""飞书多表格匹配 — Blueprint"""

import os, uuid, json
from flask import Blueprint
from shared import (
    requests as _requests,
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str,
)

feishu_bp = Blueprint('feishu', __name__)


# ── 飞书 API ────────────────────────────────────────────────

def _hq_get_sheets(base_url, origin, user_id, token):
    url = f"{base_url.rstrip('/')}/fs/sheet/v1/spreadsheetsMetainfo"
    r = _requests.get(url, params={
        "origin": origin, "userId": user_id,
        "spreadsheetToken": token,
    }, timeout=15)
    r.raise_for_status()
    d = r.json()
    if d.get("code") not in (0, 200):
        raise RuntimeError(f"获取 Sheet 列表失败：{d.get('msg')}（code={d.get('code')}）")
    return [s for s in d["data"]["sheets"] if s.get("title")]


def _hq_read_sheet(base_url, origin, user_id, token,
                   sheet_id, row_count=200000, col_count=100,
                   batch_size=3000, progress_cb=None):
    end_col = get_column_letter(max(col_count, 26))
    all_rows, start = [], 1
    while start <= max(row_count, 1):
        end = min(start + batch_size - 1, row_count)
        params = {
            "origin": origin, "userId": user_id,
            "spreadsheetToken": token,
            "range": f"{sheet_id}!A{start}:{end_col}{end}",
        }
        r = _requests.get(
            f"{base_url.rstrip('/')}/fs/sheet/v1/getSheetsValue",
            params=params, timeout=60,
        )
        r.raise_for_status()
        d = r.json()
        if d.get("code") not in (0, 200):
            raise RuntimeError(f"读取失败：{d.get('msg')}")
        batch = d["data"]["valueRange"].get("values") or []
        if all_rows and batch:
            batch = batch[1:]
        if not batch:
            break
        all_rows.extend(batch)
        if progress_cb:
            progress_cb(len(all_rows))
        expected = end - start + 1
        skip = 1 if start > 1 else 0
        if len(batch) < expected - skip:
            break
        start = end + 1
    while all_rows and not any(_cell_str(v) for v in all_rows[-1]):
        all_rows.pop()
    return all_rows


def _hq_read_table(base_url, origin, user_id, token, sheets_meta, active_ids):
    combined, headers_set = [], False
    for s in sheets_meta:
        if s["sheetId"] not in active_ids:
            continue
        rows = _hq_read_sheet(
            base_url, origin, user_id, token,
            s["sheetId"],
            row_count=s.get("rowCount", 200000),
            col_count=s.get("columnCount", 100),
            batch_size=3000,
        )
        if not rows:
            continue
        if not headers_set:
            combined = rows
            headers_set = True
        else:
            combined.extend(rows[1:])
    return combined


def _do_match_multi(local_ws, local_header_row, prepared_tables, all_fetch_cols, out_file):
    max_local_col = local_ws.max_column
    local_header = [local_ws.cell(row=local_header_row, column=ci).value
                    for ci in range(1, max_local_col + 1)]

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "匹配结果"
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="D9D9D9")
    hq_fill = PatternFill("solid", fgColor="FFFF00")
    src_fill = PatternFill("solid", fgColor="BDD7EE")

    out_hdrs = list(local_header) + all_fetch_cols + ["来源表格"]
    for ci, h in enumerate(out_hdrs, 1):
        c = ws_out.cell(row=1, column=ci, value=h or "")
        c.font = Font(bold=True)
        c.fill = (PatternFill("solid", fgColor="FFC000") if ci > max_local_col else hdr_fill)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr
    ws_out.row_dimensions[1].height = 22
    for ci in range(1, max_local_col + 1):
        ws_out.column_dimensions[get_column_letter(ci)].width = 18
    for ci in range(max_local_col + 1, len(out_hdrs) + 1):
        ws_out.column_dimensions[get_column_letter(ci)].width = 24

    dr = 2
    matched = unmatched = 0

    for ri in range(local_header_row + 1, local_ws.max_row + 1):
        row_vals = [local_ws.cell(row=ri, column=ci).value
                    for ci in range(1, max_local_col + 1)]
        if not any(v is not None and str(v).strip() for v in row_vals):
            continue

        found = False
        for pt in prepared_tables:
            key = tuple(_cell_str(row_vals[lc - 1]) for lc in pt["local_key_cols"])
            matches = pt["lookup"].get(key, [])
            if not matches:
                continue
            first = True
            for mdict in matches:
                for ci, val in enumerate(row_vals, 1):
                    c = ws_out.cell(row=dr, column=ci, value=val if first else None)
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    c.border = bdr
                for j, col_name in enumerate(all_fetch_cols):
                    c = ws_out.cell(row=dr, column=max_local_col + j + 1,
                                    value=mdict.get(col_name, ""))
                    c.fill = hq_fill
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    c.border = bdr
                src_col = max_local_col + len(all_fetch_cols) + 1
                c = ws_out.cell(row=dr, column=src_col, value=pt["name"] if first else "")
                c.fill = src_fill
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = bdr
                first = False
                dr += 1
            found = True
            matched += 1
            break

        if not found:
            for ci, val in enumerate(row_vals, 1):
                c = ws_out.cell(row=dr, column=ci, value=val)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = bdr
            for j in range(len(all_fetch_cols)):
                ws_out.cell(row=dr, column=max_local_col + j + 1).border = bdr
            c = ws_out.cell(row=dr, column=max_local_col + len(all_fetch_cols) + 1, value="未匹配")
            c.border = bdr
            unmatched += 1
            dr += 1

    wb_out.save(out_file)
    total = dr - 2
    return total, matched, unmatched


# ── 路由 ─────────────────────────────────────────────────────

@feishu_bp.route('/api/feishu/sheets', methods=['POST'])
def api_feishu_sheets():
    """获取飞书表格的 Sheet 列表"""
    data = request.get_json(silent=True) or {}
    base_url = data.get('base_url', 'https://mcenter.huaqin.com')
    origin = data.get('origin', '')
    user_id = data.get('user_id', '')
    token = data.get('token', '').strip()
    if not token:
        return jsonify({'success': False, 'error': '请填写 Token'})
    try:
        sheets = _hq_get_sheets(base_url, origin, user_id, token)
        return jsonify({'success': True, 'sheets': sheets})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@feishu_bp.route('/api/feishu/headers', methods=['POST'])
def api_feishu_headers():
    """获取飞书表格的列标题（读第一行）"""
    data = request.get_json(silent=True) or {}
    base_url = data.get('base_url', 'https://mcenter.huaqin.com')
    origin = data.get('origin', '')
    user_id = data.get('user_id', '')
    token = data.get('token', '').strip()
    sheet_id = data.get('sheet_id', '')
    col_count = data.get('col_count', 100)
    if not token or not sheet_id:
        return jsonify({'success': False, 'error': '缺少参数'})
    try:
        rows = _hq_read_sheet(base_url, origin, user_id, token,
                              sheet_id, row_count=2, col_count=col_count, batch_size=10)
        headers = [_cell_str(v) for v in (rows[0] if rows else [])]
        return jsonify({'success': True, 'headers': headers})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@feishu_bp.route('/api/feishu/match', methods=['POST'])
def api_feishu_match():
    """执行飞书多表格匹配"""
    local_file = request.files.get('file')
    if not local_file:
        return jsonify({'success': False, 'error': '请上传本地 Excel 文件'})

    config_str = request.form.get('config', '{}')
    try:
        config = json.loads(config_str)
    except Exception:
        return jsonify({'success': False, 'error': 'config 参数格式错误'})

    base_url = config.get('base_url', 'https://mcenter.huaqin.com')
    origin = config.get('origin', '')
    user_id = config.get('user_id', '')
    sheet_name = config.get('sheet_name', '')
    header_row = int(config.get('header_row', 1))
    tables_cfg = config.get('tables', [])

    uid = str(uuid.uuid4())[:8]
    local_path = os.path.join(UPLOAD_DIR, f"fs_local_{uid}.xlsx")
    local_file.save(local_path)

    # Load local workbook
    wb_local = openpyxl.load_workbook(local_path, data_only=True)
    sheets = wb_local.sheetnames
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]
    ws_local = wb_local[sheet_name]
    local_header = [ws_local.cell(row=header_row, column=ci).value
                    for ci in range(1, ws_local.max_column + 1)]
    local_header_strs = [_cell_str(h) for h in local_header]

    logs = []
    prepared_tables = []
    all_fetch_cols_ordered = []
    seen_fetch_cols = set()

    for tcfg in tables_cfg:
        if not tcfg.get('enabled'):
            continue
        name = tcfg.get('name', '')
        token = tcfg.get('token', '').strip()
        active_sheet_ids = tcfg.get('active_sheet_ids', [])
        local_key_names = [k for k in tcfg.get('local_key_names', []) if k]
        feishu_key_names = [k for k in tcfg.get('feishu_key_names', []) if k]
        fetch_col_names = [c for c in tcfg.get('fetch_col_names', []) if c]

        if not token:
            logs.append(f"[{name}] 跳过：Token 为空")
            continue
        if not local_key_names or not feishu_key_names:
            logs.append(f"[{name}] 跳过：匹配键未配置")
            continue
        if len(local_key_names) != len(feishu_key_names):
            logs.append(f"[{name}] 跳过：本地键与飞书键数量不匹配")
            continue

        # Map local key names to column indices
        local_key_cols = []
        for kn in local_key_names:
            try:
                idx = local_header_strs.index(kn) + 1
                local_key_cols.append(idx)
            except ValueError:
                logs.append(f"[{name}] 跳过：本地列「{kn}」不存在")
                local_key_cols = []
                break
        if not local_key_cols:
            continue

        # Fetch feishu data
        try:
            logs.append(f"[{name}] 正在连接...")
            sheets_meta = _hq_get_sheets(base_url, origin, user_id, token)
            if not active_sheet_ids:
                active_sheet_ids = [s["sheetId"] for s in sheets_meta]
            logs.append(f"[{name}] 正在读取数据...")
            rows = _hq_read_table(base_url, origin, user_id, token, sheets_meta, active_sheet_ids)
            if not rows:
                logs.append(f"[{name}] 读取到 0 行，跳过")
                continue
            fs_headers = [_cell_str(v) for v in rows[0]]
            logs.append(f"[{name}] 读取 {len(rows)-1} 行，{len(fs_headers)} 列")
        except Exception as e:
            logs.append(f"[{name}] 读取失败：{e}")
            continue

        # Verify feishu key columns exist
        bad_keys = [k for k in feishu_key_names if k not in fs_headers]
        if bad_keys:
            logs.append(f"[{name}] 跳过：飞书列 {bad_keys} 不存在")
            continue

        # Build lookup dict
        feishu_key_indices = [fs_headers.index(k) for k in feishu_key_names]
        lookup = {}
        for row in rows[1:]:
            padded = list(row) + [""] * max(0, len(fs_headers) - len(row))
            key = tuple(_cell_str(padded[i]) if i < len(padded) else "" for i in feishu_key_indices)
            row_dict = {fs_headers[i]: _cell_str(padded[i]) if i < len(padded) else ""
                        for i in range(len(fs_headers))}
            lookup.setdefault(key, []).append(row_dict)

        for col in fetch_col_names:
            if col not in seen_fetch_cols:
                all_fetch_cols_ordered.append(col)
                seen_fetch_cols.add(col)

        prepared_tables.append({
            "name": name,
            "local_key_cols": local_key_cols,
            "lookup": lookup,
            "fetch_col_names": fetch_col_names,
        })
        logs.append(f"[{name}] 就绪，{len(lookup)} 个唯一键")

    if not prepared_tables:
        wb_local.close()
        return jsonify({'success': False, 'error': '没有可用的表格，请检查配置', 'logs': logs})

    out_name = f"飞书匹配结果_{uid}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    try:
        total, matched, unmatched = _do_match_multi(
            ws_local, header_row, prepared_tables, all_fetch_cols_ordered, out_path)
        wb_local.close()
        logs.append(f"完成：共 {total} 行，匹配 {matched}，未匹配 {unmatched}")
        return jsonify({
            'success': True,
            'download': f'/download/{out_name}',
            'total': total,
            'matched': matched,
            'unmatched': unmatched,
            'logs': logs,
        })
    except Exception as e:
        wb_local.close()
        logs.append(f"匹配出错：{e}")
        return jsonify({'success': False, 'error': str(e), 'logs': logs})


@feishu_bp.route('/api/feishu/local_sheets', methods=['POST'])
def api_feishu_local_sheets():
    """获取本地 Excel 的 Sheet 列表和列标题"""
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    path = os.path.join(UPLOAD_DIR, f"fs_pre_{uid}.xlsx")
    file.save(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    wb.close()

    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    header_row = int(request.form.get('header_row', 1))

    wb2 = openpyxl.load_workbook(path, data_only=True)
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    headers = [_cell_str(ws.cell(row=header_row, column=ci).value)
               for ci in range(1, ws.max_column + 1)]
    # Preview rows
    preview = []
    for ri in range(header_row + 1, min(header_row + 4, ws.max_row + 1)):
        row = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
        if any(row):
            preview.append(row)
    wb2.close()

    return jsonify({
        'success': True,
        'uid': uid,
        'sheets': sheets,
        'current_sheet': sheet_name,
        'headers': headers,
        'preview': preview,
    })
