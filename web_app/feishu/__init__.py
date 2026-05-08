# -*- coding: utf-8 -*-
"""飞书多表格匹配工具 — Blueprint（参照 feishu_multi_matcher.py v3.0）"""

import os, uuid, json, traceback

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter, column_index_from_string,
    render_template, request, jsonify, requests,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str, _unique_path,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint, send_file

feishu_bp = Blueprint('feishu_tool', __name__)

# ─────────────────────── 飞书 API（与 feishu_multi_matcher.py 一致）─

def hq_get_sheets(base_url, origin, user_id, token):
    url = f"{base_url}/fs/sheet/v1/spreadsheetsMetainfo"
    r = requests.get(url, params={
        "origin": origin, "userId": user_id, "spreadsheetToken": token,
    }, timeout=15)
    r.raise_for_status()
    d = r.json()
    if d.get("code") not in (0, 200):
        raise RuntimeError(f"获取 Sheet 列表失败：{d.get('msg')}（code={d.get('code')}）")
    return [s for s in d["data"]["sheets"] if s.get("title")]


def hq_read_sheet(base_url, origin, user_id, token,
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
        r = requests.get(f"{base_url}/fs/sheet/v1/getSheetsValue",
                         params=params, timeout=60)
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


def hq_read_table(base_url, origin, user_id, token,
                  sheets_meta, active_ids, progress_cb=None):
    """读取并合并多个 Sheet 的数据。progress_cb(sheet_title, n_rows)。"""
    combined, headers_set = [], False
    for s in sheets_meta:
        if s["sheetId"] not in active_ids:
            continue
        title = s.get("title", s["sheetId"])

        def _pcb(n, _t=title):
            if progress_cb:
                progress_cb(_t, n)

        rows = hq_read_sheet(base_url, origin, user_id, token,
                             s["sheetId"],
                             row_count=s.get("rowCount", 200000),
                             col_count=s.get("columnCount", 100),
                             batch_size=3000, progress_cb=_pcb)
        if not rows:
            continue
        if not headers_set:
            combined = rows
            headers_set = True
        else:
            combined.extend(rows[1:])
    return combined


# ─────────────────────── 匹配核心（与 feishu_multi_matcher.py 一致）─

def do_match_multi(local_ws, local_header_row,
                   prepared_tables, all_fetch_cols,
                   out_file, log_cb):
    """
    多表格匹配：每行本地数据依次尝试各表格，首个命中为准。
    prepared_tables: [{name, local_key_cols, lookup, fetch_col_names}, ...]
    all_fetch_cols:  所有表格提取列名的有序并集
    """
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
        c.fill = (PatternFill("solid", fgColor="FFC000")
                  if ci > max_local_col else hdr_fill)
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
                c = ws_out.cell(row=dr, column=src_col,
                                value=pt["name"] if first else "")
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
            c = ws_out.cell(row=dr, column=max_local_col + len(all_fetch_cols) + 1,
                            value="未匹配")
            c.border = bdr
            unmatched += 1
            dr += 1

    wb_out.save(out_file)
    total = dr - 2
    log_cb(f"写入 {total} 行：{matched} 个匹配成功，{unmatched} 个未匹配")
    return total, matched, unmatched


# ─────────────────────── 构建 prepared_tables（从多表配置）─────

def build_prepared_tables(enabled_tables, gateway, log_cb, local_headers):
    """
    连接、读取、构建 prepared_tables。
    enabled_tables: [{name, token, category, active_sheet_ids,
                      local_key_names, feishu_key_names, fetch_col_names}, ...]
    返回: (prepared_tables, all_fetch_cols)
    """
    base_url, origin, user_id = gateway

    for t in enabled_tables:
        log_cb(f"[{t['name']}] 连接中...")
        try:
            sheets = hq_get_sheets(base_url, origin, user_id, t["token"])
        except Exception as e:
            log_cb(f"[{t['name']}] 连接失败：{e}，已跳过")
            t["_error"] = str(e)
            continue

        online_ids = {s["sheetId"] for s in sheets}
        saved_ids = t.get("active_sheet_ids", [])
        if not saved_ids:
            active_ids = list(online_ids)
            t["active_sheet_ids"] = active_ids
            log_cb(f"[{t['name']}] 首次加载，自动选择全部 {len(active_ids)} 个 Sheet")
        else:
            stale_ids = [sid for sid in saved_ids if sid not in online_ids]
            active_ids = [sid for sid in saved_ids if sid in online_ids]
            t["active_sheet_ids"] = active_ids
            if stale_ids:
                stale_names = "、".join(
                    s.get("title", s["sheetId"])
                    for s in sheets if s["sheetId"] in stale_ids)
                log_cb(f"[{t['name']}] ⚠ Sheet 已失效并自动剔除：{stale_names}")
            if not active_ids:
                log_cb(f"[{t['name']}] ✗ 所有 Sheet 均已失效，已跳过")
                continue

        log_cb(f"[{t['name']}] 读取数据中...")
        try:
            rows = hq_read_table(base_url, origin, user_id, t["token"],
                                 sheets, active_ids, log_cb)
            t["_rows"] = rows
            t["_headers"] = [_cell_str(v) for v in rows[0]] if rows else []
            t["_loaded"] = True
            cnt = max(len(rows) - 1, 0)
            log_cb(f"[{t['name']}] ✓ 加载完成，{cnt} 行")
        except Exception as e:
            log_cb(f"[{t['name']}] 读取失败：{e}，已跳过")
            continue

    # 构建 prepared_tables
    prepared_tables = []
    for t in enabled_tables:
        if not t.get("_loaded"):
            continue
        tname = t["name"]
        fs_header_set = set(t["_headers"])
        local_key_cols, feishu_key_cols = [], []
        for lk, fk in zip(t.get("local_key_names", []), t.get("feishu_key_names", [])):
            if not lk or not fk:
                continue
            if fk not in fs_header_set:
                log_cb(f"[{tname}] ⚠ 飞书匹配键「{fk}」已不存在，此键对跳过")
                continue
            try:
                lc = next(ci + 1 for ci, h in enumerate(local_headers)
                          if _cell_str(h) == lk)
            except StopIteration:
                log_cb(f"[{tname}] ⚠ 本地匹配键「{lk}」在表头中未找到，此键对跳过")
                continue
            fc = t["_headers"].index(fk)
            local_key_cols.append(lc)
            feishu_key_cols.append(fc)
        if not local_key_cols:
            log_cb(f"[{tname}] ✗ 无有效匹配键，跳过")
            continue
        fetch_cols = t.get("fetch_col_names", [])
        if not fetch_cols:
            log_cb(f"[{tname}] ✗ 未选择提取列，跳过")
            continue

        fetch_idxs = []
        stale_fetch = []
        for col_name in fetch_cols:
            if col_name in fs_header_set:
                fetch_idxs.append(t["_headers"].index(col_name))
            else:
                fetch_idxs.append(-1)
                stale_fetch.append(col_name)
        if stale_fetch:
            log_cb(f"[{tname}] ⚠ 提取列已失效（输出空白）：{' / '.join(stale_fetch)}")

        lookup = {}
        for row in t["_rows"][1:]:
            key = tuple(_cell_str(row[fc]) if fc < len(row) else ""
                        for fc in feishu_key_cols)
            if not any(key):
                continue
            vals = {col_name: (_cell_str(row[idx]) if 0 <= idx < len(row) else "")
                    for col_name, idx in zip(fetch_cols, fetch_idxs)}
            lookup.setdefault(key, []).append(vals)

        prepared_tables.append({
            "name": tname,
            "local_key_cols": local_key_cols,
            "lookup": lookup,
            "fetch_col_names": fetch_cols,
        })

    seen, all_fetch_cols = set(), []
    for pt in prepared_tables:
        for col_name in pt["fetch_col_names"]:
            if col_name not in seen:
                seen.add(col_name)
                all_fetch_cols.append(col_name)

    return prepared_tables, all_fetch_cols


# ─────────────────────── 路由 ─────────────────────────────

@feishu_bp.route('/feishu', methods=['GET', 'POST'])
def tool_feishu():
    if request.method == 'POST':
        action = request.form.get('action', 'match')

        # ── 获取 Sheet 列表 ─────────────────────────────
        if action == 'load_sheets':
            token = request.form.get('token', '')
            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            try:
                sheets = hq_get_sheets(base_url, origin, user_id, token)
                return jsonify({
                    'success': True,
                    'sheets': [{'sheetId': s['sheetId'], 'title': s['title']}
                               for s in sheets],
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

        # ── 读取选中 Sheet 的数据（仅返回列名+行数）─────
        elif action == 'load_data':
            token = request.form.get('token', '')
            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            sheet_ids = json.loads(request.form.get('sheet_ids', '[]'))
            try:
                sheets = hq_get_sheets(base_url, origin, user_id, token)
                rows = []
                for s in sheets:
                    if s['sheetId'] not in sheet_ids:
                        continue
                    batch = hq_read_sheet(
                        base_url, origin, user_id, token, s['sheetId'],
                        row_count=s.get('rowCount', 200000),
                        col_count=s.get('columnCount', 100),
                    )
                    if batch:
                        if not rows:
                            rows = batch
                        else:
                            rows.extend(batch[1:])
                headers = [_cell_str(v) for v in (rows[0] if rows else [])]
                return jsonify({
                    'success': True,
                    'headers': headers,
                    'row_count': max(len(rows) - 1, 0),
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

        # ── 执行匹配 ──────────────────────────────────
        elif action == 'match':
            file = request.files.get('file')
            if not file:
                return "请上传文件", 400

            base_url = request.form.get('base_url', 'https://mcenter.huaqin.com')
            origin = request.form.get('origin', 'cli_a96ac38049f8d0e5')
            user_id = request.form.get('user_id', '100448405')
            header_row = int(request.form.get('header_row', 1))
            sheet_name = request.form.get('sheet_name', '')
            tables_json = request.form.get('tables', '[]')

            try:
                tables = json.loads(tables_json)
            except Exception:
                return "表格配置格式错误", 400

            if not tables:
                return jsonify({'success': False, 'error': '请至少启用一个表格'})

            uid = str(uuid.uuid4())[:8]
            in_path = os.path.join(UPLOAD_DIR, f"feishu_in_{uid}.xlsx")
            out_path = os.path.join(OUTPUT_DIR, f"飞书匹配结果_{uid}.xlsx")
            file.save(in_path)

            wb = openpyxl.load_workbook(in_path, data_only=True)
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb[wb.sheetnames[0]]
            local_headers = [_cell_str(ws.cell(row=header_row, column=ci).value)
                             for ci in range(1, ws.max_column + 1)]

            logs = []

            def log_cb(msg):
                logs.append(msg)
                print(msg)

            gateway = (base_url, origin, user_id)
            try:
                prepared_tables, all_fetch_cols = build_prepared_tables(
                    tables, gateway, log_cb, local_headers)
                if not prepared_tables:
                    return jsonify({
                        'success': False,
                        'error': '没有可用的已启用表格（请检查列映射配置）',
                        'logs': logs,
                    })
                total, matched, unmatched = do_match_multi(
                    ws, header_row, prepared_tables, all_fetch_cols, out_path, log_cb)
                return jsonify({
                    'success': True,
                    'total': total,
                    'matched': matched,
                    'unmatched': unmatched,
                    'logs': logs,
                    'download': f'/download/飞书匹配结果_{uid}.xlsx',
                })
            except Exception as e:
                logs.append(f"\n❌ 错误：{e}")
                logs.append(traceback.format_exc())
                return jsonify({'success': False, 'error': str(e), 'logs': logs})

    return render_template('feishu.html', tables=FEISHU_PRESET_TABLES)
