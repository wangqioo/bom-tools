# -*- coding: utf-8 -*-
"""Format-free BOM comparison endpoints."""

import json
import os
import re
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation

from flask import Blueprint

from shared import (
    OUTPUT_DIR,
    _cell_str,
    _open_workbook,
    _save_uploaded_excel,
    _to_int,
    PLATFORM_VERSION,
    TOOL_VERSIONS,
    get_column_letter,
    request,
    jsonify,
    Workbook,
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)

free_bom_compare_bp = Blueprint("free_bom_compare", __name__)

_NUMERIC_RE = re.compile(r"^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?$|^[+-]?0?\.\d+$")


def _headers(ws, header_row):
    result = []
    seen = {}
    for ci in range(1, ws.max_column + 1):
        value = _cell_str(ws.cell(row=header_row, column=ci).value)
        if not value:
            value = f"未命名列{get_column_letter(ci)}"
        if value in seen:
            seen[value] += 1
            value = f"{value}_{seen[value]}"
        else:
            seen[value] = 1
        result.append(value)
    return result


def _pick_sheet(wb, sheet_name):
    return sheet_name if sheet_name in wb.sheetnames else wb.sheetnames[0]


def _read_headers(path, sheet_name, header_row):
    wb = _open_workbook(path, data_only=True)
    try:
        sheet_name = _pick_sheet(wb, sheet_name)
        headers = _headers(wb[sheet_name], header_row)
        return wb.sheetnames, sheet_name, headers
    finally:
        wb.close()


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
                "row_number": ri,
                "values": [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, col_limit + 1)],
            })
        return {
            "sheets": wb.sheetnames,
            "current_sheet": sheet_name,
            "rows": rows,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "shown_columns": col_limit,
        }
    finally:
        wb.close()


def _normalize_header(value):
    return "".join(str(value or "").lower().split()).replace("_", "").replace("-", "")


def _detect_key(headers):
    normalized = {_normalize_header(h): h for h in headers if h}
    candidates = (
        "料号", "hq料号", "物料编码", "物料编号", "partnumber", "partno", "pn",
        "客户料号", "编码", "型号", "规格型号", "refdes", "reference", "位号",
    )
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    for header in headers:
        norm = _normalize_header(header)
        if any(token in norm for token in ("料号", "编码", "part", "pn", "型号", "位号", "refdes")):
            return header
    return headers[0] if headers else ""


def _detect_common_key(left_headers, right_headers):
    left_norm = {_normalize_header(h): h for h in left_headers if h}
    right_norm = {_normalize_header(h): h for h in right_headers if h}
    pairs = (
        ("料号", "料号"), ("物料编码", "物料编码"), ("partnumber", "partnumber"),
        ("pn", "pn"), ("型号", "型号"), ("规格型号", "规格型号"),
        ("位号", "位号"), ("refdes", "refdes"),
    )
    for left_key, right_key in pairs:
        left = left_norm.get(_normalize_header(left_key))
        right = right_norm.get(_normalize_header(right_key))
        if left and right:
            return left, right
    common = [h for h in left_headers if h and h in right_headers]
    if common:
        key = _detect_key(common)
        return key, key
    return _detect_key(left_headers), _detect_key(right_headers)


def _numeric(value):
    text = _cell_str(value)
    if not text or not _NUMERIC_RE.match(text):
        return None
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def _values_equal(left, right):
    left_num = _numeric(left)
    right_num = _numeric(right)
    if left_num is not None and right_num is not None:
        return left_num == right_num
    return _cell_str(left) == _cell_str(right)


def _is_refdes_header(value):
    norm = _normalize_header(value)
    return "位号" in norm or "refdes" in norm or "reference" in norm


def _split_refdes(value):
    text = _cell_str(value)
    if not text:
        return []
    text = re.sub(r"[\s,;，；、]+", ",", text)
    return [part.strip().upper() for part in text.split(",") if part.strip()]


def _field_values_equal(left_col, right_col, left_value, right_value):
    if _is_refdes_header(left_col) or _is_refdes_header(right_col):
        return set(_split_refdes(left_value)) == set(_split_refdes(right_value))
    return _values_equal(left_value, right_value)


def _field_change_cells(left_col, right_col, left_value, right_value):
    label = _pair_label(left_col, right_col)
    if not (_is_refdes_header(left_col) or _is_refdes_header(right_col)):
        return [(label, f"{left_value} -> {right_value}")]
    left_refs = _split_refdes(left_value)
    right_refs = _split_refdes(right_value)
    right_set = set(right_refs)
    left_set = set(left_refs)
    removed = "、".join(ref for ref in left_refs if ref not in right_set)
    added = "、".join(ref for ref in right_refs if ref not in left_set)
    return [(f"{label} 删除位号", removed), (f"{label} 新增位号", added)]


def _load_rows(path, sheet_name, header_row, key_col, compare_cols):
    wb = _open_workbook(path, data_only=True)
    try:
        sheet_name = _pick_sheet(wb, sheet_name)
        ws = wb[sheet_name]
        headers = _headers(ws, header_row)
        if key_col not in headers:
            raise ValueError(f'匹配键列 "{key_col}" 不存在')
        key_idx = headers.index(key_col) + 1
        compare_indices = [(col, headers.index(col) + 1) for col in compare_cols if col in headers]
        rows = {}
        duplicates = {}
        blank_keys = 0
        for ri in range(header_row + 1, ws.max_row + 1):
            raw = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
            if not any(raw):
                continue
            key = _cell_str(ws.cell(row=ri, column=key_idx).value)
            if not key:
                blank_keys += 1
                continue
            values = {col: _cell_str(ws.cell(row=ri, column=idx).value) for col, idx in compare_indices}
            all_values = {header: raw[i] if i < len(raw) else "" for i, header in enumerate(headers)}
            if key in rows:
                duplicates.setdefault(key, [rows[key]["row"]]).append(ri)
                continue
            rows[key] = {"row": ri, "values": values, "all_values": all_values, "headers": headers}
        return rows, duplicates, blank_keys
    finally:
        wb.close()


def _field_pairs(config_pairs, left_headers, right_headers):
    pairs = []
    seen = set()
    for pair in config_pairs or []:
        left = str((pair or {}).get("left") or "").strip()
        right = str((pair or {}).get("right") or "").strip()
        if not left or not right or left not in left_headers or right not in right_headers:
            continue
        key = (left, right)
        if key not in seen:
            pairs.append(key)
            seen.add(key)
    if pairs:
        return pairs
    return [(h, h) for h in left_headers if h and h in right_headers]


def _safe_title(value):
    text = str(value or "Sheet")[:31]
    for ch in r"\/*?:[]":
        text = text.replace(ch, "_")
    return text or "Sheet"


def _excel_text_width(value):
    text = _cell_str(value)
    if not text:
        return 0
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 127 else 1
    return width


def _auto_fit_columns(sheet, min_width=10, max_width=60):
    for ci in range(1, sheet.max_column + 1):
        max_len = 0
        for row in range(1, sheet.max_row + 1):
            max_len = max(max_len, _excel_text_width(sheet.cell(row=row, column=ci).value))
        width = min(max(max_len + 2, min_width), max_width)
        sheet.column_dimensions[get_column_letter(ci)].width = width


def _pair_label(left_col, right_col):
    return left_col if left_col == right_col else f"{left_col} <-> {right_col}"


def _write_export_info(ws, title, meta, title_fill, title_font, header_fill, border, left_align):
    meta = meta or {}
    ws.merge_cells("A1:D1")
    ws["A1"] = "BOM Tools 导出报告"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    rows = [
        ("报告名称", title),
        ("导出来源", f"BOM Tools 平台 v{PLATFORM_VERSION}"),
        ("平台版本", f"v{PLATFORM_VERSION}"),
        ("工具名称", "通用 BOM 对比"),
        ("工具版本", f"v{TOOL_VERSIONS['free-bom-compare']}"),
        ("导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("基准 BOM 文件", meta.get("left_filename", "")),
        ("对比 BOM 文件", meta.get("right_filename", "")),
        ("基准 Sheet / 表头行", f"{meta.get('left_sheet', '')} / {meta.get('left_header_row', '')}"),
        ("对比 Sheet / 表头行", f"{meta.get('right_sheet', '')} / {meta.get('right_header_row', '')}"),
        ("匹配键", f"{meta.get('left_key_col', '')} <-> {meta.get('right_key_col', '')}"),
        ("比对字段", meta.get("field_pairs", "")),
        ("报告说明", "本报告由 BOM Tools 自动生成，结果依赖上传文件内容和用户选择的匹配键及比对字段。"),
    ]
    for offset, (name, value) in enumerate(rows, 2):
        key = ws.cell(row=offset, column=1, value=name)
        val = ws.cell(row=offset, column=2, value=value)
        key.font = Font(bold=True)
        key.fill = header_fill
        key.border = border
        val.border = border
        val.alignment = left_align
    return len(rows) + 3


def _write_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, meta=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "差异总览"
    bdr = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(bold=True, color="FFFFFF", size=14)
    fills = {
        "新增物料": PatternFill("solid", fgColor="E8F5E9"),
        "删除物料": PatternFill("solid", fgColor="FFEBEE"),
        "变更物料": PatternFill("solid", fgColor="FFF9C4"),
    }
    change_group_fills = [PatternFill("solid", fgColor="FFF9C4"), PatternFill("solid", fgColor="EAF2F8")]
    center = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    summary_start_row = _write_export_info(
        ws,
        "通用 BOM 对比差异总览",
        meta,
        title_fill,
        title_font,
        header_fill,
        bdr,
        left_align,
    )
    ws.cell(row=summary_start_row, column=1, value="指标")
    ws.cell(row=summary_start_row, column=2, value="数量")
    summary = [
        ("基准BOM唯一键数量", stats["left_total"]),
        ("对比BOM唯一键数量", stats["right_total"]),
        ("新增物料", stats["right_only"]),
        ("删除物料", stats["left_only"]),
        ("变更物料", stats["changed"]),
        ("完全一致", stats["same"]),
        ("基准BOM重复键", stats["left_duplicates"]),
        ("对比BOM重复键", stats["right_duplicates"]),
        ("基准BOM空匹配键行", stats["left_blank_keys"]),
        ("对比BOM空匹配键行", stats["right_blank_keys"]),
    ]
    summary_row_by_name = {}
    for ri, (name, value) in enumerate(summary, summary_start_row + 1):
        ws.cell(row=ri, column=1, value=name)
        ws.cell(row=ri, column=2, value=value)
        summary_row_by_name[name] = ri

    all_keys = sorted(set(left_rows) | set(right_rows))
    added_keys = []
    removed_keys = []
    changed_items = []
    changed_labels = []
    changed_label_set = set()
    for key in all_keys:
        left = left_rows.get(key)
        right = right_rows.get(key)
        if left and not right:
            removed_keys.append(key)
            continue
        if right and not left:
            added_keys.append(key)
            continue
        changes = {}
        changed_field_count = 0
        for left_col, right_col in field_pairs:
            left_value = left["values"].get(left_col, "")
            right_value = right["values"].get(right_col, "")
            if not _field_values_equal(left_col, right_col, left_value, right_value):
                changed_field_count += 1
                for label, value in _field_change_cells(left_col, right_col, left_value, right_value):
                    changes[label] = value
                    if label not in changed_label_set:
                        changed_labels.append(label)
                        changed_label_set.add(label)
        if changes:
            changed_items.append([key, left["row"], right["row"], changed_field_count, changes])

    def source_headers(keys, rows):
        result = []
        seen = set()
        for key in keys:
            row = rows.get(key) or {}
            values = row.get("all_values") or {}
            for header in row.get("headers") or values.keys():
                if header in values and header not in seen:
                    result.append(header)
                    seen.add(header)
        return result

    def write_source_sheet(sheet, keys, rows, diff_type, side_label):
        source_cols = source_headers(keys, rows)
        headers = ["差异类型", "匹配键", f"{side_label}行号"] + source_cols
        sheet.append(headers)
        for key in keys:
            row = rows[key]
            values = row.get("all_values") or {}
            sheet.append([diff_type, key, row.get("row", "")] + [values.get(col, "") for col in source_cols])

    added_sheet = wb.create_sheet("新增物料")
    write_source_sheet(added_sheet, added_keys, right_rows, "新增物料", "对比BOM")

    removed_sheet = wb.create_sheet("删除物料")
    write_source_sheet(removed_sheet, removed_keys, left_rows, "删除物料", "基准BOM")

    changed_sheet = wb.create_sheet("变更物料")
    changed_headers = ["匹配键", "基准BOM行号", "对比BOM行号", "变更字段数"] + changed_labels
    changed_sheet.append(changed_headers)
    for key, left_row, right_row, changed_field_count, changes in changed_items:
        changed_sheet.append([key, left_row, right_row, changed_field_count] + [changes.get(label, "") for label in changed_labels])

    dup = wb.create_sheet("重复和空键")
    dup.append(["类型", "说明"])
    for note in duplicate_notes:
        dup.append(note)

    for name, sheet in [
        ("新增物料", added_sheet),
        ("删除物料", removed_sheet),
        ("变更物料", changed_sheet),
        ("基准BOM重复键", dup),
        ("对比BOM重复键", dup),
        ("基准BOM空匹配键行", dup),
        ("对比BOM空匹配键行", dup),
    ]:
        row_idx = summary_row_by_name.get(name)
        if row_idx:
            cell = ws.cell(row=row_idx, column=2)
            quoted = sheet.title.replace("'", "''")
            cell.hyperlink = f"#'{quoted}'!A1"
            cell.style = "Hyperlink"

    for sheet in wb.worksheets:
        fill_by_key = {}
        for row in sheet.iter_rows():
            row_fill = fills.get(row[0].value)
            if sheet.title == "变更物料" and row[0].row > 1:
                key = row[0].value
                if key not in fill_by_key:
                    fill_by_key[key] = change_group_fills[len(fill_by_key) % len(change_group_fills)]
                row_fill = fill_by_key[key]
            for cell in row:
                cell.border = bdr
                cell.alignment = center if cell.column <= 3 else left_align
                is_summary_header = sheet.title == "差异总览" and cell.row == summary_start_row
                is_table_header = sheet.title != "差异总览" and cell.row == 1
                if is_summary_header or is_table_header:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                elif row_fill:
                    cell.fill = row_fill
        _auto_fit_columns(sheet, min_width=10, max_width=60)
        if sheet.max_row >= 1:
            sheet.freeze_panes = None if sheet.title == "差异总览" else "A2"
            sheet.auto_filter.ref = f"A{summary_start_row}:B{sheet.max_row}" if sheet.title == "差异总览" else sheet.dimensions
        sheet.title = _safe_title(sheet.title)
    wb.save(out_path)


@free_bom_compare_bp.route("/api/bom_compare/free_preview", methods=["POST"])
def free_preview():
    uid = str(uuid.uuid4())[:8]
    result = {"success": True}
    try:
        left_file = request.files.get("left_file")
        right_file = request.files.get("right_file")
        if left_file:
            left_path = _save_uploaded_excel(left_file, "bomcmp_free_preview_left", uid)
            result["left"] = _preview_rows(left_path, request.form.get("left_sheet", ""))
        if right_file:
            right_path = _save_uploaded_excel(right_file, "bomcmp_free_preview_right", uid)
            result["right"] = _preview_rows(right_path, request.form.get("right_sheet", ""))
        if not left_file and not right_file:
            return jsonify({"success": False, "error": "\u8bf7\u5148\u4e0a\u4f20 BOM \u6587\u4ef6"})
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})
    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)})


@free_bom_compare_bp.route("/api/bom_compare/free_sheets", methods=["POST"])
def free_sheets():
    left_file = request.files.get("left_file")
    right_file = request.files.get("right_file")
    if not left_file or not right_file:
        return jsonify({"success": False, "error": "请上传两份 BOM 文件"})
    left_header_row = _to_int(request.form.get("left_header_row", 1), 1)
    right_header_row = _to_int(request.form.get("right_header_row", 1), 1)
    if left_header_row is None or right_header_row is None:
        return jsonify({"success": False, "error": "表头行必须是大于等于 1 的数字"})
    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_excel(left_file, "bomcmp_free_left", uid)
        right_path = _save_uploaded_excel(right_file, "bomcmp_free_right", uid)
        left_sheets, left_sheet, left_headers = _read_headers(left_path, request.form.get("left_sheet", ""), left_header_row)
        right_sheets, right_sheet, right_headers = _read_headers(right_path, request.form.get("right_sheet", ""), right_header_row)
        left_key, right_key = _detect_common_key(left_headers, right_headers)
        return jsonify({
            "success": True,
            "left_sheets": left_sheets,
            "left_current_sheet": left_sheet,
            "left_headers": left_headers,
            "left_header_row": left_header_row,
            "right_sheets": right_sheets,
            "right_current_sheet": right_sheet,
            "right_headers": right_headers,
            "right_header_row": right_header_row,
            "detected_left_key": left_key,
            "detected_right_key": right_key,
            "left_format": "generic",
            "right_format": "generic",
        })
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})
    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)})


@free_bom_compare_bp.route("/api/bom_compare/free", methods=["POST"])
def free_compare():
    left_file = request.files.get("left_file")
    right_file = request.files.get("right_file")
    if not left_file or not right_file:
        return jsonify({"success": False, "error": "请上传两份 BOM 文件"})
    try:
        config = json.loads(request.form.get("config", "{}"))
    except Exception:
        return jsonify({"success": False, "error": "config 参数格式错误"})

    left_header_row = _to_int(config.get("left_header_row", 1), 1)
    right_header_row = _to_int(config.get("right_header_row", 1), 1)
    left_key_col = str(config.get("left_key_col") or "").strip()
    right_key_col = str(config.get("right_key_col") or "").strip()
    if left_header_row is None or right_header_row is None:
        return jsonify({"success": False, "error": "表头行必须是大于等于 1 的数字"})
    if not left_key_col or not right_key_col:
        return jsonify({"success": False, "error": "请先选择两份 BOM 的匹配键"})

    uid = str(uuid.uuid4())[:8]
    try:
        left_path = _save_uploaded_excel(left_file, "bomcmp_free_left", uid)
        right_path = _save_uploaded_excel(right_file, "bomcmp_free_right", uid)
        _, left_sheet, left_headers = _read_headers(left_path, config.get("left_sheet", ""), left_header_row)
        _, right_sheet, right_headers = _read_headers(right_path, config.get("right_sheet", ""), right_header_row)
        field_pairs = _field_pairs(config.get("field_pairs", []), left_headers, right_headers)
        field_pairs = [(l, r) for l, r in field_pairs if (l, r) != (left_key_col, right_key_col)]
        if not field_pairs:
            return jsonify({"success": False, "error": "请至少选择一组需要比对的字段"})
        left_compare_cols = [l for l, _ in field_pairs]
        right_compare_cols = [r for _, r in field_pairs]
        left_rows, left_dups, left_blank = _load_rows(left_path, config.get("left_sheet", ""), left_header_row, left_key_col, left_compare_cols)
        right_rows, right_dups, right_blank = _load_rows(right_path, config.get("right_sheet", ""), right_header_row, right_key_col, right_compare_cols)
        common = sorted(set(left_rows) & set(right_rows))
        changed = [
            key for key in common
            if any(
                not _field_values_equal(l, r, left_rows[key]["values"].get(l, ""), right_rows[key]["values"].get(r, ""))
                for l, r in field_pairs
            )
        ]
        stats = {
            "left_total": len(left_rows),
            "right_total": len(right_rows),
            "left_only": len(set(left_rows) - set(right_rows)),
            "right_only": len(set(right_rows) - set(left_rows)),
            "changed": len(changed),
            "same": len(common) - len(changed),
            "left_duplicates": len(left_dups),
            "right_duplicates": len(right_dups),
            "left_blank_keys": left_blank,
            "right_blank_keys": right_blank,
        }
        duplicate_notes = (
            [["基准BOM", f"重复键 {key}: 行 {', '.join(map(str, rows))}"] for key, rows in left_dups.items()] +
            [["对比BOM", f"重复键 {key}: 行 {', '.join(map(str, rows))}"] for key, rows in right_dups.items()] +
            ([["基准BOM", f"空匹配键行数：{left_blank}"]] if left_blank else []) +
            ([["对比BOM", f"空匹配键行数：{right_blank}"]] if right_blank else [])
        )
        out_name = f"Generic_BOM_Compare_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        _write_report(out_path, left_rows, right_rows, field_pairs, stats, duplicate_notes, meta={
            "left_filename": left_file.filename or "",
            "right_filename": right_file.filename or "",
            "left_sheet": left_sheet,
            "right_sheet": right_sheet,
            "left_header_row": left_header_row,
            "right_header_row": right_header_row,
            "left_key_col": left_key_col,
            "right_key_col": right_key_col,
            "field_pairs": "；".join(_pair_label(l, r) for l, r in field_pairs),
        })
        return jsonify({"success": True, "download": f"/download/{out_name}", **stats})
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})
    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)})
