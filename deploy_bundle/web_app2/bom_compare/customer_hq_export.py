# -*- coding: utf-8 -*-
"""Detailed customer BOM vs HQ BOM Excel report."""

from datetime import datetime

from shared import Workbook, Font, PatternFill, Alignment, Border, Side, get_column_letter, _cell_str, PLATFORM_VERSION, TOOL_VERSIONS
from .customer_hq import normalize_customer, normalize_hq


FIELDS = [
    ("customer_part", "\u5ba2\u6237\u6599\u53f7", "part_no", "HQ\u6599\u53f7", "\u6599\u53f7\u5dee\u5f02"),
    ("model", "\u5ba2\u6237\u578b\u53f7", "model", "HQ\u578b\u53f7", "\u578b\u53f7\u5dee\u5f02"),
    ("manufacturer_mapped", "\u5ba2\u6237\u6620\u5c04\u540e\u5382\u5bb6", "manufacturer_mapped", "HQ\u6620\u5c04\u540e\u5382\u5bb6", "\u5382\u5bb6\u5dee\u5f02"),
    ("quantity", "\u5ba2\u6237\u6570\u91cf", "quantity", "HQ\u5355\u8017", "\u6570\u91cf\u5dee\u5f02"),
    ("refdes", "\u5ba2\u6237\u4f4d\u53f7", "refdes", "HQ\u4f4d\u53f7", "\u4f4d\u53f7\u5dee\u5f02"),
    ("", "", "alternate", "HQ\u66ff\u4ee3\u5173\u7cfb", "\u66ff\u4ee3\u5173\u7cfb\u5dee\u5f02"),
]


def _norm_text(value):
    return "".join(_cell_str(value).upper().split())


def _norm_number(value):
    text = _cell_str(value)
    if not text:
        return ""
    try:
        return str(float(text)).rstrip("0").rstrip(".")
    except Exception:
        return text


def _norm_refdes(value):
    parts = [p.strip().upper() for p in _cell_str(value).replace("\uff0c", ",").replace("\u3001", ",").replace(";", ",").split(",")]
    return ",".join(sorted(p for p in parts if p))


def _split_refdes_delta(left, right):
    def split(value):
        refs = []
        seen = set()
        for part in _cell_str(value).replace("\uff0c", ",").replace("\u3001", ",").replace(";", ",").split(","):
            ref = part.strip().upper()
            if ref and ref not in seen:
                refs.append(ref)
                seen.add(ref)
        return refs
    left_refs = split(left)
    right_refs = split(right)
    right_set = set(right_refs)
    left_set = set(left_refs)
    return (
        ",".join(ref for ref in left_refs if ref not in right_set),
        ",".join(ref for ref in right_refs if ref not in left_set),
    )


def _field_equal(field_key, left, right):
    if field_key in ("quantity",):
        return _norm_number(left) == _norm_number(right)
    if field_key in ("refdes",):
        return _norm_refdes(left) == _norm_refdes(right)
    return _norm_text(left) == _norm_text(right)


def _index_rows(rows):
    indexed = {}
    duplicates = {}
    invalid = []
    for row in rows:
        if row.get("status") != "\u6709\u6548":
            invalid.append(row)
            continue
        key = row.get("match_key") or ""
        if not key:
            invalid.append(row)
            continue
        if key in indexed:
            duplicates.setdefault(key, [indexed[key]["row"]]).append(row.get("row"))
            continue
        indexed[key] = row
    return indexed, duplicates, invalid


def _changed_fields(customer, hq):
    changed = []
    for left_key, left_label, right_key, right_label, diff_label in FIELDS:
        if not left_key:
            continue
        if not _field_equal(left_key, customer.get(left_key, ""), hq.get(right_key, "")):
            changed.append((left_key, left_label, right_key, right_label, diff_label))
    return changed


def _style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="D9EAF7")
    font = Font(bold=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _excel_text_width(value):
    text = _cell_str(value)
    if not text:
        return 0
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 127 else 1
    return width


def _auto_fit_sheet(ws, min_width=10, max_width=60, base_row_height=18, max_row_height=96):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, ws.max_row + 1):
            max_len = max(max_len, _excel_text_width(ws.cell(row=row, column=col).value))
        width = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[get_column_letter(col)].width = width

    for row in range(1, ws.max_row + 1):
        visual_lines = 1
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            text = _cell_str(cell.value)
            if not text:
                continue
            col_width = ws.column_dimensions[get_column_letter(col)].width or min_width
            wrapped = max(1, int((_excel_text_width(text) + col_width - 1) // col_width))
            explicit = text.count("\n") + 1
            visual_lines = max(visual_lines, wrapped, explicit)
        ws.row_dimensions[row].height = min(max(base_row_height, visual_lines * base_row_height), max_row_height)


def _style_body(ws):
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    _auto_fit_sheet(ws)

def _quote_sheet_name(name):
    return "'" + str(name).replace("'", "''") + "'"


def _link_summary_row(ws, summary_row_by_name, name, target_ws):
    row_idx = summary_row_by_name.get(name)
    if not row_idx:
        return
    cell = ws.cell(row_idx, 1)
    cell.hyperlink = f"#{_quote_sheet_name(target_ws.title)}!A1"
    cell.style = "Hyperlink"


def _append_dict_rows(ws, rows, headers, source):
    for row in rows:
        ws.append([
            source,
            row.get("row", ""),
            row.get("match_key", ""),
            row.get("customer_part", ""),
            row.get("part_no", ""),
            row.get("model", ""),
            row.get("manufacturer_raw", ""),
            row.get("manufacturer_mapped", ""),
            row.get("quantity", ""),
            row.get("refdes", ""),
            row.get("alternate", ""),
            row.get("status", ""),
            row.get("issue", ""),
        ])


def build_report(out_path, left_path, right_path, left_sheet, right_sheet, left_header_row, mapping, match_mode, helpers, meta=None):
    customer_rows = normalize_customer(left_path, left_sheet, left_header_row, mapping, match_mode, helpers)
    hq_rows = normalize_hq(right_path, right_sheet, match_mode, helpers)
    customer_by_key, customer_dups, customer_invalid = _index_rows(customer_rows)
    hq_by_key, hq_dups, hq_invalid = _index_rows(hq_rows)

    all_keys = sorted(set(customer_by_key) | set(hq_by_key))
    matched = []
    changed = []
    same = []
    customer_only = []
    hq_only = []
    field_diffs = []
    for key in all_keys:
        c = customer_by_key.get(key)
        h = hq_by_key.get(key)
        if c and h:
            diffs = _changed_fields(c, h)
            item = (key, c, h, diffs)
            matched.append(item)
            if diffs:
                changed.append(item)
                for left_key, left_label, right_key, right_label, diff_label in diffs:
                    left_value = c.get(left_key, "")
                    right_value = h.get(right_key, "")
                    if left_key == "refdes":
                        left_value, right_value = _split_refdes_delta(left_value, right_value)
                    field_diffs.append([key, diff_label, left_label, left_value, right_label, right_value, c.get("row", ""), h.get("row", "")])
            else:
                same.append(item)
        elif c:
            customer_only.append(c)
        elif h:
            hq_only.append(h)

    wb = Workbook()
    ws = wb.active
    ws.title = "\u5dee\u5f02\u603b\u89c8"
    title_fill = PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells("A1:D1")
    ws["A1"] = "BOM Tools \u5bfc\u51fa\u62a5\u544a"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center")
    meta = meta or {}
    match_label = "\u6309\u4f4d\u53f7\u5339\u914d" if match_mode == "refdes" else "\u6309\u578b\u53f7+\u5236\u9020\u5546\u5339\u914d"
    mapping = mapping or {}
    mapping_text = "\uff1b".join(
        f"{label}: {mapping.get(key, '') or '\u672a\u6620\u5c04'}"
        for key, label in [
            ("model", "\u89c4\u683c\u578b\u53f7"),
            ("manufacturer", "\u5236\u9020\u5546"),
            ("refdes", "\u4f4d\u53f7"),
            ("quantity", "\u6570\u91cf/\u7528\u91cf"),
            ("customer_part", "\u5ba2\u6237\u6599\u53f7"),
        ]
    )
    summary = [
        ("\u62a5\u544a\u540d\u79f0", "\u5ba2\u6237 BOM \u5bf9\u6bd4 HQ BOM \u5dee\u5f02\u603b\u89c8"),
        ("\u5bfc\u51fa\u6765\u6e90", f"BOM Tools \u5e73\u53f0 v{PLATFORM_VERSION}"),
        ("\u5e73\u53f0\u7248\u672c", f"v{PLATFORM_VERSION}"),
        ("\u5de5\u5177\u540d\u79f0", "\u5ba2\u6237 BOM \u5bf9\u6bd4 HQ BOM"),
        ("\u5de5\u5177\u7248\u672c", f"v{TOOL_VERSIONS['customer-hq-compare']}"),
        ("\u5bfc\u51fa\u65f6\u95f4", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("\u5ba2\u6237 BOM \u6587\u4ef6", meta.get("left_filename", "")),
        ("HQ BOM \u6587\u4ef6", meta.get("right_filename", "")),
        ("\u5ba2\u6237 Sheet / \u8868\u5934\u884c", f"{meta.get('left_sheet') or '\u81ea\u52a8\u9009\u62e9'} / {left_header_row}"),
        ("HQ Sheet", meta.get("right_sheet") or "\u81ea\u52a8\u8bc6\u522b"),
        ("\u5339\u914d\u6a21\u5f0f", match_label),
        ("\u5b57\u6bb5\u6620\u5c04", mapping_text),
        ("\u62a5\u544a\u8bf4\u660e", "\u672c\u62a5\u544a\u7531 BOM Tools \u81ea\u52a8\u751f\u6210\uff0c\u7ed3\u679c\u4f9d\u8d56\u4e0a\u4f20\u6587\u4ef6\u5185\u5bb9\u3001\u5b57\u6bb5\u6620\u5c04\u548c\u5339\u914d\u6a21\u5f0f\u3002"),
        ("\u5ba2\u6237 BOM \u6709\u6548\u884c", len(customer_by_key)),
        ("HQ BOM \u6709\u6548\u884c", len(hq_by_key)),
        ("\u5339\u914d\u6210\u529f", len(matched)),
        ("\u5b8c\u5168\u4e00\u81f4", len(same)),
        ("\u5b57\u6bb5\u5dee\u5f02", len(changed)),
        ("\u4ec5\u5ba2\u6237\u5b58\u5728", len(customer_only)),
        ("\u4ec5 HQ \u5b58\u5728", len(hq_only)),
        ("\u5ba2\u6237\u5f02\u5e38\u884c", len(customer_invalid)),
        ("HQ \u5f02\u5e38\u884c", len(hq_invalid)),
        ("\u5ba2\u6237\u91cd\u590d\u5339\u914d\u952e", len(customer_dups)),
        ("HQ \u91cd\u590d\u5339\u914d\u952e", len(hq_dups)),
    ]
    summary_row_by_name = {}
    for idx, (name, value) in enumerate(summary, 3):
        summary_row_by_name[name] = idx
        ws.cell(idx, 1, name).font = Font(bold=True)
        ws.cell(idx, 2, value)
    _style_body(ws)

    detail = wb.create_sheet("\u5339\u914d\u660e\u7ec6")
    detail.append(["\u5339\u914d\u72b6\u6001", "\u5339\u914d\u952e", "\u5ba2\u6237\u884c\u53f7", "HQ\u884c\u53f7", "\u5dee\u5f02\u5b57\u6bb5", "\u5ba2\u6237\u6599\u53f7", "HQ\u6599\u53f7", "\u5ba2\u6237\u578b\u53f7", "HQ\u578b\u53f7", "\u5ba2\u6237\u539f\u59cb\u5382\u5bb6", "\u5ba2\u6237\u6620\u5c04\u540e\u5382\u5bb6", "HQ\u751f\u4ea7\u5382\u5bb6", "\u5ba2\u6237\u6570\u91cf", "HQ\u5355\u8017", "\u5ba2\u6237\u4f4d\u53f7", "HQ\u4f4d\u53f7", "HQ\u66ff\u4ee3\u5173\u7cfb"])
    for key, c, h, diffs in matched:
        detail.append(["\u5b57\u6bb5\u5dee\u5f02" if diffs else "\u4e00\u81f4", key, c.get("row", ""), h.get("row", ""), "\n".join(d[4] for d in diffs), c.get("customer_part", ""), h.get("part_no", ""), c.get("model", ""), h.get("model", ""), c.get("manufacturer_raw", ""), c.get("manufacturer_mapped", ""), h.get("manufacturer_raw", ""), c.get("quantity", ""), h.get("quantity", ""), c.get("refdes", ""), h.get("refdes", ""), h.get("alternate", "")])
    for c in customer_only:
        detail.append(["\u4ec5\u5ba2\u6237\u5b58\u5728", c.get("match_key", ""), c.get("row", ""), "", "", c.get("customer_part", ""), "", c.get("model", ""), "", c.get("manufacturer_raw", ""), c.get("manufacturer_mapped", ""), "", c.get("quantity", ""), "", c.get("refdes", ""), "", ""])
    for h in hq_only:
        detail.append(["\u4ec5 HQ \u5b58\u5728", h.get("match_key", ""), "", h.get("row", ""), "", "", h.get("part_no", ""), "", h.get("model", ""), "", "", h.get("manufacturer_raw", ""), "", h.get("quantity", ""), "", h.get("refdes", ""), h.get("alternate", "")])
    _style_header(detail)
    _style_body(detail)

    fd = wb.create_sheet("\u5b57\u6bb5\u5dee\u5f02")
    fd.append(["\u5339\u914d\u952e", "\u5dee\u5f02\u7c7b\u578b", "\u5ba2\u6237\u5b57\u6bb5", "\u5ba2\u6237\u503c", "HQ\u5b57\u6bb5", "HQ\u503c", "\u5ba2\u6237\u884c\u53f7", "HQ\u884c\u53f7"])
    diff_group_fills = [PatternFill("solid", fgColor="FFF9C4"), PatternFill("solid", fgColor="EAF2F8")]
    fill_by_key = {}
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for row in field_diffs:
        fd.append(row)
        key = row[0]
        if key not in fill_by_key:
            fill_by_key[key] = diff_group_fills[len(fill_by_key) % len(diff_group_fills)]
        for cell in fd[fd.max_row]:
            cell.fill = fill_by_key[key]
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    _style_header(fd)
    _style_body(fd)

    co = wb.create_sheet("\u4ec5\u5ba2\u6237\u5b58\u5728")
    co.append(["\u6765\u6e90", "\u539f\u59cb\u884c\u53f7", "\u5339\u914d\u952e", "\u5ba2\u6237\u6599\u53f7", "HQ\u6599\u53f7", "\u578b\u53f7", "\u539f\u59cb\u5382\u5bb6", "\u6620\u5c04\u540e\u5382\u5bb6", "\u6570\u91cf", "\u4f4d\u53f7", "\u66ff\u4ee3\u5173\u7cfb", "\u72b6\u6001", "\u95ee\u9898\u8bf4\u660e"])
    _append_dict_rows(co, customer_only, None, "\u5ba2\u6237")
    _style_header(co)
    _style_body(co)

    ho = wb.create_sheet("\u4ec5HQ\u5b58\u5728")
    ho.append(["\u6765\u6e90", "\u539f\u59cb\u884c\u53f7", "\u5339\u914d\u952e", "\u5ba2\u6237\u6599\u53f7", "HQ\u6599\u53f7", "\u578b\u53f7", "\u539f\u59cb\u5382\u5bb6", "\u6620\u5c04\u540e\u5382\u5bb6", "\u6570\u91cf", "\u4f4d\u53f7", "\u66ff\u4ee3\u5173\u7cfb", "\u72b6\u6001", "\u95ee\u9898\u8bf4\u660e"])
    _append_dict_rows(ho, hq_only, None, "HQ")
    _style_header(ho)
    _style_body(ho)

    bad = wb.create_sheet("\u5f02\u5e38\u884c")
    bad.append(["\u6765\u6e90", "\u539f\u59cb\u884c\u53f7", "\u5339\u914d\u952e", "\u5ba2\u6237\u6599\u53f7", "HQ\u6599\u53f7", "\u578b\u53f7", "\u539f\u59cb\u5382\u5bb6", "\u6620\u5c04\u540e\u5382\u5bb6", "\u6570\u91cf", "\u4f4d\u53f7", "\u66ff\u4ee3\u5173\u7cfb", "\u72b6\u6001", "\u95ee\u9898\u8bf4\u660e"])
    _append_dict_rows(bad, customer_invalid, None, "\u5ba2\u6237")
    _append_dict_rows(bad, hq_invalid, None, "HQ")
    _style_header(bad)
    _style_body(bad)

    dup = wb.create_sheet("\u91cd\u590d\u5339\u914d\u952e")
    dup.append(["\u6765\u6e90", "\u5339\u914d\u952e", "\u884c\u53f7"])
    for key, rows in customer_dups.items():
        dup.append(["\u5ba2\u6237", key, ", ".join(map(str, rows))])
    for key, rows in hq_dups.items():
        dup.append(["HQ", key, ", ".join(map(str, rows))])
    _style_header(dup)
    _style_body(dup)

    _link_summary_row(ws, summary_row_by_name, "\u5339\u914d\u6210\u529f", detail)
    _link_summary_row(ws, summary_row_by_name, "\u5b8c\u5168\u4e00\u81f4", detail)
    _link_summary_row(ws, summary_row_by_name, "\u5b57\u6bb5\u5dee\u5f02", fd)
    _link_summary_row(ws, summary_row_by_name, "\u4ec5\u5ba2\u6237\u5b58\u5728", co)
    _link_summary_row(ws, summary_row_by_name, "\u4ec5 HQ \u5b58\u5728", ho)
    _link_summary_row(ws, summary_row_by_name, "\u5ba2\u6237\u5f02\u5e38\u884c", bad)
    _link_summary_row(ws, summary_row_by_name, "HQ \u5f02\u5e38\u884c", bad)
    _link_summary_row(ws, summary_row_by_name, "\u5ba2\u6237\u91cd\u590d\u5339\u914d\u952e", dup)
    _link_summary_row(ws, summary_row_by_name, "HQ \u91cd\u590d\u5339\u914d\u952e", dup)

    wb.save(out_path)
    return {
        "customer_total": len(customer_rows),
        "hq_total": len(hq_rows),
        "matched": len(matched),
        "same": len(same),
        "changed": len(changed),
        "customer_only": len(customer_only),
        "hq_only": len(hq_only),
        "customer_invalid": len(customer_invalid),
        "hq_invalid": len(hq_invalid),
    }
