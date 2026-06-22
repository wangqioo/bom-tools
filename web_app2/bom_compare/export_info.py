# -*- coding: utf-8 -*-
"""Shared Excel export information block for BOM compare reports."""

from datetime import datetime

from shared import Alignment, Font, PatternFill, PLATFORM_VERSION, TOOL_VERSIONS


def write_export_info(
    ws,
    title,
    tool_name,
    tool_version_key,
    rows=None,
    note=None,
    title_fill=None,
    title_font=None,
    header_fill=None,
    border=None,
    value_alignment=None,
):
    title_fill = title_fill or PatternFill("solid", fgColor="1F4E78")
    title_font = title_font or Font(bold=True, color="FFFFFF", size=14)
    header_fill = header_fill or PatternFill("solid", fgColor="D9EAF7")
    value_alignment = value_alignment or Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("A1:D1")
    ws["A1"] = "BOM Tools 导出报告"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    info_rows = [
        ("报告名称", title),
        ("导出来源", f"BOM Tools 平台 v{PLATFORM_VERSION}"),
        ("平台版本", f"v{PLATFORM_VERSION}"),
        ("工具名称", tool_name),
        ("工具版本", f"v{TOOL_VERSIONS.get(tool_version_key, '')}"),
        ("导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    info_rows.extend(rows or [])
    if note:
        info_rows.append(("报告说明", note))

    for offset, (name, value) in enumerate(info_rows, 2):
        key = ws.cell(row=offset, column=1, value=name)
        val = ws.cell(row=offset, column=2, value=value)
        key.font = Font(bold=True)
        key.fill = header_fill
        key.border = border
        val.border = border
        val.alignment = value_alignment

    return len(info_rows) + 3
