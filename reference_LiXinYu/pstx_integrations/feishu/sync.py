# -*- coding: utf-8 -*-
"""Online preview and sync orchestration for Feishu BOM libraries."""

from __future__ import annotations

import json
from datetime import datetime
from typing import List, Optional, Sequence

from pstx_integrations.diagnostics import new_diagnostic_request_id
from pstx_integrations.feishu.cache_store import (
    _connect_cache_for_write,
    _library_id_from_name,
    _upsert_library_config,
    get_saved_feishu_field_order,
)
from pstx_integrations.feishu.client import FeishuBomClient
from pstx_integrations.feishu.common import (
    FeishuBomError,
    _column_index_to_name,
    _cache_path,
    _config_path,
    _duplicate_headers_debug,
    _feishu_parse_log_row_limit,
    _field_indexes_debug,
    _headers_position_debug,
    _log_feishu_event,
    _log_feishu_parse_event,
    _resolve_connection,
    _row_cells_debug,
    _row_field_values_debug,
    _row_non_empty_positions,
    _rows_preview,
    _rows_shape,
    _safe_cell_str,
    _save_config,
    feishu_debug_log_path,
    feishu_parse_debug_log_path,
    extract_spreadsheet_token,
)
from pstx_integrations.feishu.mapping import (
    _align_data_row_to_headers,
    _header_index,
    _normalize_optional_fields,
    _normalize_sheet_cfgs,
    _to_text_rows,
    suggest_feishu_mapping_from_preview,
)


def fetch_feishu_sheet_list(*,
                            spreadsheet_token_or_url: str,
                            base_url: str = "",
                            origin: str = "",
                            user_id: str = "",
                            data_dir: str = "",
                            client: Optional[FeishuBomClient] = None) -> dict:
    operation_id = new_diagnostic_request_id()
    token = extract_spreadsheet_token(spreadsheet_token_or_url)
    if not token:
        raise FeishuBomError("缺少 spreadsheet_token_or_url。")
    resolved_base_url, resolved_origin, resolved_user_id, root, _config = _resolve_connection(
        base_url=base_url,
        origin=origin,
        user_id=user_id,
        data_dir=data_dir,
    )
    _log_feishu_event(
        "sheet_list.start",
        {
            "spreadsheet_token": token,
            "base_url": resolved_base_url,
            "origin": resolved_origin,
            "user_id": resolved_user_id,
            "data_dir": str(root),
        },
        request_id=operation_id,
    )
    api = client or FeishuBomClient(resolved_base_url, resolved_origin, resolved_user_id)
    sheets = api.get_sheets(token)
    _log_feishu_event(
        "sheet_list.success",
        {
            "spreadsheet_token": token,
            "sheet_count": len(sheets),
            "sheets": [
                {
                    "sheet_id": sheet.get("sheet_id"),
                    "title": sheet.get("title"),
                    "row_count": sheet.get("row_count"),
                    "column_count": sheet.get("column_count"),
                    "column_range": sheet.get("column_range"),
                }
                for sheet in sheets[:40]
            ],
        },
        request_id=operation_id,
    )
    return {
        "ok": True,
        "data_dir": str(root),
        "online_debug_log_file": feishu_debug_log_path(),
        "online_parse_debug_log_file": feishu_parse_debug_log_path(),
        "spreadsheet_token": token,
        "sheet_count": len(sheets),
        "sheets": sheets,
    }


def preview_feishu_sheet(*,
                         spreadsheet_token_or_url: str,
                         sheet_id: str,
                         base_url: str = "",
                         origin: str = "",
                         user_id: str = "",
                         data_dir: str = "",
                         row_count: int = 50,
                         column_range: str = "A:Z",
                         header_row: int = 1,
                         client: Optional[FeishuBomClient] = None) -> dict:
    operation_id = new_diagnostic_request_id()
    token = extract_spreadsheet_token(spreadsheet_token_or_url)
    if not token:
        raise FeishuBomError("缺少 spreadsheet_token_or_url。")
    if not _safe_cell_str(sheet_id):
        raise FeishuBomError("缺少 sheet_id。")
    try:
        safe_row_count = min(max(int(row_count or 50), 1), 200)
        safe_header_row = max(1, int(header_row or 1))
    except (TypeError, ValueError) as exc:
        raise FeishuBomError("row_count/header_row 必须是数字。") from exc
    resolved_base_url, resolved_origin, resolved_user_id, root, _config = _resolve_connection(
        base_url=base_url,
        origin=origin,
        user_id=user_id,
        data_dir=data_dir,
    )
    api = client or FeishuBomClient(resolved_base_url, resolved_origin, resolved_user_id)
    _log_feishu_event(
        "preview.start",
        {
            "spreadsheet_token": token,
            "sheet_id": _safe_cell_str(sheet_id),
            "row_count": safe_row_count,
            "column_range": column_range,
            "header_row": safe_header_row,
            "base_url": resolved_base_url,
            "origin": resolved_origin,
            "user_id": resolved_user_id,
            "data_dir": str(root),
        },
        request_id=operation_id,
    )
    values = api.read_sheet(
        token,
        _safe_cell_str(sheet_id),
        row_count=max(safe_row_count, 50),
        column_range=column_range,
    )
    rows = _to_text_rows(values)[:safe_row_count]
    header_values = rows[safe_header_row - 1] if len(rows) >= safe_header_row else []
    saved_order = get_saved_feishu_field_order(data_dir=str(root)).get("optional_field_order", [])
    mapping_suggestion = suggest_feishu_mapping_from_preview(
        rows,
        saved_optional_order=saved_order,
        header_row=safe_header_row,
    )
    _log_feishu_event(
        "preview.success",
        {
            "spreadsheet_token": token,
            "sheet_id": _safe_cell_str(sheet_id),
            "row_shape": _rows_shape(rows),
            "row_preview": _rows_preview(rows),
            "header_row": safe_header_row,
            "headers": [header for header in header_values if header],
            "mapping": mapping_suggestion.get("mapping", {}),
            "confidence": mapping_suggestion.get("confidence"),
            "notes": mapping_suggestion.get("notes", []),
        },
        request_id=operation_id,
    )
    _log_feishu_parse_event(
        "preview.mapping_diagnostics",
        {
            "spreadsheet_token": token,
            "sheet_id": _safe_cell_str(sheet_id),
            "row_shape": _rows_shape(rows),
            "header_row": safe_header_row,
            "header_positions": _headers_position_debug(header_values),
            "duplicate_headers": _duplicate_headers_debug(header_values),
            "mapping": mapping_suggestion.get("mapping", {}),
            "confidence": mapping_suggestion.get("confidence"),
            "notes": mapping_suggestion.get("notes", []),
            "sample_rows": [
                {
                    "sheet_row_number": index + 1,
                    "non_empty_positions": _row_non_empty_positions(row, limit=30),
                    "cells": _row_cells_debug(row, header_values, limit=30),
                }
                for index, row in enumerate(rows[:min(len(rows), 8)])
            ],
        },
        request_id=operation_id,
    )
    return {
        "ok": True,
        "data_dir": str(root),
        "online_debug_log_file": feishu_debug_log_path(),
        "online_parse_debug_log_file": feishu_parse_debug_log_path(),
        "spreadsheet_token": token,
        "sheet_id": _safe_cell_str(sheet_id),
        "row_count": len(rows),
        "column_range": _safe_cell_str(column_range or "A:Z").upper(),
        "header_row": safe_header_row,
        "headers": [header for header in header_values if header],
        "rows": rows,
        "mapping_suggestion": mapping_suggestion,
    }


def sync_feishu_library(*,
                         library_name: str,
                         spreadsheet_token_or_url: str,
                         sheets: Sequence[dict],
                         base_url: str = "",
                         origin: str = "",
                         user_id: str = "",
                         data_dir: str = "",
                         library_id: str = "",
                         client: Optional[FeishuBomClient] = None) -> dict:
    operation_id = new_diagnostic_request_id()
    library_name = _safe_cell_str(library_name)
    token = extract_spreadsheet_token(spreadsheet_token_or_url)
    if not library_name:
        raise FeishuBomError("缺少 library_name。")
    if not token:
        raise FeishuBomError("缺少 spreadsheet_token_or_url。")

    resolved_base_url, resolved_origin, resolved_user_id, root, config = _resolve_connection(
        base_url=base_url,
        origin=origin,
        user_id=user_id,
        data_dir=data_dir,
    )
    config.update({
        "base_url": resolved_base_url,
        "origin": resolved_origin,
        "user_id": resolved_user_id,
    })
    sheet_cfgs = _normalize_sheet_cfgs(sheets)
    if not sheet_cfgs:
        raise FeishuBomError("缺少 sheets 配置。")

    _log_feishu_event(
        "sync.start",
        {
            "library_name": library_name,
            "library_id": library_id,
            "spreadsheet_token": token,
            "base_url": resolved_base_url,
            "origin": resolved_origin,
            "user_id": resolved_user_id,
            "data_dir": str(root),
            "sheet_config_count": len(sheet_cfgs),
            "enabled_sheet_count": len([cfg for cfg in sheet_cfgs if cfg.get("enabled", True)]),
            "sheet_configs": [
                {
                    "sheet_id": cfg.get("sheet_id"),
                    "title": cfg.get("title"),
                    "enabled": cfg.get("enabled", True),
                    "row_count": cfg.get("row_count"),
                    "column_range": cfg.get("column_range"),
                    "header_row": cfg.get("header_row"),
                    "hq_code_col": cfg.get("hq_code_col"),
                    "spec_model_col": cfg.get("spec_model_col"),
                    "pi_col": cfg.get("pi_col"),
                    "selection_order_col": cfg.get("selection_order_col"),
                    "optional_field_count": len(cfg.get("optional_fields") or []),
                }
                for cfg in sheet_cfgs[:60]
            ],
        },
        request_id=operation_id,
    )
    api = client or FeishuBomClient(resolved_base_url, resolved_origin, resolved_user_id)
    resolved_library_id = _safe_cell_str(library_id) or _library_id_from_name(library_name)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows_to_insert: List[tuple] = []
    per_sheet: List[dict] = []

    for sheet_cfg in [cfg for cfg in sheet_cfgs if cfg.get("enabled", True)]:
        sheet_id = sheet_cfg["sheet_id"]
        title = sheet_cfg["title"]
        spec_model_col = _safe_cell_str(sheet_cfg.get("spec_model_col") or sheet_cfg.get("key_col"))
        hq_code_col = _safe_cell_str(sheet_cfg.get("hq_code_col") or sheet_cfg.get("hq_no_col"))
        if not spec_model_col and not hq_code_col:
            _log_feishu_event(
                "sync.sheet.skipped",
                {
                    "sheet_id": sheet_id,
                    "title": title,
                    "reason": "未配置规格型号列或 HQ 料号列",
                    "sheet_config": sheet_cfg,
                },
                level="warning",
                request_id=operation_id,
            )
            per_sheet.append({
                "sheet_id": sheet_id,
                "title": title,
                "status": "skipped",
                "reason": "未配置规格型号列或 HQ 料号列",
                "row_count": 0,
            })
            continue

        values = api.read_sheet(
            token,
            sheet_id,
            row_count=int(sheet_cfg.get("row_count") or 5000),
            column_range=_safe_cell_str(sheet_cfg.get("column_range") or "A:Z"),
        )
        _log_feishu_event(
            "sync.sheet.read",
            {
                "sheet_id": sheet_id,
                "title": title,
                "row_shape": _rows_shape(values),
                "row_preview": _rows_preview(values),
                "configured_header_row": sheet_cfg.get("header_row"),
                "configured_columns": {
                    "hq_code_col": hq_code_col,
                    "spec_model_col": spec_model_col,
                    "pi_col": _safe_cell_str(sheet_cfg.get("pi_col")),
                    "selection_order_col": _safe_cell_str(sheet_cfg.get("selection_order_col")),
                },
            },
            request_id=operation_id,
        )
        if not values:
            _log_feishu_event(
                "sync.sheet.skipped",
                {
                    "sheet_id": sheet_id,
                    "title": title,
                    "reason": "无数据",
                },
                level="warning",
                request_id=operation_id,
            )
            per_sheet.append({
                "sheet_id": sheet_id,
                "title": title,
                "status": "skipped",
                "reason": "无数据",
                "row_count": 0,
            })
            continue

        header_row = int(sheet_cfg.get("header_row") or 1)
        header_index = header_row - 1
        if len(values) <= header_index:
            _log_feishu_event(
                "sync.sheet.skipped",
                {
                    "sheet_id": sheet_id,
                    "title": title,
                    "reason": f"表格不足 {header_row} 行，无法读取表头",
                    "row_shape": _rows_shape(values),
                },
                level="warning",
                request_id=operation_id,
            )
            per_sheet.append({
                "sheet_id": sheet_id,
                "title": title,
                "status": "skipped",
                "reason": f"表格不足 {header_row} 行，无法读取表头",
                "row_count": 0,
            })
            continue

        headers = [_safe_cell_str(cell) for cell in values[header_index]]
        spec_index = _header_index(headers, spec_model_col)
        hq_index = _header_index(headers, hq_code_col)
        if spec_index < 0 and hq_index < 0:
            _log_feishu_event(
                "sync.sheet.skipped",
                {
                    "sheet_id": sheet_id,
                    "title": title,
                    "reason": f"找不到规格型号列 {spec_model_col or '未配置'} 或 HQ 料号列 {hq_code_col or '未配置'}",
                    "headers": headers,
                    "configured_columns": {
                        "hq_code_col": hq_code_col,
                        "spec_model_col": spec_model_col,
                    },
                },
                level="warning",
                request_id=operation_id,
            )
            per_sheet.append({
                "sheet_id": sheet_id,
                "title": title,
                "status": "skipped",
                "reason": f"找不到规格型号列 {spec_model_col or '未配置'} 或 HQ 料号列 {hq_code_col or '未配置'}",
                "row_count": 0,
                "headers": headers,
            })
            continue

        brand_index = _header_index(headers, _safe_cell_str(sheet_cfg.get("brand_col")))
        desc_index = _header_index(headers, _safe_cell_str(sheet_cfg.get("desc_col")))
        pi_index = _header_index(headers, _safe_cell_str(sheet_cfg.get("pi_col")))
        selection_order_index = _header_index(headers, _safe_cell_str(sheet_cfg.get("selection_order_col")))
        optional_fields = _normalize_optional_fields(sheet_cfg.get("optional_fields"), headers)
        optional_indexes = [
            (field["label"], field["column"], _header_index(headers, field["column"]))
            for field in optional_fields
        ]
        header_width = len(headers)
        field_indexes = {
            "hq": hq_index,
            "spec": spec_index,
            "brand": brand_index,
            "description": desc_index,
            "pi": pi_index,
            "selection_order": selection_order_index,
        }
        configured_columns = {
            "hq_code_col": hq_code_col,
            "spec_model_col": spec_model_col,
            "brand_col": _safe_cell_str(sheet_cfg.get("brand_col")),
            "desc_col": _safe_cell_str(sheet_cfg.get("desc_col")),
            "pi_col": _safe_cell_str(sheet_cfg.get("pi_col")),
            "selection_order_col": _safe_cell_str(sheet_cfg.get("selection_order_col")),
        }
        missing_configured_columns = [
            {"field": field, "column": column}
            for field, column in configured_columns.items()
            if column and _header_index(headers, column) < 0
        ]
        _log_feishu_parse_event(
            "sync.sheet.mapping_resolved",
            {
                "spreadsheet_token": token,
                "library_id": resolved_library_id,
                "library_name": library_name,
                "sheet_id": sheet_id,
                "title": title,
                "header_row": header_row,
                "header_index": header_index,
                "header_width": header_width,
                "row_shape": _rows_shape(values),
                "configured_columns": configured_columns,
                "field_indexes": _field_indexes_debug(headers, field_indexes),
                "optional_indexes": [
                    {
                        "label": label,
                        "column": column,
                        "index": index,
                        "excel_column": _column_index_to_name(index + 1) if index >= 0 else "",
                        "found": index >= 0,
                    }
                    for label, column, index in optional_indexes
                ],
                "missing_configured_columns": missing_configured_columns,
                "duplicate_headers": _duplicate_headers_debug(headers),
                "header_positions": _headers_position_debug(headers),
            },
            request_id=operation_id,
        )
        row_alignment_adjusted = 0
        row_alignment_examples: List[dict] = []
        parse_row_limit = _feishu_parse_log_row_limit()
        row_parse_samples: List[dict] = []
        row_parse_counters = {
            "total_data_rows": 0,
            "inserted_rows": 0,
            "skipped_blank_or_missing_key": 0,
            "rows_with_a_empty": 0,
            "rows_with_alignment_adjusted": 0,
            "row_samples_suppressed": 0,
        }

        def get_cell(row: Sequence[object], index: int) -> str:
            return _safe_cell_str(row[index]) if 0 <= index < len(row) else ""

        def append_row_parse_sample(*,
                                    sheet_row_number: int,
                                    row: Sequence[object],
                                    original_width: int,
                                    row_alignment: dict,
                                    status: str,
                                    skip_reason: str = "") -> None:
            if parse_row_limit <= 0:
                row_parse_counters["row_samples_suppressed"] += 1
                return
            if len(row_parse_samples) >= parse_row_limit:
                row_parse_counters["row_samples_suppressed"] += 1
                return
            row_parse_samples.append({
                "sheet_row_number": sheet_row_number,
                "status": status,
                "skip_reason": skip_reason,
                "original_width": original_width,
                "aligned_width": len(row or []),
                "a_column_empty": not bool(get_cell(row, 0)),
                "row_alignment": row_alignment,
                "field_values": _row_field_values_debug(row, headers, field_indexes),
                "non_empty_positions": _row_non_empty_positions(row, limit=40),
                "cells": _row_cells_debug(row, headers, limit=40),
            })

        sheet_rows = 0
        for data_index, row in enumerate(values[header_row:], start=0):
            if not isinstance(row, list):
                continue
            row_parse_counters["total_data_rows"] += 1
            sheet_row_number = header_row + data_index + 1
            original_width = len(row)
            row, row_alignment = _align_data_row_to_headers(
                row,
                header_width=header_width,
                hq_index=hq_index,
                spec_index=spec_index,
                pi_index=pi_index,
                selection_order_index=selection_order_index,
            )
            if row_alignment.get("applied"):
                row_alignment_adjusted += 1
                row_parse_counters["rows_with_alignment_adjusted"] += 1
                if len(row_alignment_examples) < 5:
                    row_alignment_examples.append({
                        **row_alignment,
                        "preview": [_safe_cell_str(cell) for cell in row[:min(header_width, 12)]],
                    })
            if not get_cell(row, 0):
                row_parse_counters["rows_with_a_empty"] += 1
            spec_model = get_cell(row, spec_index)
            hq_code = get_cell(row, hq_index)
            key_value = spec_model or hq_code
            if not key_value:
                row_parse_counters["skipped_blank_or_missing_key"] += 1
                if len(row_parse_samples) < parse_row_limit:
                    append_row_parse_sample(
                        sheet_row_number=sheet_row_number,
                        row=row,
                        original_width=original_width,
                        row_alignment=row_alignment,
                        status="skipped",
                        skip_reason="missing_spec_and_hq",
                    )
                continue
            extra_fields = {
                label: get_cell(row, index)
                for label, _column, index in optional_indexes
                if index >= 0 and get_cell(row, index)
            }
            raw_data = json.dumps(
                {
                    headers[index]: get_cell(row, index)
                    for index in range(min(len(headers), len(row)))
                },
                ensure_ascii=False,
            )
            rows_to_insert.append((
                resolved_library_id,
                library_name,
                title,
                key_value,
                hq_code,
                get_cell(row, brand_index),
                spec_model,
                get_cell(row, desc_index),
                get_cell(row, pi_index),
                get_cell(row, selection_order_index),
                json.dumps(extra_fields, ensure_ascii=False),
                raw_data,
                now,
            ))
            sheet_rows += 1
            row_parse_counters["inserted_rows"] += 1
            interesting = (
                sheet_rows <= 5
                or row_alignment.get("applied")
                or not get_cell(row, 0)
                or not hq_code
                or not spec_model
            )
            if interesting:
                append_row_parse_sample(
                    sheet_row_number=sheet_row_number,
                    row=row,
                    original_width=original_width,
                    row_alignment=row_alignment,
                    status="inserted",
                )

        _log_feishu_event(
            "sync.sheet.synced",
            {
                "sheet_id": sheet_id,
                "title": title,
                "row_count": sheet_rows,
                "header_row": header_row,
                "headers": headers,
                "resolved_indexes": {
                    "hq_index": hq_index,
                    "spec_index": spec_index,
                    "brand_index": brand_index,
                    "desc_index": desc_index,
                    "pi_index": pi_index,
                    "selection_order_index": selection_order_index,
                },
                "optional_indexes": optional_indexes,
                "row_alignment_adjusted": row_alignment_adjusted,
                "row_alignment_examples": row_alignment_examples,
            },
            request_id=operation_id,
        )
        _log_feishu_parse_event(
            "sync.sheet.row_parse_summary",
            {
                "spreadsheet_token": token,
                "library_id": resolved_library_id,
                "library_name": library_name,
                "sheet_id": sheet_id,
                "title": title,
                "header_row": header_row,
                "header_width": header_width,
                "counters": row_parse_counters,
                "row_alignment_adjusted": row_alignment_adjusted,
                "row_alignment_examples": row_alignment_examples,
                "row_samples_limit": parse_row_limit,
                "row_samples": row_parse_samples,
            },
            request_id=operation_id,
        )
        per_sheet.append({
            "sheet_id": sheet_id,
            "title": title,
            "status": "synced",
            "row_count": sheet_rows,
            "header_row": header_row,
            "headers": headers,
            "row_alignment_adjusted": row_alignment_adjusted,
            "field_mapping": {
                "hq_code_col": hq_code_col,
                "spec_model_col": spec_model_col,
                "pi_col": _safe_cell_str(sheet_cfg.get("pi_col")),
                "selection_order_col": _safe_cell_str(sheet_cfg.get("selection_order_col")),
                "optional_fields": optional_fields,
            },
        })

    conn = _connect_cache_for_write(_cache_path(root))
    try:
        _log_feishu_event(
            "sync.cache_write.start",
            {
                "library_id": resolved_library_id,
                "cache_file": str(_cache_path(root)),
                "rows_to_insert": len(rows_to_insert),
            },
            request_id=operation_id,
        )
        conn.execute("DELETE FROM materials WHERE lib_id=?", (resolved_library_id,))
        if rows_to_insert:
            conn.executemany(
                "INSERT INTO materials(lib_id,lib_name,sheet_name,key_value,hq_no,brand,spec,description,pi,selection_order,extra_fields,raw_data,synced_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                rows_to_insert,
            )
        conn.commit()
        _log_feishu_event(
            "sync.cache_write.success",
            {
                "library_id": resolved_library_id,
                "cache_file": str(_cache_path(root)),
                "inserted_rows": len(rows_to_insert),
            },
            request_id=operation_id,
        )
    finally:
        conn.close()

    _upsert_library_config(
        config,
        library_id=resolved_library_id,
        library_name=library_name,
        token=token,
        sheets=sheet_cfgs,
        last_sync=now,
    )
    _save_config(root, config)

    _log_feishu_event(
        "sync.success",
        {
            "library_id": resolved_library_id,
            "library_name": library_name,
            "sheet_count": len([sheet for sheet in sheet_cfgs if sheet.get("enabled", True)]),
            "synced_rows": len(rows_to_insert),
            "skipped_sheets": len([sheet for sheet in per_sheet if sheet.get("status") == "skipped"]),
            "per_sheet": per_sheet,
            "config_file": str(_config_path(root)),
            "cache_file": str(_cache_path(root)),
        },
        request_id=operation_id,
    )
    return {
        "ok": True,
        "data_dir": str(root),
        "cache_file": str(_cache_path(root)),
        "config_file": str(_config_path(root)),
        "online_debug_log_file": feishu_debug_log_path(),
        "online_parse_debug_log_file": feishu_parse_debug_log_path(),
        "library_id": resolved_library_id,
        "library_name": library_name,
        "spreadsheet_token": token,
        "sheet_count": len([sheet for sheet in sheet_cfgs if sheet.get("enabled", True)]),
        "synced_rows": len(rows_to_insert),
        "skipped_sheets": len([sheet for sheet in per_sheet if sheet.get("status") == "skipped"]),
        "synced_at": now,
        "per_sheet": per_sheet,
    }
