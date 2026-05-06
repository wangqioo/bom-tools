# -*- coding: utf-8 -*-
"""HTTP client for the internal Feishu sheet gateway."""

from __future__ import annotations

import json
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Callable, List, Optional

from pstx_integrations.diagnostics import new_diagnostic_request_id
from pstx_integrations.feishu.common import (
    FeishuBomError,
    _align_values_to_requested_range,
    _default_column_range,
    _include_feishu_payload,
    _log_feishu_event,
    _log_feishu_parse_event,
    _payload_shape,
    _row_cells_debug,
    _row_non_empty_positions,
    _rows_preview,
    _rows_shape,
    _safe_cell_str,
    _sheet_meta_debug,
    build_sheet_value_range,
    extract_spreadsheet_token,
)


HttpTransport = Callable[[str, dict, int], dict]


def _http_get_json(url: str, params: dict, timeout: int) -> dict:
    query = urllib.parse.urlencode(params)
    request_url = f"{url}?{query}" if query else url
    request = urllib.request.Request(request_url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise FeishuBomError(f"飞书网关 HTTP {exc.code}: {body[:500]}") from exc
    except urllib.error.URLError as exc:
        raise FeishuBomError(f"飞书网关连接失败：{exc.reason}") from exc
    try:
        payload = json.loads(body)
    except json.JSONDecodeError as exc:
        raise FeishuBomError(f"飞书网关返回非 JSON 内容：{body[:300]}") from exc
    if not isinstance(payload, dict):
        raise FeishuBomError("飞书网关返回 JSON 不是对象。")
    return payload


class FeishuBomClient:
    """Minimal client for the internal Feishu sheet gateway."""

    def __init__(self,
                 base_url: str,
                 origin: str,
                 user_id: str,
                 *,
                 timeout: int = 30,
                 transport: Optional[HttpTransport] = None):
        self.base_url = _safe_cell_str(base_url).rstrip("/")
        self.origin = _safe_cell_str(origin)
        self.user_id = _safe_cell_str(user_id)
        self.timeout = int(timeout or 30)
        self._transport = transport or _http_get_json
        if not self.base_url.startswith("http"):
            raise FeishuBomError("飞书网关地址必须以 http/https 开头。")
        if not self.origin:
            raise FeishuBomError("缺少 origin / AppId。")
        if not self.user_id:
            raise FeishuBomError("缺少工号 user_id。")

    def _get(self, path: str, params: dict, *, timeout: Optional[int] = None, request_id: str = "") -> dict:
        url = f"{self.base_url}{path}"
        request_id = request_id or new_diagnostic_request_id()
        timeout_seconds = int(timeout or self.timeout)
        started_at = time.time()
        _log_feishu_event(
            "request.start",
            {
                "method": "GET",
                "url": url,
                "path": path,
                "params": params,
                "timeout_seconds": timeout_seconds,
            },
            request_id=request_id,
        )
        try:
            payload = self._transport(url, params, timeout_seconds)
        except Exception as exc:
            _log_feishu_event(
                "request.error",
                {
                    "method": "GET",
                    "url": url,
                    "path": path,
                    "params": params,
                    "elapsed_ms": int((time.time() - started_at) * 1000),
                    "error_type": exc.__class__.__name__,
                    "error": str(exc),
                },
                level="error",
                request_id=request_id,
            )
            raise
        details = {
            "method": "GET",
            "url": url,
            "path": path,
            "params": params,
            "elapsed_ms": int((time.time() - started_at) * 1000),
            "payload_shape": _payload_shape(payload),
        }
        if _include_feishu_payload():
            details["payload"] = payload
        _log_feishu_event(
            "request.response",
            details,
            level="info" if payload.get("code") in (0, 200) else "warning",
            request_id=request_id,
        )
        if payload.get("code") not in (0, 200):
            _log_feishu_event(
                "request.api_error",
                {
                    "path": path,
                    "code": payload.get("code"),
                    "msg": payload.get("msg") or payload.get("message") or "unknown",
                    "payload_shape": _payload_shape(payload),
                },
                level="error",
                request_id=request_id,
            )
            raise FeishuBomError(
                f"飞书接口失败：{payload.get('msg') or payload.get('message') or 'unknown'} "
                f"(code={payload.get('code')})"
            )
        return payload

    def get_sheets(self, spreadsheet_token: str) -> List[dict]:
        request_id = new_diagnostic_request_id()
        token = extract_spreadsheet_token(spreadsheet_token)
        if not token:
            raise FeishuBomError("spreadsheet_token 不能为空。")
        _log_feishu_event(
            "metainfo.start",
            {
                "spreadsheet_token": token,
                "base_url": self.base_url,
                "origin": self.origin,
                "user_id": self.user_id,
            },
            request_id=request_id,
        )
        payload = self._get(
            "/fs/sheet/v1/spreadsheetsMetainfo",
            {
                "origin": self.origin,
                "userId": self.user_id,
                "spreadsheetToken": token,
            },
            timeout=min(max(self.timeout, 15), 60),
            request_id=request_id,
        )
        data = payload.get("data") or {}
        sheets = data.get("sheets") or []
        output: List[dict] = []
        skipped: List[dict] = []
        for sheet in sheets:
            if not isinstance(sheet, dict):
                skipped.append({"reason": "not_object", "type": type(sheet).__name__})
                continue
            if not sheet.get("title"):
                skipped.append({"reason": "empty_title", **_sheet_meta_debug(sheet)})
                continue
            if sheet.get("blockInfo"):
                skipped.append({"reason": "block_info", **_sheet_meta_debug(sheet)})
                continue
            output.append({
                "sheet_id": _safe_cell_str(sheet.get("sheetId")),
                "sheetId": _safe_cell_str(sheet.get("sheetId")),
                "title": _safe_cell_str(sheet.get("title")),
                "row_count": int(sheet.get("rowCount") or sheet.get("row_count") or 0),
                "column_count": int(sheet.get("columnCount") or sheet.get("column_count") or 0),
                "column_range": _default_column_range(sheet.get("columnCount") or sheet.get("column_count")),
                "index": int(sheet.get("index") or 0),
                "revision": int(data.get("properties", {}).get("revision") or data.get("revision") or 0),
                "frozen_row_count": int(sheet.get("frozenRowCount") or 0),
                "frozen_col_count": int(sheet.get("frozenColCount") or 0),
                "protected_range": sheet.get("protectedRange") or [],
                "raw": sheet,
            })
        _log_feishu_event(
            "metainfo.parsed",
            {
                "spreadsheet_token": token,
                "raw_sheet_count": len(sheets) if isinstance(sheets, list) else 0,
                "returned_sheet_count": len(output),
                "skipped_sheet_count": len(skipped),
                "skipped_sheets": skipped[:20],
                "sheets": [
                    {
                        "sheet_id": item.get("sheet_id"),
                        "title": item.get("title"),
                        "row_count": item.get("row_count"),
                        "column_count": item.get("column_count"),
                        "column_range": item.get("column_range"),
                        "index": item.get("index"),
                    }
                    for item in output[:40]
                ],
            },
            request_id=request_id,
        )
        return output

    def read_sheet(self,
                   spreadsheet_token: str,
                   sheet_id: str,
                   *,
                   row_count: int = 5000,
                   column_range: str = "A:Z") -> List[List[object]]:
        request_id = new_diagnostic_request_id()
        token = extract_spreadsheet_token(spreadsheet_token)
        value_range = build_sheet_value_range(sheet_id, row_count, column_range)
        _log_feishu_event(
            "read_sheet.start",
            {
                "spreadsheet_token": token,
                "sheet_id": sheet_id,
                "requested_row_count": row_count,
                "column_range": column_range,
                "value_range": value_range,
            },
            request_id=request_id,
        )
        payload = None
        endpoint = "v2-formatted"
        try:
            payload = self._get(
                "/fs/sheets",
                {
                    "origin": self.origin,
                    "empNo": self.user_id,
                    "spreadsheetToken": token,
                    "valueRenderOption": "FormattedValue",
                    "range": value_range,
                    "appId": self.origin,
                },
                timeout=max(self.timeout, 60),
                request_id=request_id,
            )
        except Exception as exc:
            endpoint = "v1-fallback"
            _log_feishu_event(
                "read_sheet.v2_fallback",
                {
                    "spreadsheet_token": token,
                    "sheet_id": sheet_id,
                    "value_range": value_range,
                    "error_type": exc.__class__.__name__,
                    "error": str(exc)[:500],
                },
                level="warning",
                request_id=request_id,
            )
            payload = self._get(
                "/fs/sheet/v1/getSheetsValue",
                {
                    "origin": self.origin,
                    "userId": self.user_id,
                    "spreadsheetToken": token,
                    "range": value_range,
                    "valueRenderOption": "FormattedValue",
                },
                timeout=max(self.timeout, 60),
                request_id=request_id,
            )
        data = payload.get("data") or {}
        value_range_payload = data.get("valueRange") or {}
        if not value_range_payload and isinstance(data.get("valueRanges"), list) and data.get("valueRanges"):
            value_range_payload = data["valueRanges"][0] or {}
        returned_range = _safe_cell_str(value_range_payload.get("range") or "")
        values = value_range_payload.get("values") or []
        if not isinstance(values, list):
            _log_feishu_event(
                "read_sheet.invalid_values",
                {
                    "spreadsheet_token": token,
                    "sheet_id": sheet_id,
                    "value_range": value_range,
                    "endpoint": endpoint,
                    "values_type": type(values).__name__,
                    "payload_shape": _payload_shape(payload),
                },
                level="warning",
                request_id=request_id,
            )
            return []
        raw_shape = _rows_shape(values)
        values, alignment = _align_values_to_requested_range(
            values,
            requested_range=value_range,
            returned_range=returned_range,
        )
        aligned_shape = _rows_shape(values)
        aligned_row_count = len(values)
        while values and not any(_safe_cell_str(v) for v in values[-1]):
            values.pop()
        trimmed_trailing_empty_rows = aligned_row_count - len(values)
        # The sheet API can return rich cell objects (links, mentions, etc.).
        # Normalize here so all downstream preview/sync logic sees the same
        # text shape as CSV-like rows.
        normalized = [
            [_safe_cell_str(cell) for cell in row]
            for row in values
            if isinstance(row, list)
        ]
        normalized_shape = _rows_shape(normalized)
        _log_feishu_event(
            "read_sheet.parsed",
            {
                "spreadsheet_token": token,
                "sheet_id": sheet_id,
                "value_range": value_range,
                "endpoint": endpoint,
                "returned_range": returned_range,
                "range_alignment": alignment,
                "raw_shape": raw_shape,
                "aligned_shape": aligned_shape,
                "normalized_shape": normalized_shape,
                "row_preview": _rows_preview(normalized),
            },
            request_id=request_id,
        )
        _log_feishu_parse_event(
            "read_sheet.rows_aligned",
            {
                "spreadsheet_token": token,
                "sheet_id": sheet_id,
                "value_range": value_range,
                "endpoint": endpoint,
                "returned_range": returned_range,
                "range_alignment": alignment,
                "raw_shape": raw_shape,
                "aligned_shape": aligned_shape,
                "normalized_shape": normalized_shape,
                "trimmed_trailing_empty_rows": trimmed_trailing_empty_rows,
                "row_preview": _rows_preview(normalized, row_limit=8, col_limit=20),
                "first_rows_cells": [
                    {
                        "sheet_row_number": index + 1,
                        "non_empty_positions": _row_non_empty_positions(row, limit=20),
                        "cells": _row_cells_debug(row, limit=20),
                    }
                    for index, row in enumerate(normalized[:5])
                ],
            },
            request_id=request_id,
        )
        return normalized
