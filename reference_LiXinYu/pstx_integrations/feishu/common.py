# -*- coding: utf-8 -*-
"""Common constants, diagnostics, text normalization, range and config helpers for Feishu BOM."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import List, Optional, Sequence

from pstx_integrations.diagnostics import write_diagnostic_event


BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR_NAME = "feishu_bom_data"
DATA_DIR_ENV = "PSTX_FEISHU_DATA_DIR"
CONFIG_FILE_NAME = "feishu_libraries.json"
CACHE_FILE_NAME = "feishu_cache.db"
FEISHU_LOG_FILE = BASE_DIR / "logs" / "feishu_bom_debug.log"
FEISHU_PARSE_LOG_FILE = BASE_DIR / "logs" / "feishu_bom_parse_debug.log"
FEISHU_LOG_FILE_ENV = "PSTX_FEISHU_LOG_FILE"
FEISHU_PARSE_LOG_FILE_ENV = "PSTX_FEISHU_PARSE_LOG_FILE"
FEISHU_LOG_PAYLOAD_ENV = "PSTX_FEISHU_LOG_PAYLOAD"
FEISHU_PARSE_LOG_ROWS_ENV = "PSTX_FEISHU_PARSE_LOG_ROWS"
DEFAULT_BASE_URL = "https://mcenter.huaqin.com"
DEFAULT_ORIGIN = "cli_a96ac38049f8d0e5"
HQ_FIELD_ALIASES = {
    "hq_code",
    "hq code",
    "hqcode",
    "hq no",
    "hq_no",
    "hq料号",
    "hq编码",
    "料号",
    "物料号",
    "物料编码",
}
SPEC_FIELD_ALIASES = {
    "规格型号",
    "part number",
    "part no",
    "p/n",
    "厂家型号",
    "制造商型号",
    "描述",
}


class FeishuBomError(RuntimeError):
    """Raised when the Feishu BOM adapter cannot complete a requested action."""


def feishu_debug_log_path(environ: Optional[dict] = None) -> str:
    env = environ if environ is not None else os.environ
    raw = str(env.get(FEISHU_LOG_FILE_ENV) or "").strip().strip('"')
    return str(Path(raw).expanduser()) if raw else str(FEISHU_LOG_FILE)


def feishu_parse_debug_log_path(environ: Optional[dict] = None) -> str:
    env = environ if environ is not None else os.environ
    raw = str(env.get(FEISHU_PARSE_LOG_FILE_ENV) or "").strip().strip('"')
    return str(Path(raw).expanduser()) if raw else str(FEISHU_PARSE_LOG_FILE)


def _include_feishu_payload(environ: Optional[dict] = None) -> bool:
    env = environ if environ is not None else os.environ
    return str(env.get(FEISHU_LOG_PAYLOAD_ENV) or "").strip().lower() in {"1", "true", "yes", "on"}


def _feishu_parse_log_row_limit(environ: Optional[dict] = None) -> int:
    env = environ if environ is not None else os.environ
    try:
        return min(max(int(env.get(FEISHU_PARSE_LOG_ROWS_ENV) or 120), 0), 2000)
    except (TypeError, ValueError):
        return 120


def _log_feishu_event(event: str,
                      details: Optional[dict] = None,
                      *,
                      level: str = "info",
                      request_id: str = "") -> dict:
    payload = {
        "component": "feishu_bom",
        **(details or {}),
    }
    return write_diagnostic_event(
        f"feishu_bom.{event}",
        payload,
        level=level,
        request_id=request_id,
        log_file=feishu_debug_log_path(),
    )


def _log_feishu_parse_event(event: str,
                            details: Optional[dict] = None,
                            *,
                            level: str = "info",
                            request_id: str = "") -> dict:
    payload = {
        "component": "feishu_bom_parse",
        **(details or {}),
    }
    return write_diagnostic_event(
        f"feishu_bom_parse.{event}",
        payload,
        level=level,
        request_id=request_id,
        log_file=feishu_parse_debug_log_path(),
    )


def _rows_shape(rows: Sequence[Sequence[object]]) -> dict:
    normalized = [row for row in (rows or []) if isinstance(row, (list, tuple))]
    non_empty = [row for row in normalized if any(_safe_cell_str(cell) for cell in row)]
    return {
        "row_count": len(normalized),
        "non_empty_row_count": len(non_empty),
        "max_column_count": max((len(row) for row in normalized), default=0),
        "first_non_empty_row": (normalized.index(non_empty[0]) + 1) if non_empty else 0,
    }


def _rows_preview(rows: Sequence[Sequence[object]], *, row_limit: int = 5, col_limit: int = 12) -> list:
    preview = []
    for row in list(rows or [])[:row_limit]:
        if not isinstance(row, (list, tuple)):
            continue
        preview.append([_safe_cell_str(cell) for cell in list(row)[:col_limit]])
    return preview


def _payload_shape(payload: object) -> dict:
    if not isinstance(payload, dict):
        return {"type": type(payload).__name__}
    data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
    value_range = data.get("valueRange") if isinstance(data.get("valueRange"), dict) else {}
    value_ranges = data.get("valueRanges") if isinstance(data.get("valueRanges"), list) else []
    first_value_range = value_range or (value_ranges[0] if value_ranges and isinstance(value_ranges[0], dict) else {})
    values = first_value_range.get("values") if isinstance(first_value_range, dict) else []
    return {
        "code": payload.get("code"),
        "msg": payload.get("msg") or payload.get("message"),
        "keys": sorted(str(key) for key in payload.keys())[:30],
        "data_keys": sorted(str(key) for key in data.keys())[:30] if data else [],
        "sheet_count": len(data.get("sheets") or []) if isinstance(data, dict) else 0,
        "value_range_count": len(value_ranges),
        "values_shape": _rows_shape(values if isinstance(values, list) else []),
    }


def _sheet_meta_debug(sheet: dict) -> dict:
    return {
        "sheet_id": _safe_cell_str(sheet.get("sheetId") or sheet.get("sheet_id")),
        "title": _safe_cell_str(sheet.get("title")),
        "index": sheet.get("index"),
        "row_count": sheet.get("rowCount") or sheet.get("row_count"),
        "column_count": sheet.get("columnCount") or sheet.get("column_count"),
        "has_block_info": bool(sheet.get("blockInfo")),
        "block_type": _safe_cell_str((sheet.get("blockInfo") or {}).get("blockType")) if isinstance(sheet.get("blockInfo"), dict) else "",
        "frozen_row_count": sheet.get("frozenRowCount"),
        "frozen_col_count": sheet.get("frozenColCount"),
        "merge_count": len(sheet.get("merges") or []) if isinstance(sheet.get("merges"), list) else 0,
        "protected_range_count": len(sheet.get("protectedRange") or []) if isinstance(sheet.get("protectedRange"), list) else 0,
    }


def _safe_cell_str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        for key in ("text", "link", "value", "name"):
            text = _safe_cell_str(value.get(key))
            if text:
                return text
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    if isinstance(value, list):
        parts = [_safe_cell_str(item) for item in value]
        return " ".join(part for part in parts if part).strip()
    return str(value).strip()


def _normalized_name(value: object) -> str:
    return re.sub(r"\s+", " ", _safe_cell_str(value).replace("_", " ")).strip().lower()


def extract_spreadsheet_token(raw: str) -> str:
    """Extract token from a Feishu sheet/base URL, or return the raw token."""

    raw = (raw or "").strip()
    if not raw:
        return ""
    match = re.search(r"/(?:sheets|base)/([A-Za-z0-9]+)", raw)
    return match.group(1) if match else raw


def _column_index_to_name(index: int) -> str:
    value = max(1, int(index or 1))
    chars: List[str] = []
    while value:
        value, remainder = divmod(value - 1, 26)
        chars.append(chr(65 + remainder))
    return "".join(reversed(chars))


def _column_name_to_index(name: str) -> int:
    value = 0
    for char in _safe_cell_str(name).upper():
        if not ("A" <= char <= "Z"):
            return 0
        value = value * 26 + (ord(char) - 64)
    return value


def _default_column_range(column_count: object) -> str:
    try:
        count = int(column_count or 0)
    except (TypeError, ValueError):
        count = 0
    if count <= 0:
        return "A:Z"
    return f"A:{_column_index_to_name(min(max(count, 1), 702))}"


def _headers_position_debug(headers: Sequence[object], *, limit: int = 120) -> list:
    output = []
    for index, header in enumerate(list(headers or [])[:limit]):
        output.append({
            "index": index,
            "column": _column_index_to_name(index + 1),
            "header": _safe_cell_str(header),
        })
    if len(headers or []) > limit:
        output.append({"truncated": True, "remaining": len(headers or []) - limit})
    return output


def _duplicate_headers_debug(headers: Sequence[object]) -> list:
    seen: dict[str, List[int]] = {}
    for index, header in enumerate(headers or []):
        text = _safe_cell_str(header)
        if not text:
            continue
        seen.setdefault(text, []).append(index)
    return [
        {
            "header": header,
            "positions": [
                {"index": index, "column": _column_index_to_name(index + 1)}
                for index in indexes
            ],
        }
        for header, indexes in seen.items()
        if len(indexes) > 1
    ]


def _row_non_empty_positions(row: Sequence[object], *, limit: int = 60) -> list:
    output = []
    for index, value in enumerate(row or []):
        text = _safe_cell_str(value)
        if not text:
            continue
        output.append({
            "index": index,
            "column": _column_index_to_name(index + 1),
            "value": text,
        })
        if len(output) >= limit:
            output.append({"truncated": True})
            break
    return output


def _row_cells_debug(row: Sequence[object], headers: Sequence[object] = (), *, limit: int = 40) -> list:
    output = []
    width = min(max(len(row or []), len(headers or [])), limit)
    for index in range(width):
        header = _safe_cell_str(headers[index]) if index < len(headers or []) else ""
        value = _safe_cell_str(row[index]) if index < len(row or []) else ""
        output.append({
            "index": index,
            "column": _column_index_to_name(index + 1),
            "header": header,
            "value": value,
            "empty": not bool(value),
        })
    if max(len(row or []), len(headers or [])) > limit:
        output.append({"truncated": True, "remaining": max(len(row or []), len(headers or [])) - limit})
    return output


def _field_indexes_debug(headers: Sequence[object], indexes: dict) -> dict:
    output = {}
    for name, index in indexes.items():
        try:
            position = int(index)
        except (TypeError, ValueError):
            position = -1
        output[name] = {
            "index": position,
            "column": _column_index_to_name(position + 1) if position >= 0 else "",
            "header": _safe_cell_str(headers[position]) if 0 <= position < len(headers or []) else "",
            "found": position >= 0,
        }
    return output


def _row_field_values_debug(row: Sequence[object], headers: Sequence[object], indexes: dict) -> dict:
    output = {}
    for name, index in indexes.items():
        try:
            position = int(index)
        except (TypeError, ValueError):
            position = -1
        output[name] = {
            "index": position,
            "column": _column_index_to_name(position + 1) if position >= 0 else "",
            "header": _safe_cell_str(headers[position]) if 0 <= position < len(headers or []) else "",
            "value": _safe_cell_str(row[position]) if 0 <= position < len(row or []) else "",
        }
    return output


def build_sheet_value_range(sheet_id: str, row_count: int, column_range: str = "A:Z") -> str:
    sheet_id = (sheet_id or "").strip()
    if not sheet_id:
        raise FeishuBomError("sheet_id 不能为空。")

    column_range = (column_range or "A:Z").strip().upper()
    match = re.fullmatch(r"([A-Z]+)\s*:\s*([A-Z]+)", column_range)
    if not match:
        raise FeishuBomError("column_range 必须是类似 A:Z 或 A:AZ 的列范围。")

    safe_row_count = min(max(int(row_count or 0), 50), 10000)
    start_col, end_col = match.groups()
    return f"{sheet_id}!{start_col}1:{end_col}{safe_row_count}"


def _parse_a1_range_start(value: object) -> tuple[int, int]:
    text = _safe_cell_str(value)
    if not text:
        return 0, 0
    coord = text.rsplit("!", 1)[-1].strip().upper()
    match = re.match(r"^\$?([A-Z]+)\$?(\d+)", coord)
    if not match:
        return 0, 0
    return _column_name_to_index(match.group(1)), int(match.group(2))


def _parse_a1_range_bounds(value: object) -> tuple[int, int, int, int]:
    text = _safe_cell_str(value)
    if not text:
        return 0, 0, 0, 0
    coord = text.rsplit("!", 1)[-1].strip().upper()
    match = re.match(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", coord)
    if not match:
        start_col, start_row = _parse_a1_range_start(value)
        return start_col, start_row, start_col, start_row
    start_col = _column_name_to_index(match.group(1))
    start_row = int(match.group(2))
    end_col = _column_name_to_index(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    return start_col, start_row, end_col, end_row


def _align_values_to_requested_range(values: Sequence[object],
                                     *,
                                     requested_range: str,
                                     returned_range: str) -> tuple[List[List[object]], dict]:
    requested_col, requested_row, requested_end_col, _requested_end_row = _parse_a1_range_bounds(requested_range)
    returned_range_text = _safe_cell_str(returned_range)
    returned_col, returned_row = _parse_a1_range_start(returned_range_text)
    column_offset = max(0, returned_col - requested_col) if requested_col and returned_col else 0
    row_offset = max(0, returned_row - requested_row) if requested_row and returned_row else 0
    requested_width = max(0, requested_end_col - requested_col + 1) if requested_col and requested_end_col else 0
    aligned_rows: List[List[object]] = []
    for row in values or []:
        if isinstance(row, (list, tuple)):
            aligned = [""] * column_offset + list(row)
        else:
            aligned = [""] * column_offset + [row]
        if requested_width and len(aligned) < requested_width:
            aligned = aligned + [""] * (requested_width - len(aligned))
        aligned_rows.append(aligned)
    if row_offset and aligned_rows:
        width = max(requested_width, max((len(row) for row in aligned_rows), default=column_offset))
        aligned_rows = [[""] * width for _ in range(row_offset)] + aligned_rows
    return aligned_rows, {
        "requested_range": requested_range,
        "returned_range": _safe_cell_str(returned_range),
        "requested_start": {"column": requested_col, "row": requested_row},
        "returned_start": {"column": returned_col, "row": returned_row},
        "requested_width": requested_width,
        "column_offset": column_offset,
        "row_offset": row_offset,
        "applied": bool(column_offset or row_offset),
    }


def _pad_row(row: Sequence[object], width: int) -> List[object]:
    values = list(row or [])
    safe_width = max(0, int(width or 0))
    if safe_width and len(values) < safe_width:
        values.extend([""] * (safe_width - len(values)))
    return values


def resolve_data_dir(data_dir: str = "") -> Path:
    raw = str(data_dir or "").strip().strip('"')
    if not raw:
        raw = os.environ.get(DATA_DIR_ENV, "").strip().strip('"')
    if raw:
        return Path(raw).expanduser()
    return BASE_DIR / DEFAULT_DATA_DIR_NAME


def _config_path(data_dir: Path) -> Path:
    return data_dir / CONFIG_FILE_NAME


def _cache_path(data_dir: Path) -> Path:
    return data_dir / CACHE_FILE_NAME


def _default_config() -> dict:
    return {
        "base_url": DEFAULT_BASE_URL,
        "origin": DEFAULT_ORIGIN,
        "user_id": "",
        "libraries": [],
    }


def _load_config(config_path: Path) -> dict:
    data = _default_config()
    if not config_path.is_file():
        return data
    try:
        with config_path.open("r", encoding="utf-8") as handle:
            loaded = json.load(handle)
        if isinstance(loaded, dict):
            data.update(loaded)
    except Exception:
        pass
    if not _safe_cell_str(data.get("base_url")) or "example.com" in _safe_cell_str(data.get("base_url")):
        data["base_url"] = DEFAULT_BASE_URL
    if not _safe_cell_str(data.get("origin")):
        data["origin"] = DEFAULT_ORIGIN
    if not isinstance(data.get("libraries"), list):
        data["libraries"] = []
    return data


def _save_config(root: Path, config: dict) -> None:
    root.mkdir(parents=True, exist_ok=True)
    with _config_path(root).open("w", encoding="utf-8") as handle:
        json.dump(config, handle, ensure_ascii=False, indent=2)


def _is_configured(config: dict) -> bool:
    base_url = _safe_cell_str(config.get("base_url", ""))
    origin = _safe_cell_str(config.get("origin", ""))
    user_id = _safe_cell_str(config.get("user_id", ""))
    return base_url.startswith("http") and bool(origin) and bool(user_id)


def _resolve_connection(*,
                        base_url: str = "",
                        origin: str = "",
                        user_id: str = "",
                        data_dir: str = "") -> tuple[str, str, str, Path, dict]:
    root = resolve_data_dir(data_dir)
    config = _load_config(_config_path(root))
    resolved_base_url = _safe_cell_str(base_url) or _safe_cell_str(config.get("base_url")) or DEFAULT_BASE_URL
    resolved_origin = _safe_cell_str(origin) or _safe_cell_str(config.get("origin")) or DEFAULT_ORIGIN
    resolved_user_id = _safe_cell_str(user_id) or _safe_cell_str(config.get("user_id"))
    if not resolved_base_url.startswith("http"):
        raise FeishuBomError("飞书网关地址必须以 http/https 开头。")
    if not resolved_origin:
        raise FeishuBomError("缺少 origin / AppId。")
    if not resolved_user_id:
        raise FeishuBomError("缺少工号 user_id。")
    return resolved_base_url.rstrip("/"), resolved_origin, resolved_user_id, root, config
