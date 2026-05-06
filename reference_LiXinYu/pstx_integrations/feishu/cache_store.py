# -*- coding: utf-8 -*-
"""SQLite cache, local CRUD, and project-row matching for Feishu BOM data."""

from __future__ import annotations

import json
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence

from pstx_integrations.feishu.common import (
    CACHE_FILE_NAME,
    DATA_DIR_ENV,
    FEISHU_PARSE_LOG_FILE_ENV,
    HQ_FIELD_ALIASES,
    SPEC_FIELD_ALIASES,
    FeishuBomError,
    _cache_path,
    _config_path,
    _is_configured,
    _load_config,
    _normalized_name,
    _safe_cell_str,
    _save_config,
    feishu_debug_log_path,
    feishu_parse_debug_log_path,
    resolve_data_dir,
)
from pstx_integrations.feishu.mapping import _normalize_optional_fields, _ordered_unique


def _connect_cache(cache_path: Path) -> Optional[sqlite3.Connection]:
    if not cache_path.is_file():
        return None
    conn = sqlite3.connect(str(cache_path))
    conn.row_factory = sqlite3.Row
    _ensure_materials_schema(conn)
    return conn


def _connect_cache_for_write(cache_path: Path) -> sqlite3.Connection:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(cache_path))
    conn.row_factory = sqlite3.Row
    _ensure_materials_schema(conn)
    return conn


def _ensure_materials_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lib_id TEXT,
            lib_name TEXT,
            sheet_name TEXT,
            key_value TEXT COLLATE NOCASE,
            hq_no TEXT,
            brand TEXT,
            spec TEXT,
            description TEXT,
            pi TEXT,
            selection_order TEXT,
            extra_fields TEXT,
            raw_data TEXT,
            synced_at TEXT
        )
        """
    )
    existing = _table_columns(conn, "materials")
    for column_name, column_type in {
        "pi": "TEXT",
        "selection_order": "TEXT",
        "extra_fields": "TEXT",
    }.items():
        if column_name not in existing:
            conn.execute(f"ALTER TABLE materials ADD COLUMN {column_name} {column_type}")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_key ON materials(key_value COLLATE NOCASE)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_hq_no ON materials(hq_no COLLATE NOCASE)")
    conn.commit()


def _table_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    try:
        rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    except sqlite3.DatabaseError:
        return set()
    return {_safe_cell_str(row[1]) for row in rows}


def _table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone()
    return row is not None


def _cache_count(conn: Optional[sqlite3.Connection]) -> int:
    if conn is None or not _table_exists(conn, "materials"):
        return 0
    return int(conn.execute("SELECT COUNT(*) FROM materials").fetchone()[0] or 0)


def _cache_stats(conn: Optional[sqlite3.Connection]) -> List[dict]:
    if conn is None or not _table_exists(conn, "materials"):
        return []
    rows = conn.execute(
        "SELECT lib_id, lib_name, COUNT(*) AS count FROM materials GROUP BY lib_id, lib_name"
    ).fetchall()
    return [
        {
            "lib_id": _safe_cell_str(row["lib_id"]),
            "lib_name": _safe_cell_str(row["lib_name"]),
            "count": int(row["count"] or 0),
        }
        for row in rows
    ]


def _library_id_from_name(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", _safe_cell_str(name)).strip("_")
    return f"lib_{cleaned or int(time.time() * 1000)}"


def _upsert_library_config(config: dict,
                           *,
                           library_id: str,
                           library_name: str,
                           token: str,
                           sheets: Sequence[dict],
                           last_sync: str) -> None:
    libraries = list(config.get("libraries") or [])
    payload = {
        "id": library_id,
        "name": library_name,
        "token": token,
        "sheets": [dict(sheet) for sheet in sheets],
        "last_sync": last_sync,
    }
    replaced = False
    for index, library in enumerate(libraries):
        if _safe_cell_str(library.get("id")) == library_id:
            libraries[index] = payload
            replaced = True
            break
    if not replaced:
        libraries.append(payload)
    config["libraries"] = libraries


def build_feishu_bom_status(*, data_dir: str = "", load_runtime: bool = True) -> dict:
    root = resolve_data_dir(data_dir)
    config_path = _config_path(root)
    cache_path = _cache_path(root)
    config = _load_config(config_path)
    libraries = list(config.get("libraries", []) or []) if isinstance(config.get("libraries", []), list) else []

    status = {
        "ok": True,
        "available": root.is_dir() or config_path.is_file() or cache_path.is_file(),
        "data_dir": str(root),
        "config_file": str(config_path),
        "cache_file": str(cache_path),
        "online_debug_log_file": feishu_debug_log_path(),
        "online_parse_debug_log_file": feishu_parse_debug_log_path(),
        "configured": _is_configured(config),
        "library_count": len(libraries),
        "cache_count": 0,
        "cache_stats": [],
        "capabilities": ["status", "cache_match_preview", "online_sheet_list", "online_sync", "online_debug_log", "online_parse_debug_log"],
        "error": "",
    }
    if not status["available"]:
        status["error"] = f"未找到飞书 BOM 数据目录：{root}。可通过 {DATA_DIR_ENV} 指定。"
        return status
    if not load_runtime:
        return status

    conn = None
    try:
        conn = _connect_cache(cache_path)
        status["cache_count"] = _cache_count(conn)
        status["cache_stats"] = _cache_stats(conn)
    except Exception as exc:
        status["ok"] = False
        status["error"] = f"飞书 BOM 缓存状态读取失败：{exc}"
    finally:
        if conn is not None:
            conn.close()
    return status


def get_saved_feishu_field_order(*, data_dir: str = "") -> dict:
    """Return optional field titles seen in saved configs/cache, preserving first-seen order."""

    root = resolve_data_dir(data_dir)
    config = _load_config(_config_path(root))
    optional_titles: List[str] = []
    for library in config.get("libraries") or []:
        for sheet in library.get("sheets") or []:
            if not isinstance(sheet, dict):
                continue
            for field in _normalize_optional_fields(sheet.get("optional_fields")):
                optional_titles.append(field["label"])

    conn = _connect_cache(_cache_path(root))
    if conn is not None:
        try:
            if _table_exists(conn, "materials"):
                for row in conn.execute("SELECT raw_data, extra_fields FROM materials LIMIT 500").fetchall():
                    for key in _parse_raw_data(_safe_cell_str(row["extra_fields"])).keys():
                        optional_titles.append(key)
                    for key in _parse_raw_data(_safe_cell_str(row["raw_data"])).keys():
                        if key not in {"HQ料号", "HQ编码", "物料编码", "规格型号", "Part Number", "PI", "Pi", "选型顺序"}:
                            optional_titles.append(key)
        finally:
            conn.close()

    return {
        "ok": True,
        "data_dir": str(root),
        "optional_field_order": _ordered_unique(optional_titles),
    }


def _material_row_to_match(row: sqlite3.Row) -> dict:
    keys = set(row.keys()) if hasattr(row, "keys") else set()

    def get(name: str) -> str:
        return _safe_cell_str(row[name]) if name in keys else ""

    try:
        cache_row_id = int(row["id"] or 0) if "id" in keys else 0
    except (TypeError, ValueError):
        cache_row_id = 0
    key_value = get("key_value")
    hq_no = get("hq_no")
    spec = get("spec")
    return {
        "缓存行ID": cache_row_id,
        "来源库": get("lib_name"),
        "来源Sheet": get("sheet_name"),
        "匹配关键值": key_value,
        "规格型号": key_value,
        "HQ料号": hq_no,
        "飞书HQ料号": hq_no,
        "HQ制造商": get("brand"),
        "HQ规格型号": spec,
        "飞书规格型号": spec or key_value,
        "HQ描述": get("description"),
        "PI": get("pi"),
        "选型顺序": get("selection_order"),
        "扩展字段": get("extra_fields"),
        "原始数据": get("raw_data"),
    }


def _search_materials(conn: sqlite3.Connection, key_value: str, *, match_mode: str = "key_value") -> List[dict]:
    if not _table_exists(conn, "materials"):
        return []
    column_name = "hq_no" if match_mode == "hq_no" else "key_value"
    rows = conn.execute(
        "SELECT id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data "
        f"FROM materials WHERE {column_name}=? COLLATE NOCASE",
        (key_value,),
    ).fetchall()
    return [_material_row_to_match(row) for row in rows]


def _parse_raw_data(raw_data: str) -> dict:
    try:
        loaded = json.loads(raw_data or "{}")
        return loaded if isinstance(loaded, dict) else {}
    except Exception:
        return {}


def _json_object_text(value: object, *, field_name: str) -> str:
    if value is None or value == "":
        return "{}"
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, list):
        raise FeishuBomError(f"{field_name} 必须是 JSON 对象。")
    text = _safe_cell_str(value)
    if not text:
        return "{}"
    try:
        loaded = json.loads(text)
    except json.JSONDecodeError as exc:
        raise FeishuBomError(f"{field_name} 必须是合法 JSON 对象。") from exc
    if not isinstance(loaded, dict):
        raise FeishuBomError(f"{field_name} 必须是 JSON 对象。")
    return json.dumps(loaded, ensure_ascii=False)


def _build_manual_raw_data(payload: dict, extra_fields_text: str, raw_data: object = None) -> str:
    if raw_data not in (None, ""):
        return _json_object_text(raw_data, field_name="raw_data")
    raw_fields = {
        "规格型号": _safe_cell_str(payload.get("key_value")),
        "HQ料号": _safe_cell_str(payload.get("hq_no")),
        "制造商": _safe_cell_str(payload.get("brand")),
        "规格": _safe_cell_str(payload.get("spec")),
        "描述": _safe_cell_str(payload.get("description")),
        "PI": _safe_cell_str(payload.get("pi")),
        "选型顺序": _safe_cell_str(payload.get("selection_order")),
    }
    raw_fields.update(_parse_raw_data(extra_fields_text))
    return json.dumps({key: value for key, value in raw_fields.items() if value}, ensure_ascii=False)


def _material_row_to_cache_preview(row: sqlite3.Row) -> dict:
    raw_data = _safe_cell_str(row["raw_data"])
    return {
        "id": int(row["id"] or 0),
        "lib_id": _safe_cell_str(row["lib_id"]),
        "lib_name": _safe_cell_str(row["lib_name"]),
        "sheet_name": _safe_cell_str(row["sheet_name"]),
        "key_value": _safe_cell_str(row["key_value"]),
        "hq_no": _safe_cell_str(row["hq_no"]),
        "brand": _safe_cell_str(row["brand"]),
        "spec": _safe_cell_str(row["spec"]),
        "description": _safe_cell_str(row["description"]),
        "pi": _safe_cell_str(row["pi"]),
        "selection_order": _safe_cell_str(row["selection_order"]),
        "extra_fields": _safe_cell_str(row["extra_fields"]),
        "extra_field_values": _parse_raw_data(_safe_cell_str(row["extra_fields"])),
        "synced_at": _safe_cell_str(row["synced_at"]),
        "raw_data": raw_data,
        "raw_fields": _parse_raw_data(raw_data),
    }


def build_feishu_database_overview(*, data_dir: str = "") -> dict:
    root = resolve_data_dir(data_dir)
    config_path = _config_path(root)
    cache_path = _cache_path(root)
    config = _load_config(config_path)
    libraries = list(config.get("libraries") or [])
    conn = _connect_cache(cache_path)
    stats_by_lib: dict[str, dict] = {}
    sheet_stats_by_lib: dict[str, list] = {}
    if conn is not None:
        try:
            if _table_exists(conn, "materials"):
                for row in conn.execute(
                    "SELECT lib_id, lib_name, COUNT(*) AS count, MAX(synced_at) AS synced_at "
                    "FROM materials GROUP BY lib_id, lib_name"
                ).fetchall():
                    lib_id = _safe_cell_str(row["lib_id"])
                    stats_by_lib[lib_id] = {
                        "lib_id": lib_id,
                        "lib_name": _safe_cell_str(row["lib_name"]),
                        "cache_count": int(row["count"] or 0),
                        "last_synced_at": _safe_cell_str(row["synced_at"]),
                    }
                for row in conn.execute(
                    "SELECT lib_id, sheet_name, COUNT(*) AS count, MAX(synced_at) AS synced_at "
                    "FROM materials GROUP BY lib_id, sheet_name ORDER BY lib_id, sheet_name"
                ).fetchall():
                    lib_id = _safe_cell_str(row["lib_id"])
                    sheet_stats_by_lib.setdefault(lib_id, []).append({
                        "sheet_name": _safe_cell_str(row["sheet_name"]),
                        "count": int(row["count"] or 0),
                        "last_synced_at": _safe_cell_str(row["synced_at"]),
                    })
        finally:
            conn.close()

    output = []
    seen = set()
    for library in libraries:
        lib_id = _safe_cell_str(library.get("id"))
        if not lib_id:
            continue
        seen.add(lib_id)
        sheet_configs = list(library.get("sheets") or [])
        stat = stats_by_lib.get(lib_id, {})
        output.append({
            "lib_id": lib_id,
            "lib_name": _safe_cell_str(library.get("name")),
            "cache_count": int(stat.get("cache_count") or 0),
            "last_synced_at": _safe_cell_str(library.get("last_sync") or stat.get("last_synced_at")),
            "sheet_config_count": len(sheet_configs),
            "enabled_sheet_count": len([sheet for sheet in sheet_configs if str(sheet.get("enabled", True)).lower() not in {"0", "false", "no", "off"}]),
            "configured_sheets": [
                {
                    "sheet_id": _safe_cell_str(sheet.get("sheet_id") or sheet.get("sheetId")),
                    "title": _safe_cell_str(sheet.get("title") or sheet.get("sheet_title")),
                    "header_row": int(sheet.get("header_row") or 1),
                    "hq_code_col": _safe_cell_str(sheet.get("hq_code_col") or sheet.get("hq_no_col")),
                    "spec_model_col": _safe_cell_str(sheet.get("spec_model_col") or sheet.get("key_col") or sheet.get("spec_col")),
                    "pi_col": _safe_cell_str(sheet.get("pi_col")),
                    "selection_order_col": _safe_cell_str(sheet.get("selection_order_col")),
                    "optional_fields": _normalize_optional_fields(sheet.get("optional_fields")),
                    "key_col": _safe_cell_str(sheet.get("key_col") or sheet.get("spec_model_col")),
                    "hq_no_col": _safe_cell_str(sheet.get("hq_no_col") or sheet.get("hq_code_col")),
                    "brand_col": _safe_cell_str(sheet.get("brand_col")),
                    "spec_col": _safe_cell_str(sheet.get("spec_col")),
                    "desc_col": _safe_cell_str(sheet.get("desc_col")),
                }
                for sheet in sheet_configs
            ],
            "sheet_stats": sheet_stats_by_lib.get(lib_id, []),
            "token_present": bool(_safe_cell_str(library.get("token"))),
            "token_length": len(_safe_cell_str(library.get("token"))),
        })

    for lib_id, stat in stats_by_lib.items():
        if lib_id in seen:
            continue
        output.append({
            "lib_id": lib_id,
            "lib_name": _safe_cell_str(stat.get("lib_name")),
            "cache_count": int(stat.get("cache_count") or 0),
            "last_synced_at": _safe_cell_str(stat.get("last_synced_at")),
            "sheet_config_count": 0,
            "enabled_sheet_count": 0,
            "configured_sheets": [],
            "sheet_stats": sheet_stats_by_lib.get(lib_id, []),
            "token_present": False,
            "token_length": 0,
        })

    status = build_feishu_bom_status(data_dir=data_dir)
    status.update({
        "libraries": output,
        "saved_field_order": get_saved_feishu_field_order(data_dir=data_dir).get("optional_field_order", []),
    })
    return status


def get_feishu_cache_rows(*,
                          lib_id: str = "",
                          sheet_name: str = "",
                          query: str = "",
                          data_dir: str = "",
                          limit: int = 100,
                          offset: int = 0) -> dict:
    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache(cache_path)
    if conn is None or not _table_exists(conn, "materials"):
        return {
            "ok": False,
            "error": f"未找到飞书 BOM 缓存：{cache_path}。",
            "rows": [],
        }

    conditions = []
    params: List[object] = []
    lib_id = _safe_cell_str(lib_id)
    sheet_name = _safe_cell_str(sheet_name)
    query = _safe_cell_str(query)
    if lib_id:
        conditions.append("lib_id = ?")
        params.append(lib_id)
    if sheet_name:
        conditions.append("sheet_name = ?")
        params.append(sheet_name)
    if query:
        conditions.append("(key_value LIKE ? OR hq_no LIKE ? OR brand LIKE ? OR spec LIKE ? OR description LIKE ? OR pi LIKE ? OR selection_order LIKE ? OR extra_fields LIKE ?)")
        like = f"%{query}%"
        params.extend([like, like, like, like, like, like, like, like])
    where_clause = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    safe_limit = min(max(int(limit or 100), 1), 5000)
    safe_offset = max(int(offset or 0), 0)

    try:
        total = int(conn.execute(f"SELECT COUNT(*) FROM materials {where_clause}", params).fetchone()[0] or 0)
        rows = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data, synced_at "
            f"FROM materials {where_clause} ORDER BY lib_name, sheet_name, key_value LIMIT ? OFFSET ?",
            params + [safe_limit, safe_offset],
        ).fetchall()
        return {
            "ok": True,
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "total": total,
            "limit": safe_limit,
            "offset": safe_offset,
            "has_more": safe_offset + len(rows) < total,
            "next_offset": safe_offset + len(rows),
            "rows": [_material_row_to_cache_preview(row) for row in rows],
        }
    finally:
        conn.close()


def get_feishu_cache_row(row_id: int, *, data_dir: str = "") -> dict:
    try:
        safe_row_id = int(row_id)
    except (TypeError, ValueError) as exc:
        raise FeishuBomError("row_id 必须是数字。") from exc
    if safe_row_id <= 0:
        raise FeishuBomError("row_id 必须大于 0。")

    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache(cache_path)
    if conn is None or not _table_exists(conn, "materials"):
        return {
            "ok": False,
            "error": f"未找到飞书 BOM 缓存：{cache_path}。",
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "row": None,
        }

    try:
        row = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data, synced_at "
            "FROM materials WHERE id=?",
            (safe_row_id,),
        ).fetchone()
        if row is None:
            return {
                "ok": False,
                "error": f"未找到缓存行 id={safe_row_id}。",
                "data_dir": str(root),
                "cache_file": str(cache_path),
                "row_id": safe_row_id,
                "row": None,
            }
        return {
            "ok": True,
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "row_id": safe_row_id,
            "row": _material_row_to_cache_preview(row),
        }
    finally:
        conn.close()


def create_feishu_cache_row(payload: dict, *, data_dir: str = "") -> dict:
    if not isinstance(payload, dict):
        raise FeishuBomError("请求体必须是 JSON 对象。")
    lib_id = _safe_cell_str(payload.get("lib_id") or "manual")
    lib_name = _safe_cell_str(payload.get("lib_name") or lib_id or "手工维护")
    sheet_name = _safe_cell_str(payload.get("sheet_name") or "手工维护")
    key_value = _safe_cell_str(payload.get("key_value"))
    hq_no = _safe_cell_str(payload.get("hq_no"))
    if not key_value and not hq_no:
        raise FeishuBomError("规格型号或 HQ 料号至少填写一项。")
    if not key_value:
        key_value = hq_no
    extra_fields = _json_object_text(payload.get("extra_fields"), field_name="extra_fields")
    raw_data = _build_manual_raw_data(payload, extra_fields, payload.get("raw_data"))
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache_for_write(cache_path)
    try:
        cur = conn.execute(
            "INSERT INTO materials(lib_id,lib_name,sheet_name,key_value,hq_no,brand,spec,description,pi,selection_order,extra_fields,raw_data,synced_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                lib_id,
                lib_name,
                sheet_name,
                key_value,
                hq_no,
                _safe_cell_str(payload.get("brand")),
                _safe_cell_str(payload.get("spec") or key_value),
                _safe_cell_str(payload.get("description")),
                _safe_cell_str(payload.get("pi")),
                _safe_cell_str(payload.get("selection_order")),
                extra_fields,
                raw_data,
                now,
            ),
        )
        conn.commit()
        row_id = int(cur.lastrowid)
        row = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data, synced_at "
            "FROM materials WHERE id=?",
            (row_id,),
        ).fetchone()
        return {
            "ok": True,
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "row_id": row_id,
            "row": _material_row_to_cache_preview(row),
        }
    finally:
        conn.close()


def update_feishu_cache_row(row_id: int, payload: dict, *, data_dir: str = "") -> dict:
    try:
        safe_row_id = int(row_id)
    except (TypeError, ValueError) as exc:
        raise FeishuBomError("row_id 必须是数字。") from exc
    if safe_row_id <= 0:
        raise FeishuBomError("row_id 必须大于 0。")
    if not isinstance(payload, dict):
        raise FeishuBomError("请求体必须是 JSON 对象。")

    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache_for_write(cache_path)
    try:
        existing = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data, synced_at "
            "FROM materials WHERE id=?",
            (safe_row_id,),
        ).fetchone()
        if existing is None:
            return {
                "ok": False,
                "error": f"未找到缓存行 id={safe_row_id}。",
                "data_dir": str(root),
                "cache_file": str(cache_path),
                "row_id": safe_row_id,
                "row": None,
            }
        current = _material_row_to_cache_preview(existing)
        merged = dict(current)
        for key in [
            "lib_id", "lib_name", "sheet_name", "key_value", "hq_no",
            "brand", "spec", "description", "pi", "selection_order",
        ]:
            if key in payload:
                merged[key] = _safe_cell_str(payload.get(key))
        if not merged.get("key_value") and not merged.get("hq_no"):
            raise FeishuBomError("规格型号或 HQ 料号至少填写一项。")
        if not merged.get("key_value"):
            merged["key_value"] = merged.get("hq_no", "")
        extra_fields = (
            _json_object_text(payload.get("extra_fields"), field_name="extra_fields")
            if "extra_fields" in payload
            else _safe_cell_str(existing["extra_fields"])
        )
        raw_data = _build_manual_raw_data(merged, extra_fields, payload.get("raw_data") if "raw_data" in payload else existing["raw_data"])
        conn.execute(
            "UPDATE materials SET lib_id=?, lib_name=?, sheet_name=?, key_value=?, hq_no=?, brand=?, spec=?, description=?, pi=?, selection_order=?, extra_fields=?, raw_data=?, synced_at=? "
            "WHERE id=?",
            (
                merged.get("lib_id", ""),
                merged.get("lib_name", ""),
                merged.get("sheet_name", ""),
                merged.get("key_value", ""),
                merged.get("hq_no", ""),
                merged.get("brand", ""),
                merged.get("spec", ""),
                merged.get("description", ""),
                merged.get("pi", ""),
                merged.get("selection_order", ""),
                extra_fields,
                raw_data,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                safe_row_id,
            ),
        )
        conn.commit()
        row = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no, brand, spec, description, pi, selection_order, extra_fields, raw_data, synced_at "
            "FROM materials WHERE id=?",
            (safe_row_id,),
        ).fetchone()
        return {
            "ok": True,
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "row_id": safe_row_id,
            "row": _material_row_to_cache_preview(row),
        }
    finally:
        conn.close()


def delete_feishu_cache_library(lib_id: str, *, data_dir: str = "") -> dict:
    lib_id = _safe_cell_str(lib_id)
    if not lib_id:
        raise FeishuBomError("缺少 lib_id。")
    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    deleted_rows = 0
    conn = _connect_cache_for_write(cache_path)
    try:
        cur = conn.execute("DELETE FROM materials WHERE lib_id=?", (lib_id,))
        deleted_rows = int(cur.rowcount or 0)
        conn.commit()
    finally:
        conn.close()

    config = _load_config(_config_path(root))
    before = len(config.get("libraries") or [])
    config["libraries"] = [
        library
        for library in (config.get("libraries") or [])
        if _safe_cell_str(library.get("id")) != lib_id
    ]
    removed_config = before - len(config["libraries"])
    _save_config(root, config)
    return {
        "ok": True,
        "data_dir": str(root),
        "cache_file": str(cache_path),
        "lib_id": lib_id,
        "deleted_rows": deleted_rows,
        "removed_config": removed_config,
    }


def delete_feishu_cache_row(row_id: int, *, data_dir: str = "") -> dict:
    try:
        safe_row_id = int(row_id)
    except (TypeError, ValueError) as exc:
        raise FeishuBomError("row_id 必须是数字。") from exc
    if safe_row_id <= 0:
        raise FeishuBomError("row_id 必须大于 0。")

    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache_for_write(cache_path)
    try:
        target = conn.execute(
            "SELECT id, lib_id, lib_name, sheet_name, key_value, hq_no "
            "FROM materials WHERE id=?",
            (safe_row_id,),
        ).fetchone()
        if target is None:
            return {
                "ok": False,
                "error": f"未找到缓存行 id={safe_row_id}。",
                "data_dir": str(root),
                "cache_file": str(cache_path),
                "row_id": safe_row_id,
                "deleted_rows": 0,
            }
        cur = conn.execute("DELETE FROM materials WHERE id=?", (safe_row_id,))
        conn.commit()
        return {
            "ok": True,
            "data_dir": str(root),
            "cache_file": str(cache_path),
            "row_id": safe_row_id,
            "deleted_rows": int(cur.rowcount or 0),
            "deleted_row": {
                "id": int(target["id"] or 0),
                "lib_id": _safe_cell_str(target["lib_id"]),
                "lib_name": _safe_cell_str(target["lib_name"]),
                "sheet_name": _safe_cell_str(target["sheet_name"]),
                "key_value": _safe_cell_str(target["key_value"]),
                "hq_no": _safe_cell_str(target["hq_no"]),
            },
        }
    finally:
        conn.close()


def _is_hq_key_field(key_field: str) -> bool:
    return _normalized_name(key_field) in HQ_FIELD_ALIASES


def _is_spec_key_field(key_field: str) -> bool:
    return _normalized_name(key_field) in SPEC_FIELD_ALIASES


def _resolve_match_mode(key_field: str, match_mode: str) -> str:
    raw = _normalized_name(match_mode or "auto")
    if raw in {"hq", "hq no", "hq_no", "hq code", "hq_code", "hq料号", "料号"}:
        return "hq_no"
    if raw in {"key", "key value", "key_value", "spec", "规格", "规格型号"}:
        return "key_value"
    return "hq_no" if _is_hq_key_field(key_field) else "key_value"


def _row_value_by_alias(row: dict, key_field: str) -> str:
    if key_field in row:
        return _safe_cell_str(row.get(key_field))
    target = _normalized_name(key_field)
    for name, value in row.items():
        if _normalized_name(name) == target:
            return _safe_cell_str(value)
    if _is_hq_key_field(key_field):
        return _row_project_hq_value(row)
    if _is_spec_key_field(key_field):
        return _row_project_spec_value(row)
    return ""


def _row_project_hq_value(row: dict) -> str:
    for name in ("料号", "HQ料号", "HQ_CODE", "hq_code", "HQ编码", "物料编码", "物料号"):
        value = _safe_cell_str(row.get(name))
        if value:
            return value
    return ""


def _row_project_spec_value(row: dict) -> str:
    for name in ("规格型号", "Part Number", "part_number", "描述", "值", "part_name"):
        value = _safe_cell_str(row.get(name))
        if value:
            return value
    return ""


def match_rows_with_feishu_cache(rows: Sequence[dict],
                                 key_field: str,
                                 *,
                                 data_dir: str = "",
                                 limit: int = 200,
                                 match_mode: str = "auto") -> dict:
    key_field = _safe_cell_str(key_field)
    if not key_field:
        return {"ok": False, "error": "缺少 key_field。", "rows": []}
    resolved_match_mode = _resolve_match_mode(key_field, match_mode)
    match_mode_label = "HQ料号直连" if resolved_match_mode == "hq_no" else "关键值匹配"

    root = resolve_data_dir(data_dir)
    cache_path = _cache_path(root)
    conn = _connect_cache(cache_path)
    if conn is None:
        return {
            "ok": False,
            "error": f"未找到飞书 BOM 缓存：{cache_path}。可通过 {DATA_DIR_ENV} 指定。",
            "rows": [],
        }

    checked_rows = list(rows or [])[:max(0, int(limit or 0))]
    output_rows: List[dict] = []
    matched_count = 0
    unmatched_count = 0
    skipped_count = 0
    try:
        for index, row in enumerate(checked_rows, start=1):
            key_value = _row_value_by_alias(row, key_field)
            base = {
                "序号": index,
                "位号": _safe_cell_str(row.get("位号", "")),
                "BOM状态": _safe_cell_str(row.get("BOM状态", "")),
                "项目HQ料号": _row_project_hq_value(row),
                "项目规格型号": _row_project_spec_value(row),
                "项目值": _safe_cell_str(row.get("值", "")),
                "项目封装": _safe_cell_str(row.get("封装", "")),
                "项目类型": _safe_cell_str(row.get("类型", "")),
                "匹配字段": key_field,
                "匹配方式": match_mode_label,
                "匹配关键值": key_value,
                "匹配状态": "",
                "匹配数量": 0,
            }
            if not key_value:
                skipped_count += 1
                base["匹配状态"] = "跳过：关键值为空"
                output_rows.append(base)
                continue

            matches = _search_materials(conn, key_value, match_mode=resolved_match_mode)
            if not matches:
                unmatched_count += 1
                base["匹配状态"] = "未匹配"
                output_rows.append(base)
                continue

            matched_count += 1
            first = matches[0]
            base.update({
                "匹配状态": "已匹配",
                "匹配数量": len(matches),
                "缓存行ID": first.get("缓存行ID", ""),
                "HQ料号": first.get("HQ料号", ""),
                "飞书HQ料号": first.get("飞书HQ料号", first.get("HQ料号", "")),
                "HQ制造商": first.get("HQ制造商", ""),
                "HQ规格型号": first.get("HQ规格型号", ""),
                "飞书规格型号": first.get("飞书规格型号", first.get("HQ规格型号", "")),
                "HQ描述": first.get("HQ描述", ""),
                "PI": first.get("PI", ""),
                "选型顺序": first.get("选型顺序", ""),
                "来源库": first.get("来源库", ""),
                "来源Sheet": first.get("来源Sheet", ""),
                "全部匹配": matches,
            })
            output_rows.append(base)
    finally:
        conn.close()

    return {
        "ok": True,
        "data_dir": str(root),
        "cache_file": str(cache_path),
        "key_field": key_field,
        "match_mode": resolved_match_mode,
        "match_mode_label": match_mode_label,
        "total_rows": len(rows or []),
        "checked_rows": len(checked_rows),
        "matched_count": matched_count,
        "unmatched_count": unmatched_count,
        "skipped_count": skipped_count,
        "rows": output_rows,
    }
