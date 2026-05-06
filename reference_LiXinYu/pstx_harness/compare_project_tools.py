# -*- coding: utf-8 -*-
"""Project file and Cadence page tools for the compare agent."""

from __future__ import annotations

import re
from pathlib import Path
from typing import List

from pstx_core.cadence.page_model import compare_page_models, load_cadence_page_model
from pstx_harness.tool_core import HarnessToolError


BATCH_MAX_ITEMS = 20
ALLOWED_COMPARE_PROJECT_DIRS = {"packaged", "sch_1"}
ALLOWED_COMPARE_ROOT_FILES = {"module_order", "module_order.dat", "page.map"}
ALLOWED_COMPARE_TEXT_SUFFIXES = {".dat", ".csv", ".csa", ".map", ".txt"}
COMPARE_PAGE_RANGE_HARD_MAX = 60

SENSITIVE_TEXT_RE = re.compile(
    r"(?i)(secret|token|apikey|api_key|appsecret|app_secret|authorization|ciphertext|password)"
    r"(\s*[:=]\s*)"
    r"([^\s,;'\"]+)"
)


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value if value is not None else default)
    except (TypeError, ValueError):
        return default


def _redact_sensitive_text(text: str) -> str:
    return SENSITIVE_TEXT_RE.sub(lambda match: f"{match.group(1)}{match.group(2)}<redacted>", text)


def _batch_input_items(args: dict, key: str, *, max_items: int = BATCH_MAX_ITEMS) -> tuple[List[object], bool]:
    raw = args.get(key)
    if not isinstance(raw, list):
        raise HarnessToolError(f"批量 compare 工具需要数组参数：{key}。")
    items = list(raw)
    return items[:max_items], len(items) > max_items


def _batch_summary(title: str, items: List[dict], *, truncated: bool = False) -> str:
    found = sum(1 for item in items if item.get("status") == "found")
    missing = sum(1 for item in items if item.get("status") == "missing")
    errors = sum(1 for item in items if item.get("status") == "error")
    suffix = "，输入已按上限截断" if truncated else ""
    return f"{title}完成：{len(items)} 项，命中 {found} 项，缺失 {missing} 项，错误 {errors} 项{suffix}。"


def _side_payload(context, side: str) -> dict:
    normalized = str(side or "").strip().lower()
    if normalized == "left":
        return context.left_payload
    if normalized == "right":
        return context.right_payload
    raise HarnessToolError("side 必须是 left 或 right。")


def _project_root_for_side(context, side: str) -> Path:
    payload = _side_payload(context, side)
    bundle = payload.get("bundle") if isinstance(payload.get("bundle"), dict) else {}
    raw = str(bundle.get("project_root") or "").strip()
    if not raw:
        raise HarnessToolError(f"{side} 项目缺少 project_root，无法读取项目文件。")
    root = Path(raw).expanduser().resolve()
    if not root.is_dir():
        raise HarnessToolError(f"{side} project_root 不存在或不是目录：{root}")
    return root


def _is_allowed_compare_project_file(rel: Path) -> bool:
    parts = rel.parts
    if not parts:
        return False
    name = rel.name
    suffix = rel.suffix.lower()
    if len(parts) == 1 and name in ALLOWED_COMPARE_ROOT_FILES:
        return True
    if name == "page.map" and (len(parts) == 1 or parts[0] == "sch_1"):
        return True
    if parts[0] in ALLOWED_COMPARE_PROJECT_DIRS and suffix in ALLOWED_COMPARE_TEXT_SUFFIXES:
        return True
    return False


def _resolve_compare_project_file(context, side: str, raw_path: str) -> Path:
    root = _project_root_for_side(context, side)
    if not str(raw_path or "").strip():
        raise HarnessToolError("缺少文件路径。")
    candidate = Path(str(raw_path).strip().strip('"')).expanduser()
    if not candidate.is_absolute():
        candidate = root / candidate
    resolved = candidate.resolve()
    try:
        rel = resolved.relative_to(root)
    except ValueError as exc:
        raise HarnessToolError("禁止读取项目根目录之外的文件。") from exc
    if not _is_allowed_compare_project_file(rel):
        raise HarnessToolError(f"文件不在 compare agent 允许读取范围内：{rel.as_posix()}")
    if not resolved.is_file():
        raise HarnessToolError(f"文件不存在：{rel.as_posix()}")
    return resolved


def _list_files_for_side(context, side: str, limit: int) -> List[dict]:
    root = _project_root_for_side(context, side)
    candidates = []
    for dirname in sorted(ALLOWED_COMPARE_PROJECT_DIRS):
        folder = root / dirname
        if folder.is_dir():
            candidates.extend(path for path in folder.rglob("*") if path.is_file())
    for name in sorted(ALLOWED_COMPARE_ROOT_FILES):
        path = root / name
        if path.is_file():
            candidates.append(path)
    files = []
    for path in sorted(set(candidates), key=lambda item: item.as_posix()):
        rel = path.resolve().relative_to(root)
        if not _is_allowed_compare_project_file(rel):
            continue
        try:
            size = path.stat().st_size
        except OSError:
            size = 0
        files.append({
            "side": side,
            "path": rel.as_posix(),
            "name": path.name,
            "size": size,
        })
        if len(files) >= limit:
            break
    return files


def list_compare_project_files_tool(context, args: dict) -> dict:
    side = str(args.get("side") or "both").strip().lower()
    limit = _as_int(args.get("limit", 120), 120)
    files = []
    sides = ["left", "right"] if side == "both" else [side]
    for item_side in sides:
        files.extend(_list_files_for_side(context, item_side, max(1, limit - len(files))))
        if len(files) >= limit:
            break
    return {
        "id": "list_compare_project_files",
        "title": "A/B 项目只读文件清单",
        "target": "compare_file",
        "summary": f"返回 {len(files)} 个允许读取的 A/B 项目文本文件。",
        "side": side,
        "files": files,
        "readonly": True,
    }


def read_compare_project_text_tool(context, args: dict) -> dict:
    side = str(args.get("side") or "").strip().lower()
    path = _resolve_compare_project_file(context, side, str(args.get("path") or ""))
    root = _project_root_for_side(context, side)
    max_chars = _as_int(args.get("max_chars", 12000), 12000)
    raw = path.read_bytes()
    encoding = "utf-8"
    for candidate in ["utf-8", "gb18030", "latin-1"]:
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    text = _redact_sensitive_text(text)
    rel = path.resolve().relative_to(root).as_posix()
    truncated = len(text) > max_chars
    return {
        "id": "read_compare_project_text",
        "title": f"{side}:{rel}",
        "target": "compare_file",
        "summary": f"读取 {side} 项目 {rel}，编码 {encoding}，返回 {min(len(text), max_chars)} 字符。",
        "side": side,
        "path": rel,
        "encoding": encoding,
        "chars": len(text),
        "truncated": truncated,
        "content": text[:max_chars],
        "readonly": True,
    }


def _parse_page_range_args(args: dict) -> tuple[int, int]:
    page_start = _as_int(args.get("page_start"), 0)
    page_end = _as_int(args.get("page_end"), 0)
    raw_range = str(args.get("page_range") or "").strip()
    if raw_range and (page_start <= 0 or page_end <= 0):
        nums = [int(value) for value in re.findall(r"(?i)(?:PAGE)?\s*(\d+)", raw_range)]
        if len(nums) >= 2:
            page_start, page_end = nums[0], nums[1]
        elif len(nums) == 1:
            page_start = page_end = nums[0]
    if page_start <= 0 or page_end <= 0:
        raise HarnessToolError("需要 page_start/page_end 或 page_range，例如 1-30。")
    if page_start > page_end:
        page_start, page_end = page_end, page_start
    page_count = page_end - page_start + 1
    if page_count > COMPARE_PAGE_RANGE_HARD_MAX:
        raise HarnessToolError(f"页范围过大：{page_count} 页，硬上限 {COMPARE_PAGE_RANGE_HARD_MAX} 页。")
    return page_start, page_end


def resolve_compare_page_range_tool(context, args: dict) -> dict:
    page_start, page_end = _parse_page_range_args(args)
    pages = list(range(page_start, page_end + 1))
    return {
        "id": "resolve_compare_page_range",
        "title": f"解析页码范围 PAGE{page_start}-PAGE{page_end}",
        "target": "cadence_page",
        "summary": (
            f"已将用户页码范围解析为 PAGE{page_start}-PAGE{page_end}，"
            f"对应 A/B 项目 sch_1/pageX.csv 与 pageX.csa。"
        ),
        "page_start": page_start,
        "page_end": page_end,
        "page_count": len(pages),
        "pages": pages,
        "page_semantics": "页码直接对应项目根目录 sch_1/pageX.csv|csa 文件名中的 X；内部主模块页另在映射表中复核。",
        "readonly": True,
    }


def _model_for_side(context,
                    side: str,
                    page: int,
                    *,
                    coordinate_tolerance: int = 0,
                    include_raw_unknown: bool = True):
    root = _project_root_for_side(context, side)
    return load_cadence_page_model(
        root,
        side,
        page,
        coordinate_tolerance=coordinate_tolerance,
        include_raw_unknown=include_raw_unknown,
    )


def compare_cadence_page_semantics_tool(context, args: dict) -> dict:
    page_start, page_end = _parse_page_range_args(args)
    include_raw_unknown = bool(args.get("include_raw_unknown", True))
    coordinate_tolerance = _as_int(args.get("coordinate_tolerance", 0), 0)
    max_diff_items = _as_int(args.get("max_diff_items", 40), 40)
    page_results = []
    changed_pages = []
    total_diff_count = 0
    for page in range(page_start, page_end + 1):
        left = _model_for_side(
            context,
            "left",
            page,
            coordinate_tolerance=coordinate_tolerance,
            include_raw_unknown=include_raw_unknown,
        )
        right = _model_for_side(
            context,
            "right",
            page,
            coordinate_tolerance=coordinate_tolerance,
            include_raw_unknown=include_raw_unknown,
        )
        diff = compare_page_models(left, right, max_diff_items=max_diff_items)
        total_diff_count += _as_int(diff.get("diff_count"), 0)
        if diff.get("status") != "same":
            changed_pages.append(page)
        page_results.append(diff)
    return {
        "id": "compare_cadence_page_semantics",
        "title": f"Cadence 页级语义比对 PAGE{page_start}-PAGE{page_end}",
        "target": "cadence_page",
        "summary": (
            f"完成 {len(page_results)} 个页码的 CSV/CSA 语义比对，"
            f"{len(changed_pages)} 页存在差异，总差异 {total_diff_count} 项。"
        ),
        "page_start": page_start,
        "page_end": page_end,
        "page_count": len(page_results),
        "changed_pages": changed_pages,
        "total_diff_count": total_diff_count,
        "coordinate_tolerance": coordinate_tolerance,
        "include_raw_unknown": include_raw_unknown,
        "page_results": page_results,
        "readonly": True,
    }


def get_cadence_page_object_tool(context, args: dict) -> dict:
    side = str(args.get("side") or "").strip().lower()
    page = _as_int(args.get("page"), 0)
    object_id = str(args.get("object_id") or "").strip()
    if page <= 0:
        raise HarnessToolError("page 必须是正整数。")
    if not object_id:
        raise HarnessToolError("缺少 object_id。")
    model = _model_for_side(context, side, page)
    obj = model.object_by_id(object_id)
    if obj:
        payload = obj.to_dict(include_raw=True)
        object_kind = "graphic_object"
    else:
        conn = model.connectivity_by_id(object_id)
        if not conn:
            raise HarnessToolError(f"{side} PAGE{page} 不存在对象或连接组件：{object_id}")
        payload = conn.to_dict()
        object_kind = "connectivity"
    return {
        "id": "get_cadence_page_object",
        "title": f"{side}:PAGE{page}:{object_id}",
        "target": "cadence_page",
        "summary": f"读取 {side} 项目页码 PAGE{page} 的 {object_kind} 详情。",
        "side": side,
        "page": page,
        "object_id": object_id,
        "object_kind": object_kind,
        "object": payload,
        "page_digest": model.digest(),
        "readonly": True,
    }


def batch_get_cadence_page_objects_tool(context, args: dict) -> dict:
    raw_items, input_truncated = _batch_input_items(args, "objects")
    items = []
    for index, raw_item in enumerate(raw_items, start=1):
        if not isinstance(raw_item, dict):
            items.append({
                "index": index,
                "status": "error",
                "summary": "批量 Cadence 对象请求必须是对象。",
                "missing_reason": "bad_request_item",
            })
            continue
        side = str(raw_item.get("side") or "").strip()
        page = _as_int(raw_item.get("page"), 0)
        object_id = str(raw_item.get("object_id") or "").strip()
        try:
            result = get_cadence_page_object_tool(context, {"side": side, "page": page, "object_id": object_id})
            items.append({
                "index": index,
                "side": side,
                "page": page,
                "object_id": object_id,
                "status": "found",
                "summary": result.get("summary", ""),
                "object_kind": result.get("object_kind", ""),
                "object": result.get("object") or {},
                "missing_reason": "",
            })
        except Exception as exc:
            items.append({
                "index": index,
                "side": side,
                "page": page,
                "object_id": object_id,
                "status": "error",
                "summary": str(exc),
                "object": {},
                "missing_reason": str(exc),
            })
    return {
        "id": "batch_get_cadence_page_objects",
        "title": "批量读取 Cadence 页对象详情",
        "target": "cadence_page",
        "summary": _batch_summary("批量读取 Cadence 页对象", items, truncated=input_truncated),
        "input_count": len(raw_items),
        "input_truncated": input_truncated,
        "items": items,
        "readonly": True,
    }


def get_cadence_page_raw_excerpt_tool(context, args: dict) -> dict:
    side = str(args.get("side") or "").strip().lower()
    page = _as_int(args.get("page"), 0)
    file_type = str(args.get("file_type") or "").strip().lower()
    offset = _as_int(args.get("offset", 0), 0)
    max_chars = _as_int(args.get("max_chars", 12000), 12000)
    if page <= 0:
        raise HarnessToolError("page 必须是正整数。")
    if file_type not in {"csv", "csa"}:
        raise HarnessToolError("file_type 必须是 csv 或 csa。")
    path = _resolve_compare_project_file(context, side, f"sch_1/page{page}.{file_type}")
    root = _project_root_for_side(context, side)
    raw = path.read_bytes()
    encoding = "utf-8"
    for candidate in ["utf-8-sig", "utf-16", "gb18030", "latin-1"]:
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    offset = max(0, offset)
    snippet = text[offset:offset + max_chars]
    rel = path.resolve().relative_to(root).as_posix()
    return {
        "id": "get_cadence_page_raw_excerpt",
        "title": f"{side}:{rel}@{offset}",
        "target": "cadence_page",
        "summary": f"读取 {side} 项目 {rel} 原始片段，编码 {encoding}，offset={offset}。",
        "side": side,
        "page": page,
        "file_type": file_type,
        "path": rel,
        "encoding": encoding,
        "offset": offset,
        "max_chars": max_chars,
        "chars": len(text),
        "truncated": offset + max_chars < len(text),
        "content": snippet,
        "readonly": True,
    }
