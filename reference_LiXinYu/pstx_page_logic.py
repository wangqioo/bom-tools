# -*- coding: utf-8 -*-
"""PSTX page parsing and resolution helpers."""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple


_PAGE_TOKEN_RE = re.compile(
    r"(?<![A-Z0-9])PAGE(?:[_\-/ ]*)(\d+)([A-Z]?)(?![A-Z0-9])",
    re.IGNORECASE,
)
_PATH_SEGMENT_RE = re.compile(
    r"^(?P<head>.+?)\((?P<view>[^)]+)\)\s*:\s*(?P<tail>.+)$",
    re.IGNORECASE,
)
_SECTION_PATH_RE = re.compile(
    r"(?ims)^\s*SECTION_NUMBER\s+(?P<num>\d+)\s*\n\s*'(?P<path>[^']+)'\s*:",
)
_MODULE_ORDER_LINE_RE = re.compile(
    r"^\s*(?P<path>@\S+)\s+(?P<unk1>\d+)\s+(?P<unk2>\d+)\s+(?P<start>\d+)\s+(?P<count>\d+)\s+(?P<flag>\d+)\s*$",
    re.IGNORECASE,
)
_PAGE_NUMBER_LINE_RE = re.compile(
    r"""^\s*["']?PAGE_NUMBER["']?\s*(?:=|:)\s*["']?(?P<value>[A-Z0-9_./ -]+?)["']?\s*[;,]?\s*$""",
    re.IGNORECASE,
)


def _natural_sort_key(value: str):
    parts = re.split(r"(\d+)", str(value or "").upper())
    return [int(part) if part.isdigit() else part for part in parts]


def _normalize_page_token(match: re.Match) -> str:
    num = str(int(match.group(1)))
    suffix = match.group(2).upper()
    return f"PAGE{num}{suffix}"


def normalize_page_label(page_label: str) -> str:
    value = str(page_label or "").strip().upper()
    if not value:
        return ""
    matches = list(_PAGE_TOKEN_RE.finditer(value))
    if not matches:
        return value
    normalized = [_normalize_page_token(match) for match in matches]
    if len(normalized) == 1:
        return normalized[0]
    return " / ".join(normalized)


def _coerce_page_number(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if not text.upper().startswith("PAGE"):
        text = f"PAGE{text}"
    return normalize_page_label(text)


def _clean_page_csv_value(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = text.rstrip(";,").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1].strip()
    return text


def _parse_page_map_line(raw_line: str) -> Optional[Dict[str, str]]:
    line = str(raw_line or "").strip()
    if not line:
        return None

    parts = re.split(r"\s+", line, maxsplit=2)
    if len(parts) < 3:
        return None

    logical_page = _coerce_page_number(parts[0])
    real_page = _coerce_page_number(parts[1])
    page_name = parts[2].strip()
    if not logical_page or not real_page or not page_name:
        return None

    return {
        "logical_page": logical_page,
        "real_page": real_page,
        "page_name": page_name,
    }


def _iter_text_with_fallback_encodings(file_path: Path) -> List[str]:
    try:
        raw_bytes = file_path.read_bytes()
    except OSError:
        return []

    texts: List[str] = []
    seen = set()
    for encoding in [
        "utf-8-sig",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
        "utf-8",
        "gb18030",
        "cp936",
    ]:
        try:
            text = raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
        if text not in seen:
            seen.add(text)
            texts.append(text)
    if not texts:
        texts.append(raw_bytes.decode("utf-8", errors="replace"))
    return texts


def extract_path_segments(path_text: str) -> List[Dict[str, str]]:
    raw = str(path_text or "").strip()
    if not raw:
        return []

    segments: List[Dict[str, str]] = []
    for chunk in [segment.strip() for segment in raw.split("@") if segment.strip()]:
        match = _PATH_SEGMENT_RE.match(chunk)
        if not match:
            continue
        head = match.group("head").strip()
        view = match.group("view").strip()
        tail = match.group("tail").strip()
        page_match = _PAGE_TOKEN_RE.search(tail)
        if not page_match:
            continue
        lib, _, cell = head.rpartition(".")
        segments.append(
            {
                "raw": chunk,
                "head": head,
                "lib": lib.strip(),
                "cell": (cell or head).strip(),
                "view": view,
                "tail": tail,
                "raw_page": _normalize_page_token(page_match),
            }
        )
    return segments


def extract_top_level_page(path_text: str) -> str:
    segments = extract_path_segments(path_text)
    for segment in segments:
        if segment.get("view", "").upper() == "SCH_1":
            return segment.get("raw_page", "")
    if segments:
        return segments[0].get("raw_page", "")
    normalized = normalize_page_label(path_text)
    if " / " in normalized:
        return normalized.split(" / ", 1)[0]
    return normalized


def extract_submodule_page(path_text: str) -> str:
    sch_segments = [
        segment
        for segment in extract_path_segments(path_text)
        if segment.get("view", "").upper() == "SCH_1"
    ]
    if len(sch_segments) >= 2:
        # Return the local physical page of the innermost schematic module.
        # For two-level reuse this is the second SCH_1 segment; for deeper
        # reuse it is the page used together with the deepest module_order key.
        return sch_segments[-1].get("raw_page", "")
    return ""


def extract_section_paths(block_text: str) -> List[Dict[str, str]]:
    entries: List[Dict[str, str]] = []
    for match in _SECTION_PATH_RE.finditer(str(block_text or "")):
        entries.append(
            {
                "section_number": match.group("num"),
                "path": match.group("path").strip(),
            }
        )
    return entries


def select_component_page_sources(block_text: str, attrs: Dict[str, str]) -> Dict[str, str]:
    section_paths = extract_section_paths(block_text)
    logical_path_raw = ""
    logical_path_source = "none"
    if section_paths:
        preferred = next(
            (entry for entry in section_paths if entry.get("section_number") == "1"),
            section_paths[0],
        )
        logical_path_raw = preferred.get("path", "").strip()
        logical_path_source = "section_path" if logical_path_raw else "none"
    if not logical_path_raw:
        logical_path_raw = str(attrs.get("C_PATH", "")).strip()
        logical_path_source = "c_path" if logical_path_raw else "none"
    if not logical_path_raw:
        logical_path_raw = str(attrs.get("DRAWING", "")).strip()
        logical_path_source = "drawing" if logical_path_raw else "none"

    real_path_raw = str(attrs.get("P_PATH", "")).strip()
    real_path_source = "p_path" if real_path_raw else "none"
    return {
        "logical_path_raw": logical_path_raw,
        "logical_path_source": logical_path_source,
        "real_path_raw": real_path_raw,
        "real_path_source": real_path_source,
    }


def _iter_page_csv_paths(project_root: Path) -> List[Path]:
    candidates: Dict[str, Path] = {}
    direct_sch = project_root / "sch_1"
    if direct_sch.is_dir():
        for csv_path in direct_sch.iterdir():
            if (
                csv_path.is_file()
                and csv_path.suffix.lower() == ".csv"
                and csv_path.stem.lower().startswith("page")
            ):
                candidates[str(csv_path.resolve())] = csv_path
    for csv_path in project_root.rglob("page*.csv"):
        if csv_path.is_file() and csv_path.parent.name.lower() == "sch_1":
            candidates[str(csv_path.resolve())] = csv_path
    return sorted(candidates.values(), key=lambda path: _natural_sort_key(str(path)))


def _extract_page_number_from_text(text: str) -> str:
    if not text:
        return ""

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = _PAGE_NUMBER_LINE_RE.match(line)
        if match:
            page_number = _coerce_page_number(_clean_page_csv_value(match.group("value")))
            if page_number:
                return page_number

    rows = []
    for raw_line in text.splitlines():
        parts = [_clean_page_csv_value(part) for part in raw_line.split(",")]
        rows.append(parts)
        for idx, part in enumerate(parts):
            if part.upper() != "PAGE_NUMBER":
                continue
            for follower in parts[idx + 1 :]:
                page_number = _coerce_page_number(_clean_page_csv_value(follower))
                if page_number:
                    return page_number

    for row_idx, parts in enumerate(rows):
        header_indexes = [idx for idx, part in enumerate(parts) if part.upper() == "PAGE_NUMBER"]
        if not header_indexes:
            continue
        for col_idx in header_indexes:
            for data_row in rows[row_idx + 1 :]:
                if col_idx >= len(data_row):
                    continue
                page_number = _coerce_page_number(_clean_page_csv_value(data_row[col_idx]))
                if page_number:
                    return page_number

    for regex in [
        re.compile(
            r'(?im)["\']?PAGE_NUMBER["\']?\s*[,=:\t;]\s*["\']?([A-Z0-9_./ -]+?)["\']?\s*[;,]?(?:$|\r|\n)'
        ),
        re.compile(
            r'(?im)^["\']?PAGE_NUMBER["\']?\s*[,;\t]\s*["\']?([A-Z0-9_./ -]+?)["\']?\s*[;,]?(?:$|\r|\n)'
        ),
    ]:
        match = regex.search(text)
        if not match:
            continue
        page_number = _coerce_page_number(_clean_page_csv_value(match.group(1)))
        if page_number:
            return page_number
    return ""


def _read_page_number_from_csv(csv_path: Path) -> str:
    for text in _iter_text_with_fallback_encodings(csv_path):
        page_number = _extract_page_number_from_text(text)
        if page_number:
            return page_number
    return ""


def build_page_csv_index(project_root: str) -> Dict[str, object]:
    root = Path(project_root).expanduser()
    index = {
        "root": str(root),
        "by_logical_page": defaultdict(list),
        "root_by_logical_page": defaultdict(list),
        "warnings": [],
        "count": 0,
        "scanned": 0,
        "matched_root_sch1": 0,
        "skipped_paths": [],
    }
    if not project_root:
        return index
    if not root.exists():
        index["warnings"].append(f"项目根路径不存在，无法建立 page*.csv 映射：{root}")
        return index

    csv_paths = _iter_page_csv_paths(root)
    index["scanned"] = len(csv_paths)
    index["matched_root_sch1"] = sum(1 for path in csv_paths if path.parent == (root / "sch_1"))

    for csv_path in csv_paths:
        real_page = _coerce_page_number(csv_path.stem)
        if not real_page:
            if len(index["skipped_paths"]) < 10:
                index["skipped_paths"].append(str(csv_path))
            continue
        logical_page = _read_page_number_from_csv(csv_path)
        if not logical_page:
            if len(index["skipped_paths"]) < 10:
                index["skipped_paths"].append(str(csv_path))
            continue
        entry = {
            "path": str(csv_path),
            "resolved_page": real_page,
            "page_name": csv_path.stem,
            "is_root_sch1": csv_path.parent == (root / "sch_1"),
        }
        index["by_logical_page"][logical_page].append(entry)
        if entry["is_root_sch1"]:
            index["root_by_logical_page"][logical_page].append(entry)
        index["count"] += 1

    if index["scanned"] == 0:
        index["warnings"].append(f"未在项目根路径下找到任何 sch_1/page*.csv：{root}")
    elif index["count"] == 0:
        samples = "；".join(index["skipped_paths"][:3])
        suffix = f"；例如：{samples}" if samples else ""
        index["warnings"].append(
            f"已扫描 {index['scanned']} 个 sch_1/page*.csv，但没有读出任何 PAGE_NUMBER{suffix}"
        )
    return index


def build_page_map_index(project_root: str) -> Dict[str, object]:
    root = Path(project_root).expanduser()
    index = {
        "root": str(root),
        "by_logical_page": defaultdict(list),
        "root_by_logical_page": defaultdict(list),
        "warnings": [],
        "count": 0,
        "files": [],
    }
    if not project_root:
        return index
    if not root.exists():
        index["warnings"].append(f"项目根路径不存在，无法读取 page.map：{root}")
        return index

    file_paths: List[Path] = []
    direct = root / "sch_1" / "page.map"
    if direct.is_file():
        file_paths.append(direct)
    for path in root.rglob("page.map"):
        if path.is_file() and path not in file_paths:
            file_paths.append(path)
    file_paths = sorted(file_paths, key=lambda item: (0 if item == direct else 1, _natural_sort_key(str(item))))
    index["files"] = [str(path) for path in file_paths]

    for path in file_paths:
        matched_in_file = False
        for text in _iter_text_with_fallback_encodings(path):
            for raw_line in text.splitlines():
                parsed = _parse_page_map_line(raw_line)
                if not parsed:
                    continue
                logical_page = parsed.get("logical_page", "")
                real_page = parsed.get("real_page", "")
                if not logical_page or not real_page:
                    continue
                entry = {
                    "path": str(path),
                    "logical_page": logical_page,
                    "resolved_page": real_page,
                    "page_name": parsed.get("page_name", ""),
                    "is_root_sch1": path.parent == (root / "sch_1"),
                }
                index["by_logical_page"][logical_page].append(entry)
                if entry["is_root_sch1"]:
                    index["root_by_logical_page"][logical_page].append(entry)
                index["count"] += 1
                matched_in_file = True
            if matched_in_file:
                break

    if file_paths and index["count"] == 0:
        index["warnings"].append(f"已扫描 {len(file_paths)} 个 page.map，但没有读出有效映射：{root}")
    return index


def _iter_named_files(project_root: Path, filename: str) -> List[Path]:
    candidates: Dict[str, Path] = {}
    direct = project_root / filename
    if direct.is_file():
        candidates[str(direct.resolve())] = direct
    for path in project_root.rglob(filename):
        if path.is_file():
            candidates[str(path.resolve())] = path
    return sorted(candidates.values(), key=lambda item: _natural_sort_key(str(item)))


def _iter_module_order_files(project_root: Path) -> List[Path]:
    candidates: Dict[str, Path] = {}
    for filename in ("module_order.dat", "module_order"):
        for path in _iter_named_files(project_root, filename):
            candidates[str(path.resolve())] = path
    return sorted(candidates.values(), key=lambda item: _natural_sort_key(str(item)))


def _normalize_module_order_key(path_text: str) -> str:
    return str(path_text or "").strip().upper()


def build_module_order_index(project_root: str) -> Dict[str, object]:
    root = Path(project_root).expanduser()
    index = {
        "root": str(root),
        "by_key": defaultdict(list),
        "warnings": [],
        "count": 0,
        "duplicate_count": 0,
        "files": [],
    }
    if not project_root:
        return index
    if not root.exists():
        index["warnings"].append(f"项目根路径不存在，无法读取 module_order：{root}")
        return index

    file_paths = _iter_module_order_files(root)
    index["files"] = [str(path) for path in file_paths]
    seen_entries = set()
    for path in file_paths:
        matched_in_file = False
        for text in _iter_text_with_fallback_encodings(path):
            in_section = False
            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    continue
                upper_line = line.upper()
                if upper_line == "START_MODULEORDER":
                    in_section = True
                    continue
                if upper_line == "END_MODULEORDER":
                    in_section = False
                    continue
                if not in_section or not line.startswith("@"):
                    continue
                match = _MODULE_ORDER_LINE_RE.match(line)
                if not match:
                    continue
                key = _normalize_module_order_key(match.group("path"))
                entry = {
                    "path": match.group("path"),
                    "path_key": key,
                    "start_real_page": _coerce_page_number(match.group("start")),
                    "page_count": int(match.group("count")),
                    "flag": int(match.group("flag")),
                    "source_file": str(path),
                }
                signature = (
                    entry["path_key"],
                    entry["start_real_page"],
                    entry["page_count"],
                    entry["flag"],
                )
                if signature in seen_entries:
                    index["duplicate_count"] = int(index.get("duplicate_count", 0) or 0) + 1
                    matched_in_file = True
                    continue
                seen_entries.add(signature)
                index["by_key"][key].append(entry)
                index["count"] += 1
                matched_in_file = True
            if matched_in_file:
                break

    if file_paths and index["count"] == 0:
        index["warnings"].append(f"已扫描 {len(file_paths)} 个 module_order，但没有读出有效映射：{root}")
    return index


def _page_index_root_name(*indexes: Optional[Dict[str, object]]) -> str:
    for index in indexes:
        root = str((index or {}).get("root", "")).strip()
        if root:
            try:
                return Path(root).name.upper()
            except Exception:
                continue
    return ""


def _root_logical_pages(index: Optional[Dict[str, object]]) -> set:
    if not index:
        return set()
    pages = set()
    for logical_page, entries in index.get("by_logical_page", {}).items():
        if any(entry.get("is_root_sch1") for entry in entries):
            pages.add(logical_page)
    return pages


def pick_top_schematic_segment(
    path_text: str,
    *,
    page_map_index: Optional[Dict[str, object]] = None,
    page_csv_index: Optional[Dict[str, object]] = None,
) -> Dict[str, str]:
    sch_segments = [
        segment
        for segment in extract_path_segments(path_text)
        if segment.get("view", "").upper() == "SCH_1"
    ]
    if not sch_segments:
        return {}

    root_name = _page_index_root_name(page_map_index, page_csv_index)
    if root_name:
        exact_root_matches = [
            segment for segment in sch_segments if segment.get("cell", "").upper() == root_name
        ]
        if exact_root_matches:
            return exact_root_matches[0]

    root_pages = _root_logical_pages(page_map_index) | _root_logical_pages(page_csv_index)
    if root_pages:
        root_matches = [segment for segment in sch_segments if segment.get("raw_page", "") in root_pages]
        if len(root_matches) == 1:
            return root_matches[0]
        if root_matches:
            return root_matches[0]

    return sch_segments[0]


def _build_module_order_lookup_candidates_from_path(path_text: str, source: str) -> List[Dict[str, str]]:
    sch_segments = [
        segment
        for segment in extract_path_segments(path_text)
        if segment.get("view", "").upper() == "SCH_1"
    ]
    if len(sch_segments) < 2:
        return []

    candidates: List[Dict[str, str]] = []
    # Prefer the deepest exact module_order key. A nested component can be
    # inside TOP -> REUSE_A -> REUSE_B; in that case the REUSE_B entry maps the
    # component's local page to the final top-level physical page. If that key
    # is unavailable, the resolver may still fall back to the outer reuse key.
    for child_index in range(len(sch_segments) - 1, 0, -1):
        parent_chain = "@".join(segment["raw"] for segment in sch_segments[:child_index])
        child = sch_segments[child_index]
        raw_key = f"@{parent_chain}@{child['head']}({child['view']})"
        candidates.append(
            {
                "key": _normalize_module_order_key(raw_key),
                "raw_key": raw_key,
                "local_page": child.get("raw_page", ""),
                "segment_depth": str(child_index + 1),
                "source": source,
            }
        )
    return candidates


def build_module_order_lookup_candidates(logical_path: str, real_path: str) -> List[Dict[str, str]]:
    candidates: List[Dict[str, str]] = []
    seen = set()
    # `module_order` keys are primarily instance-path identifiers. In current
    # DE-HDL exports they align more reliably with the logical hierarchy path
    # (`SECTION_NUMBER` / `C_PATH`) than with `P_PATH`. Keep `P_PATH` only as a
    # conservative fallback for projects that already emitted physical-page
    # keys.
    source_paths: List[Tuple[str, str]] = []
    if str(logical_path or "").strip():
        source_paths.append(("logical_path", logical_path))
    if str(real_path or "").strip():
        source_paths.append(("real_path", real_path))
    for source, path_text in source_paths:
        for candidate in _build_module_order_lookup_candidates_from_path(path_text, source):
            key = candidate.get("key", "")
            if not key or key in seen:
                continue
            seen.add(key)
            candidates.append(candidate)
    return candidates


def _entries_for_logical_page(
    index: Optional[Dict[str, object]],
    logical_page: str,
    *,
    prefer_root: bool = True,
) -> List[Dict[str, str]]:
    if not index or not logical_page:
        return []
    entries = list(index.get("by_logical_page", {}).get(logical_page, []))
    if not entries:
        return []
    if prefer_root:
        root_entries = list(index.get("root_by_logical_page", {}).get(logical_page, []))
        if not root_entries:
            root_entries = [entry for entry in entries if entry.get("is_root_sch1")]
        if root_entries:
            return root_entries
    return entries


def _resolve_unique_real_page(
    index: Optional[Dict[str, object]],
    logical_page: str,
    *,
    prefer_root: bool = True,
) -> Tuple[str, str]:
    if not index or not logical_page:
        return "", "none"
    entries = _entries_for_logical_page(index, logical_page, prefer_root=prefer_root)
    if not entries:
        return "", "none"
    real_pages = sorted(
        {entry.get("resolved_page", "") for entry in entries if entry.get("resolved_page", "")},
        key=_natural_sort_key,
    )
    if len(real_pages) != 1:
        return "", "ambiguous"
    return real_pages[0], "unique"


def _resolve_module_order_entry(
    module_order_index: Optional[Dict[str, object]],
    lookup_candidates: List[Dict[str, str]],
):
    if not module_order_index:
        return {}, "none", {}
    for candidate in lookup_candidates:
        key = candidate.get("key", "")
        entries = module_order_index.get("by_key", {}).get(key, [])
        if len(entries) == 1:
            return entries[0], "unique", candidate
        if len(entries) > 1:
            return {}, "ambiguous", candidate
    return {}, "none", {}


def resolve_component_page_info(
    comp: Dict[str, object],
    *,
    page_map_index: Optional[Dict[str, object]] = None,
    page_csv_index: Optional[Dict[str, object]] = None,
    module_order_index: Optional[Dict[str, object]] = None,
) -> Dict[str, str]:
    logical_path = str(comp.get("page_path_logical_raw", "") or comp.get("page_path_raw", "") or comp.get("drawing", ""))
    real_path = str(comp.get("page_path_real_raw", ""))

    top_segment = pick_top_schematic_segment(
        logical_path,
        page_map_index=page_map_index,
        page_csv_index=page_csv_index,
    )
    top_logical_page = top_segment.get("raw_page", "") or extract_top_level_page(logical_path)
    top_real_page_from_path = extract_top_level_page(real_path)
    submodule_real_page = extract_submodule_page(real_path)

    page_map_real, page_map_state = _resolve_unique_real_page(page_map_index, top_logical_page, prefer_root=True)
    page_csv_real, page_csv_state = _resolve_unique_real_page(page_csv_index, top_logical_page, prefer_root=True)

    final_real_page = top_real_page_from_path or page_map_real or page_csv_real
    if top_real_page_from_path:
        real_source = "p_path"
    elif page_map_real:
        real_source = "page_map"
    elif page_csv_real:
        real_source = "page_csv"
    elif page_map_state == "ambiguous":
        real_source = "page_map_ambiguous"
    elif page_csv_state == "ambiguous":
        real_source = "page_csv_ambiguous"
    else:
        real_source = "none"

    seen_sources = []
    if top_real_page_from_path:
        seen_sources.append(("p_path", top_real_page_from_path))
    if page_map_real:
        seen_sources.append(("page.map", page_map_real))
    if page_csv_real:
        seen_sources.append(("page*.csv", page_csv_real))

    unique_values = {value for _, value in seen_sources if value}
    if not final_real_page and not seen_sources and not (page_map_index or page_csv_index or real_path):
        validation_status = ""
        mapping_ok = ""
        validation_note = ""
    elif not final_real_page:
        validation_status = "未命中真实页"
        mapping_ok = "否"
        validation_note = f"{top_logical_page or 'UNKNOWN'} 未命中 P_PATH/page.map/page*.csv"
    elif len(unique_values) <= 1:
        validation_status = "一致"
        mapping_ok = "是"
        validation_note = (
            "P_PATH/page.map/page*.csv 一致"
            if len(seen_sources) > 1
            else f"仅命中 {real_source}"
        )
    else:
        validation_status = "冲突"
        mapping_ok = "否"
        validation_note = "；".join(f"{source}={value}" for source, value in seen_sources)

    lookup_candidates = build_module_order_lookup_candidates(logical_path, real_path)
    module_entry, module_state, matched_candidate = _resolve_module_order_entry(
        module_order_index,
        lookup_candidates,
    )
    submodule_mapped_page = ""
    submodule_mapping_note = ""
    module_order_local_page = submodule_real_page or matched_candidate.get("local_page", "")
    if module_entry and module_order_local_page:
        try:
            start_page_match = re.search(r"(\d+)", str(module_entry.get("start_real_page", "")))
            local_page_match = re.search(r"(\d+)", module_order_local_page)
            if start_page_match and local_page_match:
                start_page = int(start_page_match.group(1))
                local_page = int(local_page_match.group(1))
                page_count = int(module_entry.get("page_count", 0) or 0)
                if local_page < 1:
                    module_state = "local_page_invalid"
                    submodule_mapping_note = f"子模块本地页无效：{module_order_local_page}"
                elif page_count and local_page > page_count:
                    module_state = "local_page_out_of_range"
                    submodule_mapping_note = (
                        f"子模块本地页 {module_order_local_page} 超出 module_order 页数 {page_count}"
                    )
                else:
                    submodule_mapped_page = _coerce_page_number(str(start_page + local_page - 1))
                    submodule_mapping_note = (
                        f"{module_entry.get('start_real_page', '')} + ({module_order_local_page} - 1)"
                    )
        except Exception:
            submodule_mapped_page = ""
            module_state = "calculation_error"
            submodule_mapping_note = "module_order 页码计算失败"

    top_segments = [
        segment
        for segment in extract_path_segments(logical_path)
        if segment.get("view", "").upper() == "SCH_1"
    ]
    top_cell = top_segments[0].get("cell", "") if top_segments else ""

    return {
        "page_logical": top_logical_page,
        "page_real": final_real_page,
        "page_submodule_real": submodule_real_page,
        "page_submodule_mapped": submodule_mapped_page,
        "page_logical_source": str(
            comp.get("page_path_logical_source", "")
            or comp.get("page_path_source", "")
            or ("drawing" if comp.get("drawing") else "none")
        ),
        "page_real_source": real_source,
        "page_validation_status": validation_status,
        "page_validation_note": validation_note,
        "page_mapping_ok": mapping_ok,
        "page_context": f"{top_cell}:{top_logical_page}" if top_cell and top_logical_page else top_logical_page,
        "page_context_real": f"{top_cell}:{final_real_page}" if top_cell and final_real_page else final_real_page,
        "page_map_real": page_map_real,
        "page_map_state": page_map_state,
        "page_csv_real": page_csv_real,
        "page_csv_state": page_csv_state,
        "module_order_key": matched_candidate.get("raw_key", ""),
        "module_order_state": module_state,
        "module_order_local_page": module_order_local_page if module_entry else "",
        "module_order_start_page": str(module_entry.get("start_real_page", "")) if module_entry else "",
        "module_order_page_count": str(module_entry.get("page_count", "")) if module_entry else "",
        "page_submodule_mapping_note": submodule_mapping_note,
    }


def build_page_mapping_rows(
    page_map_index: Optional[Dict[str, object]],
    page_csv_index: Optional[Dict[str, object]],
) -> Dict[str, object]:
    rows = []
    warnings: List[str] = []
    logical_meta: Dict[str, Dict[str, str]] = {}
    logical_pages = set()
    if page_map_index:
        logical_pages.update(page_map_index.get("by_logical_page", {}).keys())
    if page_csv_index:
        logical_pages.update(page_csv_index.get("by_logical_page", {}).keys())

    resolved_by_logical: Dict[str, Dict[str, str]] = {}
    reverse_by_real: Dict[str, set] = defaultdict(set)
    for logical_page in sorted(logical_pages, key=_natural_sort_key):
        page_map_real, page_map_state = _resolve_unique_real_page(page_map_index, logical_page, prefer_root=True)
        page_csv_real, page_csv_state = _resolve_unique_real_page(page_csv_index, logical_page, prefer_root=True)
        final_real = page_map_real or page_csv_real
        state = {
            "page_map_real": page_map_real,
            "page_map_state": page_map_state,
            "page_csv_real": page_csv_real,
            "page_csv_state": page_csv_state,
            "final_real": final_real,
        }
        resolved_by_logical[logical_page] = state
        if final_real:
            reverse_by_real[final_real].add(logical_page)

    warned_reverse: set = set()
    for logical_page in sorted(logical_pages, key=_natural_sort_key):
        state = resolved_by_logical[logical_page]
        page_map_real = state["page_map_real"]
        page_map_state = state["page_map_state"]
        page_csv_real = state["page_csv_real"]
        page_csv_state = state["page_csv_state"]
        final_real = state["final_real"]
        reverse_logicals = sorted(reverse_by_real.get(final_real, set()), key=_natural_sort_key) if final_real else []

        if page_map_state == "ambiguous" or page_csv_state == "ambiguous":
            ok = "否"
            status = "冲突"
            note = f"{logical_page} 同时命中多个真实页"
            warnings.append(note)
        elif page_map_real and page_csv_real and page_map_real != page_csv_real:
            ok = "否"
            status = "冲突"
            note = f"{logical_page} 在 page.map={page_map_real} 与 page*.csv={page_csv_real} 之间冲突"
            warnings.append(note)
        elif final_real and len(reverse_logicals) > 1:
            ok = "否"
            status = "真实页对应多个逻辑页"
            note = f"{final_real} 同时被多个逻辑页复用：{', '.join(reverse_logicals)}"
            if final_real not in warned_reverse:
                warnings.append(note)
                warned_reverse.add(final_real)
        elif page_map_real and page_csv_real and page_map_real == page_csv_real:
            ok = "是"
            status = "一致"
            note = f"{logical_page} 在 page.map 与 page*.csv 中一致映射到 {final_real}"
        elif final_real:
            ok = "是"
            status = "单源"
            note = f"{logical_page} 仅在 {'page.map' if page_map_real else 'page*.csv'} 中命中 {final_real}"
        else:
            ok = "否"
            status = "未命中"
            note = f"{logical_page} 未命中 page.map/page*.csv"

        rows.append(
            {
                "逻辑页": logical_page,
                "真实页": final_real,
                "page.map真实页": page_map_real,
                "page*.csv真实页": page_csv_real,
                "是否一一对应": ok,
                "状态": status,
                "说明": note,
            }
        )
        logical_meta[logical_page] = {
            "real_page": final_real if ok == "是" else "" if status == "冲突" else final_real,
            "mapping_ok": ok,
            "status": status,
            "note": note,
            "page_map_real": page_map_real,
            "page_csv_real": page_csv_real,
            "page_map_state": page_map_state,
            "page_csv_state": page_csv_state,
        }
    return {
        "rows": rows,
        "warnings": warnings,
        "logical_meta": logical_meta,
    }
