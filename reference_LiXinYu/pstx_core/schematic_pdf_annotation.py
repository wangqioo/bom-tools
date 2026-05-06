# -*- coding: utf-8 -*-
"""Schematic PDF page/coordinate annotation helpers.

This module builds a conservative overlay model for schematic PDFs. It does
not mutate PDFs and it does not treat raw schematic XY as PDF coordinates
unless a page calibration is supplied.
"""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from pstx_core.page_resolution import component_user_visible_page
from pstx_core import pages as page_logic


SCHEMA_VERSION = "pstx-schematic-pdf-annotation.v1"
DEFAULT_MARKER_SIZE = 18.0

_PAGE_TYPE_RE = re.compile(rb"/Type\s*/Page(?!s)\b")
_MEDIABOX_RE = re.compile(
    rb"/MediaBox\s*\[\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*\]",
    re.IGNORECASE,
)
_PAGE_NUMBER_RE = re.compile(r"(\d+)")


def _json_default(value: Any) -> str:
    if isinstance(value, Path):
        return str(value)
    return str(value)


def _coerce_float(value: Any) -> Optional[float]:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _coerce_int(value: Any) -> int:
    try:
        if value is None or value == "":
            return 0
        return int(value)
    except (TypeError, ValueError):
        return 0


def _safe_text(value: Any) -> str:
    return str(value or "").strip()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalize_page_label(value: Any) -> str:
    text = _safe_text(value)
    if not text:
        return ""
    if text.isdigit():
        text = f"PAGE{text}"
    return page_logic.normalize_page_label(text)


def _page_number_from_label(value: Any) -> int:
    text = _normalize_page_label(value)
    if not text:
        text = _safe_text(value)
    match = _PAGE_NUMBER_RE.search(text)
    return int(match.group(1)) if match else 0


def _normalize_token(value: Any) -> str:
    text = _safe_text(value).upper()
    text = text.strip(" \t\r\n.,;:()[]{}<>\"'")
    return re.sub(r"\s+", "", text)


def _bbox_from_value(value: Any) -> List[float]:
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            parts = [part.strip() for part in value.replace("，", ",").split(",")]
            value = parts
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return []
    numbers = [_coerce_float(item) for item in value]
    if any(item is None for item in numbers):
        return []
    x0, y0, x1, y1 = [float(item) for item in numbers if item is not None]
    left, right = sorted([x0, x1])
    top, bottom = sorted([y0, y1])
    return [left, top, right, bottom]


def _point_from_values(x_value: Any, y_value: Any) -> Optional[Tuple[float, float]]:
    x = _coerce_float(x_value)
    y = _coerce_float(y_value)
    if x is None or y is None:
        return None
    return x, y


def _normalized_bbox(bbox: Sequence[float], page: Mapping[str, Any]) -> List[float]:
    width = _coerce_float(page.get("width")) or 0.0
    height = _coerce_float(page.get("height")) or 0.0
    if not bbox or width <= 0 or height <= 0:
        return []
    x0, top, x1, bottom = [float(item) for item in bbox]
    return [
        max(0.0, min(1.0, x0 / width)),
        max(0.0, min(1.0, top / height)),
        max(0.0, min(1.0, x1 / width)),
        max(0.0, min(1.0, bottom / height)),
    ]


def _marker_bbox(x: float, y: float, *, size: float = DEFAULT_MARKER_SIZE) -> List[float]:
    half = max(1.0, float(size) / 2.0)
    return [x - half, y - half, x + half, y + half]


def _severity_color(severity: str) -> str:
    key = _safe_text(severity).lower()
    if key in {"error", "critical", "danger", "removed", "删除", "严重"}:
        return "#dc2626"
    if key in {"warning", "warn", "changed", "变更", "风险"}:
        return "#d97706"
    if key in {"added", "新增", "ok", "success"}:
        return "#16a34a"
    return "#2563eb"


def _read_pdf_metadata_with_pdfplumber(path: Path) -> Tuple[List[Dict[str, Any]], List[str]]:
    try:
        import pdfplumber  # type: ignore
    except Exception as exc:
        return [], [f"pdfplumber unavailable: {exc}"]
    try:
        with pdfplumber.open(str(path)) as pdf:
            pages = [
                {
                    "pdf_page_number": index,
                    "width": float(page.width),
                    "height": float(page.height),
                    "rotation": int(getattr(page, "rotation", 0) or 0),
                    "source": "pdfplumber",
                }
                for index, page in enumerate(pdf.pages, start=1)
            ]
        return pages, []
    except Exception as exc:
        return [], [f"pdfplumber metadata failed: {exc}"]


def _read_pdf_metadata_with_pypdf(path: Path) -> Tuple[List[Dict[str, Any]], List[str]]:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception as exc:
        return [], [f"pypdf unavailable: {exc}"]
    try:
        reader = PdfReader(str(path))
        pages: List[Dict[str, Any]] = []
        for index, page in enumerate(reader.pages, start=1):
            box = page.mediabox
            width = float(box.right) - float(box.left)
            height = float(box.top) - float(box.bottom)
            pages.append({
                "pdf_page_number": index,
                "width": width,
                "height": height,
                "rotation": int(getattr(page, "rotation", 0) or 0),
                "source": "pypdf",
            })
        return pages, []
    except Exception as exc:
        return [], [f"pypdf metadata failed: {exc}"]


def _read_pdf_metadata_fallback(path: Path) -> Tuple[List[Dict[str, Any]], List[str]]:
    warnings: List[str] = []
    try:
        data = path.read_bytes()
    except OSError as exc:
        return [], [f"PDF metadata fallback failed: {exc}"]
    page_count = len(_PAGE_TYPE_RE.findall(data))
    media_boxes = [
        [float(match.group(idx)) for idx in range(1, 5)]
        for match in _MEDIABOX_RE.finditer(data)
    ]
    if page_count <= 0 and media_boxes:
        page_count = len(media_boxes)
    if page_count <= 0:
        warnings.append("无法读取 PDF 页数；请安装 pdfplumber 或 pypdf 以获得完整 metadata。")
        return [], warnings
    default_box = media_boxes[0] if media_boxes else [0.0, 0.0, 595.0, 842.0]
    pages: List[Dict[str, Any]] = []
    for index in range(1, page_count + 1):
        box = media_boxes[index - 1] if index - 1 < len(media_boxes) else default_box
        width = abs(box[2] - box[0]) or 595.0
        height = abs(box[3] - box[1]) or 842.0
        pages.append({
            "pdf_page_number": index,
            "width": width,
            "height": height,
            "rotation": 0,
            "source": "fallback-regex",
        })
    warnings.append("PDF metadata 使用轻量 fallback 读取；安装 pdfplumber/pypdf 后可获得更可靠的页尺寸。")
    return pages, warnings


def read_pdf_metadata(pdf_path: str | Path) -> Dict[str, Any]:
    """Read PDF digest and page dimensions with graceful dependency fallback."""
    path = Path(pdf_path).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f"PDF 文件不存在：{path}")
    warnings: List[str] = []
    pages, page_warnings = _read_pdf_metadata_with_pdfplumber(path)
    warnings.extend(page_warnings)
    if not pages:
        pages, page_warnings = _read_pdf_metadata_with_pypdf(path)
        warnings.extend(page_warnings)
    if not pages:
        pages, page_warnings = _read_pdf_metadata_fallback(path)
        warnings.extend(page_warnings)
    stat = path.stat()
    return {
        "path": str(path),
        "filename": path.name,
        "size": stat.st_size,
        "sha256": _sha256_file(path),
        "page_count": len(pages),
        "pages": pages,
        "warnings": warnings,
    }


def extract_pdf_text_boxes(pdf_path: str | Path,
                           *,
                           terms: Optional[Iterable[str]] = None,
                           pages: Optional[Iterable[int]] = None) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Extract word-level text boxes from a PDF using optional pdfplumber."""
    terms_set = {_normalize_token(term) for term in (terms or []) if _normalize_token(term)}
    page_filter = {int(page) for page in (pages or []) if _coerce_int(page) > 0}
    try:
        import pdfplumber  # type: ignore
    except Exception as exc:
        return [], {
            "backend": "pdfplumber",
            "available": False,
            "ok": False,
            "error": str(exc),
        }
    boxes: List[Dict[str, Any]] = []
    try:
        with pdfplumber.open(str(Path(pdf_path).expanduser())) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                if page_filter and page_number not in page_filter:
                    continue
                for word in page.extract_words() or []:
                    token = _normalize_token(word.get("text", ""))
                    if terms_set and token not in terms_set:
                        continue
                    boxes.append({
                        "text": str(word.get("text", "") or ""),
                        "normalized_text": token,
                        "pdf_page_number": page_number,
                        "bbox": [
                            float(word.get("x0", 0.0) or 0.0),
                            float(word.get("top", word.get("y0", 0.0)) or 0.0),
                            float(word.get("x1", 0.0) or 0.0),
                            float(word.get("bottom", word.get("y1", 0.0)) or 0.0),
                        ],
                        "source": "pdfplumber.extract_words",
                    })
        return boxes, {
            "backend": "pdfplumber",
            "available": True,
            "ok": True,
            "box_count": len(boxes),
        }
    except Exception as exc:
        return [], {
            "backend": "pdfplumber",
            "available": True,
            "ok": False,
            "error": str(exc),
        }


def _page_map_lookup(pdf_page_map: Mapping[str, Any]) -> Dict[str, int]:
    lookup: Dict[str, int] = {}
    for raw_key, raw_value in dict(pdf_page_map or {}).items():
        page_number = _coerce_int(raw_value)
        if page_number <= 0:
            continue
        for key in {
            _safe_text(raw_key).upper(),
            _normalize_page_label(raw_key).upper(),
            str(_page_number_from_label(raw_key)),
        }:
            if key and key != "0":
                lookup[key] = page_number
    return lookup


def _normalize_pdf_page_map_payload(
    pdf_page_map: Mapping[str, Any],
    *,
    actual_pdf_sha256: str,
) -> Tuple[Dict[str, Any], Dict[str, Any], List[str]]:
    raw = dict(pdf_page_map or {})
    expected_sha = _safe_text(
        raw.get("_pdf_sha256")
        or raw.get("pdf_sha256")
        or raw.get("pdf_digest")
        or raw.get("pdf_sha")
    )
    mapping_value = raw.get("pages")
    if not isinstance(mapping_value, Mapping):
        mapping_value = raw.get("pdf_page_map")
    if isinstance(mapping_value, Mapping):
        mapping = dict(mapping_value)
    else:
        reserved = {"_pdf_sha256", "pdf_sha256", "pdf_digest", "pdf_sha", "pages", "pdf_page_map"}
        mapping = {key: value for key, value in raw.items() if str(key) not in reserved}

    meta = {
        "expected_pdf_sha256": expected_sha,
        "actual_pdf_sha256": actual_pdf_sha256,
        "sha256_checked": bool(expected_sha),
        "sha256_match": True,
    }
    warnings: List[str] = []
    if expected_sha and actual_pdf_sha256 and expected_sha.lower() != actual_pdf_sha256.lower():
        meta["sha256_match"] = False
        warnings.append(
            "pdf_page_map 携带的 PDF sha256 与当前 PDF 不一致；已拒绝使用该页码映射，请重新生成。"
        )
        return {}, meta, warnings
    return mapping, meta, warnings


def _pdf_text_token_page_label(value: Any) -> str:
    token = _normalize_token(value)
    match = re.fullmatch(r"PAGE(\d+)", token)
    if not match:
        return ""
    return f"PAGE{match.group(1)}"


def _same_pdf_text_line(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    left_bbox = _bbox_from_value(left.get("bbox"))
    right_bbox = _bbox_from_value(right.get("bbox"))
    if not left_bbox or not right_bbox:
        return True
    left_top = float(left_bbox[1])
    right_top = float(right_bbox[1])
    left_height = max(1.0, float(left_bbox[3]) - float(left_bbox[1]))
    right_height = max(1.0, float(right_bbox[3]) - float(right_bbox[1]))
    tolerance = max(3.0, min(left_height, right_height) * 0.6)
    return abs(left_top - right_top) <= tolerance


def _build_pdf_text_page_label_map(
    text_boxes: Sequence[Mapping[str, Any]],
    project_pages: Iterable[Any],
) -> Tuple[Dict[str, int], List[str]]:
    """Build a conservative PAGE<N> -> PDF page map from PDF text evidence.

    Bare numbers are intentionally ignored. Only explicit PAGE<N> tokens, or
    adjacent PAGE + N tokens on the same line, are accepted.
    """

    desired_labels = {
        _normalize_page_label(page).upper()
        for page in project_pages or []
        if _normalize_page_label(page)
    }
    if not desired_labels:
        return {}, []

    candidates: Dict[str, set[int]] = {label: set() for label in desired_labels}
    by_pdf_page: Dict[int, List[Dict[str, Any]]] = {}
    for raw_box in text_boxes or []:
        if not isinstance(raw_box, Mapping):
            continue
        page_number = _coerce_int(raw_box.get("pdf_page_number") or raw_box.get("page"))
        if page_number <= 0:
            continue
        item = dict(raw_box)
        item["_token"] = _normalize_token(raw_box.get("normalized_text") or raw_box.get("text"))
        item["_bbox"] = _bbox_from_value(raw_box.get("bbox"))
        by_pdf_page.setdefault(page_number, []).append(item)

    for page_number, page_boxes in by_pdf_page.items():
        page_boxes.sort(key=lambda item: (
            float((item.get("_bbox") or [0.0, 0.0, 0.0, 0.0])[1]),
            float((item.get("_bbox") or [0.0, 0.0, 0.0, 0.0])[0]),
        ))
        for index, box in enumerate(page_boxes):
            token = str(box.get("_token") or "")
            label = _pdf_text_token_page_label(token)
            if label in candidates:
                candidates[label].add(page_number)
            if token == "PAGE" and index + 1 < len(page_boxes):
                next_box = page_boxes[index + 1]
                next_token = str(next_box.get("_token") or "")
                if next_token.isdigit() and _same_pdf_text_line(box, next_box):
                    combined_label = f"PAGE{next_token}"
                    if combined_label in candidates:
                        candidates[combined_label].add(page_number)

    lookup: Dict[str, int] = {}
    warnings: List[str] = []
    for label in sorted(desired_labels):
        pages = sorted(candidates.get(label, set()))
        if len(pages) == 1:
            lookup[label] = pages[0]
        elif len(pages) > 1:
            warnings.append(
                f"PDF 文本页标 {label} 命中多个 PDF 页 {pages}；已拒绝自动映射，请提供 pdf_page_map。"
            )
    return lookup, warnings


def _resolve_pdf_page(project_page: Any,
                      *,
                      explicit_pdf_page: Any = None,
                      pdf_page_map: Mapping[str, Any],
                      pdf_text_page_map: Optional[Mapping[str, Any]] = None,
                      page_count: int,
                      allow_page_number_fallback: bool = False) -> Tuple[int, str]:
    explicit = _coerce_int(explicit_pdf_page)
    if explicit > 0:
        return explicit, "target_pdf_page"
    label = _normalize_page_label(project_page)
    lookup = _page_map_lookup(pdf_page_map)
    for key in {label.upper(), _safe_text(project_page).upper(), str(_page_number_from_label(project_page))}:
        if key and key in lookup:
            return lookup[key], "pdf_page_map"
    text_lookup = _page_map_lookup(pdf_text_page_map or {})
    for key in {label.upper(), _safe_text(project_page).upper(), str(_page_number_from_label(project_page))}:
        if key and key in text_lookup:
            return text_lookup[key], "pdf_text_page_label"
    number = _page_number_from_label(project_page)
    if allow_page_number_fallback and number > 0 and (page_count <= 0 or number <= page_count):
        return number, "page_label_number_weak"
    return 0, "unresolved"


def _page_by_number(pages: Sequence[Mapping[str, Any]], page_number: int) -> Dict[str, Any]:
    for page in pages:
        if _coerce_int(page.get("pdf_page_number")) == page_number:
            return dict(page)
    return {}


def _component_pages(comp: Mapping[str, Any]) -> List[Dict[str, Any]]:
    candidates: List[Dict[str, Any]] = []

    def add_candidate(source: str, item: Mapping[str, Any]) -> None:
        project_page = component_user_visible_page(dict(item)) or item.get("page_real") or item.get("page") or item.get("page_logical")
        if not project_page:
            return
        xy = _point_from_values(item.get("xy_x"), item.get("xy_y"))
        candidates.append({
            "source": source,
            "project_page": _normalize_page_label(project_page),
            "schematic_xy": list(xy) if xy else [],
            "section_number": _safe_text(item.get("section_number")),
            "raw_xy": _safe_text(item.get("xy")),
        })

    add_candidate("component", comp)
    for section in comp.get("sections", []) or []:
        if isinstance(section, Mapping):
            add_candidate("section", section)

    deduped: List[Dict[str, Any]] = []
    seen = set()
    for candidate in candidates:
        key = (
            candidate.get("project_page", ""),
            tuple(candidate.get("schematic_xy") or []),
            candidate.get("section_number", ""),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def _calibration_lookup(page_calibrations: Sequence[Mapping[str, Any]]) -> Dict[str, Dict[str, Any]]:
    lookup: Dict[str, Dict[str, Any]] = {}
    for calibration in page_calibrations or []:
        if not isinstance(calibration, Mapping):
            continue
        page = _normalize_page_label(calibration.get("project_page") or calibration.get("page") or calibration.get("page_label"))
        pdf_page = _coerce_int(calibration.get("pdf_page_number") or calibration.get("pdf_page"))
        schematic_bbox = _bbox_from_value(calibration.get("schematic_bbox"))
        pdf_bbox = _bbox_from_value(calibration.get("pdf_bbox"))
        if not schematic_bbox or not pdf_bbox:
            continue
        item = {
            "project_page": page,
            "pdf_page_number": pdf_page,
            "schematic_bbox": schematic_bbox,
            "pdf_bbox": pdf_bbox,
            "invert_y": bool(calibration.get("invert_y", True)),
        }
        for key in {page.upper(), str(_page_number_from_label(page)), str(pdf_page) if pdf_page else ""}:
            if key and key != "0":
                lookup[key] = item
    return lookup


def _map_schematic_xy_to_pdf_bbox(xy: Sequence[float],
                                  *,
                                  project_page: str,
                                  pdf_page_number: int,
                                  page_calibrations: Sequence[Mapping[str, Any]]) -> Tuple[List[float], str]:
    if not xy or len(xy) < 2:
        return [], "missing_xy"
    lookup = _calibration_lookup(page_calibrations)
    calibration = None
    for key in {project_page.upper(), str(_page_number_from_label(project_page)), str(pdf_page_number)}:
        if key and key in lookup:
            calibration = lookup[key]
            break
    if not calibration:
        return [], "missing_calibration"
    sx0, sy0, sx1, sy1 = calibration["schematic_bbox"]
    px0, py0, px1, py1 = calibration["pdf_bbox"]
    if sx1 == sx0 or sy1 == sy0:
        return [], "invalid_calibration"
    x, y = float(xy[0]), float(xy[1])
    tx = (x - sx0) / (sx1 - sx0)
    ty = (y - sy0) / (sy1 - sy0)
    if calibration.get("invert_y", True):
        ty = 1.0 - ty
    pdf_x = px0 + tx * (px1 - px0)
    pdf_y = py0 + ty * (py1 - py0)
    return _marker_bbox(pdf_x, pdf_y), "page_calibration"


def _target_base(target: Mapping[str, Any], index: int) -> Dict[str, Any]:
    kind = _safe_text(target.get("kind") or target.get("type") or "refdes").lower()
    severity = _safe_text(target.get("severity") or target.get("level") or "info") or "info"
    return {
        "id": _safe_text(target.get("id")) or f"ann-{index}",
        "kind": kind,
        "label": _safe_text(target.get("label") or target.get("title")),
        "severity": severity,
        "color": _safe_text(target.get("color")) or _severity_color(severity),
        "message": _safe_text(target.get("message") or target.get("note")),
        "source": _safe_text(target.get("source")),
        "raw": dict(target),
    }


def _expand_targets(targets: Sequence[Mapping[str, Any]], bundle: Mapping[str, Any]) -> List[Dict[str, Any]]:
    components = bundle.get("components", {}) or {}
    nets = bundle.get("nets", {}) or {}
    expanded: List[Dict[str, Any]] = []
    for index, raw_target in enumerate(targets or [], start=1):
        if not isinstance(raw_target, Mapping):
            continue
        target = _target_base(raw_target, index)
        kind = target["kind"]
        if kind == "refdes":
            refdes = _safe_text(raw_target.get("refdes") or raw_target.get("name") or raw_target.get("value")).upper()
            comp = components.get(refdes) or components.get(_safe_text(raw_target.get("refdes") or raw_target.get("name") or raw_target.get("value")))
            pages = _component_pages(comp) if isinstance(comp, Mapping) else []
            if not pages:
                pages = [{
                    "source": "target",
                    "project_page": _normalize_page_label(raw_target.get("page") or raw_target.get("project_page")),
                    "schematic_xy": [],
                    "section_number": "",
                    "raw_xy": "",
                }]
            for part_index, page in enumerate(pages, start=1):
                expanded.append({
                    **target,
                    "id": target["id"] if len(pages) == 1 else f"{target['id']}-{part_index}",
                    "refdes": refdes,
                    "label": target["label"] or refdes,
                    "project_page": page.get("project_page", ""),
                    "schematic_xy": list(page.get("schematic_xy") or []),
                    "section_number": page.get("section_number", ""),
                    "component_found": bool(comp),
                    "component": {
                        "refdes": refdes,
                        "part_name": _safe_text((comp or {}).get("part_name")) if isinstance(comp, Mapping) else "",
                        "value": _safe_text((comp or {}).get("value")) if isinstance(comp, Mapping) else "",
                        "hq_code": _safe_text((comp or {}).get("hq_code")) if isinstance(comp, Mapping) else "",
                        "comp_type": _safe_text((comp or {}).get("comp_type")) if isinstance(comp, Mapping) else "",
                    },
                })
        elif kind == "net":
            net_name = _safe_text(raw_target.get("net") or raw_target.get("name") or raw_target.get("value"))
            nodes = list((nets or {}).get(net_name, []) or [])
            if not nodes:
                nodes = [{"refdes": "", "pin": "", "pin_name": ""}]
            for node_index, node in enumerate(nodes, start=1):
                refdes = _safe_text((node or {}).get("refdes")).upper()
                comp = components.get(refdes) if refdes else {}
                page_candidates = _component_pages(comp) if isinstance(comp, Mapping) else []
                if not page_candidates:
                    page_candidates = [{
                        "source": "net_node",
                        "project_page": _normalize_page_label(raw_target.get("page") or raw_target.get("project_page")),
                        "schematic_xy": [],
                        "section_number": "",
                    }]
                for part_index, page in enumerate(page_candidates, start=1):
                    expanded.append({
                        **target,
                        "id": f"{target['id']}-{node_index}" if len(nodes) > 1 else target["id"],
                        "kind": "refdes",
                        "source_kind": "net",
                        "net": net_name,
                        "refdes": refdes,
                        "label": target["label"] or (f"{net_name}:{refdes}" if refdes else net_name),
                        "project_page": page.get("project_page", ""),
                        "schematic_xy": list(page.get("schematic_xy") or []),
                        "section_number": page.get("section_number", ""),
                        "component_found": bool(comp),
                        "net_node": dict(node or {}),
                    })
        elif kind == "page":
            page = _normalize_page_label(raw_target.get("page") or raw_target.get("project_page") or raw_target.get("value"))
            expanded.append({
                **target,
                "label": target["label"] or page,
                "project_page": page,
                "schematic_xy": [],
            })
        elif kind == "coordinate":
            page = _normalize_page_label(raw_target.get("page") or raw_target.get("project_page"))
            xy = _point_from_values(raw_target.get("x") or raw_target.get("schematic_x"), raw_target.get("y") or raw_target.get("schematic_y"))
            expanded.append({
                **target,
                "label": target["label"] or _safe_text(raw_target.get("refdes") or raw_target.get("net") or "coordinate"),
                "project_page": page,
                "schematic_xy": list(xy) if xy else [],
                "pdf_bbox": _bbox_from_value(raw_target.get("pdf_bbox") or raw_target.get("bbox")),
                "coordinate_space": _safe_text(raw_target.get("coordinate_space") or ("pdf" if raw_target.get("pdf_bbox") or raw_target.get("bbox") else "schematic")),
            })
        else:
            expanded.append({
                **target,
                "label": target["label"] or kind,
                "project_page": _normalize_page_label(raw_target.get("page") or raw_target.get("project_page")),
                "schematic_xy": [],
            })
    return expanded


def _match_text_box(target: Mapping[str, Any],
                    text_boxes: Sequence[Mapping[str, Any]],
                    *,
                    preferred_page: int) -> Dict[str, Any]:
    refdes = _normalize_token(target.get("refdes") or target.get("label"))
    if not refdes:
        return {}
    matches = [
        dict(box)
        for box in text_boxes or []
        if _normalize_token(box.get("normalized_text") or box.get("text")) == refdes
    ]
    if preferred_page > 0:
        page_matches = [
            item for item in matches
            if _coerce_int(item.get("pdf_page_number")) == preferred_page
        ]
        if page_matches:
            return page_matches[0]
    return matches[0] if matches else {}


def _build_overlay(annotation: Mapping[str, Any], page: Mapping[str, Any]) -> Dict[str, Any]:
    bbox = list(annotation.get("pdf_bbox") or [])
    overlay = {
        "annotation_id": annotation.get("id", ""),
        "pdf_page_number": annotation.get("pdf_page_number", 0),
        "shape": "rect" if bbox else "page_note",
        "label": annotation.get("label", ""),
        "severity": annotation.get("severity", "info"),
        "color": annotation.get("color", "#2563eb"),
        "pdf_bbox": bbox,
        "normalized_bbox": _normalized_bbox(bbox, page) if bbox else [],
        "message": annotation.get("message", ""),
    }
    if annotation.get("confidence") == "page_only":
        overlay["shape"] = "page_note"
    return overlay


def _annotation_status(confidence: str) -> str:
    if confidence in {"explicit_pdf_bbox", "calibrated_xy", "pdf_text_match"}:
        return "matched"
    if confidence == "page_only":
        return "page_only"
    return "unmatched"


def build_schematic_pdf_annotation_payload(
    pdf_path: str | Path,
    bundle: Mapping[str, Any],
    targets: Sequence[Mapping[str, Any]],
    *,
    pdf_page_map: Optional[Mapping[str, Any]] = None,
    page_calibrations: Optional[Sequence[Mapping[str, Any]]] = None,
    stdout: str = "summary",
    limit: int = 200,
    text_boxes: Optional[Sequence[Mapping[str, Any]]] = None,
    allow_page_number_fallback: bool = False,
) -> Dict[str, Any]:
    """Build schematic PDF annotation overlay payload for a report bundle."""
    pdf_meta = read_pdf_metadata(pdf_path)
    pages = list(pdf_meta.get("pages", []) or [])
    page_count = int(pdf_meta.get("page_count", 0) or 0)
    warnings = list(pdf_meta.get("warnings", []) or [])
    raw_pdf_page_map = dict(pdf_page_map or {})
    pdf_page_map, pdf_page_map_meta, page_map_warnings = _normalize_pdf_page_map_payload(
        raw_pdf_page_map,
        actual_pdf_sha256=str(pdf_meta.get("sha256", "") or ""),
    )
    warnings.extend(page_map_warnings)
    page_calibrations = list(page_calibrations or [])
    expanded_targets = _expand_targets(targets, bundle or {})

    project_pages_for_pdf_labels = [
        _normalize_page_label(target.get("project_page"))
        for target in expanded_targets
        if _normalize_page_label(target.get("project_page"))
    ]
    page_label_search_terms: List[str] = []
    for page_label in project_pages_for_pdf_labels:
        page_number = _page_number_from_label(page_label)
        page_label_search_terms.append(page_label)
        if page_number > 0:
            page_label_search_terms.extend(["PAGE", str(page_number)])
    search_terms = [
        target.get("refdes") or target.get("label")
        for target in expanded_targets
        if target.get("kind") == "refdes"
    ]
    search_terms.extend(page_label_search_terms)
    dependency_status: Dict[str, Any] = {"text_backend": "not_requested"}
    if text_boxes is None and search_terms:
        text_boxes, dependency_status = extract_pdf_text_boxes(pdf_path, terms=search_terms)
        if dependency_status.get("error"):
            warnings.append(f"PDF 文本 bbox 不可用：{dependency_status.get('error')}")
    elif text_boxes is not None:
        dependency_status = {
            "text_backend": "provided",
            "available": True,
            "ok": True,
            "box_count": len(text_boxes),
        }
    else:
        text_boxes = []

    pdf_text_page_map, page_label_warnings = _build_pdf_text_page_label_map(
        text_boxes or [],
        project_pages_for_pdf_labels,
    )
    warnings.extend(page_label_warnings)
    if allow_page_number_fallback:
        warnings.append(
            "已启用 PAGE<N> -> PDF 第 N 页弱兜底；PDF 重排、插入封面或分册导出时可能漂移，建议提供 pdf_page_map。"
        )

    annotations: List[Dict[str, Any]] = []
    for target in expanded_targets:
        project_page = _normalize_page_label(target.get("project_page"))
        pdf_page_number, page_source = _resolve_pdf_page(
            project_page,
            explicit_pdf_page=(target.get("raw") or {}).get("pdf_page_number") or (target.get("raw") or {}).get("pdf_page"),
            pdf_page_map=pdf_page_map,
            pdf_text_page_map=pdf_text_page_map,
            page_count=page_count,
            allow_page_number_fallback=allow_page_number_fallback,
        )
        page = _page_by_number(pages, pdf_page_number)
        pdf_bbox = list(target.get("pdf_bbox") or [])
        locator_source = ""
        confidence = ""
        text_match = {}

        if pdf_bbox:
            locator_source = "target.pdf_bbox"
            confidence = "explicit_pdf_bbox" if pdf_page_number > 0 else "explicit_pdf_bbox_missing_page"
        elif target.get("schematic_xy"):
            mapped_bbox, mapped_source = _map_schematic_xy_to_pdf_bbox(
                target.get("schematic_xy") or [],
                project_page=project_page,
                pdf_page_number=pdf_page_number,
                page_calibrations=page_calibrations,
            )
            if mapped_bbox:
                pdf_bbox = mapped_bbox
                confidence = "calibrated_xy"
                locator_source = mapped_source

        if not pdf_bbox and target.get("kind") == "refdes":
            text_match = _match_text_box(target, text_boxes or [], preferred_page=pdf_page_number)
            if text_match:
                pdf_bbox = _bbox_from_value(text_match.get("bbox"))
                matched_page = _coerce_int(text_match.get("pdf_page_number"))
                if matched_page > 0:
                    pdf_page_number = matched_page
                    page = _page_by_number(pages, pdf_page_number)
                confidence = "pdf_text_match"
                locator_source = _safe_text(text_match.get("source")) or "pdf_text"

        if not confidence:
            confidence = "page_only" if pdf_page_number > 0 else "unmatched"
            locator_source = page_source

        annotation = {
            "id": target.get("id", ""),
            "kind": target.get("source_kind") or target.get("kind", ""),
            "target_kind": target.get("kind", ""),
            "status": _annotation_status(confidence),
            "confidence": confidence,
            "locator_source": locator_source,
            "label": target.get("label", ""),
            "severity": target.get("severity", "info"),
            "color": target.get("color", "#2563eb"),
            "message": target.get("message", ""),
            "refdes": target.get("refdes", ""),
            "net": target.get("net", ""),
            "project_page": project_page,
            "pdf_page_number": pdf_page_number,
            "pdf_page_source": page_source,
            "pdf_bbox": pdf_bbox,
            "normalized_bbox": _normalized_bbox(pdf_bbox, page) if pdf_bbox and page else [],
            "schematic_xy": list(target.get("schematic_xy") or []),
            "section_number": target.get("section_number", ""),
            "component_found": bool(target.get("component_found", False)),
            "component": dict(target.get("component") or {}),
            "net_node": dict(target.get("net_node") or {}),
            "text_match": dict(text_match or {}),
            "target": dict(target.get("raw") or {}),
        }
        annotation["overlay"] = _build_overlay(annotation, page)
        annotations.append(annotation)

    unresolved_project_pages = sorted({
        str(item.get("project_page") or "")
        for item in annotations
        if item.get("project_page") and not item.get("pdf_page_number")
    })
    if unresolved_project_pages:
        sample = ", ".join(unresolved_project_pages[:5])
        more = "" if len(unresolved_project_pages) <= 5 else f" 等 {len(unresolved_project_pages)} 个"
        warnings.append(
            f"有标注只定位到项目真实页码但未找到可靠 PDF 页：{sample}{more}；"
            "请提供 pdf_page_map，或确保 PDF 文本中存在唯一 PAGE<N> 页标。"
        )

    limit = max(1, min(int(limit or 200), 5000))
    returned_annotations = annotations[:limit]
    truncated = len(annotations) > limit
    pages_by_number = {int(page.get("pdf_page_number") or 0): dict(page) for page in pages}
    page_overlays: List[Dict[str, Any]] = []
    for page_number in sorted({int(item.get("pdf_page_number") or 0) for item in returned_annotations if item.get("pdf_page_number")}):
        page_items = [item for item in returned_annotations if int(item.get("pdf_page_number") or 0) == page_number]
        page = pages_by_number.get(page_number, {})
        page_overlays.append({
            "pdf_page_number": page_number,
            "width": page.get("width", 0),
            "height": page.get("height", 0),
            "annotation_count": len(page_items),
            "overlays": [item.get("overlay", {}) for item in page_items],
        })

    confidence_counts: Dict[str, int] = {}
    for item in annotations:
        key = str(item.get("confidence") or "unknown")
        confidence_counts[key] = confidence_counts.get(key, 0) + 1
    matched_count = sum(1 for item in annotations if item.get("status") == "matched")
    page_only_count = sum(1 for item in annotations if item.get("status") == "page_only")
    unmatched_count = sum(1 for item in annotations if item.get("status") == "unmatched")
    payload: Dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "digest": {
            "pdf_sha256": pdf_meta.get("sha256", ""),
            "pdf_size": pdf_meta.get("size", 0),
            "pdf_page_count": page_count,
            "target_count": len(expanded_targets),
            "page_mapping_policy": "weak_page_number_fallback" if allow_page_number_fallback else "strict",
        },
        "pdf": {
            "path": pdf_meta.get("path", ""),
            "filename": pdf_meta.get("filename", ""),
            "page_count": page_count,
            "pages": pages if stdout in {"annotations", "full"} else [],
        },
        "summary": {
            "target_count": len(expanded_targets),
            "annotation_count": len(annotations),
            "returned_annotation_count": len(returned_annotations),
            "matched_count": matched_count,
            "page_only_count": page_only_count,
            "unmatched_count": unmatched_count,
            "confidence_counts": confidence_counts,
        },
        "annotations": returned_annotations if stdout in {"annotations", "full"} else [],
        "page_overlays": page_overlays if stdout in {"annotations", "full"} else [],
        "warnings": warnings,
        "dependency_status": dependency_status,
        "truncated": truncated,
    }
    if stdout == "full":
        payload["inputs"] = {
            "targets": [dict(target) for target in targets or [] if isinstance(target, Mapping)],
            "pdf_page_map": dict(pdf_page_map),
            "pdf_page_map_meta": dict(pdf_page_map_meta),
            "page_calibrations": [dict(item) for item in page_calibrations if isinstance(item, Mapping)],
            "pdf_text_page_map": dict(pdf_text_page_map),
            "allow_page_number_fallback": bool(allow_page_number_fallback),
        }
    return payload


def load_targets_json(value: str | Path | Mapping[str, Any] | Sequence[Any]) -> List[Dict[str, Any]]:
    """Load annotation targets from a JSON path, object, array or JSON string."""
    if isinstance(value, Mapping):
        raw = value.get("targets", [])
    elif isinstance(value, (list, tuple)):
        raw = value
    else:
        text = _safe_text(value)
        if not text:
            raw = []
        else:
            path = Path(text).expanduser()
            if path.is_file():
                loaded = json.loads(path.read_text(encoding="utf-8"))
            else:
                loaded = json.loads(text)
            raw = loaded.get("targets", []) if isinstance(loaded, Mapping) else loaded
    if not isinstance(raw, list):
        raise ValueError("targets JSON 必须是数组，或包含 targets 数组的对象。")
    targets = [dict(item) for item in raw if isinstance(item, Mapping)]
    if not targets:
        raise ValueError("请提供至少一个 PDF annotation target。")
    return targets


def load_json_mapping_or_sequence(value: Any, *, default: Any) -> Any:
    """Load small JSON option from object, path, string or empty default."""
    if value is None or value == "":
        return default
    if isinstance(value, (Mapping, list, tuple)):
        return value
    text = _safe_text(value)
    if not text:
        return default
    path = Path(text).expanduser()
    if path.is_file():
        return json.loads(path.read_text(encoding="utf-8"))
    return json.loads(text)
