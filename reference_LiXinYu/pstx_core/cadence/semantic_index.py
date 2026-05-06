# -*- coding: utf-8 -*-
"""Project-level Cadence page semantic index."""

from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

from pstx_core.cadence.page_model import (
    CadenceObject,
    CadencePageModel,
    ConnectivityComponent,
    iter_cadence_page_numbers,
    load_cadence_page_model,
)


CADENCE_INDEX_SCHEMA_VERSION = "pstx-cadence-index.v1"
VALID_KINDS = {"all", "net", "port", "offpage", "bus", "no_connect", "unbound"}
VALID_STDOUTS = {"summary", "nets", "ports", "links", "full"}


def _normalized_name(name: object) -> str:
    return str(name or "").strip().upper()


def _display_name(name: object) -> str:
    return str(name or "").strip()


def _unique(values: Iterable[object]) -> List[object]:
    result: List[object] = []
    seen = set()
    for value in values:
        marker = repr(value)
        if marker in seen:
            continue
        seen.add(marker)
        result.append(value)
    return result


def _page_label(page: int) -> str:
    return f"PAGE{int(page)}"


def _object_name(item: CadenceObject) -> str:
    return _display_name(item.attributes.get("name") or item.attributes.get("value") or "")


def _object_direction(item: CadenceObject) -> str:
    return _display_name(item.attributes.get("direction") or "")


def _coords(item: CadenceObject) -> List[List[int]]:
    return [list(point) for point in item.coords]


def _object_brief(item: CadenceObject, *, include_raw: bool = False) -> dict:
    payload = {
        "object_id": item.object_id,
        "type": item.object_type,
        "name": _object_name(item),
        "normalized_name": _normalized_name(_object_name(item)),
        "direction": _object_direction(item),
        "line_no": item.line_no,
        "coords": _coords(item),
        "bbox": list(item.bbox) if item.bbox else None,
    }
    if include_raw:
        payload["raw"] = item.raw
    return payload


def _component_by_semantic_id(model: CadencePageModel) -> Dict[str, ConnectivityComponent]:
    result: Dict[str, ConnectivityComponent] = {}
    for component in model.connectivity:
        for object_id in component.semantic_object_ids:
            result[object_id] = component
    return result


def _component_brief(component: ConnectivityComponent) -> dict:
    return {
        "component_id": component.component_id,
        "signal_names": list(component.signal_names),
        "labels": list(component.labels),
        "ports": list(component.ports),
        "offpage_connectors": list(component.offpage_connectors),
        "bus_names": list(component.bus_names),
        "no_connect_points": [list(point) for point in component.no_connect_points],
        "semantic_object_ids": list(component.semantic_object_ids),
        "object_ids": list(component.object_ids),
        "bbox": list(component.bbox) if component.bbox else None,
    }


def _new_group(name: str, kind: str) -> dict:
    return {
        "kind": kind,
        "name": name,
        "normalized_name": _normalized_name(name),
        "display_names": [],
        "pages": [],
        "page_labels": [],
        "page_numbers": [],
        "component_ids": [],
        "object_ids": [],
        "occurrences": [],
    }


def _append_group_occurrence(group: dict,
                             *,
                             page: int,
                             page_number: str,
                             component_id: str = "",
                             object_id: str = "",
                             occurrence: Optional[dict] = None) -> None:
    if page not in group["pages"]:
        group["pages"].append(page)
    page_label = _page_label(page)
    if page_label not in group["page_labels"]:
        group["page_labels"].append(page_label)
    if page_number and page_number not in group["page_numbers"]:
        group["page_numbers"].append(page_number)
    if component_id and component_id not in group["component_ids"]:
        group["component_ids"].append(component_id)
    if object_id and object_id not in group["object_ids"]:
        group["object_ids"].append(object_id)
    if occurrence:
        group["occurrences"].append(occurrence)


def _final_group(group: dict) -> dict:
    pages = sorted(int(item) for item in group["pages"])
    group["pages"] = pages
    group["page_labels"] = [_page_label(page) for page in pages]
    group["page_count"] = len(pages)
    group["occurrence_count"] = len(group["occurrences"])
    group["component_ids"] = sorted(group["component_ids"])
    group["object_ids"] = sorted(group["object_ids"])
    group["display_names"] = _unique(group["display_names"] or [group["name"]])
    return group


def _pstx_net_names(pstx_nets: Optional[Mapping[str, object]]) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for name in (pstx_nets or {}).keys():
        display = _display_name(name)
        normalized = _normalized_name(display)
        if normalized and normalized not in result:
            result[normalized] = display
    return result


def _matches_query(row: Mapping[str, object], query: str) -> bool:
    normalized_query = _normalized_name(query)
    if not normalized_query:
        return True
    candidates = [
        row.get("name", ""),
        row.get("normalized_name", ""),
        " ".join(str(item) for item in row.get("display_names", []) or []),
        " ".join(str(item) for item in row.get("page_labels", []) or []),
    ]
    return any(normalized_query in _normalized_name(candidate) for candidate in candidates)


def _matches_page(row: Mapping[str, object], page: int) -> bool:
    return page <= 0 or page in set(int(item) for item in row.get("pages", []) or [])


def _filter_rows(rows: Sequence[dict], *, query: str, page: int) -> List[dict]:
    return [
        row for row in rows
        if _matches_page(row, page) and _matches_query(row, query)
    ]


def _limited(rows: Sequence[dict], limit: int) -> Tuple[List[dict], bool]:
    limit = max(1, min(5000, int(limit or 200)))
    return list(rows[:limit]), len(rows) > limit


def _component_occurrence(model: CadencePageModel,
                          component: ConnectivityComponent,
                          *,
                          source: str,
                          object_id: str = "") -> dict:
    return {
        "page": model.page_no,
        "page_label": _page_label(model.page_no),
        "page_number": model.page_number,
        "source": source,
        "component_id": component.component_id,
        "object_id": object_id,
    }


def _object_occurrence(model: CadencePageModel,
                       item: CadenceObject,
                       component: Optional[ConnectivityComponent]) -> dict:
    return {
        "page": model.page_no,
        "page_label": _page_label(model.page_no),
        "page_number": model.page_number,
        "source": item.object_type,
        "component_id": component.component_id if component else "",
        "object_id": item.object_id,
        "direction": _object_direction(item),
        "coords": _coords(item),
        "line_no": item.line_no,
        "binding_status": "bound" if component else "unbound",
    }


def _add_named_group(groups: Dict[str, dict],
                     *,
                     name: str,
                     kind: str,
                     model: CadencePageModel,
                     component: Optional[ConnectivityComponent] = None,
                     item: Optional[CadenceObject] = None,
                     source: str) -> None:
    display = _display_name(name)
    normalized = _normalized_name(display)
    if not normalized:
        return
    group = groups.setdefault(normalized, _new_group(display, kind))
    if display and display not in group["display_names"]:
        group["display_names"].append(display)
    occurrence = (
        _object_occurrence(model, item, component)
        if item is not None else
        _component_occurrence(model, component, source=source)
    )
    _append_group_occurrence(
        group,
        page=model.page_no,
        page_number=model.page_number,
        component_id=component.component_id if component else "",
        object_id=item.object_id if item else "",
        occurrence=occurrence,
    )


def _semantic_rows_for_type(model: CadencePageModel,
                            object_type: str,
                            component_by_id: Mapping[str, ConnectivityComponent]) -> List[dict]:
    rows = []
    for item in model.objects:
        if item.object_type != object_type:
            continue
        component = component_by_id.get(item.object_id)
        rows.append({
            "page": model.page_no,
            "page_label": _page_label(model.page_no),
            "page_number": model.page_number,
            "name": _object_name(item),
            "normalized_name": _normalized_name(_object_name(item)),
            "direction": _object_direction(item),
            "object_id": item.object_id,
            "component_id": component.component_id if component else "",
            "binding_status": "bound" if component else "unbound",
            "line_no": item.line_no,
            "coords": _coords(item),
            "bbox": list(item.bbox) if item.bbox else None,
        })
    return rows


def _build_raw_index(project_root: str | Path,
                     *,
                     pstx_nets: Optional[Mapping[str, object]] = None) -> dict:
    root = Path(project_root).expanduser()
    warnings: List[str] = []
    pages = iter_cadence_page_numbers(root)
    if not str(project_root or "").strip():
        warnings.append("cadence-index 需要 project_root。")
    elif not pages:
        warnings.append(f"未在项目根路径下找到 sch_1/page*.csv 或 page*.csa：{root}")

    net_groups: Dict[str, dict] = {}
    offpage_groups: Dict[str, dict] = {}
    bus_groups: Dict[str, dict] = {}
    port_groups: Dict[str, dict] = {}
    port_rows: List[dict] = []
    no_connect_rows: List[dict] = []
    unbound_rows: List[dict] = []
    loaded_pages = 0
    error_count = 0
    semantic_counts: Counter[str] = Counter()

    for page in pages:
        try:
            model = load_cadence_page_model(
                root,
                "project",
                page,
                include_raw_unknown=False,
                collect_junctions=False,
            )
        except Exception as exc:
            warnings.append(f"PAGE{page} 连接语义加载失败：{exc}")
            error_count += 1
            continue
        loaded_pages += 1
        if model.csv_error or model.csa_error:
            detail = "; ".join(item for item in [model.csa_error, model.csv_error] if item)
            warnings.append(f"PAGE{page} 解析受限：{detail}")
            error_count += 1
        component_by_id = _component_by_semantic_id(model)
        for component in model.connectivity:
            for signal_name in component.signal_names:
                _add_named_group(
                    net_groups,
                    name=signal_name,
                    kind="net",
                    model=model,
                    component=component,
                    source="SIG_NAME",
                )
            for label in component.labels:
                _add_named_group(
                    net_groups,
                    name=label,
                    kind="net",
                    model=model,
                    component=component,
                    source="NET_LABEL",
                )
            for offpage in component.offpage_connectors:
                _add_named_group(
                    offpage_groups,
                    name=offpage,
                    kind="offpage",
                    model=model,
                    component=component,
                    source="OFFPAGE",
                )
            for bus_name in component.bus_names:
                _add_named_group(
                    bus_groups,
                    name=bus_name,
                    kind="bus",
                    model=model,
                    component=component,
                    source="BUS",
                )

        for item in model.objects:
            semantic_counts[item.object_type] += 1
            component = component_by_id.get(item.object_id)
            if item.object_type == "NET_LABEL" and component is None:
                continue
            if item.object_type == "PORT":
                _add_named_group(
                    port_groups,
                    name=_object_name(item),
                    kind="port",
                    model=model,
                    component=component,
                    item=item,
                    source="PORT",
                )
            elif item.object_type == "OFFPAGE" and component is None:
                _add_named_group(
                    offpage_groups,
                    name=_object_name(item),
                    kind="offpage",
                    model=model,
                    component=None,
                    item=item,
                    source="OFFPAGE",
                )
            elif item.object_type == "BUS" and component is None:
                _add_named_group(
                    bus_groups,
                    name=_object_name(item),
                    kind="bus",
                    model=model,
                    component=None,
                    item=item,
                    source="BUS",
                )

        port_rows.extend(_semantic_rows_for_type(model, "PORT", component_by_id))
        no_connect_rows.extend(_semantic_rows_for_type(model, "NO_CONNECT", component_by_id))
        for item in model.unbound_semantics:
            unbound_rows.append({
                "page": model.page_no,
                "page_label": _page_label(model.page_no),
                "page_number": model.page_number,
                "semantic_type": item.object_type,
                "name": _object_name(item),
                "normalized_name": _normalized_name(_object_name(item)),
                "object_id": item.object_id,
                "line_no": item.line_no,
                "coords": _coords(item),
                "bbox": list(item.bbox) if item.bbox else None,
                "binding_status": "unbound",
            })

    pstx_names = _pstx_net_names(pstx_nets)
    net_rows = []
    for row in [_final_group(item) for item in net_groups.values()]:
        normalized = str(row.get("normalized_name") or "")
        row["pstx_net_match"] = normalized in pstx_names
        row["pstx_net_name"] = pstx_names.get(normalized, "")
        source_counts = Counter(str(item.get("source") or "") for item in row.get("occurrences", []))
        row["source_counts"] = dict(sorted(source_counts.items()))
        net_rows.append(row)
    offpage_link_rows = []
    for row in [_final_group(item) for item in offpage_groups.values()]:
        row["link_status"] = "same_name_multi_page_evidence" if row["page_count"] > 1 else "single_page_evidence"
        row["connection_claim"] = False
        offpage_link_rows.append(row)
    bus_rows = [_final_group(item) for item in bus_groups.values()]
    port_group_rows = [_final_group(item) for item in port_groups.values()]
    for row in port_group_rows:
        directions = [
            str(item.get("direction") or "")
            for item in row.get("occurrences", [])
            if str(item.get("direction") or "")
        ]
        row["directions"] = sorted(set(directions))

    return {
        "schema_version": CADENCE_INDEX_SCHEMA_VERSION,
        "enabled": bool(pages),
        "root": str(root),
        "page_count": len(pages),
        "loaded_page_count": loaded_pages,
        "error_count": error_count,
        "semantic_counts": dict(sorted(semantic_counts.items())),
        "net_rows": sorted(net_rows, key=lambda item: str(item.get("normalized_name") or "")),
        "port_rows": sorted(port_group_rows, key=lambda item: str(item.get("normalized_name") or "")),
        "port_object_rows": sorted(port_rows, key=lambda item: (int(item.get("page", 0) or 0), str(item.get("object_id") or ""))),
        "offpage_link_rows": sorted(offpage_link_rows, key=lambda item: str(item.get("normalized_name") or "")),
        "bus_rows": sorted(bus_rows, key=lambda item: str(item.get("normalized_name") or "")),
        "no_connect_rows": sorted(no_connect_rows, key=lambda item: (int(item.get("page", 0) or 0), str(item.get("object_id") or ""))),
        "unbound_semantic_rows": sorted(unbound_rows, key=lambda item: (int(item.get("page", 0) or 0), str(item.get("object_id") or ""))),
        "warnings": warnings,
        "readonly": True,
    }


def build_cadence_index_payload(project_root: str | Path,
                                *,
                                pstx_nets: Optional[Mapping[str, object]] = None,
                                stdout: str = "summary",
                                query: str = "",
                                kind: str = "all",
                                page: int = 0,
                                limit: int = 200) -> dict:
    """Build a filtered project-level Cadence semantic index payload."""

    mode = str(stdout or "summary").strip().lower()
    if mode not in VALID_STDOUTS:
        raise ValueError("stdout 必须是 summary、nets、ports、links 或 full。")
    kind_value = str(kind or "all").strip().lower()
    if kind_value not in VALID_KINDS:
        raise ValueError("kind 必须是 all、net、port、offpage、bus、no_connect 或 unbound。")
    page_no = max(0, int(page or 0))
    limit_value = max(1, min(5000, int(limit or 200)))

    raw = _build_raw_index(project_root, pstx_nets=pstx_nets)
    row_sets = {
        "net_rows": _filter_rows(raw["net_rows"], query=query, page=page_no),
        "port_rows": _filter_rows(raw["port_rows"], query=query, page=page_no),
        "offpage_link_rows": _filter_rows(raw["offpage_link_rows"], query=query, page=page_no),
        "bus_rows": _filter_rows(raw["bus_rows"], query=query, page=page_no),
        "no_connect_rows": _filter_rows(raw["no_connect_rows"], query=query, page=page_no),
        "unbound_semantic_rows": _filter_rows(raw["unbound_semantic_rows"], query=query, page=page_no),
    }
    if kind_value != "all":
        keep = {
            "net": "net_rows",
            "port": "port_rows",
            "offpage": "offpage_link_rows",
            "bus": "bus_rows",
            "no_connect": "no_connect_rows",
            "unbound": "unbound_semantic_rows",
        }[kind_value]
        row_sets = {key: (value if key == keep else []) for key, value in row_sets.items()}

    mode_rows = {
        "summary": set(),
        "nets": {"net_rows"},
        "ports": {"port_rows"},
        "links": {"offpage_link_rows"},
        "full": set(row_sets.keys()),
    }[mode]
    payload = {
        "schema_version": CADENCE_INDEX_SCHEMA_VERSION,
        "digest": {
            "schema_version": CADENCE_INDEX_SCHEMA_VERSION,
            "enabled": bool(raw.get("enabled")),
            "root": raw.get("root", ""),
            "page_count": raw.get("page_count", 0),
            "loaded_page_count": raw.get("loaded_page_count", 0),
            "error_count": raw.get("error_count", 0),
            "net_count": len(row_sets["net_rows"]),
            "port_count": len(row_sets["port_rows"]),
            "offpage_link_count": len(row_sets["offpage_link_rows"]),
            "bus_count": len(row_sets["bus_rows"]),
            "no_connect_count": len(row_sets["no_connect_rows"]),
            "unbound_semantic_count": len(row_sets["unbound_semantic_rows"]),
            "warning_count": len(raw.get("warnings", []) or []),
            "page_filter": page_no,
            "query": str(query or ""),
            "kind": kind_value,
            "stdout": mode,
        },
        "filters": {
            "stdout": mode,
            "query": str(query or ""),
            "kind": kind_value,
            "page": page_no,
            "limit": limit_value,
        },
        "warnings": list(raw.get("warnings", []) or []),
        "net_rows": [],
        "port_rows": [],
        "offpage_link_rows": [],
        "bus_rows": [],
        "no_connect_rows": [],
        "unbound_semantic_rows": [],
        "truncated": False,
        "truncation": {
            "net_rows": False,
            "port_rows": False,
            "offpage_link_rows": False,
            "bus_rows": False,
            "no_connect_rows": False,
            "unbound_semantic_rows": False,
        },
        "readonly": True,
    }
    for key, rows in row_sets.items():
        if key not in mode_rows:
            continue
        payload[key], payload["truncation"][key] = _limited(rows, limit_value)
    payload["truncated"] = any(payload["truncation"].values())
    return payload
