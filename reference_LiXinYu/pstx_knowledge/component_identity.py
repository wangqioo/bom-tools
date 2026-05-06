# -*- coding: utf-8 -*-
"""Component identity cards for DFMEA preparation harness tools."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Tuple

from pstx_knowledge.feishu_cache import get_feishu_cache_rows


USER_VISIBLE_REAL_PAGE_LABEL = "页码"
POWER_NET_RE = re.compile(
    r"(?i)(^|[_\-/])(?:P\d|P[0-9A-Z]*V|VDD|VCC|VIN|VOUT|VSYS|VBAT|GND|AGND|DGND|PGND|VSS|0V)"
)
INTERFACE_NET_RE = re.compile(
    r"(?i)(I2C|SCL|SDA|SPI|MOSI|MISO|SCLK|UART|TX|RX|GPIO|PCIE|USB|DP|HDMI|MDI|CLK|RESET|RST|INT|ALERT|FAULT|ENABLE|EN)"
)


@dataclass
class ComponentIdentityCard:
    refdes: str
    category: str
    candidate_chip_type: str
    hq_no: str = ""
    spec: str = ""
    pi: str = ""
    selection_order: str = ""
    package: str = ""
    value: str = ""
    bom_option: str = ""
    user_visible_page: str = ""
    pin_net_summary: List[dict] = field(default_factory=list)
    power_nets: List[str] = field(default_factory=list)
    interface_nets: List[str] = field(default_factory=list)
    feishu_match: dict = field(default_factory=dict)
    datasheet_match: dict = field(default_factory=dict)
    datasheet_evidence_refs: List[dict] = field(default_factory=list)
    datasheet_missing_reason: str = ""
    missing_fields: List[str] = field(default_factory=list)
    confidence: str = "medium"
    evidence_refs: List[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "refdes": self.refdes,
            "category": self.category,
            "candidate_chip_type": self.candidate_chip_type,
            "hq_no": self.hq_no,
            "spec": self.spec,
            "pi": self.pi,
            "selection_order": self.selection_order,
            "package": self.package,
            "value": self.value,
            "bom_option": self.bom_option,
            USER_VISIBLE_REAL_PAGE_LABEL: self.user_visible_page,
            "user_visible_page": self.user_visible_page,
            "pin_net_summary": list(self.pin_net_summary),
            "power_nets": list(self.power_nets),
            "interface_nets": list(self.interface_nets),
            "feishu_match": dict(self.feishu_match),
            "datasheet_match": dict(self.datasheet_match),
            "datasheet_evidence_refs": list(self.datasheet_evidence_refs),
            "datasheet_missing_reason": self.datasheet_missing_reason,
            "missing_fields": list(self.missing_fields),
            "confidence": self.confidence,
            "evidence_refs": list(self.evidence_refs),
        }


def _safe_text(value, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _ref_prefix(refdes: str) -> str:
    match = re.match(r"^[A-Za-z]+", str(refdes or "").strip())
    return match.group(0).upper() if match else ""


def classify_refdes(refdes: str, comp: Optional[dict] = None) -> Tuple[str, str, str]:
    ref = str(refdes or "").strip().upper()
    prefix = _ref_prefix(ref)
    if ref.startswith("PU"):
        return "power_ic", "power_management_ic", "high"
    if ref.startswith("XU"):
        return "large_ic", "large_ic_or_module", "medium"
    if ref.startswith("U"):
        return "chip", "ic", "medium"
    if prefix in {"J", "P", "CN", "CON", "X"}:
        return "connector", "connector", "medium"
    if prefix == "R":
        return "passive", "resistor", "high"
    if prefix == "C":
        return "passive", "capacitor", "high"
    if prefix == "L":
        return "passive", "inductor", "high"
    if prefix in {"D", "Q", "Y", "FB", "F"}:
        return "discrete", prefix.lower(), "medium"
    return "unknown", "needs_context", "low"


def _component_hq_no(comp: dict) -> str:
    for key in ("hq_code", "HQ_CODE", "料号", "part_number", "PART_NUMBER"):
        value = _safe_text(comp.get(key, ""), 160)
        if value:
            return value
    return ""


def _component_spec(comp: dict) -> str:
    for key in ("part_type", "cds_part_name", "CDS_PART_NAME", "description", "型号", "规格型号"):
        value = _safe_text(comp.get(key, ""), 260)
        if value:
            return value
    return ""


def _component_page(comp: dict) -> str:
    for key in ("page_submodule_mapped", "user_visible_page", "page_real", USER_VISIBLE_REAL_PAGE_LABEL, "用户看到的真实页", "真实页", "page"):
        value = _safe_text(comp.get(key, ""), 80)
        if value:
            return value
    return ""


def _iter_component_pin_nets(refdes: str, comp: dict, nets: dict) -> List[dict]:
    rows = []
    comp_nets = comp.get("nets") if isinstance(comp.get("nets"), dict) else {}
    for pin, net_name in sorted(comp_nets.items(), key=lambda item: str(item[0])):
        rows.append({
            "pin": _safe_text(pin, 80),
            "pin_name": _safe_text(pin, 160),
            "net": _safe_text(net_name, 180),
            "source": "component.nets",
        })
    for net_name, nodes in (nets or {}).items():
        for node in nodes or []:
            if not isinstance(node, dict) or str(node.get("refdes") or "") != refdes:
                continue
            row = {
                "pin": _safe_text(node.get("pin", ""), 80),
                "pin_name": _safe_text(node.get("pin_name", ""), 160),
                "net": _safe_text(net_name, 180),
                "source": "nets",
            }
            if not any(existing["pin"] == row["pin"] and existing["net"] == row["net"] for existing in rows):
                rows.append(row)
    return rows


def _summarize_pin_nets(pin_nets: List[dict], limit: int = 16) -> List[dict]:
    return pin_nets[:max(0, limit)]


def _unique_matching_nets(pin_nets: Iterable[dict], pattern: re.Pattern, limit: int = 12) -> List[str]:
    result = []
    for row in pin_nets:
        name = _safe_text(row.get("net", ""), 180)
        if not name or not pattern.search(name):
            continue
        if name not in result:
            result.append(name)
        if len(result) >= limit:
            break
    return result


def _find_feishu_match(hq_no: str, spec: str = "", cache: Optional[Dict[str, dict]] = None) -> dict:
    query = hq_no or spec
    if not query:
        return {"status": "missing_query"}
    cache_key = query.strip().upper()
    if cache is not None and cache_key in cache:
        return dict(cache[cache_key])
    result = get_feishu_cache_rows(query=query, limit=5)
    if not result.get("ok"):
        match = {"status": "unavailable", "error": _safe_text(result.get("error", ""), 240)}
        if cache is not None:
            cache[cache_key] = dict(match)
        return match
    rows = list(result.get("rows") or [])
    if not rows:
        match = {"status": "not_found", "query": query}
        if cache is not None:
            cache[cache_key] = dict(match)
        return match
    preferred = None
    if hq_no:
        normalized = hq_no.strip().upper()
        for row in rows:
            if str(row.get("hq_no") or "").strip().upper() == normalized:
                preferred = row
                break
    preferred = preferred or rows[0]
    match = {
        "status": "matched",
        "row_id": preferred.get("id", ""),
        "hq_no": _safe_text(preferred.get("hq_no", ""), 160),
        "spec": _safe_text(preferred.get("spec") or preferred.get("key_value") or "", 260),
        "pi": _safe_text(preferred.get("pi", ""), 160),
        "selection_order": _safe_text(preferred.get("selection_order", ""), 120),
        "lib_name": _safe_text(preferred.get("lib_name", ""), 160),
        "sheet_name": _safe_text(preferred.get("sheet_name", ""), 160),
    }
    if cache is not None:
        cache[cache_key] = dict(match)
    return match


def _missing_fields(card: ComponentIdentityCard) -> List[str]:
    missing = []
    if not card.hq_no:
        missing.append("hq_no")
    if not card.spec:
        missing.append("spec")
    if card.category in {"chip", "power_ic", "large_ic", "connector"} and not card.pin_net_summary:
        missing.append("pin_net_summary")
    if card.category in {"chip", "power_ic", "large_ic"} and not card.power_nets:
        missing.append("power_nets")
    if card.feishu_match.get("status") != "matched":
        missing.append("feishu_match")
    return missing


def build_component_identity_cards(report: dict, bundle: dict) -> List[dict]:
    components = bundle.get("components") if isinstance(bundle.get("components"), dict) else {}
    nets = bundle.get("nets") if isinstance(bundle.get("nets"), dict) else {}
    feishu_match_cache: Dict[str, dict] = {}
    cards: List[ComponentIdentityCard] = []
    for refdes in sorted(components.keys(), key=lambda item: str(item).upper()):
        comp = components.get(refdes) if isinstance(components.get(refdes), dict) else {}
        category, candidate, confidence = classify_refdes(refdes, comp)
        hq_no = _component_hq_no(comp)
        spec = _component_spec(comp)
        feishu_match = _find_feishu_match(hq_no, spec, feishu_match_cache)
        if not spec and feishu_match.get("spec"):
            spec = str(feishu_match.get("spec") or "")
        pin_nets = _iter_component_pin_nets(refdes, comp, nets)
        card = ComponentIdentityCard(
            refdes=str(refdes),
            category=category,
            candidate_chip_type=candidate,
            hq_no=hq_no or str(feishu_match.get("hq_no") or ""),
            spec=spec,
            pi=_safe_text(comp.get("PI") or comp.get("pi") or feishu_match.get("pi") or "", 160),
            selection_order=_safe_text(comp.get("选型顺序") or comp.get("selection_order") or feishu_match.get("selection_order") or "", 120),
            package=_safe_text(comp.get("package") or comp.get("PACKAGE") or comp.get("封装") or "", 120),
            value=_safe_text(comp.get("value") or comp.get("VALUE") or comp.get("值") or "", 120),
            bom_option=_safe_text(comp.get("bom_option") or comp.get("BOM_OPTION") or "", 120),
            user_visible_page=_component_page(comp),
            pin_net_summary=_summarize_pin_nets(pin_nets),
            power_nets=_unique_matching_nets(pin_nets, POWER_NET_RE),
            interface_nets=_unique_matching_nets(pin_nets, INTERFACE_NET_RE),
            feishu_match=feishu_match,
            confidence=confidence,
            evidence_refs=[
                {"source": "bundle.components", "refdes": str(refdes)},
                {"source": "bundle.nets", "count": len(pin_nets)},
            ],
        )
        card.missing_fields = _missing_fields(card)
        if card.missing_fields and card.confidence == "high":
            card.confidence = "medium"
        if card.category == "unknown":
            card.confidence = "low"
        cards.append(card)
    return [card.to_dict() for card in cards]


def filter_component_identity_cards(cards: Iterable[dict],
                                    *,
                                    category: str = "",
                                    refdes_prefix: str = "",
                                    hq_no: str = "",
                                    feishu_status: str = "",
                                    query: str = "") -> List[dict]:
    result = []
    query_lower = str(query or "").strip().lower()
    for card in cards:
        if category and str(card.get("category") or "") != category:
            continue
        if refdes_prefix and not str(card.get("refdes") or "").upper().startswith(refdes_prefix.upper()):
            continue
        if hq_no and str(card.get("hq_no") or "").upper() != hq_no.upper():
            continue
        if feishu_status and str((card.get("feishu_match") or {}).get("status") or "") != feishu_status:
            continue
        if query_lower:
            haystack = " ".join([
                str(card.get("refdes") or ""),
                str(card.get("hq_no") or ""),
                str(card.get("spec") or ""),
                str(card.get("pi") or ""),
                str(card.get("selection_order") or ""),
                " ".join(str(row.get("net") or "") for row in card.get("pin_net_summary") or [] if isinstance(row, dict)),
            ]).lower()
            if query_lower not in haystack:
                continue
        result.append(card)
    return result


def summarize_dfmea_readiness(cards: Iterable[dict]) -> dict:
    cards = list(cards or [])
    category_counts: Dict[str, int] = {}
    missing_counts: Dict[str, int] = {}
    ready = []
    needs_context = []
    for card in cards:
        category = str(card.get("category") or "unknown")
        category_counts[category] = category_counts.get(category, 0) + 1
        missing_fields = list(card.get("missing_fields") or [])
        for field_name in missing_fields:
            missing_counts[field_name] = missing_counts.get(field_name, 0) + 1
        if not missing_fields and category in {"chip", "power_ic", "large_ic", "connector"}:
            ready.append(card)
        elif category in {"chip", "power_ic", "large_ic", "connector"}:
            needs_context.append(card)
    return {
        "total_components": len(cards),
        "category_counts": category_counts,
        "missing_counts": missing_counts,
        "ready_count": len(ready),
        "needs_context_count": len(needs_context),
        "ready_cards": ready[:12],
        "needs_context_cards": needs_context[:12],
    }
