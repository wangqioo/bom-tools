# -*- coding: utf-8 -*-
"""PSTX Packager-XL text parsers.

This module owns the low-level, side-effect-free parsing for ``pstxprt.dat``
and ``pstxnet.dat``. Higher layers may enrich components with page resolution,
BOM grouping, rules, or Web view models, but raw text parsing stays here.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

from pstx_core import pages as page_logic


def join_continuations(text: str) -> str:
    normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    result, buf = [], ""
    for line in lines:
        stripped = line.rstrip()
        if stripped.endswith("~"):
            buf += stripped[:-1]
        else:
            buf += line
            result.append(buf)
            buf = ""
    if buf:
        result.append(buf)
    return "\n".join(result)


def extract_attrs(text: str) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    for match in re.finditer(r"\b([A-Z][A-Z0-9_]*)\s*=\s*'([^']*)'", str(text or "")):
        key, val = match.group(1), match.group(2)
        if key not in attrs:
            attrs[key] = val
    return attrs


HQ_CODE_TEXT_RE = re.compile(r"\b(HQ[A-Z0-9]{6,})\b", re.IGNORECASE)
SECTION_NUMBER_RE = re.compile(r"(?im)^\s*SECTION_NUMBER\s+(\d+)\b")
XY_POINT_RE = re.compile(r"\((-?\d+(?:\.\d+)?)\s*,?\s*(-?\d+(?:\.\d+)?)\)")
NUMBER_VALUE_RE = re.compile(r"-?\d+(?:\.\d+)?")


def first_non_empty_attr(attr_sets: Sequence[Dict[str, str]], key: str) -> str:
    for attrs in attr_sets:
        if not isinstance(attrs, dict):
            continue
        value = str(attrs.get(key, "") or "").strip()
        if value:
            return value
    return ""


def extract_hq_code_from_text(*values: object) -> str:
    for value in values:
        text = str(value or "")
        if not text:
            continue
        match = HQ_CODE_TEXT_RE.search(text)
        if match:
            return match.group(1).upper()
    return ""


def infer_value_from_part_name(part_name: str) -> str:
    parts = [part.strip() for part in str(part_name or "").split(",") if part.strip()]
    if len(parts) == 2 and extract_hq_code_from_text(parts[0]):
        return parts[1]
    return ""


def split_section_blocks(block: str) -> List[Dict[str, object]]:
    matches = list(SECTION_NUMBER_RE.finditer(str(block or "")))
    sections: List[Dict[str, object]] = []
    for idx, match in enumerate(matches):
        start = match.start()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(block)
        sections.append({
            "section_number": match.group(1),
            "text": block[start:end],
            "start": start,
            "end": end,
        })
    return sections


def parse_xy_point(value: object) -> Optional[Tuple[float, float]]:
    text = str(value or "").strip()
    if not text:
        return None
    match = XY_POINT_RE.search(text)
    if match:
        return float(match.group(1)), float(match.group(2))
    nums = NUMBER_VALUE_RE.findall(text)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    return None


def get_comp_type(refdes: str, part_name: str) -> str:
    part_name_lower = str(part_name or "").lower()
    type_rules = [
        (["cap_pol"], "CAP_POL"),
        (["cap_hdl", "cap_"], "CAP"),
        (["res_hdl", "res_"], "RES"),
        (["ind_hdl", "ind_", "ferrite", "fer_hdl", "fb_hdl"], "IND"),
        (["osc_", "crystal", "xtal"], "XTAL"),
        (["conn_", "connector"], "CONN"),
        (["led_"], "LED"),
        (["diode", "_d_hdl"], "DIODE"),
        (["mosfet", "mos_", "nmos", "pmos", "nfet", "pfet"], "FET"),
        (["bjt", "transistor", "npn", "pnp"], "BJT"),
        (["fuse"], "FUSE"),
        (["sw_hdl", "switch"], "SWITCH"),
        (["testpoint", "test_point", "tp_hdl"], "TESTPOINT"),
        (["transformer", "xfmr"], "TRANSFORMER"),
    ]
    for keywords, comp_type in type_rules:
        if any(keyword in part_name_lower for keyword in keywords):
            return comp_type
    prefix = (re.match(r"[A-Za-z]+", str(refdes or "")) or re.match(r"", "")).group(0).upper()
    prefix_map = {
        "C": "CAP",
        "PC": "CAP",
        "R": "RES",
        "PR": "RES",
        "L": "IND",
        "PL": "IND",
        "FB": "IND",
        "PFB": "IND",
        "U": "IC",
        "J": "CONN",
        "P": "CONN",
        "CN": "CONN",
        "Q": "FET",
        "D": "DIODE",
        "LED": "LED",
        "Y": "XTAL",
        "F": "FUSE",
        "SW": "SWITCH",
        "TP": "TESTPOINT",
        "T": "TRANSFORMER",
    }
    return prefix_map.get(prefix, "IC")


def split_named_blocks(text: str, marker: str) -> List[str]:
    return re.split(rf"(?:^|\n){re.escape(marker)}\n", str(text or ""))[1:]


def parse_pstxnet(content: str) -> Dict[str, List[dict]]:
    text = join_continuations(content)
    nets: Dict[str, List[dict]] = {}
    node_re = re.compile(r"NODE_NAME\s+(\S+)\s+(\S+)")
    pin_name_re = re.compile(r"'([^']+)'\s*:")
    for block in split_named_blocks(text, "NET_NAME"):
        match = re.search(r"'([^']+)'", block)
        if not match:
            continue
        net_name = match.group(1)
        nodes = []
        matches = list(node_re.finditer(block))
        for idx, node_match in enumerate(matches):
            next_start = matches[idx + 1].start() if idx + 1 < len(matches) else len(block)
            after = block[node_match.end():next_start]
            pin_match = pin_name_re.search(after)
            nodes.append({
                "refdes": node_match.group(1),
                "pin": node_match.group(2),
                "pin_name": pin_match.group(1) if pin_match else node_match.group(2),
            })
        if nodes:
            nets[net_name] = nodes
    return nets


def _build_component_section_record(
    section_text: str,
    section_number: str,
    common_attrs: Dict[str, str],
    part_name: str = "",
) -> Dict[str, object]:
    attrs = dict(common_attrs)
    attrs.update(extract_attrs(section_text))
    page_sources = page_logic.select_component_page_sources(section_text, attrs)
    logical_path_raw = page_sources.get("logical_path_raw", "")
    logical_path_source = page_sources.get("logical_path_source", "none")
    real_path_raw = page_sources.get("real_path_raw", "")
    real_path_source = page_sources.get("real_path_source", "none")
    xy_raw = attrs.get("XY", "")
    xy_point = parse_xy_point(xy_raw)
    hq_code = attrs.get("HQ_CODE", "") or extract_hq_code_from_text(attrs.get("CDS_PART_NAME", ""), part_name, section_text)
    value = attrs.get("VALUE", "") or infer_value_from_part_name(part_name)
    return {
        "section_number": str(section_number or ""),
        "attrs": attrs,
        "part_name": part_name,
        "hq_code": hq_code,
        "value": value,
        "drawing": attrs.get("DRAWING", ""),
        "path": attrs.get("PATH", ""),
        "split_inst": attrs.get("SPLIT_INST", ""),
        "location": attrs.get("LOCATION", ""),
        "xy": xy_raw,
        "xy_x": xy_point[0] if xy_point else "",
        "xy_y": xy_point[1] if xy_point else "",
        "page_path_raw": logical_path_raw,
        "page_path_source": logical_path_source,
        "page_path_logical_raw": logical_path_raw,
        "page_path_logical_source": logical_path_source,
        "page_path_real_raw": real_path_raw,
        "page_path_real_source": real_path_source,
        "page": "",
        "page_logical": page_logic.extract_top_level_page(logical_path_raw or attrs.get("DRAWING", "")),
        "page_real": "",
        "page_submodule_real": page_logic.extract_submodule_page(real_path_raw),
        "page_submodule_mapped": "",
    }


def parse_pstxprt(content: str) -> Dict[str, dict]:
    text = join_continuations(content)
    components: Dict[str, dict] = {}
    for block in split_named_blocks(text, "PART_NAME"):
        match = re.match(r"(\S+)\s+'([^']+)'", block.split("\n")[0].strip())
        if not match:
            continue
        refdes, part_name = match.group(1), match.group(2)
        section_blocks = split_section_blocks(block)
        common_attrs = (
            extract_attrs(block[:int(section_blocks[0]["start"])])
            if section_blocks
            else extract_attrs(block)
        )
        sections = [
            _build_component_section_record(
                str(section.get("text", "")),
                str(section.get("section_number", "")),
                common_attrs,
                part_name,
            )
            for section in section_blocks
        ]
        primary_section = next(
            (section for section in section_blocks if str(section.get("section_number", "")) == "1"),
            section_blocks[0] if section_blocks else None,
        )
        if primary_section:
            attrs = dict(common_attrs)
            attrs.update(extract_attrs(str(primary_section.get("text", ""))))
            page_source_text = str(primary_section.get("text", ""))
        else:
            attrs = common_attrs
            page_source_text = block
        page_sources = page_logic.select_component_page_sources(page_source_text, attrs)
        logical_path_raw = page_sources.get("logical_path_raw", "")
        logical_path_source = page_sources.get("logical_path_source", "none")
        real_path_raw = page_sources.get("real_path_raw", "")
        real_path_source = page_sources.get("real_path_source", "none")
        xy_raw = attrs.get("XY", "")
        xy_point = parse_xy_point(xy_raw)
        section_attr_sets = [section.get("attrs", {}) for section in sections if isinstance(section, dict)]
        attr_sets = [attrs] + [section_attrs for section_attrs in section_attr_sets if isinstance(section_attrs, dict)]
        hq_code = (
            first_non_empty_attr(attr_sets, "HQ_CODE")
            or extract_hq_code_from_text(
                first_non_empty_attr(attr_sets, "CDS_PART_NAME"),
                part_name,
                block,
            )
        )
        value = first_non_empty_attr(attr_sets, "VALUE") or infer_value_from_part_name(part_name)
        components[refdes] = {
            "refdes": refdes,
            "part_name": part_name,
            "hq_code": hq_code,
            "value": value,
            "package": first_non_empty_attr(attr_sets, "PACKAGE"),
            "material": first_non_empty_attr(attr_sets, "MATERIAL"),
            "tolerance": first_non_empty_attr(attr_sets, "TOLERANCE"),
            "voltage": first_non_empty_attr(attr_sets, "VOLTAGE"),
            "current": first_non_empty_attr(attr_sets, "CURRENT"),
            "power": first_non_empty_attr(attr_sets, "POWER"),
            "bom_option": first_non_empty_attr(attr_sets, "BOM_OPTION"),
            "bom_cost": first_non_empty_attr(attr_sets, "BOM_COST"),
            "room": first_non_empty_attr(attr_sets, "ROOM"),
            "drawing": attrs.get("DRAWING", ""),
            "xy": xy_raw,
            "xy_x": xy_point[0] if xy_point else "",
            "xy_y": xy_point[1] if xy_point else "",
            "page_path_raw": logical_path_raw,
            "page_path_source": logical_path_source,
            "page_path_logical_raw": logical_path_raw,
            "page_path_logical_source": logical_path_source,
            "page_path_real_raw": real_path_raw,
            "page_path_real_source": real_path_source,
            "page": "",
            "page_logical": page_logic.extract_top_level_page(logical_path_raw or attrs.get("DRAWING", "")),
            "page_real": "",
            "page_submodule_real": page_logic.extract_submodule_page(real_path_raw),
            "page_submodule_mapped": "",
            "split_inst": attrs.get("SPLIT_INST", ""),
            "location": attrs.get("LOCATION", ""),
            "section_count": len(sections) if sections else 0,
            "sections": sections,
            "comp_type": get_comp_type(refdes, part_name),
        }
    return components


def parse_all(prt_content: str, net_content: str):
    components = parse_pstxprt(prt_content)
    nets = parse_pstxnet(net_content)
    comp_nets: Dict[str, Dict[str, str]] = {}
    for net_name, nodes in nets.items():
        for node in nodes:
            refdes = node["refdes"]
            if refdes not in comp_nets:
                comp_nets[refdes] = {}
            comp_nets[refdes][node["pin"]] = net_name
    for refdes, comp in components.items():
        comp["nets"] = comp_nets.get(refdes, {})
    return components, nets, comp_nets
