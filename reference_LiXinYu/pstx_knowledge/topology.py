# -*- coding: utf-8 -*-
"""Chip-level fuzzy topology extraction for harness agents.

The goal of this module is intentionally higher-level than a pin-accurate
schematic graph: it keeps only IC-like nodes and summarizes their shared
signal nets so an agent can reason about relations such as "main FPGA connects
to a level shifter" without dragging every passive into context.
"""

from __future__ import annotations

from collections import Counter, defaultdict
import hashlib
from itertools import combinations
import json
import re
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from pstx_core.analysis_cache import (
    ANALYSIS_CACHE_SCHEMA_VERSION,
    ANALYSIS_CACHE_VERSION,
    DISABLE_ANALYSIS_CACHE_ENV,
    analysis_cache_dir,
    analysis_cache_enabled,
)

from pstx_knowledge.business_dictionary import (
    business_dictionary_summary,
    interface_alias_snapshot,
    interface_aliases,
    load_business_dictionary,
    review_focus_for_interface,
)
from pstx_knowledge.component_identity import USER_VISIBLE_REAL_PAGE_LABEL, classify_refdes


CHIP_CATEGORIES = {"chip", "large_ic", "power_ic"}
CONNECTOR_CATEGORIES = {"connector"}
DEFAULT_LIMIT = 30
MAX_LIMIT = 100
MAX_QUERY_ITEMS = 20
MAX_NET_NODES = 12
LLM_TOPOLOGY_SCHEMA_VERSION = "llm-topology.v1"
LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION = "llm-topology-review-task.v1"
TOPOLOGY_CACHE_KIND = "llm_topology"
TOPOLOGY_CACHE_VERSION = "2026-05-06.1"
DEFAULT_SUPPLY_LIMIT = 12
MAX_SUPPLY_LIMIT = 250
PROJECT_DOMAIN_CONTEXT = "server_hardware"

INTERFACE_EXPECTATIONS: Dict[str, dict] = {
    "i2c": {
        "required_signals": {
            "scl": ["SCL", "I2C_SCL"],
            "sda": ["SDA", "I2C_SDA"],
        },
        "recommended_checks": ["上拉电压", "重复上拉", "总线电容", "跨电平转换器方向/OE"],
    },
    "spi": {
        "required_signals": {
            "clock": ["SCLK", "SPI_CLK", "CLK"],
            "data_out": ["MOSI", "SDO", "TXD"],
            "data_in": ["MISO", "SDI", "RXD"],
            "chip_select": ["CS", "CSN", "NCS", "SS"],
        },
        "recommended_checks": ["串阻", "片选默认态", "电压域", "时钟边沿/模式"],
    },
    "uart": {
        "required_signals": {
            "tx": ["TX", "TXD", "UART_TX"],
            "rx": ["RX", "RXD", "UART_RX"],
        },
        "recommended_checks": ["TX/RX 交叉方向", "电压域", "上下拉", "调试访问边界"],
    },
    "pcie": {
        "required_signals": {
            "tx_pair": ["TXP", "TXN", "TX_P", "TX_N", "PETP", "PETN"],
            "rx_pair": ["RXP", "RXN", "RX_P", "RX_N", "PERP", "PERN"],
            "refclk": ["REFCLK", "CLKREQ", "P5E_REFCLK", "PCE_REFCLK", "PCIE_REFCLK"],
            "reset": ["PERST", "RST", "RESET"],
        },
        "recommended_checks": ["差分阻抗", "AC 耦合", "REFCLK/PERST/CLKREQ", "端接/ESD", "lane 方向"],
    },
    "usb": {
        "required_signals": {
            "data_or_superspeed": ["DP", "DM", "DPLUS", "DMINUS", "TXP", "TXN", "RXP", "RXN"],
        },
        "recommended_checks": ["ESD", "VBUS/CC", "差分阻抗", "共模电感/端接", "连接器方向"],
    },
    "mipi_lvds": {
        "required_signals": {
            "clock_pair": ["CLKP", "CLKN", "CLK_P", "CLK_N"],
            "data_pair": ["DP", "DN", "D_P", "D_N", "LANE"],
        },
        "recommended_checks": ["差分阻抗", "lane 极性", "ESD", "连接器/屏端 pinout", "时钟/数据 lane 数"],
    },
    "ddr": {
        "required_signals": {
            "address_or_command": ["ADDR", "A", "BA", "RAS", "CAS", "WE", "CS"],
            "data": ["DQ", "DQS", "DM"],
            "clock": ["CK", "CLK"],
        },
        "recommended_checks": ["拓扑/端接", "VTT/VREF", "长度匹配", "ODT", "复位/CKE"],
    },
    "ethernet": {
        "required_signals": {
            "data": ["RGMII", "SGMII", "TXD", "RXD", "MDI"],
            "management": ["MDIO", "MDC"],
            "clock": ["REFCLK", "CLK"],
        },
        "recommended_checks": ["PHY strap", "RGMII delay", "MDIO/MDC 上拉", "时钟/复位", "磁性器件/ESD"],
    },
    "storage_sdio": {
        "required_signals": {
            "clock": ["CLK", "SDCLK"],
            "command": ["CMD", "SDCMD"],
            "data": ["DAT", "DATA", "D0", "D1", "D2", "D3"],
        },
        "recommended_checks": ["上拉", "时钟串阻", "位宽", "电压切换", "复位"],
    },
    "clock": {
        "required_signals": {
            "clock": ["CLK", "CLOCK", "OSC", "XO", "REFCLK"],
        },
        "recommended_checks": ["源端串阻", "端接", "扇出", "抖动/电平", "使能"],
    },
    "reset": {
        "required_signals": {
            "reset": ["RST", "RESET", "PERST", "POR"],
        },
        "recommended_checks": ["默认态", "释放时序", "上拉/下拉", "跨域复位"],
    },
    "power_control": {
        "required_signals": {
            "enable_or_good": ["EN", "ENABLE", "PGOOD", "POWERGOOD", "PWREN"],
        },
        "recommended_checks": ["默认态", "上下拉", "上电时序", "电压域"],
    },
    "jtag_debug": {
        "required_signals": {
            "clock": ["TCK"],
            "mode": ["TMS"],
            "data": ["TDI", "TDO"],
        },
        "recommended_checks": ["上下拉", "测试点", "量产可访问性", "复用状态"],
    },
    "analog_sense": {
        "required_signals": {
            "sense_or_feedback": ["SENSE", "FB", "ADC", "AIN"],
        },
        "recommended_checks": ["量程", "滤波", "地参考", "Kelvin 路径", "保护"],
    },
    "audio": {
        "required_signals": {
            "clock": ["MCLK", "BCLK", "LRCLK", "WS"],
            "data": ["SDIN", "SDOUT", "DIN", "DOUT"],
        },
        "recommended_checks": ["主从时钟", "MCLK/BCLK/LRCLK 方向", "电压域", "模拟地"],
    },
}

POWER_OR_GROUND_NET_RE = re.compile(
    r"(?i)(^|[_\-/])(?:P\d|P[0-9A-Z]*V|VDD|VCC|VIN|VOUT|VSYS|VBAT|VCORE|VIO|VREF|VTT|VPLL|VDDA|VCCA|VCCB|AVDD|DVDD|PVDD|GND|GNDA|AGND|DGND|PGND|VSS|AVSS|DVSS|VSSA|0V|0)(?:$|[_\-/])"
)
LEVEL_SHIFTER_RE = re.compile(
    r"(?i)(level.?shift|translator|transceiver|txs\d|txb\d|sn74|lsf\d|74avc|74lvc|74axp|电平转换)"
)
POWER_IC_RE = re.compile(r"(?i)(pmic|buck|boost|ldo|regulator|power|charger|switcher|dc.?dc)")
PROCESSOR_RE = re.compile(r"(?i)(fpga|cpu|gpu|soc|mcu|dsp|processor|xilinx|alter(a|ra)|lcmxo|kintex|zynq)")
MEMORY_RE = re.compile(r"(?i)(ddr|lpddr|emmc|nand|nor|flash|sdram|memory)")
CLOCK_RE = re.compile(r"(?i)(clock|clk|osc|xo|晶振)")
VOLTAGE_DOMAIN_RE = re.compile(r"(?i)(?:^|[^0-9A-Z])P?(\d{1,2})V(\d{0,3})(?:[^0-9A-Z]|$)")


def _safe_text(value, limit: int = 240) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _safe_fragment(value: object) -> str:
    text = re.sub(r"[^0-9A-Za-z_.-]+", "-", str(value or "").strip())
    return text.strip("-").lower() or "unknown"


def _net_tokens(value: object) -> List[str]:
    return [token for token in re.split(r"[^0-9A-Z]+", str(value or "").upper()) if token]


def _alias_matches(tokens: set[str], text: str, interface_id: str) -> bool:
    for alias in interface_aliases(interface_id):
        alias_text = str(alias or "").upper()
        if not alias_text:
            continue
        if alias_text in tokens or alias_text in text:
            return True
        if re.fullmatch(r"P\d+E", alias_text) and any(re.fullmatch(r"P\d+E", token) for token in tokens):
            return True
    return False


def _as_bool(value: object, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "y", "on"}:
            return True
        if lowered in {"0", "false", "no", "n", "off"}:
            return False
    return default


def _as_limit(value: object, default: int = DEFAULT_LIMIT) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default
    return max(1, min(number, MAX_LIMIT))


def _as_supply_limit(value: object, default: int = DEFAULT_SUPPLY_LIMIT) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default
    return max(0, min(number, MAX_SUPPLY_LIMIT))


def _normalize_topology_view(view: object, return_all_edges: bool) -> str:
    text = str(view or "").strip().lower()
    if return_all_edges:
        return "full"
    if text in {"full", "details", "detail", "all"}:
        return "full"
    return "summary"


def _normalize_supply_mode(supply_mode: object, view: str, return_all_edges: bool) -> str:
    text = str(supply_mode or "").strip().lower()
    if text in {"details", "detail", "full", "edges"}:
        return "details"
    if text in {"hidden", "hide", "none", "off", "0"}:
        return "hidden"
    if text in {"grouped", "group", "summary", "groups"}:
        return "grouped"
    return "details" if return_all_edges or view == "full" else "grouped"


def _stable_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _stable_hash(value: Any) -> str:
    return hashlib.sha256(_stable_json(value).encode("utf-8")).hexdigest()


def _component_text(refdes: str, comp: Mapping[str, object], pin_nets: Sequence[dict] = ()) -> str:
    keys = (
        "CDS_PART_NAME",
        "cds_part_name",
        "PART_NAME",
        "part_type",
        "VALUE",
        "value",
        "description",
        "型号",
        "规格型号",
        "HQ_CODE",
        "hq_code",
    )
    pieces = [refdes]
    pieces.extend(_safe_text(comp.get(key), 260) for key in keys if comp.get(key) not in {None, ""})
    pieces.extend(_safe_text(row.get("net"), 180) for row in list(pin_nets)[:24])
    return " ".join(piece for piece in pieces if piece)


def _component_spec(comp: Mapping[str, object]) -> str:
    for key in ("CDS_PART_NAME", "cds_part_name", "part_type", "PART_NAME", "VALUE", "value", "description", "型号", "规格型号"):
        text = _safe_text(comp.get(key), 260)
        if text:
            return text
    return ""


def _component_hq_no(comp: Mapping[str, object]) -> str:
    for key in ("HQ_CODE", "hq_code", "料号", "part_number", "PART_NUMBER"):
        text = _safe_text(comp.get(key), 160)
        if text:
            return text
    return ""


def _component_page(comp: Mapping[str, object]) -> str:
    for key in ("user_visible_page", USER_VISIBLE_REAL_PAGE_LABEL, "page_submodule_mapped", "page_real", "真实页", "page"):
        text = _safe_text(comp.get(key), 80)
        if text:
            return text
    return ""


def _component_module(comp: Mapping[str, object]) -> str:
    for key in ("module_name", "module", "模块", "scope_module_name", "module_id", "module_path"):
        text = _safe_text(comp.get(key), 160)
        if text:
            return text
    p_path = _safe_text(comp.get("P_PATH") or comp.get("p_path") or "", 260)
    match = re.search(r"@[^@.]+\.([^@()]+)\(sch_1\)", p_path, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def _component_package(comp: Mapping[str, object]) -> str:
    for key in ("PACKAGE", "package", "封装", "PCB_FOOTPRINT", "footprint"):
        text = _safe_text(comp.get(key), 120)
        if text:
            return text
    return ""


def _component_value(comp: Mapping[str, object]) -> str:
    for key in ("VALUE", "value", "值"):
        text = _safe_text(comp.get(key), 160)
        if text:
            return text
    return ""


def _component_bom_option(comp: Mapping[str, object]) -> str:
    for key in ("BOM_OPTION", "bom_option", "装配状态", "assembly_state"):
        text = _safe_text(comp.get(key), 160)
        if text:
            return text
    return ""


def _component_part_name(comp: Mapping[str, object]) -> str:
    for key in ("CDS_PART_NAME", "cds_part_name", "PART_NAME", "part_name", "part_type", "VALUE", "value"):
        text = _safe_text(comp.get(key), 360)
        if text:
            return text
    return ""


def _part_name_tokens(text: object, *, limit: int = 28) -> List[str]:
    tokens = [
        token for token in re.split(r"[^0-9A-Za-z]+", str(text or "").upper())
        if token and len(token) >= 2
    ]
    noise = {"HDL", "CHIPS", "HQ", "IC", "CAP", "RES", "IND", "SCH", "PART", "NAME"}
    return [token for token in dict.fromkeys(tokens) if token not in noise][:limit]


def _llm_device_identity_hint(refdes: str,
                              comp: Mapping[str, object],
                              *,
                              role: str,
                              category: str,
                              candidate: str,
                              confidence: str) -> dict:
    """Build non-authoritative device identity input for the Agent/LLM.

    This deliberately avoids a rigid regex "device library".  The local code
    only packages evidence and a server-hardware taxonomy; the model should
    still explain uncertainty and cite evidence before naming a concrete type.
    """

    part_name = _component_part_name(comp)
    spec = _component_spec(comp)
    tokens = _part_name_tokens(" ".join([part_name, spec, _component_value(comp)]))
    return {
        "schema_version": "llm-device-identity-hint.v1",
        "domain_context": PROJECT_DOMAIN_CONTEXT,
        "refdes": _safe_text(refdes, 80),
        "part_name": part_name,
        "spec": spec,
        "value": _component_value(comp),
        "package": _component_package(comp),
        "hq_no": _component_hq_no(comp),
        "page": _component_page(comp),
        "local_role_hint": role,
        "local_refdes_category": category,
        "local_candidate_chip_type": candidate,
        "local_confidence": confidence,
        "tokens": tokens,
        "server_device_taxonomy": [
            "processor_or_fpga",
            "bmc_or_management_controller",
            "pcie_switch_or_retimer",
            "clock_generator_or_buffer",
            "power_management_ic",
            "level_shifter_or_translator",
            "memory_or_storage",
            "ethernet_phy_or_controller",
            "usb_typec_mux_or_redriver",
            "sensor_or_monitor",
            "connector_or_external_interface",
            "unknown_needs_datasheet",
        ],
        "llm_instruction": (
            "请把这些字段作为候选证据来判断器件大类；不要仅凭位号或单个 token 下定论。"
            "如果 part_name/spec/HQ 料号不足，请标记 unknown_needs_datasheet，并建议读取 datasheet 或飞书缓存。"
        ),
    }


def _is_chip_ref(refdes: str, comp: Mapping[str, object], *, include_connectors: bool = False) -> Tuple[bool, str, str, str]:
    category, candidate, confidence = classify_refdes(refdes, dict(comp))
    allowed = CHIP_CATEGORIES | (CONNECTOR_CATEGORIES if include_connectors else set())
    return category in allowed, category, candidate, confidence


def _is_power_or_ground_net(net_name: object) -> bool:
    text = str(net_name or "").strip().upper()
    compact = re.sub(r"[^0-9A-Z]+", "", text)
    if not compact:
        return False
    if compact in {"0", "0V"}:
        return True
    ground_markers = ("GND", "GNDA", "AGND", "DGND", "PGND", "VSS", "AVSS", "DVSS", "VSSA", "GNDS")
    if compact in set(ground_markers) or compact.endswith(ground_markers) or compact.startswith(ground_markers):
        return True
    power_markers = (
        "VDD", "VCC", "VIN", "VOUT", "VSYS", "VBAT", "VCORE", "VIO", "VREF", "VTT", "VPLL",
        "VDDA", "VCCA", "VCCB", "AVDD", "DVDD", "PVDD", "VDDIO", "USBVBUS", "VBUS",
    )
    if compact.startswith(power_markers) or compact.endswith(power_markers):
        return True
    if compact.startswith("P") and "V" in compact[1:]:
        return True
    return bool(POWER_OR_GROUND_NET_RE.search(text))


def _is_ground_net(net_name: object) -> bool:
    compact = re.sub(r"[^0-9A-Z]+", "", str(net_name or "").strip().upper())
    if not compact:
        return False
    ground_markers = ("GND", "GNDA", "AGND", "DGND", "PGND", "VSS", "AVSS", "DVSS", "VSSA", "GNDS")
    return compact in {"0", "0V"} or compact in set(ground_markers) or compact.endswith(ground_markers) or compact.startswith(ground_markers)


def _interface_group(net_name: object, pin_names: Iterable[object] = ()) -> str:
    text = " ".join([str(net_name or ""), *(str(item or "") for item in pin_names)]).upper()
    tokens = set(_net_tokens(text))
    if "DDR" in tokens or "LPDDR" in text or "DQS" in tokens or any(re.fullmatch(r"DQ\d*", token) for token in tokens):
        return "ddr"
    if _alias_matches(tokens, text, "pcie"):
        return "pcie"
    if any(token in tokens for token in {"JTAG", "TCK", "TMS", "TDI", "TDO", "TRST"}):
        return "jtag_debug"
    if any(token in tokens for token in {"RGMII", "SGMII", "GMII", "MII", "MDIO", "MDC"}) or "ETH" in tokens or "ETHERNET" in text:
        return "ethernet"
    if any(token in tokens for token in {"SDIO", "SDMMC", "EMMC", "MMC", "SDCLK", "SDCMD"}):
        return "storage_sdio"
    if "SPI" in tokens or "MOSI" in tokens or "MISO" in tokens or "SCLK" in tokens or "SPI" in text or any(token in tokens for token in {"CS", "CSN", "NCS", "SS"}):
        return "spi"
    if "I2C" in tokens or "I2C" in text or "SCL" in tokens or "SDA" in tokens:
        return "i2c"
    if "UART" in tokens or "UART" in text or any(token in tokens for token in {"TX", "RX", "TXD", "RXD", "CTS", "RTS"}):
        return "uart"
    if "USB" in tokens or "USB" in text or any(token in tokens for token in {"USBP", "USBN", "UDP", "UDM", "DPLUS", "DMINUS"}):
        return "usb"
    if "MIPI" in tokens or "LVDS" in tokens or "CSI" in tokens or "DSI" in tokens or "MIPI" in text or "LVDS" in text:
        return "mipi_lvds"
    if "HDMI" in tokens or "DISPLAYPORT" in text or "DP_AUX" in text or "EDP" in tokens:
        return "high_speed"
    if any(token in tokens for token in {"I2S", "IIS", "MCLK", "BCLK", "LRCLK", "WS", "TDM", "PDM", "SPDIF"}):
        return "audio"
    if any(token in tokens for token in {"ADC", "DAC", "AIN", "AOUT", "SENSE", "FB"}):
        return "analog_sense"
    if "CLK" in tokens or "CLOCK" in tokens or "OSC" in tokens or "XO" in tokens or any(token.endswith("CLK") for token in tokens):
        return "clock"
    if "RESET" in tokens or "RST" in tokens or "RESETN" in tokens or "RSTN" in tokens or any(token.startswith("RST") for token in tokens):
        return "reset"
    if any(token in tokens for token in {"INT", "IRQ", "ALERT", "FAULT", "NMI"}):
        return "interrupt"
    if any(token in tokens for token in {"EN", "ENABLE", "PWR", "PWRON", "PGOOD", "POWERGOOD", "PWREN"}):
        return "power_control"
    if "GPIO" in tokens or "GPIO" in text:
        return "gpio"
    return "misc_signal"


def _is_clock_net(net_name: object, pin_names: Iterable[object] = ()) -> bool:
    return _interface_group(net_name, pin_names) == "clock"


def _is_reset_net(net_name: object, pin_names: Iterable[object] = ()) -> bool:
    return _interface_group(net_name, pin_names) == "reset"


def _evidence_id(kind: str, value: object) -> str:
    return f"llm-topology-{kind}-{_safe_fragment(value)}"


def _node_detail_tool(refdes: str) -> dict:
    return {"name": "get_llm_topology_node", "args": {"refdes": refdes}}


def _edge_detail_tool(edge_id: str) -> dict:
    return {"name": "get_llm_topology_edge", "args": {"edge_id": edge_id}}


def _critical_net_preview(pin_nets: Sequence[dict]) -> dict:
    power_nets: List[str] = []
    clock_nets: List[str] = []
    reset_nets: List[str] = []
    control_nets: List[str] = []
    for row in pin_nets:
        net_name = _safe_text(row.get("net", ""), 180)
        if not net_name:
            continue
        pin_names = [row.get("pin_name"), row.get("pin")]
        group = _interface_group(net_name, pin_names)
        if _is_power_or_ground_net(net_name):
            power_nets.append(net_name)
        elif group == "clock":
            clock_nets.append(net_name)
        elif group == "reset":
            reset_nets.append(net_name)
        elif group in {"power_control", "interrupt", "jtag_debug"}:
            control_nets.append(net_name)
    return {
        "power_nets": list(dict.fromkeys(power_nets))[:12],
        "clock_nets": list(dict.fromkeys(clock_nets))[:8],
        "reset_nets": list(dict.fromkeys(reset_nets))[:8],
        "control_nets": list(dict.fromkeys(control_nets))[:8],
    }


def _node_review_tags(node: Mapping[str, object]) -> List[str]:
    tags: List[str] = []
    role = str(node.get("role") or "")
    groups = set(node.get("interface_groups") or [])
    if role == "level_shifter":
        tags.append("level_shift_voltage_domain")
    if role == "power_management_ic":
        tags.append("power_sequence_and_enable")
    if "clock" in groups or role == "clock_source":
        tags.append("clock_fanout_and_termination")
    if "reset" in groups:
        tags.append("reset_pull_and_timing")
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}:
        tags.append("high_speed_interface_review")
    if groups & {"jtag_debug"}:
        tags.append("debug_port_review")
    if groups & {"analog_sense"}:
        tags.append("analog_feedback_or_sense_review")
    if not tags:
        tags.append("general_pin_net_review")
    return tags[:8]


def _edge_review_hints(edge: Mapping[str, object]) -> List[str]:
    groups = set(edge.get("interface_groups") or [])
    roles = {str(edge.get("source_role") or ""), str(edge.get("target_role") or "")}
    hints: List[str] = []
    if "level_shifter" in roles:
        hints.append("核对两侧供电电压域、OE/EN 控制、上电顺序和方向/自动方向约束。")
    if "i2c" in groups:
        hints.append("核对 I2C 两侧上拉电压、重复上拉、总线电容和跨电平转换器连接。")
    if "spi" in groups:
        hints.append("核对 SPI 串阻、片选默认态、跨页连接和电压域一致性。")
    if "jtag_debug" in groups:
        hints.append("核对 JTAG/debug 口上下拉、复用状态、量产可访问性和连接器/测试点路径。")
    if "ethernet" in groups:
        hints.append("核对 Ethernet/RGMII/SGMII/MDIO 的电压域、时钟、复位、PHY strap 和阻抗约束。")
    if "storage_sdio" in groups:
        hints.append("核对 SDIO/eMMC 的上拉、时钟串阻、位宽、复位和电压切换约束。")
    if "clock" in groups:
        hints.append("核对时钟源扇出、串阻/端接、跨页网名一致性和时钟使能。")
    if "reset" in groups:
        hints.append("核对 reset 上拉/下拉、释放时序、跨域复位和默认电平。")
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr"}:
        hints.append("核对高速接口差分/阻抗/AC 耦合/端接/长度匹配等约束。")
    if "audio" in groups:
        hints.append("核对音频/I2S/TDM 主从时钟、MCLK/BCLK/LRCLK 方向和跨域电平。")
    if "analog_sense" in groups:
        hints.append("核对模拟采样/反馈网络的量程、滤波、Kelvin/地参考和 RC 稳定性。")
    if "power_control" in groups:
        hints.append("核对 enable/PGOOD/power-control 的默认态、时序和上下拉。")
    if edge.get("passive_bridges"):
        hints.append("该连接含一跳无源桥摘要，需结合串阻/耦合/端接规则复核无源件取值和位置。")
    if not hints:
        hints.append("核对共享网络两端 pin 名、页码和是否为真实业务连接，避免公共信号误判。")
    return hints[:8]


def _edge_review_focus(edge: Mapping[str, object]) -> List[str]:
    groups = set(edge.get("interface_groups") or [])
    roles = {str(edge.get("source_role") or ""), str(edge.get("target_role") or "")}
    focus: List[str] = []
    if "level_shifter" in roles or edge.get("voltage_domain_transition"):
        focus.extend(["两侧电压域", "OE/EN", "上电顺序"])
    if "i2c" in groups:
        focus.extend(["上拉电压", "重复上拉", "总线电容"])
    if "spi" in groups:
        focus.extend(["串阻", "片选默认态", "时钟边沿"])
    if "clock" in groups:
        focus.extend(["时钟串阻", "端接", "扇出"])
    if "reset" in groups:
        focus.extend(["默认态", "释放时序", "上拉/下拉"])
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}:
        focus.extend(["阻抗", "差分/长度", "端接/AC耦合"])
    if "jtag_debug" in groups:
        focus.extend(["debug 可访问性", "上下拉", "测试点"])
    if "analog_sense" in groups:
        focus.extend(["量程", "滤波", "地参考"])
    if edge.get("passive_bridge_count"):
        focus.append("一跳无源桥")
    for group in sorted(groups):
        focus.extend(review_focus_for_interface(group))
    return list(dict.fromkeys(focus))[:12]


def _signal_presence_text(edge: Mapping[str, object], group: str) -> str:
    pieces: List[str] = []
    for shared in edge.get("shared_nets") or []:
        if not isinstance(shared, Mapping):
            continue
        if str(shared.get("interface_group") or "misc_signal") != group:
            continue
        pieces.append(str(shared.get("net") or ""))
        for side_key in ("source_pins", "target_pins"):
            for pin in shared.get(side_key) or []:
                if isinstance(pin, Mapping):
                    pieces.append(str(pin.get("pin") or ""))
                    pieces.append(str(pin.get("pin_name") or ""))
    for bridge in edge.get("passive_bridges") or []:
        if not isinstance(bridge, Mapping):
            continue
        bridge_group = _interface_group(" ".join(str(net) for net in bridge.get("nets") or []), [])
        if bridge_group != group:
            continue
        pieces.extend(str(net) for net in bridge.get("nets") or [])
        pieces.append(str(bridge.get("refdes") or ""))
        pieces.append(str(bridge.get("value") or ""))
    return " ".join(pieces).upper()


def _has_signal_alias(haystack: str, aliases: Sequence[str]) -> bool:
    tokens = set(_net_tokens(haystack))
    for alias in aliases:
        alias_text = str(alias or "").upper()
        if not alias_text:
            continue
        if alias_text in tokens or alias_text in haystack:
            return True
    return False


def _interface_completeness_for_group(edge: Mapping[str, object], group: str) -> dict:
    expectation = INTERFACE_EXPECTATIONS.get(group)
    if not expectation:
        return {
            "group": group,
            "status": "not_modeled",
            "observed_signals": [],
            "missing_signals": [],
            "recommended_checks": list(review_focus_for_interface(group))[:8],
            "summary": f"{group} 暂无内置完整性模板，按通用 pin/net evidence 复核。",
        }
    haystack = _signal_presence_text(edge, group)
    observed: List[str] = []
    missing: List[str] = []
    for signal_name, aliases in (expectation.get("required_signals") or {}).items():
        if _has_signal_alias(haystack, aliases):
            observed.append(signal_name)
        else:
            missing.append(signal_name)
    if not missing:
        status = "observed_required"
    elif observed:
        status = "partial_needs_detail"
    else:
        status = "needs_detail"
    return {
        "group": group,
        "status": status,
        "observed_signals": observed,
        "missing_signals": missing,
        "required_signals": dict(expectation.get("required_signals") or {}),
        "recommended_checks": list(expectation.get("recommended_checks") or [])[:10],
        "summary": (
            f"{group} 完整性：已观察 {', '.join(observed) or '无明确关键子信号'}；"
            f"需补查 {', '.join(missing) or '无'}。"
        ),
    }


def _edge_interface_completeness(edge: Mapping[str, object]) -> List[dict]:
    groups = [group for group in list(edge.get("interface_groups") or []) if group and group != "misc_signal"]
    return [_interface_completeness_for_group(edge, group) for group in groups][:8]


def _voltage_sort_key(domain: object) -> Tuple[float, str]:
    text = str(domain or "").upper().replace("V", ".")
    try:
        return (float(text.strip(".")), str(domain))
    except ValueError:
        return (999.0, str(domain))


def _voltage_domain_from_net(net_name: object) -> str:
    text = str(net_name or "").strip().upper()
    if not text or text in {"GND", "AGND", "DGND", "PGND", "VSS", "0V"}:
        return ""
    compact = re.sub(r"[^0-9A-Z.]+", "_", text)
    for match in VOLTAGE_DOMAIN_RE.finditer(f"_{compact}_"):
        integer, decimal = match.groups()
        if integer == "0" and not decimal:
            continue
        return f"{int(integer)}V{decimal}" if decimal else f"{int(integer)}V"
    decimal_match = re.search(r"(?<![0-9])(\d{1,2})\.(\d{1,3})V(?![0-9])", compact)
    if decimal_match:
        integer, decimal = decimal_match.groups()
        if integer != "0":
            return f"{int(integer)}V{decimal.rstrip('0') or '0'}"
    embedded_match = re.search(r"(?:VCC|VDD|VDDIO|VDDA|VCCA|VCCB|PWR|P)?(\d{1,2})V(\d{1,3})", compact)
    if embedded_match:
        integer, decimal = embedded_match.groups()
        return f"{int(integer)}V{decimal}"
    return ""


def _voltage_domains_from_nets(nets: Iterable[object]) -> List[str]:
    domains = {
        domain for domain in (_voltage_domain_from_net(net) for net in nets)
        if domain
    }
    return sorted(domains, key=_voltage_sort_key)[:12]


def _supply_edge_id(source_ref: str, target_ref: str, net_name: object) -> str:
    return f"supply-edge-{_safe_fragment(source_ref)}-{_safe_fragment(target_ref)}-{_safe_fragment(net_name)}"


def _supply_record(source_ref: str, target_ref: str, net_name: object) -> dict:
    net_text = _safe_text(net_name, 180)
    return {
        "edge_id": _supply_edge_id(source_ref, target_ref, net_text),
        "source_refdes": source_ref,
        "target_refdes": target_ref,
        "supply_net": net_text,
        "voltage_domain": _voltage_domain_from_net(net_text),
    }


def _supply_edge_from_record(record: Mapping[str, object], nodes: Mapping[str, Mapping[str, object]]) -> dict:
    source_ref = str(record.get("source_refdes") or "")
    target_ref = str(record.get("target_refdes") or "")
    net_name = _safe_text(record.get("supply_net", ""), 180)
    edge_id = _safe_text(record.get("edge_id", "") or _supply_edge_id(source_ref, target_ref, net_name), 220)
    source_node = nodes.get(source_ref, {})
    target_node = nodes.get(target_ref, {})
    source_control = list(source_node.get("control_nets") or [])[:8]
    target_control = list(target_node.get("control_nets") or [])[:8]
    voltage_domain = _safe_text(record.get("voltage_domain", "") or _voltage_domain_from_net(net_name), 80)
    review_hints = [
        f"核对 {net_name} 输出电压{f'（{voltage_domain}）' if voltage_domain else ''}、负载电流、纹波和上电时序。",
        "检查负载芯片附近去耦、电源滤波、sense/feedback 路径和跨页网名一致性。",
    ]
    if source_control or target_control:
        review_hints.append("结合 EN/PGOOD/RESET 等控制网络核对电源时序闭环。")
    return {
        "edge_id": edge_id,
        "evidence_id": _evidence_id("supply-edge", edge_id),
        "edge_kind": "supply",
        "undirected": True,
        "source_refdes": source_ref,
        "target_refdes": target_ref,
        "source_role": source_node.get("role", ""),
        "target_role": target_node.get("role", ""),
        "source_page": source_node.get("user_visible_page", ""),
        "target_page": target_node.get("user_visible_page", ""),
        "source_voltage_domains": list(source_node.get("voltage_domains") or [])[:8],
        "target_voltage_domains": list(target_node.get("voltage_domains") or [])[:8],
        "source_control_nets": source_control,
        "target_control_nets": target_control,
        "supply_net": net_name,
        "voltage_domain": voltage_domain,
        "relation_label": "电源管理供电关系",
        "review_priority": "medium",
        "review_score": 35,
        "risk_tags": ["power_rail_dependency"],
        "review_focus": ["输出电压", "负载电流", "上电时序", "去耦滤波", "EN/PGOOD/RESET"],
        "review_hints": review_hints,
        "summary": f"{source_ref} 通过 {net_name} 给 {target_ref} 提供供电关系。",
        "detail_tool": _edge_detail_tool(edge_id),
    }


def _supply_edges_from_records(records: Sequence[Mapping[str, object]],
                               nodes: Mapping[str, Mapping[str, object]]) -> List[dict]:
    return [_supply_edge_from_record(record, nodes) for record in records]


def _supply_group_sort_key(group: Mapping[str, object]) -> Tuple[int, str, str]:
    return (-int(group.get("target_count") or 0), str(group.get("source_refdes") or ""), str(group.get("supply_net") or ""))


def _build_supply_edge_groups(records: Sequence[Mapping[str, object]],
                              nodes: Mapping[str, Mapping[str, object]],
                              *,
                              limit: int) -> Tuple[List[dict], List[dict]]:
    grouped: Dict[Tuple[str, str, str], dict] = {}
    for record in records:
        source_ref = str(record.get("source_refdes") or "")
        target_ref = str(record.get("target_refdes") or "")
        net_name = _safe_text(record.get("supply_net", ""), 180)
        voltage_domain = _safe_text(record.get("voltage_domain", "") or _voltage_domain_from_net(net_name), 80)
        target_node = nodes.get(target_ref, {})
        target_role = _safe_text(target_node.get("role", ""), 80)
        key = (source_ref, net_name, voltage_domain)
        item = grouped.setdefault(key, {
            "group_id": f"supply-group-{_safe_fragment(source_ref)}-{_safe_fragment(net_name)}",
            "edge_kind": "supply_group",
            "source_refdes": source_ref,
            "source_role": (nodes.get(source_ref, {}) or {}).get("role", ""),
            "source_page": (nodes.get(source_ref, {}) or {}).get("user_visible_page", ""),
            "supply_net": net_name,
            "voltage_domain": voltage_domain,
            "target_count": 0,
            "target_refdes_list": [],
            "sample_target_refdes": "",
            "target_roles": [],
            "target_role_counts": {},
            "target_pages": [],
            "sample_edge_ids": [],
            "relation_label": "电源管理供电关系聚合",
            "review_priority": "medium",
            "review_score": 35,
            "risk_tags": ["power_rail_dependency"],
            "review_focus": ["输出电压", "负载电流", "上电时序", "去耦滤波", "EN/PGOOD/RESET"],
            "review_hints": [
                "这是供电关系聚合视图；点选样本或切换全量模式查看每个负载芯片的明细边。",
                "按电源芯片和 rail 汇总后，优先核对输出能力、上电时序、去耦和跨页网名一致性。",
            ],
            "detail_tool": {"name": "query_llm_topology_netlist", "args": {"query": f"{source_ref} {net_name}", "limit": 20}},
        })
        item["target_count"] = int(item.get("target_count") or 0) + 1
        if target_ref and target_ref not in item["target_refdes_list"]:
            item["target_refdes_list"].append(target_ref)
        if not item.get("sample_target_refdes") and target_ref:
            item["sample_target_refdes"] = target_ref
        if target_role:
            role_counts = item["target_role_counts"]
            role_counts[target_role] = int(role_counts.get(target_role) or 0) + 1
        page = _safe_text(target_node.get("user_visible_page", ""), 40)
        if page and page not in item["target_pages"] and len(item["target_pages"]) < 12:
            item["target_pages"].append(page)
        edge_id = _safe_text(record.get("edge_id", ""), 220)
        if edge_id and len(item["sample_edge_ids"]) < 8:
            item["sample_edge_ids"].append(edge_id)
    groups = []
    for item in grouped.values():
        role_counts = item.get("target_role_counts") if isinstance(item.get("target_role_counts"), Mapping) else {}
        item["target_roles"] = [
            {"role": role, "count": count}
            for role, count in sorted(role_counts.items(), key=lambda pair: (-int(pair[1]), pair[0]))[:8]
        ]
        item["target_refdes_list"] = list(item.get("target_refdes_list") or [])[:48]
        item["summary"] = (
            f"{item.get('source_refdes')} 通过 {item.get('supply_net')} "
            f"给 {item.get('target_count')} 个芯片/连接器节点提供供电关系。"
        )
        groups.append(item)
    groups.sort(key=_supply_group_sort_key)
    visible_limit = max(0, min(limit, MAX_SUPPLY_LIMIT))
    return groups, groups[:visible_limit] if visible_limit else []


def _node_review_score(node: Mapping[str, object]) -> Tuple[int, str]:
    tags = set(node.get("risk_tags") or [])
    groups = set(node.get("interface_groups") or [])
    role = str(node.get("role") or "")
    score = 0
    if role == "level_shifter" or "level_shift_voltage_domain" in tags:
        score += 45
    if role == "power_management_ic" or "power_sequence_and_enable" in tags:
        score += 35
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}:
        score += 35
    if groups & {"clock", "reset"}:
        score += 25
    if groups & {"i2c", "spi", "power_control", "jtag_debug", "audio", "analog_sense"}:
        score += 15
    if int(node.get("pin_count") or 0) >= 64:
        score += 10
    priority = "high" if score >= 55 else "medium" if score >= 25 else "low"
    return score, priority


def _edge_risk_tags(edge: Mapping[str, object]) -> List[str]:
    groups = set(edge.get("interface_groups") or [])
    roles = {str(edge.get("source_role") or ""), str(edge.get("target_role") or "")}
    tags: List[str] = []
    if "level_shifter" in roles:
        tags.append("level_shift_cross_domain_review")
    if edge.get("voltage_domain_transition"):
        tags.append("voltage_domain_transition")
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}:
        tags.append("high_speed_interface")
    if "jtag_debug" in groups:
        tags.append("debug_port_access")
    if "audio" in groups:
        tags.append("audio_clock_domain")
    if "analog_sense" in groups:
        tags.append("analog_feedback_or_sense")
    if "clock" in groups:
        tags.append("clock_distribution")
    if "reset" in groups:
        tags.append("reset_timing")
    if "power_control" in groups:
        tags.append("power_control_sequence")
    if "i2c" in groups:
        tags.append("i2c_pullup_voltage_domain")
    if edge.get("passive_bridge_count"):
        tags.append("one_hop_passive_bridge")
    return tags[:10] or ["general_topology_review"]


def _edge_review_score(edge: Mapping[str, object]) -> Tuple[int, str]:
    groups = set(edge.get("interface_groups") or [])
    roles = {str(edge.get("source_role") or ""), str(edge.get("target_role") or "")}
    score = 0
    if "level_shifter" in roles:
        score += 45
    if edge.get("voltage_domain_transition"):
        score += 25
    if groups & {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}:
        score += 35
    if groups & {"clock", "reset"}:
        score += 25
    if groups & {"i2c", "spi", "power_control", "jtag_debug", "audio", "analog_sense"}:
        score += 15
    if edge.get("passive_bridge_count"):
        score += 15
    score += min(int(edge.get("shared_net_count") or 0), 6) * 2
    priority = "high" if score >= 55 else "medium" if score >= 25 else "low"
    return score, priority


def _edge_interface_summary(edge: Mapping[str, object]) -> List[dict]:
    groups: Dict[str, dict] = {}
    for shared in edge.get("shared_nets") or []:
        if not isinstance(shared, Mapping):
            continue
        group = str(shared.get("interface_group") or "misc_signal")
        item = groups.setdefault(group, {
            "group": group,
            "net_count": 0,
            "nets": [],
            "pin_samples": [],
            "passive_bridge_count": 0,
        })
        item["net_count"] += 1
        net_name = _safe_text(shared.get("net", ""), 180)
        if net_name and net_name not in item["nets"]:
            item["nets"].append(net_name)
        if len(item["pin_samples"]) < 6:
            item["pin_samples"].append({
                "source_pins": list(shared.get("source_pins") or [])[:2],
                "target_pins": list(shared.get("target_pins") or [])[:2],
            })
    for bridge in edge.get("passive_bridges") or []:
        if not isinstance(bridge, Mapping):
            continue
        group = _interface_group(" ".join(str(net) for net in bridge.get("nets") or []), [])
        item = groups.setdefault(group, {
            "group": group,
            "net_count": 0,
            "nets": [],
            "pin_samples": [],
            "passive_bridge_count": 0,
        })
        item["passive_bridge_count"] += 1
        for net_name in bridge.get("nets") or []:
            text = _safe_text(net_name, 180)
            if text and text not in item["nets"]:
                item["nets"].append(text)
    return sorted(
        groups.values(),
        key=lambda item: (-int(item.get("net_count") or 0), -int(item.get("passive_bridge_count") or 0), str(item.get("group") or "")),
    )[:12]


def _topology_counts(nodes: Sequence[Mapping[str, object]],
                     node_list: Sequence[Mapping[str, object]],
                     edges: Sequence[Mapping[str, object]],
                     visible_edges: Sequence[Mapping[str, object]],
                     supply_edges: Sequence[Mapping[str, object]],
                     visible_supply_edges: Sequence[Mapping[str, object]],
                     supply_edge_groups: Sequence[Mapping[str, object]] = (),
                     visible_supply_edge_groups: Sequence[Mapping[str, object]] = ()) -> dict:
    return {
        "total_node_count": len(nodes),
        "returned_node_count": len(node_list),
        "total_signal_edge_count": len(edges),
        "returned_signal_edge_count": len(visible_edges),
        "total_supply_edge_count": len(supply_edges),
        "returned_supply_edge_count": len(visible_supply_edges),
        "total_supply_group_count": len(supply_edge_groups),
        "returned_supply_group_count": len(visible_supply_edge_groups),
        "visual_edge_count": len(visible_edges) + len(visible_supply_edges) + len(visible_supply_edge_groups),
    }


def _edge_nets_preview(edge: Mapping[str, object]) -> List[str]:
    nets: List[str] = []
    for shared in edge.get("shared_nets") or []:
        if isinstance(shared, Mapping):
            net = _safe_text(shared.get("net", ""), 120)
            if net and net not in nets:
                nets.append(net)
    for bridge in edge.get("passive_bridges") or []:
        if isinstance(bridge, Mapping):
            for net in bridge.get("nets") or []:
                text = _safe_text(net, 120)
                if text and text not in nets:
                    nets.append(text)
    if edge.get("supply_net"):
        text = _safe_text(edge.get("supply_net", ""), 120)
        if text and text not in nets:
            nets.append(text)
    return nets[:8]


def _business_pages_for_item(item: Mapping[str, object]) -> List[str]:
    pages = [
        _safe_text(item.get("source_page", ""), 40),
        _safe_text(item.get("target_page", ""), 40),
        _safe_text(item.get("user_visible_page", ""), 40),
        _safe_text(item.get(USER_VISIBLE_REAL_PAGE_LABEL, ""), 40),
    ]
    return [page for page in dict.fromkeys(page for page in pages if page)]


def _business_edge_item(edge: Mapping[str, object], *, kind: str) -> dict:
    return {
        "item_id": _safe_text(edge.get("edge_id", ""), 180),
        "kind": kind,
        "title": edge.get("relation_label") or edge.get("edge_id") or "拓扑关系",
        "summary": edge.get("summary", ""),
        "source_refdes": edge.get("source_refdes", ""),
        "target_refdes": edge.get("target_refdes", ""),
        "pages": _business_pages_for_item(edge),
        "interface_groups": list(edge.get("interface_groups") or []),
        "nets_preview": _edge_nets_preview(edge),
        "supply_net": edge.get("supply_net", ""),
        "voltage_domains": list(dict.fromkeys([
            *list(edge.get("source_voltage_domains") or []),
            *list(edge.get("target_voltage_domains") or []),
            *([edge.get("voltage_domain")] if edge.get("voltage_domain") else []),
        ]))[:8],
        "risk_tags": list(edge.get("risk_tags") or [])[:8],
        "review_priority": edge.get("review_priority", "low"),
        "review_score": edge.get("review_score", 0),
        "review_focus": list(edge.get("review_focus") or [])[:12],
        "review_hints": list(edge.get("review_hints") or [])[:5],
        "confidence": edge.get("confidence", "medium" if kind == "supply_edge" else "low"),
        "evidence_id": edge.get("evidence_id", ""),
        "detail_tool": edge.get("detail_tool"),
    }


def _business_node_item(node: Mapping[str, object]) -> dict:
    return {
        "item_id": _safe_text(node.get("node_id", ""), 160),
        "kind": "node",
        "title": f"{node.get('refdes')} · {node.get('role') or node.get('category') or 'chip'}",
        "summary": (
            f"{node.get('refdes')} 位于页码 {node.get('user_visible_page') or ''}，"
            f"角色={node.get('role') or ''}，接口={', '.join(node.get('interface_groups') or []) or 'misc_signal'}。"
        ),
        "refdes": node.get("refdes", ""),
        "pages": _business_pages_for_item(node),
        "interface_groups": list(node.get("interface_groups") or [])[:8],
        "nets_preview": list(node.get("signal_net_preview") or [])[:8],
        "voltage_domains": list(node.get("voltage_domains") or [])[:8],
        "risk_tags": list(node.get("risk_tags") or [])[:8],
        "review_priority": node.get("review_priority", "low"),
        "review_score": node.get("review_score", 0),
        "review_focus": list(dict.fromkeys([
            *list(node.get("risk_tags") or []),
            *[focus for group in node.get("interface_groups") or [] for focus in review_focus_for_interface(str(group))],
        ]))[:12],
        "evidence_id": node.get("evidence_id", ""),
        "detail_tool": node.get("detail_tool"),
    }


def _review_task_detail_tool(task_id: str) -> dict:
    return {"name": "get_topology_review_task", "args": {"task_id": task_id}}


def _review_task_id(kind: str, source_id: object) -> str:
    return f"topology-review-{_safe_fragment(kind)}-{_safe_fragment(source_id)}"


def _edge_review_task(edge: Mapping[str, object], *, kind: str) -> dict:
    source_id = edge.get("edge_id") or edge.get("evidence_id") or ""
    interface_completeness = list(edge.get("interface_completeness") or [])
    missing_signals = list(dict.fromkeys(
        signal
        for item in interface_completeness
        for signal in item.get("missing_signals", []) or []
    ))[:12]
    observed_signals = list(dict.fromkeys(
        signal
        for item in interface_completeness
        for signal in item.get("observed_signals", []) or []
    ))[:12]
    checklist = list(dict.fromkeys([
        *list(edge.get("review_focus") or []),
        *[check for item in interface_completeness for check in item.get("recommended_checks", []) or []],
    ]))[:16]
    task_id = _review_task_id(kind, source_id)
    return {
        "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
        "task_id": task_id,
        "source_kind": kind,
        "source_id": source_id,
        "title": edge.get("relation_label") or f"{edge.get('source_refdes')} ↔ {edge.get('target_refdes')}",
        "summary": edge.get("summary", ""),
        "review_priority": edge.get("review_priority", "low"),
        "review_score": edge.get("review_score", 0),
        "refdes": [edge.get("source_refdes", ""), edge.get("target_refdes", "")],
        "pages": _business_pages_for_item(edge),
        "interface_groups": list(edge.get("interface_groups") or [])[:8],
        "risk_tags": list(edge.get("risk_tags") or [])[:10],
        "review_focus": list(edge.get("review_focus") or [])[:12],
        "checklist": checklist,
        "interface_completeness": interface_completeness,
        "observed_signals": observed_signals,
        "missing_signals": missing_signals,
        "evidence_id": edge.get("evidence_id", ""),
        "edge_detail_tool": edge.get("detail_tool"),
        "detail_tool": _review_task_detail_tool(task_id),
    }


def _node_review_task(node: Mapping[str, object]) -> dict:
    task_id = _review_task_id("node", node.get("node_id") or node.get("refdes"))
    identity_hint = node.get("llm_device_identity_hint") if isinstance(node.get("llm_device_identity_hint"), Mapping) else {}
    checklist = list(dict.fromkeys([
        *list(node.get("risk_tags") or []),
        *[focus for group in node.get("interface_groups") or [] for focus in review_focus_for_interface(str(group))],
        "器件身份/服务器项目角色确认",
        "关键电源/时钟/reset/control 网络复核",
    ]))[:16]
    return {
        "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
        "task_id": task_id,
        "source_kind": "node",
        "source_id": node.get("node_id", ""),
        "title": f"{node.get('refdes')} · {node.get('role') or node.get('category') or 'chip'}",
        "summary": (
            f"{node.get('refdes')} 位于页码 {node.get('user_visible_page') or ''}，"
            f"本地角色提示={node.get('role') or ''}；需要结合 PART_NAME/规格/网络判断服务器项目器件职责。"
        ),
        "review_priority": node.get("review_priority", "low"),
        "review_score": node.get("review_score", 0),
        "refdes": [node.get("refdes", "")],
        "pages": _business_pages_for_item(node),
        "interface_groups": list(node.get("interface_groups") or [])[:8],
        "risk_tags": list(node.get("risk_tags") or [])[:8],
        "review_focus": checklist[:12],
        "checklist": checklist,
        "llm_device_identity_hint": dict(identity_hint),
        "evidence_id": node.get("evidence_id", ""),
        "node_detail_tool": node.get("detail_tool"),
        "detail_tool": _review_task_detail_tool(task_id),
    }


def _build_topology_review_tasks(nodes: Sequence[Mapping[str, object]],
                                 edges: Sequence[Mapping[str, object]],
                                 supply_edges: Sequence[Mapping[str, object]],
                                 *,
                                 limit: int = 80) -> List[dict]:
    tasks: List[dict] = []
    tasks.extend(_edge_review_task(edge, kind="signal_edge") for edge in edges)
    tasks.extend(_edge_review_task(edge, kind="supply_edge") for edge in supply_edges)
    tasks.extend(
        _node_review_task(node)
        for node in nodes
        if str(node.get("review_priority") or "low") in {"high", "medium"}
    )
    tasks = sorted(
        tasks,
        key=lambda item: (
            {"high": 0, "medium": 1, "low": 2}.get(str(item.get("review_priority") or "low"), 3),
            -int(item.get("review_score") or 0),
            str(item.get("task_id") or ""),
        ),
    )
    return tasks[:max(1, min(int(limit or 80), 200))]


def _partition_priority(items: Sequence[Mapping[str, object]]) -> str:
    priorities = [str(item.get("review_priority") or "low") for item in items]
    if "high" in priorities:
        return "high"
    if "medium" in priorities:
        return "medium"
    return "low"


def _business_partition(partition_id: str, title: str, items: Sequence[dict], lead: str) -> dict:
    sorted_items = sorted(
        list(items),
        key=lambda item: (
            {"high": 0, "medium": 1, "low": 2}.get(str(item.get("review_priority") or "low"), 3),
            -int(item.get("review_score") or 0),
            str(item.get("item_id") or ""),
        ),
    )
    return {
        "partition_id": partition_id,
        "title": title,
        "lead": lead,
        "priority": _partition_priority(sorted_items),
        "item_count": len(sorted_items),
        "items": sorted_items[:12],
        "truncated": len(sorted_items) > 12,
    }


def _build_topology_business_view(*,
                                  counts: Mapping[str, int],
                                  node_list: Sequence[Mapping[str, object]],
                                  edges: Sequence[Mapping[str, object]],
                                  supply_edges: Sequence[Mapping[str, object]],
                                  interface_groups: Sequence[Mapping[str, object]],
                                  skipped_power_net_count: int,
                                  skipped_global_net_count: int,
                                  skipped_power_nets_sample: Sequence[str],
                                  skipped_global_nets_sample: Sequence[str],
                                  truncated: bool,
                                  include_connectors: bool,
                                  scope_note: str) -> dict:
    all_edge_items = [_business_edge_item(edge, kind="signal_edge") for edge in edges]
    supply_items = [_business_edge_item(edge, kind="supply_edge") for edge in supply_edges]
    node_items = [_business_node_item(node) for node in node_list]

    high_speed_groups = {"pcie", "usb", "mipi_lvds", "high_speed", "ddr", "ethernet", "storage_sdio"}
    partitions = [
        _business_partition(
            "power_rails",
            "电源/上电链路",
            supply_items,
            "PMIC/电源芯片到负载芯片的供电依赖，重点看电压、负载、去耦、EN/PGOOD/RESET。",
        ),
        _business_partition(
            "level_shift_cross_domain",
            "电平转换/跨电压域",
            [
                item for item in all_edge_items
                if "level_shift_cross_domain_review" in item.get("risk_tags", [])
                or "voltage_domain_transition" in item.get("risk_tags", [])
            ],
            "疑似电平转换或跨电压域连接，重点看两侧供电、方向/OE、上拉和上电顺序。",
        ),
        _business_partition(
            "high_speed_interfaces",
            "高速/存储接口",
            [
                item for item in all_edge_items
                if set(item.get("interface_groups") or []) & high_speed_groups
            ],
            "高速和存储接口，重点看差分/阻抗/AC 耦合/端接/长度/参考时钟。",
        ),
        _business_partition(
            "clock_reset_control",
            "时钟/复位/控制",
            [
                item for item in all_edge_items
                if set(item.get("interface_groups") or []) & {"clock", "reset", "power_control", "interrupt"}
            ],
            "时钟、复位和控制类网络，重点看默认态、释放时序、扇出、上下拉和跨域。",
        ),
        _business_partition(
            "debug_test_access",
            "Debug/测试/外部访问",
            [
                item for item in all_edge_items
                if set(item.get("interface_groups") or []) & {"jtag_debug"}
            ],
            "JTAG/debug/测试访问路径，重点看上下拉、量产可访问性、测试点和复用状态。",
        ),
        _business_partition(
            "analog_sense_feedback",
            "模拟采样/反馈",
            [
                item for item in all_edge_items
                if set(item.get("interface_groups") or []) & {"analog_sense", "audio"}
            ],
            "模拟采样、反馈和音频类连接，重点看量程、滤波、地参考和时钟域。",
        ),
        _business_partition(
            "passive_bridge_review",
            "一跳无源桥",
            [
                item for item in all_edge_items
                if "one_hop_passive_bridge" in item.get("risk_tags", [])
            ],
            "R/C/L/FB 不作为主节点，但作为连接辅助证据，需要复核串阻、耦合、滤波和装配状态。",
        ),
        _business_partition(
            "major_hubs",
            "主要芯片/Hubs",
            [
                item for item in node_items
                if item.get("review_priority") in {"high", "medium"}
            ],
            "连接度或风险较高的关键芯片节点，适合从芯片视角展开二次 detail tool。",
        ),
    ]
    coverage_items: List[dict] = []
    if skipped_power_net_count:
        coverage_items.append({
            "item_id": "coverage-skipped-power-nets",
            "kind": "coverage_gap",
            "title": "电源/地网未作为普通信号边展开",
            "summary": f"已跳过 {skipped_power_net_count} 个电源/地类网络；其中 PMIC 到负载会单独进入供电关系。",
            "nets_preview": list(skipped_power_nets_sample)[:12],
            "review_priority": "medium",
            "review_score": 25,
            "review_focus": ["供电关系", "公共电源网", "必要时用原始 net detail 补查"],
        })
    if skipped_global_net_count:
        coverage_items.append({
            "item_id": "coverage-skipped-global-nets",
            "kind": "coverage_gap",
            "title": "过大公共网络未展开",
            "summary": f"已跳过 {skipped_global_net_count} 个节点数过多的公共网络，避免把全局信号误判为业务连接。",
            "nets_preview": list(skipped_global_nets_sample)[:12],
            "review_priority": "medium",
            "review_score": 25,
            "review_focus": ["公共网", "可能需要按网络单独查询"],
        })
    if truncated:
        coverage_items.append({
            "item_id": "coverage-truncated-result",
            "kind": "coverage_gap",
            "title": "拓扑结果已截断",
            "summary": "当前返回为轻量预览，完整节点/边请通过 CLI full/out 或 Harness detail tool 二次读取。",
            "review_priority": "medium",
            "review_score": 20,
            "review_focus": ["不要从 preview 推断完整统计", "按 evidence/detail tool 补取"],
        })
    if not include_connectors:
        coverage_items.append({
            "item_id": "coverage-connectors-disabled",
            "kind": "coverage_gap",
            "title": "连接器默认未纳入主节点",
            "summary": "当前芯片级拓扑默认聚焦 U/PU/XU；如需外部接口/连接器 review，请启用 include_connectors。",
            "review_priority": "low",
            "review_score": 10,
            "review_focus": ["连接器", "外部接口", "测试访问路径"],
        })
    partitions.append(_business_partition(
        "coverage_gaps",
        "覆盖范围/缺口提示",
        coverage_items,
        "说明本语义拓扑没有展开的范围，避免 LLM 或外部进程把预览当完整签核图。",
    ))
    partitions = [partition for partition in partitions if partition.get("item_count")]
    review_queue = sorted(
        [item for partition in partitions for item in partition.get("items", [])],
        key=lambda item: (
            {"high": 0, "medium": 1, "low": 2}.get(str(item.get("review_priority") or "low"), 3),
            -int(item.get("review_score") or 0),
            str(item.get("item_id") or ""),
        ),
    )[:24]
    dictionary_summary = business_dictionary_summary()
    return {
        "schema_version": "llm-topology-business-view.v1",
        "summary": (
            f"业务视角：{counts.get('total_node_count', 0)} 个芯片级节点、"
            f"{counts.get('total_signal_edge_count', 0)} 条信号关系、"
            f"{counts.get('total_supply_edge_count', 0)} 条供电关系；"
            f"形成 {len(partitions)} 个 review 分区。"
        ),
        "scope_note": scope_note,
        "counts": dict(counts),
        "dictionary": {
            "schema_version": dictionary_summary.get("schema_version"),
            "source": dictionary_summary.get("source"),
            "interface_aliases": interface_alias_snapshot(),
        },
        "interfaces": list(interface_groups)[:12],
        "review_partitions": partitions,
        "review_queue": review_queue,
        "legend": [
            "业务视角只保留摘要和证据入口，不替代完整 Cadence/net detail。",
            "高风险/定量结论必须使用 detail_tool 回拉原始 pin/net、页码和器件属性。",
            "coverage_gaps 表示未展开范围，不代表无风险。",
        ],
    }


def _infer_role(refdes: str, comp: Mapping[str, object], pin_nets: Sequence[dict]) -> str:
    text = _component_text(refdes, comp, pin_nets)
    upper = refdes.upper()
    category, candidate, _confidence = classify_refdes(refdes, dict(comp))
    if LEVEL_SHIFTER_RE.search(text):
        return "level_shifter"
    if POWER_IC_RE.search(text) or category == "power_ic" or upper.startswith("PU"):
        return "power_management_ic"
    if PROCESSOR_RE.search(text):
        return "processor_or_fpga"
    if MEMORY_RE.search(text):
        return "memory"
    if CLOCK_RE.search(text):
        return "clock_source"
    if category == "large_ic" or len(pin_nets) >= 64:
        return "large_ic"
    if category == "connector":
        return "connector"
    return candidate or category or "ic"


def _iter_component_pin_nets(refdes: str, comp: Mapping[str, object], nets: Mapping[str, object]) -> List[dict]:
    rows: List[dict] = []
    seen = set()
    comp_nets = comp.get("nets") if isinstance(comp.get("nets"), Mapping) else {}
    for pin, net_name in sorted(comp_nets.items(), key=lambda item: str(item[0])):
        key = (str(pin), str(net_name))
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "pin": _safe_text(pin, 80),
            "pin_name": _safe_text(pin, 160),
            "net": _safe_text(net_name, 180),
            "source": "component.nets",
        })
    for net_name, nodes in (nets or {}).items():
        if not isinstance(nodes, list):
            continue
        for node in nodes:
            if not isinstance(node, Mapping) or str(node.get("refdes") or "") != str(refdes):
                continue
            pin = _safe_text(node.get("pin", ""), 80)
            net_text = _safe_text(net_name, 180)
            key = (pin, net_text)
            if key in seen:
                continue
            seen.add(key)
            rows.append({
                "pin": pin,
                "pin_name": _safe_text(node.get("pin_name", pin), 160),
                "net": net_text,
                "source": "nets",
            })
    return rows


def _derive_nets_from_components(components: Mapping[str, object], nets: Mapping[str, object]) -> Dict[str, List[dict]]:
    result: Dict[str, List[dict]] = defaultdict(list)
    for net_name, nodes in (nets or {}).items():
        if isinstance(nodes, list):
            for node in nodes:
                if isinstance(node, Mapping):
                    result[str(net_name)].append(dict(node))
    seen = {
        (str(net), str(node.get("refdes") or ""), str(node.get("pin") or ""))
        for net, nodes in result.items()
        for node in nodes
    }
    for refdes, comp in (components or {}).items():
        if not isinstance(comp, Mapping):
            continue
        comp_nets = comp.get("nets") if isinstance(comp.get("nets"), Mapping) else {}
        for pin, net_name in comp_nets.items():
            net_text = str(net_name or "")
            key = (net_text, str(refdes), str(pin))
            if not net_text or key in seen:
                continue
            seen.add(key)
            result[net_text].append({"refdes": str(refdes), "pin": str(pin), "pin_name": str(pin)})
    return dict(result)


def _node_preview(refdes: str, comp: Mapping[str, object], pin_nets: Sequence[dict], *, include_connectors: bool = False) -> Optional[dict]:
    keep, category, candidate, confidence = _is_chip_ref(refdes, comp, include_connectors=include_connectors)
    if not keep:
        return None
    role = _infer_role(refdes, comp, pin_nets)
    interface_counts = Counter(_interface_group(row.get("net"), [row.get("pin_name"), row.get("pin")]) for row in pin_nets)
    signal_nets = [
        row.get("net", "")
        for row in pin_nets
        if row.get("net") and not _is_power_or_ground_net(row.get("net"))
    ]
    node_id = f"chip-node-{_safe_fragment(refdes)}"
    critical_nets = _critical_net_preview(pin_nets)
    voltage_domains = _voltage_domains_from_nets(critical_nets.get("power_nets") or [])
    node = {
        "node_id": node_id,
        "evidence_id": _evidence_id("node", refdes),
        "refdes": _safe_text(refdes, 80),
        "category": category,
        "role": role,
        "candidate_chip_type": candidate,
        "confidence": confidence,
        "hq_no": _component_hq_no(comp),
        "spec": _component_spec(comp),
        "package": _component_package(comp),
        "value": _component_value(comp),
        "module": _component_module(comp),
        USER_VISIBLE_REAL_PAGE_LABEL: _component_page(comp),
        "user_visible_page": _component_page(comp),
        "pin_count": len(pin_nets),
        "signal_net_count": len(set(signal_nets)),
        "interface_groups": [name for name, _count in interface_counts.most_common(8) if name != "misc_signal"],
        "signal_net_preview": list(dict.fromkeys(signal_nets))[:12],
        "voltage_domains": voltage_domains,
        **critical_nets,
        "detail_tool": _node_detail_tool(str(refdes)),
    }
    node["llm_device_identity_hint"] = _llm_device_identity_hint(
        str(refdes),
        comp,
        role=role,
        category=category,
        candidate=candidate,
        confidence=confidence,
    )
    node["risk_tags"] = _node_review_tags(node)
    node["review_score"], node["review_priority"] = _node_review_score(node)
    return node


def _edge_id(ref_a: str, ref_b: str) -> str:
    left, right = sorted([str(ref_a), str(ref_b)], key=lambda item: item.upper())
    return f"chip-edge-{_safe_fragment(left)}-{_safe_fragment(right)}"


def _edge_summary(edge: Mapping[str, object]) -> str:
    groups = ", ".join(edge.get("interface_groups") or []) or "misc_signal"
    nets = ", ".join(str(item.get("net") or "") for item in list(edge.get("shared_nets") or [])[:6])
    return (
        f"{edge.get('source_refdes')}({edge.get('source_role')}) 与 "
        f"{edge.get('target_refdes')}({edge.get('target_role')}) 共享 "
        f"{edge.get('shared_net_count', 0)} 条信号网络；接口类型={groups}；示例网络={nets}。"
    )


def _connection_label(role_a: str, role_b: str) -> str:
    roles = {role_a, role_b}
    if "level_shifter" in roles:
        return "芯片到电平转换连接"
    if "power_management_ic" in roles:
        return "芯片到电源管理连接"
    if "processor_or_fpga" in roles and ("memory" in roles or "large_ic" in roles):
        return "主芯片到关键芯片连接"
    return "芯片级连接"


def _edge_confidence(shared_net_count: int, interface_groups: Sequence[str]) -> str:
    strong_groups = {item for item in interface_groups if item not in {"misc_signal", ""}}
    if shared_net_count >= 4 or (shared_net_count >= 2 and strong_groups):
        return "high"
    if shared_net_count >= 2 or strong_groups:
        return "medium"
    return "low"


def _passive_bridge_type(refdes: str, comp: Mapping[str, object]) -> str:
    upper = str(refdes or "").upper()
    value = _component_value(comp).upper()
    passive_kind = _passive_kind_from_refdes(upper)
    if passive_kind == "R":
        if "0R" in value or value in {"0", "0OHM", "0Ω"}:
            return "zero_ohm_or_link"
        return "series_or_pull_resistor"
    if passive_kind == "C":
        return "coupling_or_filter_capacitor"
    if passive_kind in {"L", "FB"}:
        return "inductor_or_ferrite"
    return "passive_bridge"


def _passive_kind_from_refdes(refdes: object) -> str:
    upper = str(refdes or "").upper()
    if upper.startswith("FB"):
        return "FB"
    if upper.startswith(("RN", "RP", "RA")):
        return "R"
    if upper[:1] in {"R", "C", "L"}:
        return upper[:1]
    if len(upper) >= 2 and upper[0] == "P" and upper[1] in {"R", "C", "L"}:
        return upper[1]
    return ""


def _passive_assembly_state(comp: Mapping[str, object]) -> str:
    text = _component_bom_option(comp).upper()
    if any(token in text for token in ("DEPOP", "DNP", "NOPOP", "NC", "不贴", "不装")):
        return "not_populated"
    if text:
        return "conditional"
    return "unknown"


def _passive_bridge_semantics(refdes: str, comp: Mapping[str, object]) -> dict:
    bridge_type = _passive_bridge_type(refdes, comp)
    assembly_state = _passive_assembly_state(comp)
    value = _component_value(comp)
    value_upper = value.upper().replace(" ", "")
    dc_conductive = bridge_type in {"zero_ohm_or_link", "series_or_pull_resistor", "inductor_or_ferrite"}
    if bridge_type == "coupling_or_filter_capacitor":
        semantics = "capacitive_coupling_or_filter"
        dc_conductive = False
    elif bridge_type == "zero_ohm_or_link":
        semantics = "dc_link_or_strap"
    elif bridge_type == "inductor_or_ferrite":
        semantics = "dc_path_with_filtering"
    else:
        semantics = "resistive_or_series_path"
    if assembly_state == "not_populated":
        dc_conductive = False
    parsed_value = value
    low_ohm = bool(re.search(r"(^|[^0-9])0(?:R|OHM|Ω|$)", value_upper)) or value_upper in {"0", "0R", "0Ω", "0OHM"}
    return {
        "bridge_type": bridge_type,
        "bridge_semantics": semantics,
        "dc_conductive": dc_conductive,
        "low_ohm_link": low_ohm,
        "assembly_state": assembly_state,
        "bom_option": _component_bom_option(comp),
        "parsed_value": parsed_value,
    }


def _build_passive_bridge_edges(components: Mapping[str, object],
                                nets: Mapping[str, List[dict]],
                                nodes: Mapping[str, dict],
                                pin_index: Mapping[str, Mapping[str, List[dict]]]) -> Dict[Tuple[str, str], List[dict]]:
    """Summarize one-hop R/C/L bridges without promoting passives to graph nodes."""
    bridges: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    net_to_chip_refs: Dict[str, List[str]] = {}
    for net_name, net_nodes in nets.items():
        refs: List[str] = []
        if _is_power_or_ground_net(net_name):
            continue
        for node in net_nodes or []:
            if not isinstance(node, Mapping):
                continue
            ref = str(node.get("refdes") or "")
            if ref in nodes and ref not in refs:
                refs.append(ref)
        if refs:
            net_to_chip_refs[str(net_name)] = refs

    for refdes, comp in (components or {}).items():
        if not isinstance(comp, Mapping):
            continue
        ref = str(refdes)
        passive_kind = _passive_kind_from_refdes(ref)
        if not passive_kind:
            continue
        semantics = _passive_bridge_semantics(ref, comp)
        comp_nets = comp.get("nets") if isinstance(comp.get("nets"), Mapping) else {}
        unique_nets = [
            str(net_name)
            for net_name in dict.fromkeys(str(item or "") for item in comp_nets.values())
            if net_name and not _is_power_or_ground_net(net_name)
        ]
        if len(unique_nets) != 2:
            continue
        left_net, right_net = unique_nets
        left_refs = net_to_chip_refs.get(left_net, [])
        right_refs = net_to_chip_refs.get(right_net, [])
        if not left_refs or not right_refs:
            continue
        for left_ref in left_refs:
            for right_ref in right_refs:
                if left_ref == right_ref:
                    continue
                key = tuple(sorted([left_ref, right_ref], key=lambda item: item.upper()))
                pins_left = list((pin_index.get(left_ref) or {}).get(left_net, []))[:4]
                pins_right = list((pin_index.get(right_ref) or {}).get(right_net, []))[:4]
                if key[0] == left_ref:
                    source_net = left_net
                    target_net = right_net
                    source_pins = pins_left
                    target_pins = pins_right
                else:
                    source_net = right_net
                    target_net = left_net
                    source_pins = pins_right
                    target_pins = pins_left
                bridge = {
                    "refdes": ref,
                    **semantics,
                    "value": _component_value(comp),
                    "package": _component_package(comp),
                    "user_visible_page": _component_page(comp),
                    USER_VISIBLE_REAL_PAGE_LABEL: _component_page(comp),
                    "nets": [left_net, right_net],
                    "source_net": source_net,
                    "target_net": target_net,
                    "source_pins": [{"pin": row.get("pin", ""), "pin_name": row.get("pin_name", "")} for row in source_pins],
                    "target_pins": [{"pin": row.get("pin", ""), "pin_name": row.get("pin_name", "")} for row in target_pins],
                    "summary": (
                        f"{ref} 作为一跳无源关联 {left_net} 与 {right_net}；"
                        f"语义={semantics['bridge_semantics']}，装配={semantics['assembly_state']}。"
                    ),
                }
                bridges[key].append(bridge)
    return bridges


def build_chip_topology(report: Mapping[str, object] | None,
                        bundle: Mapping[str, object] | None,
                        *,
                        focus_refdes: str = "",
                        role_filter: str = "",
                        include_connectors: bool = False,
                        limit: int = DEFAULT_LIMIT,
                        return_all_edges: bool = False,
                        view: str = "summary",
                        supply_mode: str = "",
                        supply_limit: int = DEFAULT_SUPPLY_LIMIT) -> dict:
    bundle = bundle or {}
    components = bundle.get("components") if isinstance(bundle.get("components"), Mapping) else {}
    raw_nets = bundle.get("nets") if isinstance(bundle.get("nets"), Mapping) else {}
    nets = _derive_nets_from_components(components, raw_nets)
    focus = str(focus_refdes or "").strip().upper()
    role_filter_text = str(role_filter or "").strip().lower()
    limit = _as_limit(limit)
    view_mode = _normalize_topology_view(view, return_all_edges)
    supply_display_mode = _normalize_supply_mode(supply_mode, view_mode, return_all_edges)
    supply_limit = _as_supply_limit(supply_limit)

    nodes: Dict[str, dict] = {}
    pin_index: Dict[str, Dict[str, List[dict]]] = defaultdict(lambda: defaultdict(list))
    for refdes, comp in components.items():
        if not isinstance(comp, Mapping):
            continue
        pin_nets = _iter_component_pin_nets(str(refdes), comp, nets)
        node = _node_preview(str(refdes), comp, pin_nets, include_connectors=include_connectors)
        if not node:
            continue
        ref = str(refdes)
        nodes[ref] = node
        for row in pin_nets:
            net_name = _safe_text(row.get("net", ""), 180)
            if net_name:
                pin_index[ref][net_name].append(dict(row))

    edge_map: Dict[Tuple[str, str], dict] = {}
    supply_records: List[dict] = []
    skipped_power_net_count = 0
    skipped_global_net_count = 0
    skipped_power_nets_sample: List[str] = []
    skipped_global_nets_sample: List[str] = []
    for net_name, net_nodes in nets.items():
        if not net_name:
            continue
        if _is_power_or_ground_net(net_name):
            skipped_power_net_count += 1
            if len(skipped_power_nets_sample) < 24:
                skipped_power_nets_sample.append(_safe_text(net_name, 180))
            if _is_ground_net(net_name):
                continue
            chip_refs = []
            for node in net_nodes or []:
                if not isinstance(node, Mapping):
                    continue
                ref = str(node.get("refdes") or "")
                if ref in nodes and ref not in chip_refs:
                    chip_refs.append(ref)
            power_sources = [
                ref for ref in chip_refs
                if str(nodes[ref].get("role") or "") == "power_management_ic"
            ]
            for source_ref in power_sources:
                for target_ref in chip_refs:
                    if source_ref == target_ref:
                        continue
                    supply_records.append(_supply_record(source_ref, target_ref, net_name))
            continue
        chip_refs = []
        for node in net_nodes or []:
            if not isinstance(node, Mapping):
                continue
            ref = str(node.get("refdes") or "")
            if ref in nodes and ref not in chip_refs:
                chip_refs.append(ref)
        if len(chip_refs) < 2:
            continue
        if len(chip_refs) > MAX_NET_NODES:
            skipped_global_net_count += 1
            if len(skipped_global_nets_sample) < 24:
                skipped_global_nets_sample.append(_safe_text(net_name, 180))
            continue
        for ref_a, ref_b in combinations(sorted(chip_refs, key=lambda item: item.upper()), 2):
            key = tuple(sorted([ref_a, ref_b], key=lambda item: item.upper()))
            edge = edge_map.setdefault(key, {
                "edge_id": _edge_id(ref_a, ref_b),
                "evidence_id": _evidence_id("edge", _edge_id(ref_a, ref_b)),
                "source_refdes": key[0],
                "target_refdes": key[1],
                "source_role": nodes[key[0]].get("role", ""),
                "target_role": nodes[key[1]].get("role", ""),
                "source_page": nodes[key[0]].get("user_visible_page", ""),
                "target_page": nodes[key[1]].get("user_visible_page", ""),
                "source_power_nets": list(nodes[key[0]].get("power_nets") or [])[:8],
                "target_power_nets": list(nodes[key[1]].get("power_nets") or [])[:8],
                "source_voltage_domains": list(nodes[key[0]].get("voltage_domains") or [])[:8],
                "target_voltage_domains": list(nodes[key[1]].get("voltage_domains") or [])[:8],
                "shared_nets": [],
                "interface_groups": [],
                "passive_bridges": [],
            })
            pins_a = pin_index[ref_a].get(str(net_name), [])
            pins_b = pin_index[ref_b].get(str(net_name), [])
            pin_names = [*(row.get("pin_name") for row in pins_a[:2]), *(row.get("pin_name") for row in pins_b[:2])]
            group = _interface_group(net_name, pin_names)
            edge["shared_nets"].append({
                "net": _safe_text(net_name, 180),
                "interface_group": group,
                "source_pins": [
                    {"pin": row.get("pin", ""), "pin_name": row.get("pin_name", "")}
                    for row in pins_a[:4]
                ],
                "target_pins": [
                    {"pin": row.get("pin", ""), "pin_name": row.get("pin_name", "")}
                    for row in pins_b[:4]
                ],
            })
            if group not in edge["interface_groups"]:
                edge["interface_groups"].append(group)

    passive_bridge_map = _build_passive_bridge_edges(components, nets, nodes, pin_index)
    for key, bridges in passive_bridge_map.items():
        if not bridges:
            continue
        ref_a, ref_b = key
        edge = edge_map.setdefault(key, {
            "edge_id": _edge_id(ref_a, ref_b),
            "evidence_id": _evidence_id("edge", _edge_id(ref_a, ref_b)),
            "source_refdes": ref_a,
            "target_refdes": ref_b,
            "source_role": nodes[ref_a].get("role", ""),
            "target_role": nodes[ref_b].get("role", ""),
            "source_page": nodes[ref_a].get("user_visible_page", ""),
            "target_page": nodes[ref_b].get("user_visible_page", ""),
            "source_power_nets": list(nodes[ref_a].get("power_nets") or [])[:8],
            "target_power_nets": list(nodes[ref_b].get("power_nets") or [])[:8],
            "source_voltage_domains": list(nodes[ref_a].get("voltage_domains") or [])[:8],
            "target_voltage_domains": list(nodes[ref_b].get("voltage_domains") or [])[:8],
            "shared_nets": [],
            "interface_groups": [],
            "passive_bridges": [],
        })
        edge["passive_bridges"].extend(bridges[:8])
        for bridge in bridges[:8]:
            group = _interface_group(" ".join(bridge.get("nets") or []), [])
            if group not in edge["interface_groups"]:
                edge["interface_groups"].append(group)

    edges = []
    for edge in edge_map.values():
        shared_count = len(edge.get("shared_nets") or [])
        groups = list(edge.get("interface_groups") or [])
        edge["shared_net_count"] = shared_count
        edge["passive_bridge_count"] = len(edge.get("passive_bridges") or [])
        edge["edge_kind"] = "signal" if shared_count else "passive_bridge"
        edge["undirected"] = True
        edge["endpoint_nets_by_refdes"] = {
            str(edge.get("source_refdes") or ""): list(dict.fromkeys([
                *[str(item.get("net") or "") for item in edge.get("shared_nets") or [] if item.get("net")],
                *[str(item.get("source_net") or "") for item in edge.get("passive_bridges") or [] if item.get("source_net")],
            ]))[:16],
            str(edge.get("target_refdes") or ""): list(dict.fromkeys([
                *[str(item.get("net") or "") for item in edge.get("shared_nets") or [] if item.get("net")],
                *[str(item.get("target_net") or "") for item in edge.get("passive_bridges") or [] if item.get("target_net")],
            ]))[:16],
        }
        source_domains = set(edge.get("source_voltage_domains") or [])
        target_domains = set(edge.get("target_voltage_domains") or [])
        edge["voltage_domain_transition"] = bool(source_domains and target_domains and source_domains != target_domains)
        edge["confidence"] = _edge_confidence(shared_count, groups)
        if (
            shared_count == 0
            and edge["passive_bridge_count"]
            and any(bool(bridge.get("dc_conductive")) for bridge in edge.get("passive_bridges") or [])
            and edge["confidence"] == "low"
        ):
            edge["confidence"] = "medium"
        edge["relation_label"] = _connection_label(str(edge.get("source_role") or ""), str(edge.get("target_role") or ""))
        edge["summary"] = _edge_summary(edge)
        if edge["passive_bridge_count"]:
            edge["summary"] = f"{edge['summary']}；另含 {edge['passive_bridge_count']} 个一跳无源桥。"
        if edge["voltage_domain_transition"]:
            edge["summary"] = (
                f"{edge['summary']}；疑似跨电压域 "
                f"{'/'.join(edge.get('source_voltage_domains') or [])} -> "
                f"{'/'.join(edge.get('target_voltage_domains') or [])}。"
            )
        edge["interface_summary"] = _edge_interface_summary(edge)
        edge["interface_completeness"] = _edge_interface_completeness(edge)
        edge["risk_tags"] = _edge_risk_tags(edge)
        edge["review_score"], edge["review_priority"] = _edge_review_score(edge)
        edge["review_focus"] = _edge_review_focus(edge)
        edge["review_hints"] = _edge_review_hints(edge)
        edge["detail_tool"] = _edge_detail_tool(str(edge.get("edge_id") or ""))
        edges.append(edge)

    if focus:
        edges = [
            edge for edge in edges
            if focus in {str(edge.get("source_refdes") or "").upper(), str(edge.get("target_refdes") or "").upper()}
        ]
        supply_records = [
            record for record in supply_records
            if focus in {str(record.get("source_refdes") or "").upper(), str(record.get("target_refdes") or "").upper()}
        ]

    role_filtered_refs: set[str] = set()
    if role_filter_text:
        role_filtered_refs = {
            ref for ref, node in nodes.items()
            if role_filter_text in str(node.get("role") or "").lower()
        }
        if role_filtered_refs:
            edges = [
                edge for edge in edges
                if str(edge.get("source_refdes") or "") in role_filtered_refs
                or str(edge.get("target_refdes") or "") in role_filtered_refs
            ]
            supply_records = [
                record for record in supply_records
                if str(record.get("source_refdes") or "") in role_filtered_refs
                or str(record.get("target_refdes") or "") in role_filtered_refs
            ]
        else:
            edges = []
            supply_records = []

    edges.sort(key=lambda edge: (
        0 if edge.get("relation_label") == "芯片到电平转换连接" else 1,
        -int(edge.get("review_score") or 0),
        -int(edge.get("shared_net_count") or 0),
        str(edge.get("source_refdes") or ""),
        str(edge.get("target_refdes") or ""),
    ))
    supply_records = sorted(
        supply_records,
        key=lambda record: (
            str(record.get("source_refdes") or ""),
            str(record.get("supply_net") or ""),
            str(record.get("target_refdes") or ""),
        ),
    )
    visible_edges = edges if view_mode == "full" else edges[:limit]
    supply_edge_groups, visible_supply_edge_groups = _build_supply_edge_groups(
        supply_records,
        nodes,
        limit=supply_limit if supply_display_mode == "grouped" else 0,
    )
    if supply_display_mode == "hidden":
        visible_supply_records: Sequence[Mapping[str, object]] = []
    elif supply_display_mode == "details" and view_mode == "full":
        visible_supply_records = supply_records
    else:
        visible_supply_records = supply_records[:supply_limit]
    visible_supply_edges = _supply_edges_from_records(visible_supply_records, nodes)
    ref_degree = Counter()
    for edge in edges:
        ref_degree[str(edge.get("source_refdes") or "")] += 1
        ref_degree[str(edge.get("target_refdes") or "")] += 1
    for record in supply_records:
        ref_degree[str(record.get("source_refdes") or "")] += 1
        ref_degree[str(record.get("target_refdes") or "")] += 1
    hubs = [
        {**nodes[ref], "degree": degree}
        for ref, degree in ref_degree.most_common(12)
        if ref in nodes
    ]
    role_links = [
        edge for edge in visible_edges
        if edge.get("relation_label") == "芯片到电平转换连接"
    ][:12]
    node_list = sorted(nodes.values(), key=lambda node: str(node.get("refdes") or "").upper())
    if focus:
        node_list = [node for node in node_list if str(node.get("refdes") or "").upper() == focus]
    elif role_filter_text:
        relevant_refs = set(role_filtered_refs)
        for edge in [*visible_edges, *visible_supply_edges]:
            relevant_refs.add(str(edge.get("source_refdes") or ""))
            relevant_refs.add(str(edge.get("target_refdes") or ""))
        for group in visible_supply_edge_groups:
            relevant_refs.add(str(group.get("source_refdes") or ""))
            relevant_refs.update(str(ref) for ref in group.get("target_refdes_list") or [])
        node_list = [node for node in node_list if str(node.get("refdes") or "") in relevant_refs]
    for node in node_list:
        node["connected_edge_count"] = ref_degree.get(str(node.get("refdes") or ""), 0)

    review_tasks = _build_topology_review_tasks(
        list(nodes.values()),
        edges,
        visible_supply_edges,
        limit=100,
    )

    interface_counter = Counter()
    for edge in edges:
        for group in edge.get("interface_groups") or ["misc_signal"]:
            interface_counter[group or "misc_signal"] += 1
    interface_groups = [
        {
            "group": group,
            "edge_count": count,
            "evidence_id": _evidence_id("interface", group),
            "detail_tool": {"name": "query_llm_topology_netlist", "args": {"query": group, "limit": 30}},
        }
        for group, count in interface_counter.most_common(24)
    ]
    risk_edges = [
        {
            "edge_id": edge.get("edge_id"),
            "source_refdes": edge.get("source_refdes"),
            "target_refdes": edge.get("target_refdes"),
            "review_priority": edge.get("review_priority"),
            "review_score": edge.get("review_score"),
            "risk_tags": edge.get("risk_tags", []),
            "summary": edge.get("summary", ""),
            "detail_tool": edge.get("detail_tool"),
        }
        for edge in sorted([*edges, *visible_supply_edges], key=lambda item: -int(item.get("review_score") or 0))[:12]
        if edge.get("review_priority") in {"high", "medium"}
    ]

    node_result_truncated = len(node_list) > limit
    visible_edges = list(visible_edges)
    visible_supply_edges = list(visible_supply_edges)
    visible_supply_edge_groups = list(visible_supply_edge_groups)
    node_list = list(node_list[:limit])
    counts = _topology_counts(
        list(nodes.values()),
        node_list,
        edges,
        visible_edges,
        supply_records,
        visible_supply_edges,
        supply_edge_groups,
        visible_supply_edge_groups,
    )
    topology_truncated = (
        node_result_truncated
        or len(edges) > len(visible_edges)
        or len(supply_records) > len(visible_supply_edges)
        or len(supply_edge_groups) > len(visible_supply_edge_groups)
    )
    scope_note = "这是芯片级、无方向、模糊拓扑摘要；不包含 R/C/L 等无源器件主节点，且电源/地网默认不作为芯片间连接依据。"
    business_view = _build_topology_business_view(
        counts=counts,
        node_list=node_list,
        edges=visible_edges,
        supply_edges=visible_supply_edges,
        interface_groups=interface_groups,
        skipped_power_net_count=skipped_power_net_count,
        skipped_global_net_count=skipped_global_net_count,
        skipped_power_nets_sample=skipped_power_nets_sample,
        skipped_global_nets_sample=skipped_global_nets_sample,
        truncated=topology_truncated,
        include_connectors=include_connectors,
        scope_note=scope_note,
    )

    summary_layer = {
        "project_name": (report or {}).get("project_name", ""),
        "schema_version": LLM_TOPOLOGY_SCHEMA_VERSION,
        "counts": counts,
        "filters": {
            "focus_refdes": focus,
            "role_filter": role_filter_text,
            "include_connectors": include_connectors,
            "limit": limit,
            "view": view_mode,
            "supply_mode": supply_display_mode,
            "supply_limit": supply_limit,
        },
        "truncated": topology_truncated,
        "business_view_summary": {
            "schema_version": business_view.get("schema_version"),
            "partition_count": len(business_view.get("review_partitions") or []),
            "review_queue_count": len(business_view.get("review_queue") or []),
        },
        "review_task_count": len(review_tasks),
        "review_tasks_preview": review_tasks[:12],
        "node_count": len(nodes),
        "edge_count": len(edges),
        "supply_edge_count": len(supply_records),
        "supply_group_count": len(supply_edge_groups),
        "returned_edge_count": len(visible_edges),
        "returned_supply_edge_count": len(visible_supply_edges),
        "returned_supply_group_count": len(visible_supply_edge_groups),
        "visual_edge_count": counts.get("visual_edge_count", 0),
        "hub_count": len(hubs),
        "interface_groups": interface_groups[:12],
        "major_hubs": [
            {
                "refdes": hub.get("refdes"),
                "role": hub.get("role"),
                "degree": hub.get("degree"),
                "review_priority": hub.get("review_priority"),
                "voltage_domains": hub.get("voltage_domains", []),
                "detail_tool": hub.get("detail_tool"),
            }
            for hub in hubs[:12]
        ],
        "risk_edges": risk_edges,
        "skipped_power_net_count": skipped_power_net_count,
        "skipped_power_nets_sample": skipped_power_nets_sample,
        "skipped_global_net_count": skipped_global_net_count,
        "skipped_global_nets_sample": skipped_global_nets_sample,
        "scope_note": scope_note,
    }
    evidence_cards = {
        "nodes": [
            {
                "id": node.get("evidence_id"),
                "type": "llm_topology_node",
                "title": f"{node.get('refdes')} {node.get('role')}",
                "summary": (
                    f"{node.get('refdes')} 位于页码 {node.get('user_visible_page') or ''}，"
                    f"角色={node.get('role') or ''}，接口={', '.join(node.get('interface_groups') or []) or 'misc_signal'}。"
                ),
                "locator": {
                    "refdes": node.get("refdes"),
                    "node_id": node.get("node_id"),
                    "page": node.get("user_visible_page", ""),
                    "review_priority": node.get("review_priority"),
                },
                "risk_tags": node.get("risk_tags", []),
                "llm_device_identity_hint": node.get("llm_device_identity_hint", {}),
                "detail_tool": node.get("detail_tool"),
            }
            for node in node_list
        ],
        "edges": [
            {
                "id": edge.get("evidence_id"),
                "type": "llm_topology_edge",
                "title": edge.get("relation_label") or edge.get("edge_id"),
                "summary": edge.get("summary", ""),
                "locator": {
                    "edge_id": edge.get("edge_id"),
                    "source_refdes": edge.get("source_refdes"),
                    "target_refdes": edge.get("target_refdes"),
                    "review_priority": edge.get("review_priority"),
                    "interface_groups": edge.get("interface_groups", []),
                },
                "risk_tags": edge.get("risk_tags", []),
                "review_focus": edge.get("review_focus", []),
                "review_hints": edge.get("review_hints", []),
                "interface_completeness": edge.get("interface_completeness", []),
                "detail_tool": edge.get("detail_tool"),
            }
            for edge in visible_edges
        ],
        "supply_edges": [
            {
                "id": edge.get("evidence_id"),
                "type": "llm_topology_supply_edge",
                "title": edge.get("relation_label") or edge.get("edge_id"),
                "summary": edge.get("summary", ""),
                "locator": {
                    "edge_id": edge.get("edge_id"),
                    "source_refdes": edge.get("source_refdes"),
                    "target_refdes": edge.get("target_refdes"),
                    "supply_net": edge.get("supply_net"),
                    "voltage_domain": edge.get("voltage_domain"),
                },
                "risk_tags": edge.get("risk_tags", []),
                "review_focus": edge.get("review_focus", []),
                "review_hints": edge.get("review_hints", []),
                "detail_tool": edge.get("detail_tool"),
            }
            for edge in visible_supply_edges
        ],
        "supply_edge_groups": [
            {
                "id": group.get("group_id"),
                "type": "llm_topology_supply_group",
                "title": group.get("relation_label") or group.get("group_id"),
                "summary": group.get("summary", ""),
                "locator": {
                    "group_id": group.get("group_id"),
                    "source_refdes": group.get("source_refdes"),
                    "supply_net": group.get("supply_net"),
                    "voltage_domain": group.get("voltage_domain"),
                    "target_count": group.get("target_count"),
                },
                "risk_tags": group.get("risk_tags", []),
                "review_focus": group.get("review_focus", []),
                "review_hints": group.get("review_hints", []),
                "detail_tool": group.get("detail_tool"),
            }
            for group in visible_supply_edge_groups
        ],
        "interface_groups": interface_groups,
        "review_tasks": [
            {
                "id": task.get("task_id"),
                "type": "llm_topology_review_task",
                "title": task.get("title"),
                "summary": task.get("summary"),
                "locator": {
                    "task_id": task.get("task_id"),
                    "source_kind": task.get("source_kind"),
                    "source_id": task.get("source_id"),
                    "refdes": task.get("refdes", []),
                    "pages": task.get("pages", []),
                    "review_priority": task.get("review_priority"),
                },
                "risk_tags": task.get("risk_tags", []),
                "review_focus": task.get("review_focus", []),
                "detail_tool": task.get("detail_tool"),
            }
            for task in review_tasks[:24]
        ],
    }
    raw_layer = {
        "available": True,
        "detail_tools": ["get_llm_topology_node", "get_llm_topology_edge", "get_topology_review_task"],
        "note": "完整 component 属性、pin/net 明细和原始 net nodes 不直接塞入模型上下文，请通过 detail_tool 二次读取。",
    }

    return {
        "ok": True,
        "schema_version": LLM_TOPOLOGY_SCHEMA_VERSION,
        "project_name": (report or {}).get("project_name", ""),
        "summary": (
            f"识别芯片级节点 {len(nodes)} 个、芯片间信号连接 {len(edges)} 条、供电关系 {len(supply_records)} 条；"
            f"返回 {len(visible_edges)} 条信号连接、{len(visible_supply_edges)} 条供电样本、"
            f"{len(visible_supply_edge_groups)} 个供电聚合组。"
        ),
        "summary_layer": summary_layer,
        "business_view": business_view,
        "review_tasks": review_tasks[:80],
        "review_task_count": len(review_tasks),
        "evidence_cards": evidence_cards,
        "raw_layer": raw_layer,
        "node_count": len(nodes),
        "edge_count": len(edges),
        "supply_edge_count": len(supply_records),
        "supply_group_count": len(supply_edge_groups),
        "returned_edge_count": len(visible_edges),
        "returned_supply_edge_count": len(visible_supply_edges),
        "returned_supply_group_count": len(visible_supply_edge_groups),
        "visual_edge_count": counts.get("visual_edge_count", 0),
        "counts": counts,
        "truncated": topology_truncated,
        "focus_refdes": focus,
        "role_filter": role_filter_text,
        "include_connectors": include_connectors,
        "view": view_mode,
        "supply_mode": supply_display_mode,
        "supply_limit": supply_limit,
        "skipped_power_net_count": skipped_power_net_count,
        "skipped_power_nets_sample": skipped_power_nets_sample,
        "skipped_global_net_count": skipped_global_net_count,
        "skipped_global_nets_sample": skipped_global_nets_sample,
        "nodes": node_list,
        "edges": visible_edges,
        "supply_edges": visible_supply_edges,
        "supply_edge_groups": visible_supply_edge_groups,
        "interface_groups": interface_groups,
        "hubs": hubs,
        "role_links": role_links,
        "scope_note": summary_layer["scope_note"],
    }


def _topology_cache_status(kind: str,
                           *,
                           enabled: bool,
                           status: str,
                           cache_key: str = "",
                           path: str = "",
                           reason: str = "",
                           elapsed_s: float = 0.0) -> dict:
    return {
        "kind": kind,
        "enabled": enabled,
        "status": status,
        "hit": status == "hit",
        "cache_key": cache_key,
        "path": path,
        "reason": reason,
        "elapsed_ms": round(float(elapsed_s) * 1000.0, 3),
    }


def _topology_bundle_digest(bundle: Mapping[str, object] | None) -> str:
    bundle = bundle or {}
    components = bundle.get("components") if isinstance(bundle.get("components"), Mapping) else {}
    nets = bundle.get("nets") if isinstance(bundle.get("nets"), Mapping) else {}
    return _stable_hash({
        "components": components,
        "nets": nets,
        "project_name": bundle.get("project_name", ""),
    })


def _topology_cache_identity(bundle: Mapping[str, object] | None, params: Mapping[str, object]) -> dict:
    bundle = bundle or {}
    project_root = _safe_text(bundle.get("project_root", ""), 500)
    return {
        "schema_version": ANALYSIS_CACHE_SCHEMA_VERSION,
        "analysis_cache_version": ANALYSIS_CACHE_VERSION,
        "cache_version": TOPOLOGY_CACHE_VERSION,
        "kind": TOPOLOGY_CACHE_KIND,
        "schema": LLM_TOPOLOGY_SCHEMA_VERSION,
        "project_root": str(Path(project_root).expanduser()) if project_root else "",
        "bundle_digest": _topology_bundle_digest(bundle),
        "params": dict(params),
    }


def _attach_topology_cache_status(result: dict, status: Mapping[str, object]) -> dict:
    payload = dict(result)
    payload["topology_cache_status"] = dict(status)
    summary_layer = payload.get("summary_layer")
    if isinstance(summary_layer, Mapping):
        updated_summary = dict(summary_layer)
        updated_summary["topology_cache_status"] = dict(status)
        payload["summary_layer"] = updated_summary
    return payload


def _get_or_compute_topology_cache(bundle: Mapping[str, object] | None,
                                   params: Mapping[str, object],
                                   compute) -> dict:
    started = time.perf_counter()
    identity = _topology_cache_identity(bundle, params)
    cache_key = _stable_hash(identity)
    cache_path = analysis_cache_dir() / TOPOLOGY_CACHE_KIND / f"{cache_key}.json"
    if not analysis_cache_enabled():
        result = compute()
        status = _topology_cache_status(
            TOPOLOGY_CACHE_KIND,
            enabled=False,
            status="disabled",
            cache_key=cache_key,
            path=str(cache_path),
            reason=DISABLE_ANALYSIS_CACHE_ENV,
            elapsed_s=time.perf_counter() - started,
        )
        return _attach_topology_cache_status(result, status)

    read_error = ""
    try:
        if cache_path.is_file():
            cached = json.loads(cache_path.read_text(encoding="utf-8"))
            if cached.get("identity") == identity and isinstance(cached.get("result"), dict):
                status = _topology_cache_status(
                    TOPOLOGY_CACHE_KIND,
                    enabled=True,
                    status="hit",
                    cache_key=cache_key,
                    path=str(cache_path),
                    elapsed_s=time.perf_counter() - started,
                )
                return _attach_topology_cache_status(cached["result"], status)
    except Exception as exc:
        read_error = str(exc)

    result = compute()
    write_status = "miss"
    reason = read_error
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(
            json.dumps(
                {
                    "schema_version": ANALYSIS_CACHE_SCHEMA_VERSION,
                    "identity": identity,
                    "result": result,
                    "written_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
                },
                ensure_ascii=False,
                sort_keys=True,
            ),
            encoding="utf-8",
        )
    except Exception as exc:
        write_status = "write_error"
        reason = str(exc)
    status = _topology_cache_status(
        TOPOLOGY_CACHE_KIND,
        enabled=True,
        status=write_status,
        cache_key=cache_key,
        path=str(cache_path),
        reason=reason,
        elapsed_s=time.perf_counter() - started,
    )
    return _attach_topology_cache_status(result, status)


def build_llm_topology_netlist(report: Mapping[str, object] | None,
                               bundle: Mapping[str, object] | None,
                               *,
                               focus_refdes: str = "",
                               role_filter: str = "",
                               include_connectors: bool = False,
                               limit: int = DEFAULT_LIMIT,
                               return_all_edges: bool = False,
                               view: str = "summary",
                               supply_mode: str = "",
                               supply_limit: int = DEFAULT_SUPPLY_LIMIT,
                               use_cache: bool = True) -> dict:
    """Build the stable LLM-facing topology netlist artifact."""
    view_mode = _normalize_topology_view(view, return_all_edges)
    supply_display_mode = _normalize_supply_mode(supply_mode, view_mode, return_all_edges)
    normalized_limit = _as_limit(limit)
    normalized_supply_limit = _as_supply_limit(supply_limit)
    params = {
        "focus_refdes": str(focus_refdes or "").strip().upper(),
        "role_filter": str(role_filter or "").strip().lower(),
        "include_connectors": bool(include_connectors),
        "limit": normalized_limit,
        "return_all_edges": bool(return_all_edges),
        "view": view_mode,
        "supply_mode": supply_display_mode,
        "supply_limit": normalized_supply_limit,
    }

    def compute() -> dict:
        return build_chip_topology(
            report,
            bundle,
            focus_refdes=focus_refdes,
            role_filter=role_filter,
            include_connectors=include_connectors,
            limit=normalized_limit,
            return_all_edges=return_all_edges,
            view=view_mode,
            supply_mode=supply_display_mode,
            supply_limit=normalized_supply_limit,
        )

    if use_cache:
        return _get_or_compute_topology_cache(bundle, params, compute)
    status = _topology_cache_status(
        TOPOLOGY_CACHE_KIND,
        enabled=False,
        status="bypass",
        reason="use_cache=False",
    )
    return _attach_topology_cache_status(compute(), status)


def get_llm_topology_node(report: Mapping[str, object] | None,
                          bundle: Mapping[str, object] | None,
                          refdes: str,
                          *,
                          include_connectors: bool = False,
                          max_pin_nets: int = 240) -> dict:
    bundle = bundle or {}
    components = bundle.get("components") if isinstance(bundle.get("components"), Mapping) else {}
    raw_nets = bundle.get("nets") if isinstance(bundle.get("nets"), Mapping) else {}
    nets = _derive_nets_from_components(components, raw_nets)
    target = str(refdes or "").strip()
    if not target:
        return {"ok": False, "summary": "缺少 refdes。", "refdes": refdes}
    comp = None
    matched_ref = ""
    for candidate, value in components.items():
        if str(candidate).upper() == target.upper() and isinstance(value, Mapping):
            comp = value
            matched_ref = str(candidate)
            break
    if comp is None:
        return {"ok": False, "summary": f"未找到拓扑节点：{target}", "refdes": target}
    pin_nets = _iter_component_pin_nets(matched_ref, comp, nets)
    node = _node_preview(matched_ref, comp, pin_nets, include_connectors=include_connectors)
    if not node:
        return {"ok": False, "summary": f"{target} 不是第一版 LLM 拓扑主节点。", "refdes": target}
    topology = build_llm_topology_netlist(
        report,
        bundle,
        focus_refdes=matched_ref,
        include_connectors=include_connectors,
        limit=MAX_LIMIT,
        return_all_edges=True,
    )
    raw_net_nodes = []
    for row in pin_nets[:max_pin_nets]:
        net_name = str(row.get("net") or "")
        raw_net_nodes.append({
            "net": net_name,
            "pin": row.get("pin", ""),
            "pin_name": row.get("pin_name", ""),
            "nodes": [
                {
                    "refdes": _safe_text(node_item.get("refdes", ""), 80),
                    "pin": _safe_text(node_item.get("pin", ""), 80),
                    "pin_name": _safe_text(node_item.get("pin_name", ""), 160),
                }
                for node_item in list(nets.get(net_name, []) or [])[:MAX_NET_NODES]
                if isinstance(node_item, Mapping)
            ],
        })
    return {
        "ok": True,
        "schema_version": LLM_TOPOLOGY_SCHEMA_VERSION,
        "summary": (
            f"{matched_ref} LLM 拓扑节点详情：pin/net {len(pin_nets)} 条，"
            f"信号边 {len(topology.get('edges', []) or [])} 条，供电关系 {len(topology.get('supply_edges', []) or [])} 条。"
        ),
        "node": node,
        "component": dict(comp),
        "pin_nets": pin_nets[:max_pin_nets],
        "pin_nets_truncated": len(pin_nets) > max_pin_nets,
        "raw_net_nodes": raw_net_nodes,
        "edges": topology.get("edges", []),
        "supply_edges": topology.get("supply_edges", []),
        "scope_note": topology.get("scope_note", ""),
    }


def get_llm_topology_edge(report: Mapping[str, object] | None,
                          bundle: Mapping[str, object] | None,
                          edge_id: str,
                          *,
                          include_connectors: bool = False) -> dict:
    result = get_chip_topology_edge(
        report,
        bundle,
        edge_id,
        include_connectors=include_connectors,
    )
    if result.get("ok"):
        result["schema_version"] = LLM_TOPOLOGY_SCHEMA_VERSION
    return result


def query_llm_topology_netlist(report: Mapping[str, object] | None,
                               bundle: Mapping[str, object] | None,
                               query: str,
                               *,
                               include_connectors: bool = False,
                               limit: int = DEFAULT_LIMIT) -> dict:
    result = query_chip_topology(
        report,
        bundle,
        query,
        include_connectors=include_connectors,
        limit=limit,
    )
    result["schema_version"] = LLM_TOPOLOGY_SCHEMA_VERSION
    return result


def batch_query_llm_topology_netlist(report: Mapping[str, object] | None,
                                     bundle: Mapping[str, object] | None,
                                     queries: Sequence[object],
                                     *,
                                     include_connectors: bool = False,
                                     limit_per_query: int = 8) -> dict:
    result = batch_query_chip_topology(
        report,
        bundle,
        queries,
        include_connectors=include_connectors,
        limit_per_query=limit_per_query,
    )
    result["schema_version"] = LLM_TOPOLOGY_SCHEMA_VERSION
    return result


def summarize_topology_review_tasks(report: Mapping[str, object] | None,
                                    bundle: Mapping[str, object] | None,
                                    *,
                                    include_connectors: bool = False,
                                    focus_refdes: str = "",
                                    interface_group: str = "",
                                    priority: str = "",
                                    limit: int = DEFAULT_LIMIT) -> dict:
    """Return the prioritized topology review queue as lightweight evidence cards."""
    topology = build_llm_topology_netlist(
        report,
        bundle,
        include_connectors=include_connectors,
        limit=MAX_LIMIT,
        return_all_edges=True,
    )
    tasks = list(topology.get("review_tasks") or [])
    focus = str(focus_refdes or "").strip().upper()
    group_filter = str(interface_group or "").strip().lower()
    priority_filter = str(priority or "").strip().lower()
    if focus:
        tasks = [
            task for task in tasks
            if focus in {str(item or "").upper() for item in task.get("refdes", []) or []}
        ]
    if group_filter:
        tasks = [
            task for task in tasks
            if group_filter in {str(item or "").lower() for item in task.get("interface_groups", []) or []}
        ]
    if priority_filter:
        tasks = [
            task for task in tasks
            if str(task.get("review_priority") or "").lower() == priority_filter
        ]
    limit = _as_limit(limit)
    selected = tasks[:limit]
    return {
        "ok": True,
        "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
        "summary": f"拓扑 review 队列共 {len(tasks)} 项，返回 {len(selected)} 项；用于先排查高风险芯片/接口关系。",
        "total_count": len(tasks),
        "returned_count": len(selected),
        "truncated": len(tasks) > len(selected),
        "filters": {
            "include_connectors": include_connectors,
            "focus_refdes": focus,
            "interface_group": group_filter,
            "priority": priority_filter,
            "limit": limit,
        },
        "tasks": selected,
        "scope_note": topology.get("scope_note", ""),
        "detail_tool": {"name": "get_topology_review_task", "args": {"task_id": selected[0]["task_id"]}} if selected else None,
    }


def get_topology_review_task(report: Mapping[str, object] | None,
                             bundle: Mapping[str, object] | None,
                             task_id: str,
                             *,
                             include_connectors: bool = False) -> dict:
    """Read one topology review task plus the related raw topology detail."""
    target = str(task_id or "").strip().lower()
    if not target:
        return {"ok": False, "summary": "缺少 topology review task_id。", "task_id": task_id}
    topology = build_llm_topology_netlist(
        report,
        bundle,
        include_connectors=include_connectors,
        limit=MAX_LIMIT,
        return_all_edges=True,
    )
    for task in topology.get("review_tasks", []) or []:
        if str(task.get("task_id") or "").lower() != target:
            continue
        source_kind = str(task.get("source_kind") or "")
        source_id = str(task.get("source_id") or "")
        related_detail: dict = {}
        if source_kind == "node":
            refdes_items = [str(item or "") for item in task.get("refdes", []) or [] if str(item or "")]
            refdes = refdes_items[0] if refdes_items else source_id.replace("chip-node-", "")
            related_detail = get_llm_topology_node(
                report,
                bundle,
                refdes,
                include_connectors=include_connectors,
                max_pin_nets=320,
            )
        elif source_kind in {"signal_edge", "supply_edge"}:
            related_detail = get_llm_topology_edge(
                report,
                bundle,
                source_id,
                include_connectors=include_connectors,
            )
        return {
            "ok": True,
            "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
            "summary": task.get("summary") or f"拓扑 review task：{task.get('title') or task_id}",
            "task": task,
            "related_detail": related_detail,
            "scope_note": topology.get("scope_note", ""),
        }
    return {
        "ok": False,
        "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
        "summary": f"未找到拓扑 review task：{task_id}",
        "task_id": task_id,
    }


def batch_expand_topology_review_tasks(report: Mapping[str, object] | None,
                                       bundle: Mapping[str, object] | None,
                                       task_ids: Sequence[object],
                                       *,
                                       include_connectors: bool = False) -> dict:
    normalized = [_safe_text(item, 180) for item in list(task_ids or [])[:MAX_QUERY_ITEMS] if _safe_text(item, 180)]
    items: List[dict] = []
    found = 0
    for task_id in normalized:
        detail = get_topology_review_task(
            report,
            bundle,
            task_id,
            include_connectors=include_connectors,
        )
        status = "found" if detail.get("ok") else "missing"
        if status == "found":
            found += 1
        task = detail.get("task") if isinstance(detail.get("task"), Mapping) else {}
        items.append({
            "task_id": task_id,
            "status": status,
            "summary": detail.get("summary", ""),
            "review_priority": task.get("review_priority", ""),
            "evidence_id": task.get("evidence_id", ""),
            "task": task,
            "related_detail": detail.get("related_detail") if status == "found" else {},
            "missing_reason": "" if status == "found" else detail.get("summary", "未找到该 review task。"),
        })
    return {
        "ok": True,
        "schema_version": LLM_TOPOLOGY_REVIEW_TASK_SCHEMA_VERSION,
        "summary": f"批量读取拓扑 review task {len(normalized)} 项，命中 {found} 项，缺失 {len(normalized) - found} 项。",
        "query_count": len(normalized),
        "found_count": found,
        "missing_count": len(normalized) - found,
        "truncated": len(list(task_ids or [])) > len(normalized),
        "items": items,
        "readonly": True,
    }


def get_chip_topology_edge(report: Mapping[str, object] | None,
                           bundle: Mapping[str, object] | None,
                           edge_id: str,
                           *,
                           include_connectors: bool = False) -> dict:
    topology = build_llm_topology_netlist(
        report,
        bundle,
        include_connectors=include_connectors,
        limit=MAX_LIMIT,
        return_all_edges=True,
    )
    target = str(edge_id or "").strip().lower()
    for edge in [*(topology.get("edges", []) or []), *(topology.get("supply_edges", []) or [])]:
        if str(edge.get("edge_id") or "").lower() == target:
            return {
                "ok": True,
                "summary": edge.get("summary") or "",
                "edge": edge,
                "scope_note": topology.get("scope_note", ""),
            }
    return {
        "ok": False,
        "summary": f"未找到芯片级拓扑连接：{edge_id}",
        "edge_id": edge_id,
    }


def _matches_query(text: str, query: str) -> bool:
    normalized = text.lower()
    stopwords = {"和", "与", "到", "连接", "关系", "芯片", "芯片级", "拓扑", "的", "了", "哪些", "有关"}
    tokens = [
        token for token in re.findall(r"[0-9A-Za-z_\u4e00-\u9fff.+-]+", query or "")
        if token.lower() not in stopwords and token not in stopwords
    ]
    return bool(tokens) and all(token.lower() in normalized for token in tokens)


def _query_variants(query: str) -> List[str]:
    text = str(query or "").strip()
    if not text:
        return []
    dictionary = load_business_dictionary()
    alias_maps = {
        **(dictionary.get("interface_aliases") or {}),
        **(dictionary.get("role_aliases") or {}),
    }
    variants = [text]
    tokens = re.findall(r"[0-9A-Za-z_\u4e00-\u9fff.+-]+", text)
    normalized_tokens = []
    changed = False
    for token in tokens:
        token_upper = token.upper()
        replacement = token
        for canonical, aliases in alias_maps.items():
            alias_set = {str(alias).upper() for alias in aliases or []}
            alias_set.add(str(canonical).upper())
            if token_upper in alias_set:
                replacement = str(canonical)
                changed = True
                break
        normalized_tokens.append(replacement)
    if changed:
        variants.append(" ".join(normalized_tokens))
    for canonical, aliases in alias_maps.items():
        haystack = f" {text.upper()} "
        if str(canonical).upper() in haystack or any(str(alias).upper() in haystack for alias in aliases or []):
            variants.append(str(canonical))
    return list(dict.fromkeys(variant for variant in variants if variant))


def _matches_query_variants(text: str, query: str) -> bool:
    return any(_matches_query(text, variant) for variant in _query_variants(query))


def query_chip_topology(report: Mapping[str, object] | None,
                        bundle: Mapping[str, object] | None,
                        query: str,
                        *,
                        include_connectors: bool = False,
                        limit: int = DEFAULT_LIMIT) -> dict:
    topology = build_llm_topology_netlist(
        report,
        bundle,
        include_connectors=include_connectors,
        limit=MAX_LIMIT,
        return_all_edges=True,
    )
    query_text = str(query or "").strip()
    limit = _as_limit(limit)
    matches: List[dict] = []
    if query_text:
        for edge in topology.get("edges", []):
            haystack = " ".join([
                str(edge.get("edge_id") or ""),
                str(edge.get("source_refdes") or ""),
                str(edge.get("target_refdes") or ""),
                str(edge.get("source_role") or ""),
                str(edge.get("target_role") or ""),
                str(edge.get("relation_label") or ""),
                str(edge.get("summary") or ""),
                " ".join(edge.get("review_hints") or []),
                " ".join(edge.get("source_power_nets") or []),
                " ".join(edge.get("target_power_nets") or []),
                " ".join(str(item.get("net") or "") for item in edge.get("shared_nets") or []),
                " ".join(" ".join(str(net) for net in bridge.get("nets") or []) for bridge in edge.get("passive_bridges") or []),
                " ".join(edge.get("interface_groups") or []),
            ])
            if _matches_query_variants(haystack, query_text):
                matches.append({"kind": "edge", "edge": edge, "summary": edge.get("summary", "")})
        for edge in topology.get("supply_edges", []):
            haystack = " ".join([
                str(edge.get("edge_id") or ""),
                str(edge.get("source_refdes") or ""),
                str(edge.get("target_refdes") or ""),
                str(edge.get("supply_net") or ""),
                str(edge.get("voltage_domain") or ""),
                str(edge.get("relation_label") or ""),
                str(edge.get("summary") or ""),
                " ".join(edge.get("risk_tags") or []),
            ])
            if _matches_query_variants(haystack, query_text):
                matches.append({"kind": "supply_edge", "edge": edge, "summary": edge.get("summary", "")})
        for node in topology.get("nodes", []):
            identity_hint = node.get("llm_device_identity_hint") if isinstance(node.get("llm_device_identity_hint"), Mapping) else {}
            haystack = " ".join(str(node.get(key) or "") for key in (
                "node_id",
                "refdes",
                "category",
                "role",
                "candidate_chip_type",
                "hq_no",
                "spec",
                "user_visible_page",
            ))
            haystack = (
                f"{haystack} {' '.join(node.get('signal_net_preview') or [])} {' '.join(node.get('interface_groups') or [])} "
                f"{identity_hint.get('part_name', '')} {identity_hint.get('spec', '')} {' '.join(identity_hint.get('tokens') or [])}"
            )
            if _matches_query_variants(haystack, query_text):
                matches.append({"kind": "node", "node": node, "summary": f"{node.get('refdes')} {node.get('role')} 芯片节点。"})
        for task in topology.get("review_tasks", []) or []:
            haystack = " ".join([
                str(task.get("task_id") or ""),
                str(task.get("title") or ""),
                str(task.get("summary") or ""),
                " ".join(str(item or "") for item in task.get("refdes", []) or []),
                " ".join(str(item or "") for item in task.get("interface_groups", []) or []),
                " ".join(str(item or "") for item in task.get("risk_tags", []) or []),
                " ".join(str(item or "") for item in task.get("review_focus", []) or []),
                " ".join(str(item or "") for item in task.get("missing_signals", []) or []),
            ])
            if _matches_query_variants(haystack, query_text):
                matches.append({"kind": "review_task", "review_task": task, "summary": task.get("summary", "")})

    selected = matches[:limit]
    return {
        "ok": True,
        "query": _safe_text(query_text, 200),
        "summary": f"芯片级拓扑查询 `{query_text}` 命中 {len(matches)} 项，返回 {len(selected)} 项。",
        "total_matches": len(matches),
        "limit": limit,
        "truncated": len(matches) > len(selected),
        "items": selected,
        "scope_note": topology.get("scope_note", ""),
    }


def batch_query_chip_topology(report: Mapping[str, object] | None,
                              bundle: Mapping[str, object] | None,
                              queries: Sequence[object],
                              *,
                              include_connectors: bool = False,
                              limit_per_query: int = 8) -> dict:
    normalized_queries = [_safe_text(item, 200) for item in list(queries or [])[:MAX_QUERY_ITEMS] if _safe_text(item, 200)]
    limit_per_query = max(1, min(int(limit_per_query or 8), 20))
    items = []
    found = 0
    for query in normalized_queries:
        result = query_chip_topology(
            report,
            bundle,
            query,
            include_connectors=include_connectors,
            limit=limit_per_query,
        )
        status = "found" if result.get("total_matches", 0) else "missing"
        if status == "found":
            found += 1
        items.append({
            "query": query,
            "status": status,
            "summary": result.get("summary", ""),
            "total_matches": result.get("total_matches", 0),
            "items": result.get("items", []),
            "missing_reason": "" if status == "found" else "芯片级拓扑中未找到该关键词对应的芯片节点或芯片间连接。",
            "truncated": bool(result.get("truncated")),
        })
    return {
        "ok": True,
        "summary": f"批量芯片级拓扑查询 {len(normalized_queries)} 项，命中 {found} 项。",
        "query_count": len(normalized_queries),
        "found_count": found,
        "missing_count": len(normalized_queries) - found,
        "truncated": len(queries or []) > len(normalized_queries),
        "items": items,
    }
