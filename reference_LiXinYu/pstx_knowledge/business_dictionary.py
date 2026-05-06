# -*- coding: utf-8 -*-
"""Project business glossary and abbreviation dictionary.

This module is intentionally lightweight and deterministic.  It gives the
topology/harness layer a shared place for project-specific abbreviations such
as PCE/P5E -> PCIe, without making the model guess these names from scratch.
"""

from __future__ import annotations

from functools import lru_cache
import json
import os
import re
from pathlib import Path
from typing import Dict, Iterable, List, Mapping


BUILTIN_INTERFACE_ALIASES: Dict[str, List[str]] = {
    "pcie": ["PCIE", "PCI-E", "PCI_EXPRESS", "PCI", "PCE", "P5E", "P4E", "P3E", "P2E", "P1E"],
    "i2c": ["I2C", "IIC", "SCL", "SDA"],
    "spi": ["SPI", "MOSI", "MISO", "SCLK", "CS", "CSN", "NCS"],
    "uart": ["UART", "TXD", "RXD", "CTS", "RTS"],
    "usb": ["USB", "USBP", "USBN", "UDP", "UDM", "DPLUS", "DMINUS"],
    "mipi_lvds": ["MIPI", "LVDS", "CSI", "DSI"],
    "ethernet": ["ETH", "ETHERNET", "网口", "以太网", "RGMII", "SGMII", "GMII", "MII", "MDIO", "MDC"],
    "storage_sdio": ["SDIO", "SDMMC", "EMMC", "MMC", "SDCLK", "SDCMD"],
    "jtag_debug": ["JTAG", "TCK", "TMS", "TDI", "TDO", "TRST", "DBG", "DEBUG"],
    "audio": ["I2S", "IIS", "MCLK", "BCLK", "LRCLK", "TDM", "PDM", "SPDIF"],
    "analog_sense": ["ADC", "DAC", "AIN", "AOUT", "SENSE", "FB"],
    "clock": ["CLK", "CLOCK", "OSC", "XO", "REFCLK"],
    "reset": ["RESET", "RST", "RESETN", "RSTN", "POR"],
    "interrupt": ["INT", "IRQ", "ALERT", "FAULT", "NMI"],
    "power_control": ["EN", "ENABLE", "PWR", "PWRON", "PGOOD", "POWERGOOD", "PWREN"],
    "gpio": ["GPIO"],
}

BUILTIN_ROLE_ALIASES: Dict[str, List[str]] = {
    "processor_or_fpga": ["FPGA", "CPU", "GPU", "SOC", "MCU", "DSP", "LCMXO", "XILINX", "ZYNQ"],
    "level_shifter": ["LEVEL SHIFT", "LEVEL SHIFTER", "TRANSLATOR", "TRANSCEIVER", "TXS", "TXB", "LSF", "SN74", "电平转换", "电平转换器"],
    "power_management_ic": ["PMIC", "BUCK", "BOOST", "LDO", "REGULATOR", "POWER", "CHARGER", "DCDC"],
    "memory": ["DDR", "LPDDR", "EMMC", "NAND", "NOR", "FLASH", "SDRAM"],
    "clock_source": ["CLOCK", "CLK", "OSC", "XO", "晶振"],
}

BUILTIN_REVIEW_FOCUS: Dict[str, List[str]] = {
    "pcie": ["差分阻抗", "AC 耦合", "REFCLK", "PERST#", "电源时序"],
    "i2c": ["上拉电压", "重复上拉", "总线电容", "跨电平转换"],
    "spi": ["串阻", "片选默认态", "时钟边沿", "电压域"],
    "clock": ["串阻/端接", "扇出", "使能", "跨页网名一致性"],
    "reset": ["默认态", "释放时序", "上拉/下拉", "跨域复位"],
    "power_control": ["默认态", "EN/PGOOD", "上电顺序", "上下拉"],
    "jtag_debug": ["上下拉", "量产可访问性", "测试点", "复用状态"],
    "ethernet": ["RGMII/SGMII 电压域", "时钟", "复位", "PHY strap"],
    "storage_sdio": ["上拉", "位宽", "时钟串阻", "电压切换"],
    "analog_sense": ["量程", "滤波", "地参考", "反馈路径"],
}


def _clean_alias(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().upper())


def _merge_alias_map(base: Mapping[str, Iterable[object]], extra: Mapping[str, Iterable[object]] | None) -> Dict[str, List[str]]:
    merged: Dict[str, List[str]] = {
        str(key): list(dict.fromkeys(_clean_alias(item) for item in values if _clean_alias(item)))
        for key, values in (base or {}).items()
    }
    for key, values in (extra or {}).items():
        if not isinstance(values, list):
            continue
        current = merged.setdefault(str(key), [])
        for value in values:
            alias = _clean_alias(value)
            if alias and alias not in current:
                current.append(alias)
    return merged


def _load_external_dictionary() -> tuple[dict, List[str]]:
    path_text = os.environ.get("PSTX_BUSINESS_DICTIONARY_FILE", "").strip()
    if not path_text:
        return {}, []
    path = Path(path_text).expanduser()
    if not path.is_file():
        return {}, [f"业务词典文件不存在：{path}"]
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # pragma: no cover - defensive diagnostics
        return {}, [f"业务词典文件读取失败：{exc}"]
    if not isinstance(data, dict):
        return {}, ["业务词典文件根节点必须是 JSON object。"]
    return data, []


@lru_cache(maxsize=1)
def load_business_dictionary() -> dict:
    external, warnings = _load_external_dictionary()
    return {
        "schema_version": "pstx-business-dictionary.v1",
        "source": "builtin+external" if external else "builtin",
        "external_file": os.environ.get("PSTX_BUSINESS_DICTIONARY_FILE", "").strip(),
        "warnings": warnings,
        "interface_aliases": _merge_alias_map(BUILTIN_INTERFACE_ALIASES, external.get("interface_aliases") if isinstance(external, dict) else {}),
        "role_aliases": _merge_alias_map(BUILTIN_ROLE_ALIASES, external.get("role_aliases") if isinstance(external, dict) else {}),
        "review_focus": _merge_alias_map(BUILTIN_REVIEW_FOCUS, external.get("review_focus") if isinstance(external, dict) else {}),
    }


def business_dictionary_summary() -> dict:
    dictionary = load_business_dictionary()
    interface_aliases = dictionary.get("interface_aliases") or {}
    return {
        "schema_version": dictionary.get("schema_version"),
        "source": dictionary.get("source"),
        "external_file": dictionary.get("external_file"),
        "warning_count": len(dictionary.get("warnings") or []),
        "warnings": list(dictionary.get("warnings") or [])[:6],
        "interface_count": len(interface_aliases),
        "interface_aliases": {
            key: list(values)[:12]
            for key, values in sorted(interface_aliases.items())
        },
        "role_aliases": {
            key: list(values)[:12]
            for key, values in sorted((dictionary.get("role_aliases") or {}).items())
        },
        "review_focus": {
            key: list(values)[:10]
            for key, values in sorted((dictionary.get("review_focus") or {}).items())
        },
    }


def interface_aliases(interface_id: str) -> List[str]:
    aliases = (load_business_dictionary().get("interface_aliases") or {}).get(str(interface_id), [])
    return list(aliases)


def interface_alias_snapshot(limit_per_group: int = 8) -> dict:
    aliases = load_business_dictionary().get("interface_aliases") or {}
    return {
        key: list(values)[:limit_per_group]
        for key, values in sorted(aliases.items())
    }


def review_focus_for_interface(interface_id: str) -> List[str]:
    focus = (load_business_dictionary().get("review_focus") or {}).get(str(interface_id), [])
    return list(focus)


def clear_business_dictionary_cache() -> None:
    load_business_dictionary.cache_clear()
