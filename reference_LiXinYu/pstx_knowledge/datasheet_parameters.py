# -*- coding: utf-8 -*-
"""Deterministic parameter extraction for datasheet chunks."""

from __future__ import annotations

import html
import re
from typing import Iterable, List, Optional


PACKAGE_TOKENS = (
    "SOT-223",
    "TO-252",
    "DPAK",
    "SO-8",
    "SOIC",
    "QFN",
    "TQFP",
    "BGA",
    "LQFP",
    "DFN",
)


def _clean_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _compact_spaced_digits(match: re.Match) -> str:
    return re.sub(r"\s+", "", match.group(1))


def normalize_parameter_text(text: str) -> str:
    """Convert MinerU/pypdf text and simple HTML tables into searchable text."""

    content = html.unescape(str(text or ""))
    content = re.sub(r"</t[dh]>\s*<t[dh][^>]*>", " | ", content, flags=re.I)
    content = re.sub(r"</tr>\s*<tr[^>]*>", "\n", content, flags=re.I)
    content = re.sub(r"<br\s*/?>", "\n", content, flags=re.I)
    content = re.sub(r"</?(?:table|tbody|thead|tr|td|th)[^>]*>", " ", content, flags=re.I)
    content = re.sub(r"<[^>]+>", " ", content)
    content = content.replace("\\circ", "°").replace("\\Omega", "Ω")
    content = re.sub(r"\\(?:mathrm|tiny|text|rm)\s*\{([^{}]*)\}", r"\1", content)
    content = re.sub(r"[{}$]", " ", content)
    content = re.sub(
        r"((?:\d\s+){1,5}\d)\s*\^?\s*°\s*C\s*/\s*W",
        lambda match: f"{_compact_spaced_digits(match)}°C/W",
        content,
        flags=re.I,
    )
    content = re.sub(
        r"((?:\d\s+){1,5}\d)\s*\^?\s*°\s*C",
        lambda match: f"{_compact_spaced_digits(match)}°C",
        content,
        flags=re.I,
    )
    content = content.replace("°°C", "°C")
    content = content.replace("µ F", "µF").replace("μ F", "μF").replace("u F", "uF")
    lines = [_clean_spaces(line) for line in content.splitlines()]
    return "\n".join(line for line in lines if line)


def _excerpt(text: str, start: int, end: int, *, pad: int = 140, limit: int = 420) -> str:
    left = max(0, start - pad)
    right = min(len(text), end + pad)
    snippet = _clean_spaces(text[left:right])
    return snippet[:limit]


def _first_number(value: str) -> Optional[float]:
    match = re.search(r"[-+]?\d+(?:\.\d+)?", str(value or ""))
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _parse_number_cell(value: str, *, voltage: bool = False) -> Optional[float]:
    text = _clean_spaces(value)
    if not text or text.upper() in {"TBD", "NA", "N/A", "-", "■"}:
        return None
    # Avoid treating OCR fragments like "C.146" as 146.  These cells are
    # ambiguous enough that a parameter card should stay incomplete.
    if re.match(r"^[A-Za-z]\.\d+", text):
        return None
    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if not match:
        return None
    raw = match.group(0)
    if voltage and "." not in raw and re.match(r"^0\d{3,5}$", raw):
        raw = "0." + raw[1:]
    try:
        number = float(raw)
    except ValueError:
        return None
    # Common OCR artifact in narrow voltage tables: "10.75" for "0.75".
    if voltage and 10.0 <= number < 11.0:
        number -= 10.0
    return number


def _split_cells(line: str) -> List[str]:
    cells = [_clean_spaces(cell) for cell in str(line or "").split("|")]
    while cells and not cells[0]:
        cells.pop(0)
    while cells and not cells[-1]:
        cells.pop()
    return cells


def _normalize_signal_name(value: str) -> str:
    text = _clean_spaces(value).replace(" ", "")
    if text.upper() == "VDDI0":
        return "VDDIO"
    return text


def _normalize_theta_symbol(value: str) -> str:
    text = _clean_spaces(value).replace(" ", "")
    if text.startswith("0J"):
        return "θ" + text[1:]
    return text


def _range_value_text(value_min: Optional[float],
                      value_typ: Optional[float],
                      value_max: Optional[float],
                      unit: str = "") -> str:
    parts = []
    if value_min is not None:
        parts.append(f"min {value_min:g}{unit}")
    if value_typ is not None:
        parts.append(f"typ {value_typ:g}{unit}")
    if value_max is not None:
        parts.append(f"max {value_max:g}{unit}")
    return ", ".join(parts)


def _unit(value: str) -> str:
    match = re.search(r"(°C/W|C/W|µF|μF|uF|mV|ms|us|µs|μs|ns|s|V|mA|A|%|°C|C|Ω|ohm)", str(value or ""), re.I)
    if not match:
        return ""
    unit = match.group(1)
    return {"μF": "uF", "µF": "uF", "μs": "us", "µs": "us", "C/W": "°C/W", "C": "°C", "ohm": "Ω"}.get(unit, unit)


def _card(parameter_key: str,
          parameter_name: str,
          value_text: str,
          *,
          page: int,
          chunk_id: str,
          condition: str = "",
          value_min: Optional[float] = None,
          value_typ: Optional[float] = None,
          value_max: Optional[float] = None,
          unit: str = "",
          source_text: str = "",
          confidence: str = "medium",
          extraction_method: str = "regex_v1") -> dict:
    return {
        "parameter_key": parameter_key,
        "parameter_name": parameter_name,
        "value_text": _clean_spaces(value_text).replace("°°C", "°C"),
        "value_min": value_min,
        "value_typ": value_typ,
        "value_max": value_max,
        "unit": unit,
        "condition": _clean_spaces(condition),
        "page": int(page or 1),
        "chunk_id": str(chunk_id or ""),
        "source_text": _clean_spaces(source_text),
        "confidence": confidence,
        "extraction_method": extraction_method,
    }


def _add_numeric(cards: List[dict],
                 text: str,
                 match: re.Match,
                 *,
                 key: str,
                 name: str,
                 page: int,
                 chunk_id: str,
                 relation: str = "typ",
                 condition: str = "",
                 confidence: str = "high") -> None:
    value_text = _clean_spaces(match.group(1))
    number = _first_number(value_text)
    unit = _unit(value_text)
    values = {"value_min": None, "value_typ": None, "value_max": None}
    if relation == "max":
        values["value_max"] = number
    elif relation == "min":
        values["value_min"] = number
    else:
        values["value_typ"] = number
    cards.append(_card(
        key,
        name,
        value_text,
        page=page,
        chunk_id=chunk_id,
        condition=condition,
        unit=unit,
        source_text=_excerpt(text, match.start(), match.end()),
        confidence=confidence,
        **values,
    ))


def _extract_common_ldo_parameters(text: str, *, page: int, chunk_id: str) -> List[dict]:
    cards: List[dict] = []

    for match in re.finditer(r"Output Current(?:\s+of|\s*[:=])?\s*([0-9]+(?:\.[0-9]+)?\s*A)\b", text, re.I):
        _add_numeric(cards, text, match, key="output_current", name="Output Current", page=page, chunk_id=chunk_id, relation="typ")

    for match in re.finditer(r"Input Voltage\s*(?:\||:)?\s*([0-9]+(?:\.[0-9]+)?\s*V)\b", text, re.I):
        _add_numeric(cards, text, match, key="absolute_max_input_voltage", name="Input Voltage", page=page, chunk_id=chunk_id, relation="max", condition="Absolute Maximum Ratings")

    for match in re.finditer(r"Line Regulation\s*[:|]?\s*([0-9]+(?:\.[0-9]+)?\s*%)\s*(?:Max(?:imum)?\.?)?", text, re.I):
        _add_numeric(cards, text, match, key="line_regulation", name="Line Regulation", page=page, chunk_id=chunk_id, relation="max")

    for match in re.finditer(r"Load Regulation\s*[:|]?\s*([0-9]+(?:\.[0-9]+)?\s*%)\s*(?:Max(?:imum)?\.?)?", text, re.I):
        _add_numeric(cards, text, match, key="load_regulation", name="Load Regulation", page=page, chunk_id=chunk_id, relation="max")

    for match in re.finditer(r"Operates?\s+Down\s+to\s+([0-9]+(?:\.[0-9]+)?\s*V)\s+Dropout", text, re.I):
        _add_numeric(cards, text, match, key="dropout_voltage_operating", name="Dropout Voltage", page=page, chunk_id=chunk_id, relation="typ", condition="Operates down to dropout")

    for match in re.finditer(r"dropout voltage[^.\n]{0,120}?(?:maximum|max)\s*([0-9]+(?:\.[0-9]+)?\s*V)\b", text, re.I):
        _add_numeric(cards, text, match, key="dropout_voltage_max", name="Dropout Voltage", page=page, chunk_id=chunk_id, relation="max", condition="Guaranteed maximum")

    for match in re.finditer(r"(?:output capacitor|addition of)[^.\n]{0,140}?([0-9]+(?:\.[0-9]+)?\s*(?:µF|μF|uF))", text, re.I):
        _add_numeric(cards, text, match, key="output_capacitor", name="Output Capacitor", page=page, chunk_id=chunk_id, relation="min", condition="Stability")

    for match in re.finditer(r"Lead Temperature.{0,120}?([0-9]{2,3}\s*°?\s*C)", text, re.I):
        _add_numeric(cards, text, match, key="lead_temperature", name="Lead Temperature", page=page, chunk_id=chunk_id, relation="max", condition="Soldering information")

    fixed_match = re.search(r"Fixed Voltages?\\?\*?\s*((?:[0-9]+(?:\.[0-9]+)?\s*V(?:\s*(?:,|and)\s*)?){2,})", text, re.I)
    if fixed_match:
        voltages = re.findall(r"[0-9]+(?:\.[0-9]+)?\s*V", fixed_match.group(1), re.I)
        if voltages:
            cards.append(_card(
                "fixed_output_voltages",
                "Fixed Output Voltages",
                ", ".join(voltages),
                page=page,
                chunk_id=chunk_id,
                unit="V",
                source_text=_excerpt(text, fixed_match.start(), fixed_match.end()),
                confidence="high",
            ))

    packages = [token for token in PACKAGE_TOKENS if re.search(re.escape(token), text, re.I)]
    if packages:
        cards.append(_card(
            "packages",
            "Package Options",
            ", ".join(dict.fromkeys(packages)),
            page=page,
            chunk_id=chunk_id,
            source_text=_excerpt(text, 0, min(len(text), 1)),
            confidence="medium",
        ))

    thermal_conditions = set()
    package_pattern = re.compile("|".join(re.escape(token) for token in PACKAGE_TOKENS), re.I)
    thermal_value_pattern = re.compile(r"([0-9]{2,3}\s*°?\s*C\s*/\s*W)", re.I)
    for line in text.splitlines():
        if "package" not in line.lower() or "C/W" not in line:
            continue
        package_matches = list(package_pattern.finditer(line))
        values = [_clean_spaces(match.group(1)).replace("°°C", "°C").replace("C / W", "°C/W").replace("C/W", "°C/W") for match in thermal_value_pattern.finditer(line)]
        if not package_matches or not values:
            continue
        ordered_packages = []
        for match in package_matches:
            package = match.group(0).upper()
            canonical = next((token for token in PACKAGE_TOKENS if token.upper() == package), match.group(0))
            if canonical not in ordered_packages:
                ordered_packages.append(canonical)
        if len(values) >= len(ordered_packages):
            pairs = list(zip(ordered_packages, values))
        elif len(values) == 1:
            pairs = [(package, values[0]) for package in ordered_packages]
        else:
            continue
        for package, value in pairs:
            thermal_conditions.add(package.lower())
            cards.append(_card(
                "thermal_resistance_ja",
                "Thermal Resistance JA",
                value,
                page=page,
                chunk_id=chunk_id,
                value_typ=_first_number(value),
                unit="°C/W",
                condition=package,
                source_text=line[:520],
                confidence="medium",
            ))

    for package in PACKAGE_TOKENS:
        if package.lower() in thermal_conditions:
            continue
        for match in re.finditer(re.escape(package), text, re.I):
            window_start = max(0, match.start() - 80)
            window_end = min(len(text), match.end() + 160)
            window = text[window_start:window_end]
            value_match = re.search(r"([0-9]{2,3}\s*°?\s*C\s*/\s*W)", window, re.I)
            if value_match:
                value = _clean_spaces(value_match.group(1)).replace("C / W", "°C/W").replace("C/W", "°C/W")
                cards.append(_card(
                    "thermal_resistance_ja",
                    "Thermal Resistance JA",
                    value,
                    page=page,
                    chunk_id=chunk_id,
                    value_typ=_first_number(value),
                    unit="°C/W",
                    condition=package,
                    source_text=_clean_spaces(window)[:420],
                    confidence="medium",
                ))

    return cards


def _extract_complex_chip_parameters(text: str, *, page: int, chunk_id: str) -> List[dict]:
    cards: List[dict] = []

    operating_match = re.search(
        r"工作环境温度为\s*([-+]?\d+(?:\.\d+)?)\s*℃\s*[-~\\]+\s*([-+]?\d+(?:\.\d+)?)\s*℃.*?相对湿度\s*([0-9]+(?:\.\d+)?)\s*%\s*\\?~\s*([0-9]+(?:\.\d+)?)\s*%",
        text,
        re.I,
    )
    if operating_match:
        temp_min = float(operating_match.group(1))
        temp_max = float(operating_match.group(2))
        hum_min = float(operating_match.group(3))
        hum_max = float(operating_match.group(4))
        cards.append(_card(
            "environment_operating_temperature",
            "Operating Ambient Temperature",
            f"{temp_min:g}°C to {temp_max:g}°C",
            page=page,
            chunk_id=chunk_id,
            value_min=temp_min,
            value_max=temp_max,
            unit="°C",
            source_text=_excerpt(text, operating_match.start(), operating_match.end()),
            confidence="high",
            extraction_method="regex_chip_v1",
        ))
        cards.append(_card(
            "environment_operating_humidity",
            "Operating Relative Humidity",
            f"{hum_min:g}% to {hum_max:g}%",
            page=page,
            chunk_id=chunk_id,
            value_min=hum_min,
            value_max=hum_max,
            unit="%",
            source_text=_excerpt(text, operating_match.start(), operating_match.end()),
            confidence="high",
            extraction_method="regex_chip_v1",
        ))

    storage_match = re.search(
        r"储存环境温度为\s*([-+]?\d+(?:\.\d+)?)\s*℃\s*\\?~\s*([-+]?\d+(?:\.\d+)?)\s*℃.*?存储相对湿度[:：]?\s*([0-9]+(?:\.\d+)?)\s*%\s*\\?~\s*([0-9]+(?:\.\d+)?)\s*%",
        text,
        re.I,
    )
    if storage_match:
        temp_min = float(storage_match.group(1))
        temp_max = float(storage_match.group(2))
        hum_min = float(storage_match.group(3))
        hum_max = float(storage_match.group(4))
        cards.append(_card(
            "environment_storage_temperature",
            "Storage Temperature",
            f"{temp_min:g}°C to {temp_max:g}°C",
            page=page,
            chunk_id=chunk_id,
            value_min=temp_min,
            value_max=temp_max,
            unit="°C",
            source_text=_excerpt(text, storage_match.start(), storage_match.end()),
            confidence="high",
            extraction_method="regex_chip_v1",
        ))
        cards.append(_card(
            "environment_storage_humidity",
            "Storage Relative Humidity",
            f"{hum_min:g}% to {hum_max:g}%",
            page=page,
            chunk_id=chunk_id,
            value_min=hum_min,
            value_max=hum_max,
            unit="%",
            source_text=_excerpt(text, storage_match.start(), storage_match.end()),
            confidence="high",
            extraction_method="regex_chip_v1",
        ))

    lines = text.splitlines()
    current_power_scenario = ""
    for line in lines:
        cells = _split_cells(line)
        if len(cells) >= 7 and re.match(r"^V[A-Z0-9_\[\]-]+$", cells[0], re.I):
            rail = _normalize_signal_name(cells[0])
            value_min = _parse_number_cell(cells[2], voltage=True)
            value_typ = _parse_number_cell(cells[3], voltage=True)
            value_max = _parse_number_cell(cells[4], voltage=True)
            if value_min is not None or value_typ is not None or value_max is not None:
                cards.append(_card(
                    "power_rail_voltage",
                    f"{rail} Voltage Range",
                    _range_value_text(value_min, value_typ, value_max, "V"),
                    page=page,
                    chunk_id=chunk_id,
                    value_min=value_min,
                    value_typ=value_typ,
                    value_max=value_max,
                    unit="V",
                    condition=f"{rail}; {cells[1]}; AC噪声 {cells[5]}; 参考地 {cells[6]}",
                    source_text=line,
                    confidence="medium" if any(token in line for token in ("K", "DA", "VSs", "C.")) else "high",
                    extraction_method="table_chip_power_rail_v1",
                ))
            noise = _parse_number_cell(cells[5])
            if noise is not None and "%" in cells[5]:
                cards.append(_card(
                    "power_rail_ac_noise",
                    f"{rail} AC Noise",
                    f"{noise:g}%",
                    page=page,
                    chunk_id=chunk_id,
                    value_typ=noise,
                    unit="%",
                    condition=f"{rail}; {cells[1]}; 参考地 {cells[6]}",
                    source_text=line,
                    confidence="medium",
                    extraction_method="table_chip_power_rail_v1",
                ))

        if len(cells) >= 4:
            first = cells[0]
            if first in {"典型功耗", "最大功耗"}:
                current_power_scenario = first
                rail = _normalize_signal_name(cells[1]) if len(cells) > 1 else ""
                offset = 1
            elif current_power_scenario and re.match(r"^V[A-Z0-9_\[\]-]+$", first, re.I):
                rail = _normalize_signal_name(first)
                offset = 0
            else:
                rail = ""
                offset = 0
            if rail and len(cells) >= offset + 4:
                voltage = _parse_number_cell(cells[offset + 1], voltage=True)
                current = _parse_number_cell(cells[offset + 2])
                power = _parse_number_cell(cells[offset + 3])
                total = _parse_number_cell(cells[offset + 4]) if len(cells) > offset + 4 else None
                if any(value is not None for value in (current, power, total)):
                    detail = []
                    if voltage is not None:
                        detail.append(f"voltage {voltage:g}V")
                    if current is not None:
                        detail.append(f"current {current:g}A")
                    if power is not None:
                        detail.append(f"power {power:g}W")
                    if total is not None:
                        detail.append(f"total {total:g}W")
                    cards.append(_card(
                        "power_consumption",
                        f"{rail} Power Consumption",
                        ", ".join(detail),
                        page=page,
                        chunk_id=chunk_id,
                        value_typ=power if power is not None else current,
                        unit="W" if power is not None else "A",
                        condition=f"{current_power_scenario}; {rail}",
                        source_text=line,
                        confidence="low" if "TBD" in line.upper() else "medium",
                        extraction_method="table_chip_power_consumption_v1",
                    ))

        if cells and re.match(r"^T\d+(?:/T\d+)?$", cells[0], re.I) and len(cells) >= 4:
            name = cells[0]
            min_cell = cells[1]
            max_cell = cells[2]
            description = cells[3]
            value_min = _parse_number_cell(min_cell)
            value_max = _parse_number_cell(max_cell)
            unit = _unit(min_cell or max_cell)
            if value_min is not None or value_max is not None:
                cards.append(_card(
                    "power_sequence_timing",
                    f"Power Sequence {name}",
                    _range_value_text(value_min, None, value_max, unit),
                    page=page,
                    chunk_id=chunk_id,
                    value_min=value_min,
                    value_max=value_max,
                    unit=unit,
                    condition=description,
                    source_text=line,
                    confidence="high",
                    extraction_method="table_chip_sequence_v1",
                ))

        if len(cells) >= 6 and re.match(r"^(?:Tj|TA|0J[BC]|θJ[ABC])$", cells[0], re.I):
            symbol = _normalize_theta_symbol(cells[0])
            parameter = cells[1]
            value_min = _parse_number_cell(cells[2])
            value_typ = _parse_number_cell(cells[3])
            value_max = _parse_number_cell(cells[4])
            unit = _unit(cells[5])
            if value_min is not None or value_typ is not None or value_max is not None:
                cards.append(_card(
                    "thermal_characteristic",
                    f"{symbol} {parameter}",
                    _range_value_text(value_min, value_typ, value_max, unit),
                    page=page,
                    chunk_id=chunk_id,
                    value_min=value_min,
                    value_typ=value_typ,
                    value_max=value_max,
                    unit=unit,
                    condition=symbol,
                    source_text=line,
                    confidence="medium" if "■" in line else "high",
                    extraction_method="table_chip_thermal_v1",
                ))

    junction_limit = re.search(r"结温超过\s*([0-9]+(?:\.\d+)?)\s*℃\s*限制", text)
    if junction_limit:
        number = float(junction_limit.group(1))
        cards.append(_card(
            "junction_temperature_limit",
            "Junction Temperature Limit",
            f"{number:g}°C",
            page=page,
            chunk_id=chunk_id,
            value_max=number,
            unit="°C",
            source_text=_excerpt(text, junction_limit.start(), junction_limit.end()),
            confidence="high",
            extraction_method="regex_chip_v1",
        ))

    return cards


def extract_datasheet_parameters(chunks: Iterable[dict], *, title: str = "") -> List[dict]:
    """Extract stable parameter cards from indexed datasheet chunks."""

    cards: List[dict] = []
    seen = set()
    for chunk in chunks or []:
        if not isinstance(chunk, dict):
            continue
        text = normalize_parameter_text(str(chunk.get("text") or ""))
        if not text:
            continue
        page = int(chunk.get("page") or 1)
        chunk_id = str(chunk.get("chunk_id") or "")
        extracted = []
        extracted.extend(_extract_common_ldo_parameters(text, page=page, chunk_id=chunk_id))
        extracted.extend(_extract_complex_chip_parameters(text, page=page, chunk_id=chunk_id))
        for card in extracted:
            key = (
                card.get("parameter_key"),
                card.get("page"),
                card.get("chunk_id"),
                card.get("value_text"),
                card.get("condition"),
            )
            if key in seen:
                continue
            seen.add(key)
            card["document_title"] = title
            cards.append(card)
    return cards
