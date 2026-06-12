# -*- coding: utf-8 -*-
"""BOM Checklist checks."""

import uuid
from collections import Counter

from flask import Blueprint

from activity import track_tool_activity
from shared import _open_workbook, _request_int, _save_uploaded_excel, jsonify, request


bom_checklist_bp = Blueprint("bom_checklist", __name__)

PCBA_REQUIRED_HEADERS = [
    "序号", "料号", "上阶BOM名称", "BOM层级", "型号", "物料描述", "单耗", "替代关系", "位号", "生产厂家",
    "是否环保", "湿敏属性",
]


def _cell_text(value):
    if value is None:
        return ""
    return str(value).strip()


def _load_sheet(file, prefix="bom_checklist"):
    if not file:
        raise ValueError("请上传 BOM 文件")
    uid = str(uuid.uuid4())[:8]
    path = _save_uploaded_excel(file, prefix, uid)
    wb = _open_workbook(path, data_only=True)
    sheet_name = request.form.get("sheet_name", "")
    if not sheet_name or sheet_name not in wb.sheetnames or sheet_name == "先选择文件":
        sheet_name = wb.sheetnames[0]
    header_row = _request_int("header_row", 1)
    if header_row is None:
        wb.close()
        raise ValueError("表头行必须是大于等于 1 的数字")
    return uid, wb, wb[sheet_name], sheet_name, header_row


def _sheet_profile(ws, header_row):
    headers = [_cell_text(ws.cell(row=header_row, column=ci).value) for ci in range(1, ws.max_column + 1)]
    data_rows = []
    blank_rows = []
    for ri in range(header_row + 1, ws.max_row + 1):
        values = [_cell_text(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
        if any(values):
            data_rows.append((ri, values))
        else:
            blank_rows.append(ri)
    return headers, data_rows, blank_rows, [values for _, values in data_rows[:20]]


def _header_index(headers):
    return {name: idx for idx, name in enumerate(headers) if name}


def _rows_as_dicts(headers, data_rows):
    rows = []
    for excel_row, values in data_rows:
        item = {header: values[idx] if idx < len(values) else "" for idx, header in enumerate(headers) if header}
        item["__row__"] = excel_row
        item["__values__"] = values
        rows.append(item)
    return rows


def _result(check_id, name, status, message, count=0, rows=None):
    return {
        "id": check_id,
        "name": name,
        "status": status,
        "message": message,
        "count": count,
        "rows": list(rows or [])[:50],
    }


def _check_required_headers(headers):
    required = {
        "料号": ("料号", "HQ料号", "HQ PN", "Part Number", "PART_NUMBER"),
        "型号": ("型号", "规格型号", "厂商型号", "Manufacturer P/N", "MPN"),
        "用量": ("用量", "数量", "单耗", "Qty", "QTY", "Quantity"),
    }
    missing = [name for name, aliases in required.items() if not any(alias in headers for alias in aliases)]
    return _result(
        "required_headers",
        "关键表头检查",
        "pass" if not missing else "warn",
        "关键表头齐全" if not missing else "缺少或未识别：" + "、".join(missing),
        len(missing),
    )


def _check_pcba_standard_headers(headers):
    missing = [name for name in PCBA_REQUIRED_HEADERS if name not in headers]
    return _result(
        "pcba_standard_headers",
        "PCBA BOM 标准表头",
        "pass" if not missing else "fail",
        "PCBA BOM 标准关键列齐全" if not missing else "缺少标准关键列：" + "、".join(missing),
        len(missing),
    )


def _check_duplicate_headers(headers):
    counts = Counter(h for h in headers if h)
    duplicates = [name for name, count in counts.items() if count > 1]
    return _result(
        "duplicate_headers",
        "重复表头名",
        "pass" if not duplicates else "fail",
        "未发现重复表头" if not duplicates else "发现重复表头：" + "、".join(duplicates),
        len(duplicates),
    )


def _check_blank_rows(blank_rows):
    return _result(
        "blank_rows",
        "数据区空行",
        "pass" if not blank_rows else "warn",
        "数据区未发现空行" if not blank_rows else f"数据区发现 {len(blank_rows)} 个空行",
        len(blank_rows),
        blank_rows,
    )


def _check_depop_removed(row_dicts):
    rows = [row["__row__"] for row in row_dicts if any("DEPOP" in value.upper() for value in row.get("__values__", []))]
    return _result(
        "pcba_depop_removed",
        "DEPOP 物料移除检查",
        "pass" if not rows else "fail",
        "未发现 DEPOP 物料" if not rows else f"发现 {len(rows)} 行包含 DEPOP，请确认传 BOM 前已删除 DEPOP 物料",
        len(rows),
        rows,
    )


def _check_pcb_item(row_dicts):
    pcb_location_rows = [row["__row__"] for row in row_dicts if row.get("位号", "").strip().upper() == "PCB"]
    if pcb_location_rows:
        return _result(
            "pcba_pcb_location",
            "PCB 料号和位号",
            "pass",
            "已发现 PCB 主料行，且位号/location 填写为 PCB",
            len(pcb_location_rows),
            pcb_location_rows,
        )
    return _result(
        "pcba_pcb_location",
        "PCB 料号和位号",
        "fail",
        "未发现位号/location 为 PCB 的记录；请确认 PCB 料号是否已加入 BOM，且 PCB 主料行位号填写为 PCB",
        0,
        [],
    )


def _is_non_empty_qty(value):
    if value == "":
        return False
    try:
        return float(value) != 0
    except ValueError:
        return True


def _check_qty_by_substitute_type(row_dicts):
    main_missing = []
    alt_with_qty = []
    for row in row_dicts:
        rel = row.get("替代关系", "")
        qty = row.get("单耗", "")
        if "主料" in rel and not _is_non_empty_qty(qty):
            main_missing.append(row["__row__"])
        if "替代料" in rel and _is_non_empty_qty(qty):
            alt_with_qty.append(row["__row__"])
    rows = main_missing + alt_with_qty
    if not rows:
        return _result("pcba_qty_by_substitute_type", "主料/替代料单耗", "pass", "主料单耗和替代料单耗填写符合常规 PCBA BOM 规则")
    parts = []
    if main_missing:
        parts.append(f"主料单耗为空 {len(main_missing)} 行")
    if alt_with_qty:
        parts.append(f"替代料单耗不为空 {len(alt_with_qty)} 行")
    return _result("pcba_qty_by_substitute_type", "主料/替代料单耗", "fail", "；".join(parts), len(rows), rows)


def _check_parent_bom_references(row_dicts):
    part_numbers = {row.get("料号", "") for row in row_dicts if row.get("料号", "")}
    top_parents = {row.get("上阶BOM名称", "") for row in row_dicts if row.get("BOM层级", "") in ("1", "1.0") and row.get("上阶BOM名称", "")}
    missing_parent_rows = []
    for row in row_dicts:
        level = row.get("BOM层级", "")
        parent = row.get("上阶BOM名称", "")
        if level not in ("", "1", "1.0") and parent and parent not in part_numbers and parent not in top_parents:
            missing_parent_rows.append(row["__row__"])
    return _result(
        "pcba_parent_bom_reference",
        "BOM 层级父项引用",
        "pass" if not missing_parent_rows else "fail",
        "多层级物料的上阶 BOM 名称均可在 BOM 中找到" if not missing_parent_rows else "存在子阶物料的上阶 BOM 名称无法在本 BOM 中找到",
        len(missing_parent_rows),
        missing_parent_rows,
    )



COL_PART_NO = "\u6599\u53f7"
COL_MODEL = "\u578b\u53f7"
COL_DESC = "\u7269\u6599\u63cf\u8ff0"
COL_PARENT_BOM = "\u4e0a\u9636BOM\u540d\u79f0"
COL_RELATION = "\u66ff\u4ee3\u5173\u7cfb"
COL_QTY = "\u5355\u8017"
COL_REFDES = "\u4f4d\u53f7"
REL_MAIN = "\u4e3b\u6599"
REL_ALT = "\u66ff\u4ee3\u6599"


def _is_hqv_row(row):
    return row.get(COL_PART_NO, "").strip().upper().startswith("HQV")


def _classify_hqv_type(row):
    model = row.get(COL_MODEL, "").upper().replace(" ", "")
    desc = row.get(COL_DESC, "").upper().replace(" ", "")
    text = model + " " + desc
    if "SMT" in text and "DIMM" in text:
        return "DIMM Slot SMT\u6599\u53f7", "SMT_DIMM"
    if "DIP" in text and "DIMM" in text:
        return "DIMM Slot DIP\u6599\u53f7", "DIP_DIMM"
    if "SOCKET" in text:
        return "CPU Socket\u6599\u53f7", "SMT_CPUSOCKET"
    if "POWER" in text:
        return "Power Solution\u6599\u53f7", "SMT_POWER"
    if "ASSY" in text:
        return "Assy\u6599\u53f7", "ASSY"
    if "PRESSFIT" in text or "PRESS-FIT" in text:
        return "Press Fit\u6599\u53f7", "PRESS FIT"
    if "DIP" in text:
        return "DIP\u6599\u53f7", "DIP"
    return "", ""


def _check_hqv_structure(row_dicts):
    hqv_rows = [row for row in row_dicts if _is_hqv_row(row)]
    child_parents = Counter(row.get(COL_PARENT_BOM, "") for row in row_dicts if row.get(COL_PARENT_BOM, ""))
    no_child = []
    main_qty_bad = []
    alt_qty_bad = []
    main_refdes_empty = []
    for row in hqv_rows:
        part_no = row.get(COL_PART_NO, "")
        relation = row.get(COL_RELATION, "")
        if child_parents.get(part_no, 0) == 0:
            no_child.append(row["__row__"])
        if REL_MAIN in relation:
            if row.get(COL_QTY, "") != "1":
                main_qty_bad.append(row["__row__"])
            if not row.get(COL_REFDES, ""):
                main_refdes_empty.append(row["__row__"])
        elif REL_ALT in relation and _is_non_empty_qty(row.get(COL_QTY, "")):
            alt_qty_bad.append(row["__row__"])
    bad_rows = no_child + main_qty_bad + alt_qty_bad + main_refdes_empty
    if not bad_rows:
        return _result(
            "pcba_hqv_structure",
            "HQV \u5c42\u7ea7\u7ed3\u6784",
            "pass",
            f"\u5df2\u68c0\u67e5 {len(hqv_rows)} \u4e2a HQV \u7269\u6599\uff0c\u5b50\u9636/\u4e3b\u6599\u5355\u8017/\u66ff\u4ee3\u6599\u5355\u8017\u7b26\u5408\u89c4\u5219",
        )
    parts = []
    if no_child:
        parts.append(f"\u7f3a\u5c11\u5b50\u9636 {len(no_child)} \u884c")
    if main_qty_bad:
        parts.append(f"\u4e3b\u6599 HQV \u5355\u8017\u4e0d\u662f 1\uff1a{len(main_qty_bad)} \u884c")
    if alt_qty_bad:
        parts.append(f"\u66ff\u4ee3\u6599 HQV \u5355\u8017\u4e0d\u4e3a\u7a7a\uff1a{len(alt_qty_bad)} \u884c")
    if main_refdes_empty:
        parts.append(f"\u4e3b\u6599 HQV \u4f4d\u53f7\u4e3a\u7a7a\uff1a{len(main_refdes_empty)} \u884c")
    return _result("pcba_hqv_structure", "HQV \u5c42\u7ea7\u7ed3\u6784", "fail", "\uff1b".join(parts), len(bad_rows), bad_rows)


def _check_hqv_naming_and_refdes(row_dicts):
    hqv_rows = [row for row in row_dicts if _is_hqv_row(row)]
    unknown_type_rows = []
    refdes_bad_rows = []
    for row in hqv_rows:
        category, expected_refdes = _classify_hqv_type(row)
        if not category:
            unknown_type_rows.append(row["__row__"])
            continue
        if REL_MAIN in row.get(COL_RELATION, "") and row.get(COL_REFDES, "").strip().upper() != expected_refdes:
            refdes_bad_rows.append(row["__row__"])
    if refdes_bad_rows:
        return _result(
            "pcba_hqv_naming_refdes",
            "HQV \u547d\u540d/\u4f4d\u53f7",
            "fail",
            "\u90e8\u5206\u4e3b\u6599 HQV \u4f4d\u53f7\u4e0e\u6599\u53f7\u7c7b\u578b\u4e0d\u5339\u914d",
            len(refdes_bad_rows),
            refdes_bad_rows,
        )
    if unknown_type_rows:
        return _result(
            "pcba_hqv_naming_refdes",
            "HQV \u547d\u540d/\u4f4d\u53f7",
            "warn",
            "\u5b58\u5728\u672a\u80fd\u8bc6\u522b\u7c7b\u578b\u7684 HQV\uff0c\u8bf7\u786e\u8ba4\u578b\u53f7/\u7269\u6599\u63cf\u8ff0\u547d\u540d",
            len(unknown_type_rows),
            unknown_type_rows,
        )
    return _result(
        "pcba_hqv_naming_refdes",
        "HQV \u547d\u540d/\u4f4d\u53f7",
        "pass",
        f"\u5df2\u68c0\u67e5 {len(hqv_rows)} \u4e2a HQV \u7269\u6599\uff0c\u547d\u540d\u7c7b\u578b\u548c\u4e3b\u6599\u4f4d\u53f7\u7b26\u5408\u89c4\u5219",
    )


AUTH_LABELS = {
    "bios_ami": {"part": "HQ60410125009", "name": "BIOS AMI授权标签"},
    "bmc_ast2500": {"part": "HQ11111788009", "name": "BMC AMI授权标签(AST2500)"},
    "bmc_ast2520": {"part": "HQ60432070000", "name": "BMC Insyde授权标签(AST2520/Mega)"},
    "bmc_ast2600": {"part": "HQ11111801009", "name": "BMC AMI授权标签(AST2600)"},
}

MAC_CHIP_TOKENS = ("网卡芯片", "ETHERNET", "LAN_MAC", "_MAC_", "_MAC", "WGI210", "I210", "I350", "X710", "E810", "BCM", "CONNECTX")
MAC_LABEL_TOKENS = ("MAC LABEL", "MAC标签", "MAC 标签", "MAC地址标签", "MAC ADDRESS LABEL", "MAC")
SN_LABEL_TOKENS = ("SN LABEL", "SN标签", "SN 标签", "序列号标签", "SERIAL NUMBER", "S/N LABEL")


def _row_text(row):
    return " ".join(str(value) for value in row.get("__values__", []) if value).upper()


def _core_part_text(row):
    fields = (COL_MODEL, COL_DESC, "物料描述-英文", "MODEL")
    return " ".join(row.get(field, "") for field in fields if row.get(field, "")).upper()


def _top_level_parent_names(row_dicts):
    return {
        row.get(COL_PARENT_BOM, "").strip().upper()
        for row in row_dicts
        if row.get("BOM层级", "") in ("1", "1.0") and row.get(COL_PARENT_BOM, "")
    }


def _is_pcba_assembly_stage_row(row, top_parent_names=None):
    text = _row_text(row)
    parent = row.get(COL_PARENT_BOM, "").strip().upper()
    level = row.get("BOM层级", "")
    refdes = row.get(COL_REFDES, "").upper()
    top_parent_names = top_parent_names or set()
    if any(token in text for token in ("组装", "組裝", "虚拟", "虛擬", "ASSY", "ASSEMBLY", "VIRTUAL", "PCBA")):
        return True
    if level in ("1", "1.0") and parent and parent in top_parent_names:
        return True
    if level in ("1", "1.0") and any(token in parent + " " + refdes + " " + text for token in ("PCB", "PCBA")):
        return True
    return False


def _is_bmc_chip_row(row):
    text = _core_part_text(row)
    return any(token in text for token in ("BMC", "AST2500", "AST2520", "AST2600")) and not _is_label_text(text)


def _is_mac_chip_row(row):
    text = _core_part_text(row)
    return any(token in text for token in MAC_CHIP_TOKENS) and not _is_label_text(text)


def _is_label_text(text):
    return "LABEL" in text or "标签" in text or "標籤" in text


def _is_mac_label_row(row):
    text = _row_text(row)
    return _is_label_text(text) and any(token in text for token in MAC_LABEL_TOKENS)


def _is_sn_label_row(row):
    text = _row_text(row)
    return _is_label_text(text) and any(token in text for token in SN_LABEL_TOKENS)


def _check_bmc_bios_auth_labels(row_dicts):
    label_parts = {info["part"]: key for key, info in AUTH_LABELS.items()}
    present = {}
    ast_clues = {"bmc_ast2500": [], "bmc_ast2520": [], "bmc_ast2600": []}
    bmc_chip_rows = []
    mac_chip_rows = []
    mac_label_rows = []
    sn_label_rows = []
    for row in row_dicts:
        part_no = row.get(COL_PART_NO, "").strip().upper()
        if part_no in label_parts:
            present.setdefault(label_parts[part_no], []).append(row)
        text = _row_text(row)
        if _is_bmc_chip_row(row):
            bmc_chip_rows.append(row["__row__"])
        if _is_mac_chip_row(row):
            mac_chip_rows.append(row["__row__"])
        if _is_mac_label_row(row):
            mac_label_rows.append(row)
        if _is_sn_label_row(row):
            sn_label_rows.append(row)
        if "AST2500" in text:
            ast_clues["bmc_ast2500"].append(row["__row__"])
        if "AST2520" in text or "MEGA" in text:
            ast_clues["bmc_ast2520"].append(row["__row__"])
        if "AST2600" in text:
            ast_clues["bmc_ast2600"].append(row["__row__"])

    expected = ["bios_ami"]
    expected.extend(key for key, rows in ast_clues.items() if rows)
    missing = [key for key in expected if key not in present]
    missing_names = [f"{AUTH_LABELS[key]['name']}({AUTH_LABELS[key]['part']})" for key in missing]
    if mac_chip_rows and not mac_label_rows:
        missing_names.append("MAC 标签")
    if not sn_label_rows:
        missing_names.append("SN 标签")
    clue_rows = sorted(set([row for rows in ast_clues.values() for row in rows] + bmc_chip_rows + mac_chip_rows))
    if missing_names:
        return _result(
            "pcba_bmc_bios_auth_labels",
            "BMC/BIOS/MAC/SN 标签检查",
            "fail",
            "缺少应加入 PCBA 组装虚拟阶 BOM 的标签：" + "、".join(missing_names) + "；BMC/BIOS 标签请与项目软件确认跟随主芯片还是 FLASH",
            len(missing_names),
            clue_rows,
        )

    top_parent_names = _top_level_parent_names(row_dicts)
    label_rows = []
    for key in expected:
        label_rows.extend(present.get(key, []))
    label_rows.extend(mac_label_rows)
    label_rows.extend(sn_label_rows)
    non_assembly_rows = sorted({row["__row__"] for row in label_rows if not _is_pcba_assembly_stage_row(row, top_parent_names)})
    if non_assembly_rows:
        return _result(
            "pcba_bmc_bios_auth_labels",
            "BMC/BIOS/MAC/SN 标签检查",
            "warn",
            "标签已找到，但未能确认都挂在 PCBA 组装虚拟阶 BOM；请确认上阶 BOM/层级，BMC/BIOS 标签还需确认跟随主芯片还是 FLASH",
            len(non_assembly_rows),
            non_assembly_rows,
        )

    if not bmc_chip_rows and not mac_chip_rows:
        return _result(
            "pcba_bmc_bios_auth_labels",
            "BMC/BIOS/MAC/SN 标签检查",
            "warn",
            "SN/BIOS 标签已找到；未识别 BMC 或 MAC/网卡芯片线索，请按项目实际配置人工确认是否需要 BMC/MAC 标签",
            0,
        )

    return _result(
        "pcba_bmc_bios_auth_labels",
        "BMC/BIOS/MAC/SN 标签检查",
        "pass",
        "已根据规范性 PCBA BOM 结构检查到 BMC/BIOS/MAC/SN 相关标签，且标签位于疑似 PCBA 组装虚拟阶 BOM",
    )


FLASH_SOCKET_TOKENS = ("FLASH SOCKET", "FLASH_SOCKET", "FLASH插座", "FLASH 插座", "FLASH座", "FLASH 座", "SOCKET FLASH")
FLASH_CHIP_TOKENS = ("SPI FLASH", "NOR FLASH", "NAND FLASH", "FLASH_", "_FLASH", " W25", "W25", "MX25", "GD25", "N25", "MT25", "AT25", "EN25")


def _is_flash_socket_row(row):
    text = _core_part_text(row)
    if _is_label_text(text):
        return False
    return "FLASH" in text and any(token in text for token in ("SOCKET", "插座", "座")) or any(token in text for token in FLASH_SOCKET_TOKENS)


def _is_smt_flash_row(row):
    text = _core_part_text(row)
    if _is_label_text(text) or _is_flash_socket_row(row):
        return False
    if "FLASH" not in text and not any(token.strip() and token.strip() in text for token in ("W25", "MX25", "GD25", "MT25", "AT25", "EN25")):
        return False
    return any(token in text for token in FLASH_CHIP_TOKENS)


def _check_flash_socket_smt_conflict(row_dicts):
    socket_rows = [row["__row__"] for row in row_dicts if _is_flash_socket_row(row)]
    smt_flash_rows = [row["__row__"] for row in row_dicts if _is_smt_flash_row(row)]
    if socket_rows and smt_flash_rows:
        return _result(
            "pcba_flash_socket_smt_conflict",
            "Flash Socket/贴片 Flash 互斥",
            "fail",
            "Flash Socket 与贴片/焊接 Flash 不能同时出现在 BOM 中；去掉 Flash Socket 量产时，应将原组装虚拟阶 Flash 调整到 SMT 阶，并核对原理图 Symbol 与 Layout colay 方案",
            len(socket_rows) + len(smt_flash_rows),
            socket_rows + smt_flash_rows,
        )
    if socket_rows:
        return _result(
            "pcba_flash_socket_smt_conflict",
            "Flash Socket/贴片 Flash 互斥",
            "warn",
            "检测到 Flash Socket；请核对 Flash 与 Socket 对应位置关系，原理图需同时放置 Flash 和 Socket symbol，Layout 使用 colay，且 Flash 禁用万能 symbol",
            len(socket_rows),
            socket_rows,
        )
    if smt_flash_rows:
        return _result(
            "pcba_flash_socket_smt_conflict",
            "Flash Socket/贴片 Flash 互斥",
            "pass",
            "未发现 Flash Socket；贴片/焊接 Flash 可按 SMT 阶 BOM 管理",
            len(smt_flash_rows),
            smt_flash_rows,
        )
    return _result(
        "pcba_flash_socket_smt_conflict",
        "Flash Socket/贴片 Flash 互斥",
        "pass",
        "未发现 Flash Socket 与贴片/焊接 Flash 同时存在",
    )
def _check_debug_keywords(row_dicts):
    rows = [row["__row__"] for row in row_dicts if any("DEBUG" in value.upper() for value in row.get("__values__", []))]
    return _result(
        "pcba_debug_keyword",
        "Debug only 物料提示",
        "pass" if not rows else "warn",
        "未发现 Debug 关键词" if not rows else "发现 Debug 关键词；MP 阶段请确认 debug only 物料是否应移除",
        len(rows),
        rows,
    )


def _run_checks(headers, data_rows, blank_rows):
    row_dicts = _rows_as_dicts(headers, data_rows)
    checks = [
        _check_required_headers(headers),
        _check_pcba_standard_headers(headers),
        _check_duplicate_headers(headers),
        _check_blank_rows(blank_rows),
        _check_depop_removed(row_dicts),
        _check_pcb_item(row_dicts),
        _check_qty_by_substitute_type(row_dicts),
        _check_parent_bom_references(row_dicts),
        _check_hqv_structure(row_dicts),
        _check_hqv_naming_and_refdes(row_dicts),
        _check_bmc_bios_auth_labels(row_dicts),
        _check_flash_socket_smt_conflict(row_dicts),
        _check_debug_keywords(row_dicts),
    ]
    summary = {
        "total": len(checks),
        "pass": sum(1 for item in checks if item["status"] == "pass"),
        "warn": sum(1 for item in checks if item["status"] == "warn"),
        "fail": sum(1 for item in checks if item["status"] == "fail"),
        "data_rows": len(data_rows),
    }
    return summary, checks


@bom_checklist_bp.route("/api/bom_checklist/preview", methods=["POST"])
def api_bom_checklist_preview():
    try:
        _, wb, ws, sheet_name, header_row = _load_sheet(request.files.get("file"), "bom_checklist_pre")
        headers, data_rows, blank_rows, preview = _sheet_profile(ws, header_row)
        sheets = wb.sheetnames
        wb.close()
        return jsonify({
            "success": True,
            "sheets": sheets,
            "current_sheet": sheet_name,
            "header_row": header_row,
            "headers": headers,
            "data_rows": len(data_rows),
            "blank_rows": len(blank_rows),
            "preview": preview,
        })
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})


@bom_checklist_bp.route("/api/bom_checklist/run", methods=["POST"])
@track_tool_activity("BOM Checklist")
def api_bom_checklist_run():
    try:
        uid, wb, ws, sheet_name, header_row = _load_sheet(request.files.get("file"), "bom_checklist_run")
        headers, data_rows, blank_rows, preview = _sheet_profile(ws, header_row)
        summary, checks = _run_checks(headers, data_rows, blank_rows)
        wb.close()
        return jsonify({
            "success": True,
            "uid": uid,
            "sheet_name": sheet_name,
            "header_row": header_row,
            "headers": headers,
            "summary": summary,
            "checks": checks,
            "preview": preview,
        })
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})