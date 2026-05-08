# -*- coding: utf-8 -*-
"""BOM Tools Web v2 — 公共模块"""

import os, sys, time, re

_parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _parent not in sys.path:
    sys.path.insert(0, _parent)

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter, column_index_from_string

try:
    import requests
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
    import requests

from flask import request, jsonify, send_file

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")

FEISHU_PRESET_TABLES = [
    {"name": "MLCC",               "token": "shthq7d9W17DSo7cwuFhtIg7KPf", "category": "优选库"},
    {"name": "电阻",                "token": "shthqdJvubPLY8mrO8qkOMXmGiw", "category": "优选库"},
    {"name": "电解电容",             "token": "shthqO56eTG9DyJX60nDaaMGp0e", "category": "优选库"},
    {"name": "网络变压器/电感器",     "token": "shthquBGGQB8twAgmJWSBwFY5he", "category": "优选库"},
    {"name": "磁珠",                "token": "shthq1EJlpgHfBqBGNediRZdt8c", "category": "优选库"},
    {"name": "晶体晶振",             "token": "shthqpy6hg6hVD78VNPElmcSF6d", "category": "优选库"},
    {"name": "保险丝",               "token": "shthqpsZCFkwGjn62CjRYCpKrHg", "category": "优选库"},
    {"name": "纽扣电池",             "token": "shthqngCNYdufotGIcL6Vn6gWWh", "category": "优选库"},
    {"name": "滤波器/共模扼流圈",     "token": "shthqaZbFrblnh3V0A3ahhOj4Og", "category": "优选库"},
    {"name": "Power IC优选库",       "token": "shthqz7lKPJt9UGF4FOIU1uyTUh", "category": "优选库"},
    {"name": "功能IC优选库",          "token": "shthq4b1PTCh1HqyalUal6aTYte", "category": "优选库"},
    {"name": "DBG分立器件优选库",     "token": "shthqEZrwmemvVULwrhmyAmlfzd", "category": "优选库"},
    {"name": "连接器",               "token": "shthqE9sVI2DkIBYkSkLcUdNxvn", "category": "优选库"},
    {"name": "Cable",               "token": "shthqpYELkJAH7b0uPn1HRcEyLg",  "category": "优选库"},
    {"name": "客户物料型号与HQ料号对应关系", "token": "shthq1R9G7zSp5hvTISGNDOWjme", "category": "对应关系库"},
]


def _cell_str(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, str):
        return val.strip()
    if isinstance(val, list):
        return " ".join(
            item.get("text", "") or item.get("link", "")
            for item in val if isinstance(item, dict)
        ).strip()
    return str(val).strip()


def _cleanup_old_files(directory, minutes=30):
    now = time.time()
    for f in os.listdir(directory):
        fp = os.path.join(directory, f)
        try:
            if os.path.isfile(fp) and now - os.path.getmtime(fp) > minutes * 60:
                os.remove(fp)
        except Exception:
            pass


def _col_int(s):
    if not s:
        return None
    s = str(s).strip().upper()
    try:
        return int(s) if s.isdigit() else column_index_from_string(s)
    except Exception:
        return None
