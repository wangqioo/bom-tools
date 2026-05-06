# -*- coding: utf-8 -*-
"""Export helpers for PSTX analysis outputs.

The package keeps heavy spreadsheet dependencies lazy.  Importing
``pstx_exports`` alone should not load ``openpyxl``; only calling
``export_to_excel`` does.
"""


def export_to_excel(data: dict, out_path: str) -> str:
    from pstx_exports.excel import export_to_excel as _export_to_excel
    return _export_to_excel(data, out_path)

__all__ = ["export_to_excel"]
