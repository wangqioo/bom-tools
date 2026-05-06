# -*- coding: utf-8 -*-
"""Public facade for Feishu BOM online gateway, mapping, sync, and cache APIs.

The implementation is split across focused modules in this package. Keep this
facade as the stable import surface for Web, Harness, and tests.
"""

from __future__ import annotations

from pstx_integrations.feishu.cache_store import (
    build_feishu_bom_status,
    build_feishu_database_overview,
    create_feishu_cache_row,
    delete_feishu_cache_library,
    delete_feishu_cache_row,
    get_feishu_cache_row,
    get_feishu_cache_rows,
    get_saved_feishu_field_order,
    match_rows_with_feishu_cache,
    update_feishu_cache_row,
)
from pstx_integrations.feishu.client import FeishuBomClient
from pstx_integrations.feishu.common import (
    CACHE_FILE_NAME,
    CONFIG_FILE_NAME,
    DATA_DIR_ENV,
    DEFAULT_BASE_URL,
    DEFAULT_DATA_DIR_NAME,
    DEFAULT_ORIGIN,
    FEISHU_LOG_FILE_ENV,
    FEISHU_LOG_PAYLOAD_ENV,
    FEISHU_PARSE_LOG_FILE_ENV,
    FEISHU_PARSE_LOG_ROWS_ENV,
    FeishuBomError,
    _default_column_range,
    _safe_cell_str,
    build_sheet_value_range,
    extract_spreadsheet_token,
    feishu_debug_log_path,
    feishu_parse_debug_log_path,
    resolve_data_dir,
)
from pstx_integrations.feishu.mapping import (
    build_feishu_mapping_from_headers,
    suggest_feishu_mapping_from_preview,
)
from pstx_integrations.feishu.sync import (
    fetch_feishu_sheet_list,
    preview_feishu_sheet,
    sync_feishu_library,
)


__all__ = [
    "CACHE_FILE_NAME",
    "CONFIG_FILE_NAME",
    "DATA_DIR_ENV",
    "DEFAULT_BASE_URL",
    "DEFAULT_DATA_DIR_NAME",
    "DEFAULT_ORIGIN",
    "FEISHU_LOG_FILE_ENV",
    "FEISHU_LOG_PAYLOAD_ENV",
    "FEISHU_PARSE_LOG_FILE_ENV",
    "FEISHU_PARSE_LOG_ROWS_ENV",
    "FeishuBomClient",
    "FeishuBomError",
    "_default_column_range",
    "_safe_cell_str",
    "build_feishu_bom_status",
    "build_feishu_database_overview",
    "build_feishu_mapping_from_headers",
    "build_sheet_value_range",
    "create_feishu_cache_row",
    "delete_feishu_cache_library",
    "delete_feishu_cache_row",
    "extract_spreadsheet_token",
    "feishu_debug_log_path",
    "feishu_parse_debug_log_path",
    "fetch_feishu_sheet_list",
    "get_feishu_cache_row",
    "get_feishu_cache_rows",
    "get_saved_feishu_field_order",
    "match_rows_with_feishu_cache",
    "preview_feishu_sheet",
    "resolve_data_dir",
    "suggest_feishu_mapping_from_preview",
    "sync_feishu_library",
    "update_feishu_cache_row",
]
