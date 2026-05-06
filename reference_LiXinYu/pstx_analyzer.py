# -*- coding: utf-8 -*-
"""Historical CLI entrypoint for the PSTX schematic analysis tool.

The implementation has been split into layered packages:
- core parsing/page helpers: ``pstx_core``
- review rules and project orchestration: ``pstx_rules``
- structured query helpers: ``pstx_queries``
- Excel export: ``pstx_exports``

Keep this file thin so importing it does not pull Web, Harness, or Excel-heavy
implementation details into unrelated code paths.
"""

from pstx_core.page_resolution import (
    MAIN_MODULE_PAGE_LABEL,
    USER_VISIBLE_REAL_PAGE_LABEL,
    analyze_page_mappings,
    component_user_visible_page,
    resolve_component_pages,
    _build_page_csv_index,
    _read_page_number_from_csv,
)
from pstx_core.pstx_parser import parse_all, parse_pstxnet, parse_pstxprt
from pstx_queries.project_query import query_project_data
from pstx_rules.bom import build_bom, build_total_bom
from pstx_rules.bom_option_circle import check_bom_option_circle_coverage
from pstx_rules.common import _build_analysis_scope, _infer_project_root_from_data_paths
from pstx_rules.derating import analyze_derating
from pstx_rules.drc import check_drc
from pstx_rules.network import analyze_networks
from pstx_rules.project_analysis import analyze_project_contents
from pstx_rules.resistor_bias import (
    analyze_resistors,
    _extract_pin_submodule_info,
    _extract_refdes_suffix_group,
    _net_is_gnd,
    _parse_ohms,
)


def export_to_excel(data: dict, out_path: str) -> str:
    """Export analysis data to Excel using the split export layer."""
    from pstx_exports.excel import export_to_excel as _export_to_excel

    return _export_to_excel(data, out_path)


def main() -> int:
    """Launch the local UI historical entrypoint."""
    from pstx_apps.local_ui import main as _local_ui_main

    return _local_ui_main()


if __name__ == '__main__':
    raise SystemExit(main())
