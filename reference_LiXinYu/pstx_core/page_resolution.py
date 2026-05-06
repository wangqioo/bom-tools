# -*- coding: utf-8 -*-
"""User-visible schematic page resolution helpers.

This module keeps page/map/module_order display logic out of rule code.
"""

from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pstx_core import pages as page_logic

USER_VISIBLE_REAL_PAGE_LABEL = '页码'
MAIN_MODULE_PAGE_LABEL = '主模块页'


def _normalize_page_label(page_label: str) -> str:
    return page_logic.normalize_page_label(page_label)


def _extract_top_level_logical_page(path_text: str) -> str:
    return page_logic.extract_top_level_page(path_text)


def _extract_section_paths(block_text: str) -> List[Dict[str, str]]:
    return page_logic.extract_section_paths(block_text)


def _select_component_page_source(block_text: str, attrs: Dict[str, str]) -> Tuple[str, str]:
    sources = page_logic.select_component_page_sources(block_text, attrs)
    return sources.get('logical_path_raw', ''), sources.get('logical_path_source', 'none')


def _extract_page_number_from_text(text: str) -> str:
    return page_logic.extract_page_number_from_text(text)


def _read_page_number_from_csv(csv_path: Path) -> str:
    return page_logic.read_page_number_from_csv(Path(csv_path))


def _build_page_csv_index(project_root: str) -> Dict[str, object]:
    return page_logic.build_page_csv_index(project_root)


def analyze_page_mappings(page_index: Optional[Dict[str, object]]) -> Dict[str, object]:
    return page_logic.build_page_mapping_rows(None, page_index)


def _prepare_page_resolution(project_root: str) -> Dict[str, object]:
    page_csv_index = page_logic.build_page_csv_index(project_root) if project_root else None
    page_map_index = page_logic.build_page_map_index(project_root) if project_root else None
    module_order_index = page_logic.build_module_order_index(project_root) if project_root else None
    page_mapping = page_logic.build_page_mapping_rows(page_map_index, page_csv_index)

    warnings: List[str] = []
    for index in [page_csv_index, page_map_index, module_order_index]:
        if index:
            warnings.extend(index.get('warnings', []))
    warnings.extend(page_mapping.get('warnings', []))
    return {
        'page_csv_index': page_csv_index,
        'page_map_index': page_map_index,
        'module_order_index': module_order_index,
        'page_mapping': page_mapping,
        'warnings': warnings,
    }


def prepare_page_resolution(project_root: str) -> Dict[str, object]:
    """Build page resolution context for a project root."""
    return _prepare_page_resolution(project_root)


def _apply_page_info_fields(target: Dict[str, object], page_info: Dict[str, str]) -> None:
    display_page = page_info.get('page_real', '')
    target['page'] = display_page
    target['page_logical'] = page_info.get('page_logical', '')
    target['page_raw'] = page_info.get('page_logical', '')
    target['page_real'] = display_page
    target['page_submodule_real'] = page_info.get('page_submodule_real', '')
    target['page_submodule_mapped'] = page_info.get('page_submodule_mapped', '')
    target['page_context'] = page_info.get('page_context', '')
    target['page_context_real'] = page_info.get('page_context_real', '')
    target['page_source'] = page_info.get('page_logical_source', '')
    target['page_real_source'] = page_info.get('page_real_source', '')
    target['page_validation_status'] = page_info.get('page_validation_status', '')
    target['page_mapping_ok'] = page_info.get('page_mapping_ok', '')
    target['page_mapping_note'] = page_info.get('page_validation_note', '')
    target['page_validation_note'] = page_info.get('page_validation_note', '')
    target['page_map_real'] = page_info.get('page_map_real', '')
    target['page_map_state'] = page_info.get('page_map_state', '')
    target['page_csv_real'] = page_info.get('page_csv_real', '')
    target['page_csv_state'] = page_info.get('page_csv_state', '')
    target['module_order_key'] = page_info.get('module_order_key', '')
    target['module_order_state'] = page_info.get('module_order_state', '')
    target['module_order_local_page'] = page_info.get('module_order_local_page', '')
    target['module_order_start_page'] = page_info.get('module_order_start_page', '')
    target['module_order_page_count'] = page_info.get('module_order_page_count', '')
    target['page_submodule_mapping_note'] = page_info.get('page_submodule_mapping_note', '')


def _join_unique_page_labels(values: List[str]) -> str:
    seen = set()
    pages: List[str] = []
    for value in values:
        text = _normalize_page_label(str(value or '').strip())
        if not text:
            continue
        if text in seen:
            continue
        seen.add(text)
        pages.append(text)
    return ', '.join(pages)


def _apply_component_pages(components: Dict[str, dict],
                           page_context: Optional[Dict[str, object]]) -> None:
    context = page_context or {}
    page_map_index = context.get('page_map_index')
    page_csv_index = context.get('page_csv_index')
    module_order_index = context.get('module_order_index')
    for comp in components.values():
        page_info = page_logic.resolve_component_page_info(
            comp,
            page_map_index=page_map_index,
            page_csv_index=page_csv_index,
            module_order_index=module_order_index,
        )
        _apply_page_info_fields(comp, page_info)

        sections = comp.get('sections', [])
        if isinstance(sections, list) and sections:
            for section in sections:
                if not isinstance(section, dict):
                    continue
                section_info = page_logic.resolve_component_page_info(
                    section,
                    page_map_index=page_map_index,
                    page_csv_index=page_csv_index,
                    module_order_index=module_order_index,
                )
                _apply_page_info_fields(section, section_info)
            comp['page_section_count'] = len([section for section in sections if isinstance(section, dict)])
            comp['page_logical_pages'] = _join_unique_page_labels([
                section.get('page_logical', '') for section in sections if isinstance(section, dict)
            ])
            comp['page_real_pages'] = _join_unique_page_labels([
                section.get('page_real', '') for section in sections if isinstance(section, dict)
            ])
            comp['page_user_visible_pages'] = _join_unique_page_labels([
                component_user_visible_page(section) for section in sections if isinstance(section, dict)
            ])


def apply_component_pages(components: Dict[str, dict],
                          page_context: Optional[Dict[str, object]]) -> None:
    """Apply user-visible page fields to parsed components and split sections."""
    _apply_component_pages(components, page_context)


def resolve_component_pages(components: Dict[str, dict], project_root: str = '') -> List[str]:
    page_context = _prepare_page_resolution(project_root)
    _apply_component_pages(components, page_context)
    return list(page_context.get('warnings', []))


def _component_logical_page(comp: Dict) -> str:
    return _normalize_page_label(
        comp.get('page_logical', '')
        or comp.get('page_raw', '')
        or comp.get('page', '')
    )


def _component_display_page(comp: Dict) -> str:
    return _normalize_page_label(comp.get('page_real', '') or comp.get('page', ''))


def _component_submodule_mapped_page(comp: Dict) -> str:
    mapped_page = _normalize_page_label(comp.get('page_submodule_mapped', ''))
    if mapped_page:
        return mapped_page
    if _normalize_page_label(comp.get('page_submodule_real', '')) or comp.get('module_order_key', ''):
        return ''
    return _component_display_page(comp)


def component_user_visible_page(comp: Dict) -> str:
    """Return the schematic page number that users see in the overall design."""
    multi_page = str(comp.get('page_user_visible_pages', '') or '').strip()
    if multi_page:
        return multi_page
    return _component_submodule_mapped_page(comp)


def _component_page_fields(comp: Dict) -> Dict[str, str]:
    visible_page = component_user_visible_page(comp)
    main_module_page = _component_logical_page(comp)
    return {
        '页面': visible_page,
        USER_VISIBLE_REAL_PAGE_LABEL: visible_page,
        MAIN_MODULE_PAGE_LABEL: main_module_page,
    }


def component_page_fields(comp: Dict) -> Dict[str, str]:
    """Return standardized page columns for rule/report rows."""
    return _component_page_fields(comp)


def summarize_module_order_page_extent(project_root: str) -> Dict[str, object]:
    """Return user-visible schematic page extent derived from module_order."""
    return page_logic.summarize_module_order_page_extent(project_root)
