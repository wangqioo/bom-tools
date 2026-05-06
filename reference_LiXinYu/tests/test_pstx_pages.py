# -*- coding: utf-8 -*-
"""Tests for user-visible PSTX page resolution."""

import tempfile
import unittest
from pathlib import Path

from pstx_core import page_resolution
from pstx_core import pages as pstx_pages
from pstx_core import pstx_parser
from tests.pstx_test_fixtures import (
    deep_hierarchy_part_block,
    pex90144_part_block,
    sample_part_block,
    split_symbol_part_block,
)


class PageResolutionTests(unittest.TestCase):

    def test_core_page_entrypoint_exports_public_api(self):
        self.assertTrue(callable(pstx_pages.build_module_order_index))
        self.assertTrue(callable(pstx_pages.resolve_component_page_info))
        self.assertFalse(hasattr(pstx_pages, "_read_text_file"))
        self.assertFalse(Path("pstx_pages.py").exists())
        parsed = pstx_pages.parse_page_map_line('144 114 TOP')
        self.assertEqual({'logical_page': 'PAGE144', 'real_page': 'PAGE114', 'page_name': 'TOP'}, parsed)

    def test_resolve_component_pages_prefers_outer_schematic_segment(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': (
                    '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
                    '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
                    '@HQ_IC.TPS259261DRCR_11P_HDL(CHIPS)'
                ),
                'page': 'PAGE242 / PAGE1',
            },
        }
        warnings = page_resolution.resolve_component_pages(components)
        self.assertEqual([], warnings)
        self.assertEqual('', components['U1']['page'])
        self.assertEqual('PAGE242', components['U1']['page_logical'])
        self.assertEqual('PAGE242', components['U1']['page_raw'])
        self.assertEqual('', components['U1']['page_real'])
        self.assertEqual('GPU_2SW_BOARD:PAGE242', components['U1']['page_context'])
        self.assertEqual('', components['U1']['page_context_real'])
        self.assertEqual('drawing', components['U1']['page_source'])
        self.assertEqual('none', components['U1']['page_real_source'])
        self.assertEqual('', components['U1']['page_mapping_ok'])

    def test_resolve_component_pages_uses_page_csv_mapping_for_outer_schematic_segment(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': (
                    '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
                    '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
                    '@HQ_IC.TPS259261DRCR_11P_HDL(CHIPS)'
                ),
                'page': 'PAGE242 / PAGE1',
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            top_sch = root / 'GPU_2SW_BOARD' / 'sch_1'
            top_sch.mkdir(parents=True)
            (top_sch / 'page518.csv').write_text('NAME,PAGE_NUMBER\nTOP,242\n', encoding='utf-8')

            child_sch = root / 'HQPWR_EFUSE_TPS259260_12VIN_4A' / 'sch_1'
            child_sch.mkdir(parents=True)
            (child_sch / 'page12.csv').write_text('NAME,PAGE_NUMBER\nCHILD,1\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        self.assertEqual('PAGE518', components['U1']['page'])
        self.assertEqual('PAGE242', components['U1']['page_logical'])
        self.assertEqual('PAGE518', components['U1']['page_real'])
        self.assertEqual('drawing', components['U1']['page_source'])
        self.assertEqual('page_csv', components['U1']['page_real_source'])
        self.assertEqual('是', components['U1']['page_mapping_ok'])
        self.assertEqual('GPU_2SW_BOARD:PAGE242', components['U1']['page_context'])
        self.assertEqual('GPU_2SW_BOARD:PAGE518', components['U1']['page_context_real'])

    def test_resolve_component_pages_ignores_child_page_when_child_segment_appears_first(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': (
                    '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
                    '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
                    '@HQ_IC.TPS259261DRCR_11P_HDL(CHIPS)'
                ),
                'page': 'PAGE1 / PAGE242',
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / 'GPU_2SW_BOARD'
            root.mkdir(parents=True)
            top_sch = root / 'sch_1'
            top_sch.mkdir(parents=True)
            (top_sch / 'page518.csv').write_text('"PAGE_NUMBER" = 242;', encoding='utf-8')
            (top_sch / 'page12.csv').write_text('"PAGE_NUMBER" = 1;', encoding='utf-8')

            child_sch = root / 'HQPWR_EFUSE_TPS259260_12VIN_4A' / 'sch_1'
            child_sch.mkdir(parents=True)
            (child_sch / 'page12.csv').write_text('"PAGE_NUMBER" = 1;', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        self.assertEqual('PAGE242', components['U1']['page_logical'])
        self.assertEqual('PAGE518', components['U1']['page'])
        self.assertEqual('PAGE518', components['U1']['page_real'])
        self.assertEqual('page_csv', components['U1']['page_real_source'])

    def test_resolve_component_pages_reads_direct_project_root_sch1_pages(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1',
                'page': 'PAGE242',
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            direct_sch = root / 'sch_1'
            direct_sch.mkdir(parents=True)
            (direct_sch / 'page518.csv').write_text('NAME,PAGE_NUMBER\nTOP,242\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        self.assertEqual('PAGE518', components['U1']['page'])
        self.assertEqual('PAGE518', components['U1']['page_real'])
        self.assertEqual('page_csv', components['U1']['page_real_source'])

    def test_resolve_component_pages_uses_global_page_number_mapping_without_cell_filter(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1',
                'page': 'PAGE242',
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            unrelated_sch = root / 'UNRELATED_BLOCK' / 'sch_1'
            unrelated_sch.mkdir(parents=True)
            (unrelated_sch / 'page900.csv').write_text('NAME,PAGE_NUMBER\nX,242\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        self.assertEqual('PAGE900', components['U1']['page'])
        self.assertEqual('PAGE900', components['U1']['page_real'])
        self.assertEqual('drawing', components['U1']['page_source'])
        self.assertEqual('page_csv', components['U1']['page_real_source'])
        self.assertEqual('是', components['U1']['page_mapping_ok'])

    def test_resolve_component_pages_leaves_real_page_blank_when_logical_page_hits_multiple_real_pages(self):
        components = {
            'U1': {
                'refdes': 'U1',
                'drawing': '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1',
                'page': 'PAGE242',
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            first_sch = root / 'BLOCK_A' / 'sch_1'
            first_sch.mkdir(parents=True)
            (first_sch / 'page518.csv').write_text('NAME,PAGE_NUMBER\nA,242\n', encoding='utf-8')

            second_sch = root / 'BLOCK_B' / 'sch_1'
            second_sch.mkdir(parents=True)
            (second_sch / 'page900.csv').write_text('NAME,PAGE_NUMBER\nB,242\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertTrue(any('PAGE242 同时命中多个页码' in warning for warning in warnings))
        self.assertEqual('', components['U1']['page'])
        self.assertEqual('', components['U1']['page_real'])
        self.assertEqual('page_csv_ambiguous', components['U1']['page_real_source'])
        self.assertEqual('否', components['U1']['page_mapping_ok'])

    def test_build_page_csv_index_reports_scanned_but_unparsed_csv_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            direct_sch = root / 'sch_1'
            direct_sch.mkdir(parents=True)
            (direct_sch / 'page518.csv').write_text('NAME,VALUE\nTOP,242\n', encoding='utf-8')

            index = pstx_pages.build_page_csv_index(str(root))

        self.assertEqual(1, index['scanned'])
        self.assertEqual(1, index['matched_root_sch1'])
        self.assertEqual(0, index['count'])
        self.assertTrue(any('没有读出任何 PAGE_NUMBER' in warning for warning in index['warnings']))

    def test_read_page_number_from_csv_prefers_exact_assignment_format_with_semicolon(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = Path(temp_dir) / 'page518.csv'
            csv_path.write_text('"PAGE_NUMBER" = 242;\nNAME = TOP;\n', encoding='utf-8')

            page_number = pstx_pages.read_page_number_from_csv(csv_path)

        self.assertEqual('PAGE242', page_number)

    def test_read_page_number_from_utf16_csv_with_assignment_format(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = Path(temp_dir) / 'page518.csv'
            csv_path.write_text('"PAGE_NUMBER" = "242";\n"NAME" = "TOP";\n', encoding='utf-16')

            page_number = pstx_pages.read_page_number_from_csv(csv_path)

        self.assertEqual('PAGE242', page_number)

    def test_build_page_csv_index_reads_utf16_assignment_page_csv(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            direct_sch = root / 'sch_1'
            direct_sch.mkdir(parents=True)
            (direct_sch / 'page518.csv').write_text(
                '"PAGE_NUMBER" = 242;\n"NAME" = "TOP";\n',
                encoding='utf-16',
            )

            index = pstx_pages.build_page_csv_index(str(root))

        self.assertEqual(1, index['count'])
        self.assertIn('PAGE242', index['by_logical_page'])
        self.assertEqual('PAGE518', index['by_logical_page']['PAGE242'][0]['resolved_page'])

    def test_analyze_page_mappings_marks_reverse_conflict_as_not_one_to_one(self):
        page_index = {
            'by_logical_page': {
                'PAGE242': [{'resolved_page': 'PAGE518', 'cell': 'TOP', 'path': 'a/page518.csv'}],
                'PAGE300': [{'resolved_page': 'PAGE518', 'cell': 'TOP', 'path': 'b/page518.csv'}],
            }
        }
        result = page_resolution.analyze_page_mappings(page_index)
        row_map = {row['主模块页']: row for row in result['rows']}
        self.assertEqual('否', row_map['PAGE242']['是否一一对应'])
        self.assertEqual('页码对应多个主模块页', row_map['PAGE242']['状态'])
        self.assertTrue(any('PAGE518 同时被多个主模块页复用' in warning for warning in result['warnings']))

    def test_resolve_component_pages_aggregates_split_symbol_user_visible_pages(self):
        components = pstx_parser.parse_pstxprt(split_symbol_part_block())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page130.csv').write_text('"PAGE_NUMBER" = 151;\n', encoding='utf-8')
            (sch_dir / 'page131.csv').write_text('"PAGE_NUMBER" = 152;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('151 130 FPGA_B\n152 131 FPGA_A\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['U46']
        self.assertEqual('PAGE131', comp['page_real'])
        self.assertEqual('PAGE152, PAGE151', comp['page_logical_pages'])
        self.assertEqual('PAGE131, PAGE130', comp['page_real_pages'])
        self.assertEqual('PAGE131, PAGE130', comp['page_user_visible_pages'])
        self.assertEqual('PAGE131, PAGE130', page_resolution.component_user_visible_page(comp))
        self.assertEqual('PAGE130', comp['sections'][1]['page_real'])
        self.assertEqual('PAGE130', page_resolution.component_user_visible_page(comp['sections'][1]))

    def test_parse_page_map_line_reads_logical_then_real_then_name(self):
        parsed = pstx_pages.parse_page_map_line('144 114 TOP')
        self.assertEqual(
            {
                'logical_page': 'PAGE144',
                'real_page': 'PAGE114',
                'page_name': 'TOP',
            },
            parsed,
        )

    def test_parse_page_map_line_keeps_full_name_segment_after_real_page(self):
        parsed = pstx_pages.parse_page_map_line('144   114   TOP MAIN BLOCK')
        self.assertEqual('PAGE144', parsed['logical_page'])
        self.assertEqual('PAGE114', parsed['real_page'])
        self.assertEqual('TOP MAIN BLOCK', parsed['page_name'])

    def test_build_page_map_index_reads_name_segment_with_spaces(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page.map').write_text('144 114 TOP MAIN BLOCK\n', encoding='utf-8')

            index = pstx_pages.build_page_map_index(str(root))

        entries = index['by_logical_page']['PAGE144']
        self.assertEqual(1, len(entries))
        self.assertEqual('PAGE114', entries[0]['resolved_page'])
        self.assertEqual('TOP MAIN BLOCK', entries[0]['page_name'])

    def test_resolve_component_pages_prefers_p_path_and_computes_mapped_submodule_page(self):
        components = pstx_parser.parse_pstxprt(sample_part_block())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')
            (root / 'module_order').write_text(
                'Version 15.0\n'
                'START_MODULEORDER\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70'
                '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n'
                'END_MODULEORDER\n',
                encoding='utf-8',
            )

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['C1A104']
        self.assertEqual('PAGE144', comp['page_logical'])
        self.assertEqual('PAGE114', comp['page_real'])
        self.assertEqual('PAGE114', comp['page'])
        self.assertEqual('PAGE1', comp['page_submodule_real'])
        self.assertEqual('PAGE177', comp['page_submodule_mapped'])
        self.assertEqual('p_path', comp['page_real_source'])
        self.assertEqual('是', comp['page_mapping_ok'])

    def test_page_map_cross_check_prefers_root_sch1_over_child_sch1(self):
        components = pstx_parser.parse_pstxprt(
            "PART_NAME\n"
            "U1 'IC_CPU':\n"
            " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70@hq_ic.cpu(chips)',\n"
            " DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144'\n"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir(parents=True)
            (root / 'sch_1' / 'page.map').write_text('144 114 ROOT_TOP\n', encoding='utf-8')
            (root / 'sch_1' / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            child_sch = root / 'reuse_block' / 'sch_1'
            child_sch.mkdir(parents=True)
            (child_sch / 'page.map').write_text('144 999 CHILD_SHOULD_NOT_WIN\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['U1']
        self.assertEqual('PAGE144', comp['page_logical'])
        self.assertEqual('PAGE114', comp['page_real'])
        self.assertEqual('page_map', comp['page_real_source'])
        self.assertEqual('unique', comp['page_map_state'])
        self.assertEqual('是', comp['page_mapping_ok'])

    def test_module_order_prefers_logical_path_key_when_p_path_exists(self):
        components = pstx_parser.parse_pstxprt(sample_part_block())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')
            (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            (root / 'module_order').write_text(
                'Version 15.0\n'
                'START_MODULEORDER\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70'
                '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n'
                'END_MODULEORDER\n',
                encoding='utf-8',
            )

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['C1A104']
        self.assertEqual('PAGE114', comp['page_real'])
        self.assertEqual('PAGE177', comp['page_submodule_mapped'])
        self.assertEqual('unique', comp['module_order_state'])
        self.assertIn('@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70', comp['module_order_key'])

    def test_module_order_mapping_rejects_submodule_page_out_of_range(self):
        part_block = sample_part_block().replace('page1_i17', 'page35_i17').replace('PAGE1_I17', 'PAGE35_I17')
        components = pstx_parser.parse_pstxprt(part_block)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')
            (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            (root / 'module_order').write_text(
                'Version 15.0\n'
                'START_MODULEORDER\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70'
                '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n'
                'END_MODULEORDER\n',
                encoding='utf-8',
            )

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['C1A104']
        self.assertEqual('PAGE35', comp['page_submodule_real'])
        self.assertEqual('', comp['page_submodule_mapped'])
        self.assertEqual('local_page_out_of_range', comp['module_order_state'])
        self.assertIn('超出 module_order 页数 34', comp['page_submodule_mapping_note'])

    def test_non_submodule_component_uses_real_page_as_submodule_mapped_page(self):
        part_block = (
            "PART_NAME\n"
            "R1 'RES_HDL-HQ00000001,10K,1%,0402':\n"
            "SECTION_NUMBER 1\n"
            " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70"
            "@HQ_RES.RES_HDL(CHIPS)':\n"
            " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70"
            "@hq_res.res_hdl(chips)',\n"
            " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70"
            "@hq_res.res_hdl(chips)',\n"
            " DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144',\n"
            " PHYS_PAGE='114',\n"
            " HQ_CODE='HQ00000001',\n"
            " VALUE='10K'\n"
        )
        components = pstx_parser.parse_pstxprt(part_block)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['R1']
        self.assertEqual('PAGE144', comp['page_logical'])
        self.assertEqual('PAGE114', comp['page_real'])
        self.assertEqual('', comp['page_submodule_real'])
        self.assertEqual('PAGE114', comp['page_submodule_mapped'])
        self.assertEqual('not_submodule', comp['module_order_state'])
        self.assertIn('非子模块元件', comp['page_submodule_mapping_note'])

    def test_resolve_component_pages_maps_deepest_module_order_for_nested_reuse(self):
        components = pstx_parser.parse_pstxprt(deep_hierarchy_part_block())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')
            (root / 'module_order').write_text(
                'Version 15.0\n'
                'START_MODULEORDER\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70'
                '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70'
                '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page3_i17'
                '@gpu_2sw_board_lib.grand_child_block(sch_1) 0 1 250 10 0\n'
                'END_MODULEORDER\n',
                encoding='utf-8',
            )

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['C9A001']
        self.assertEqual('PAGE114', comp['page_real'])
        self.assertEqual('PAGE2', comp['page_submodule_real'])
        self.assertEqual('PAGE251', comp['page_submodule_mapped'])
        self.assertIn('@GPU_2SW_BOARD_LIB.GRAND_CHILD_BLOCK(SCH_1)', comp['module_order_key'])
        self.assertEqual('PAGE2', comp['module_order_local_page'])

    def test_resolve_component_pages_reads_module_order_dat_and_maps_pex90144_sample(self):
        components = pstx_parser.parse_pstxprt(pex90144_part_block())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page24.csv').write_text('"PAGE_NUMBER" = 112;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('112 24 TOP\n', encoding='utf-8')
            (root / 'module_order.dat').write_text(
                'Version 15.0\n'
                'START_MODULEORDER\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1) 0 1 1 176 0\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page112_i167'
                '@gpu_2sw_board_lib.pex90144_cbb_v1(sch_1) 0 1 177 34 0\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page4_i1'
                '@gpu_2sw_board_lib.pex90144_cbb_v1(sch_1) 0 1 211 34 1\n'
                '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page239_i64'
                '@gpu_2sw_board_lib.i2c_sw_tpt29548_000_cbb_v3(sch_1) 0 1 245 1 0\n'
                'END_MODULEORDER\n',
                encoding='utf-8',
            )

            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        comp = components['C1A101']
        self.assertEqual('PAGE112', comp['page_logical'])
        self.assertEqual('PAGE24', comp['page_real'])
        self.assertEqual('PAGE1', comp['page_submodule_real'])
        self.assertEqual('PAGE177', comp['page_submodule_mapped'])
        self.assertEqual('p_path', comp['page_real_source'])
        self.assertEqual('unique', comp['module_order_state'])
        self.assertIn('@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE112_I167', comp['module_order_key'])

    def test_module_order_deduplicates_identical_dat_and_extensionless_files(self):
        components = pstx_parser.parse_pstxprt(pex90144_part_block())
        module_order_text = (
            'Version 15.0\n'
            'START_MODULEORDER\n'
            '@gpu_2sw_board_lib.gpu_2sw_board(sch_1) 0 1 1 176 0\n'
            '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page112_i167'
            '@gpu_2sw_board_lib.pex90144_cbb_v1(sch_1) 0 1 177 34 0\n'
            'END_MODULEORDER\n'
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir(parents=True)
            (sch_dir / 'page24.csv').write_text('"PAGE_NUMBER" = 112;\n', encoding='utf-8')
            (sch_dir / 'page.map').write_text('112 24 TOP\n', encoding='utf-8')
            (root / 'module_order.dat').write_text(module_order_text, encoding='utf-8')
            (root / 'module_order').write_text(module_order_text, encoding='utf-8')

            index = pstx_pages.build_module_order_index(str(root))
            warnings = page_resolution.resolve_component_pages(components, str(root))

        self.assertEqual([], warnings)
        self.assertEqual(2, index['count'])
        self.assertEqual(2, index['duplicate_count'])
        comp = components['C1A101']
        self.assertEqual('PAGE177', comp['page_submodule_mapped'])
        self.assertEqual('unique', comp['module_order_state'])


if __name__ == '__main__':
    unittest.main()
