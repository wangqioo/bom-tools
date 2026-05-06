# -*- coding: utf-8 -*-
"""Tests for low-level PSTX Packager-XL text parsing."""

import unittest

from pstx_core import pages as pstx_pages
from pstx_core import pstx_parser
from pstx_core.pstx_parser import get_comp_type
from tests.pstx_test_fixtures import sample_part_block, split_symbol_part_block


class ParserTests(unittest.TestCase):

    def test_get_comp_type_treats_power_prefixed_passives_as_passive(self):
        self.assertEqual('CAP', get_comp_type('PC16A10', ''))
        self.assertEqual('RES', get_comp_type('PR10A1', ''))
        self.assertEqual('IND', get_comp_type('PL2A5', ''))
        self.assertEqual('IND', get_comp_type('PFB3A2', ''))
        self.assertEqual('CONN', get_comp_type('P1', ''))

    def test_parse_pstxprt_handles_marker_at_file_start(self):
        content = (
            "PART_NAME\n"
            "C1 'CAP_0402'\n"
            "VALUE='1uF'\n"
            "PACKAGE='0402'\n"
            "DRAWING='SCH_PAGE1'\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertIn('C1', components)
        self.assertEqual('1uF', components['C1']['value'])

    def test_parse_pstxprt_keeps_component_xy_center(self):
        content = (
            "PART_NAME\n"
            "R1 'RES_0402':\n"
            "VALUE='10k'\n"
            "PACKAGE='0402'\n"
            "XY='(-4600,3650)'\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertEqual('(-4600,3650)', components['R1']['xy'])
        self.assertEqual(-4600.0, components['R1']['xy_x'])
        self.assertEqual(3650.0, components['R1']['xy_y'])

    def test_parse_pstxprt_handles_crlf_newlines(self):
        content = (
            "PART_NAME\r\n"
            "C1 'CAP_0402'\r\n"
            "VALUE='1uF'\r\n"
            "PACKAGE='0402'\r\n"
            "DRAWING='SCH_PAGE1'\r\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertIn('C1', components)

    def test_parse_pstxprt_normalizes_page_tokens_with_separator_and_suffix(self):
        content = (
            "PART_NAME\n"
            "U1 'IC_CPU'\n"
            "DRAWING='ROOT/PAGE_02A'\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertEqual('', components['U1']['page'])
        self.assertEqual('PAGE2A', components['U1']['page_logical'])

    def test_parse_pstxprt_preserves_hierarchical_page_chain(self):
        content = (
            "PART_NAME\n"
            "U1 'IC_CPU'\n"
            "DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1"
            "@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertEqual('', components['U1']['page'])
        self.assertEqual('PAGE242', components['U1']['page_logical'])

    def test_parse_pstxprt_prefers_section_path_over_submodule_drawing_for_page_source(self):
        content = (
            "PART_NAME\n"
            "C1A104 'CAP_HDL-HQ17101005HS0,100NF,10%,0402,X7R,50V':\n"
            "SECTION_NUMBER 1\n"
            " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70"
            "@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1_I17"
            "@HQ_CAP.CAP_HDL(CHIPS)':\n"
            " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70"
            "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page1_i17"
            "@hq_cap.cap_hdl(chips)',\n"
            " PATH='I17',\n"
            " DRAWING='@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1',\n"
            " PHYS_PAGE='1',\n"
            " BOM_OPTION='DEPOP',\n"
            " PACKAGE='0402',\n"
            " HQ_CODE='HQ17101005HS0',\n"
            " VALUE='100NF'\n"
        )
        components = pstx_parser.parse_pstxprt(content)
        self.assertEqual('section_path', components['C1A104']['page_path_source'])
        self.assertEqual(
            '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70'
            '@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1_I17'
            '@HQ_CAP.CAP_HDL(CHIPS)',
            components['C1A104']['page_path_raw'],
        )
        self.assertEqual('PAGE144', components['C1A104']['page_logical'])

    def test_parse_pstxnet_handles_marker_at_file_start(self):
        content = (
            "NET_NAME\n"
            "'P1V8'\n"
            "NODE_NAME C1 1\n"
            "'POS':\n"
        )
        nets = pstx_parser.parse_pstxnet(content)
        self.assertIn('P1V8', nets)
        self.assertEqual('POS', nets['P1V8'][0]['pin_name'])

    def test_parse_pstxnet_finds_pin_name_beyond_fixed_window(self):
        content = (
            "\nNET_NAME\n"
            "'NET1'\n"
            "NODE_NAME U1 1\n"
            f"{'X' * 220}'GPIO1':\n"
        )
        nets = pstx_parser.parse_pstxnet(content)
        self.assertEqual('GPIO1', nets['NET1'][0]['pin_name'])

    def test_parse_pstxprt_keeps_c_path_for_logical_and_p_path_for_real(self):
        components = pstx_parser.parse_pstxprt(sample_part_block())
        comp = components['C1A104']
        self.assertEqual('section_path', comp['page_path_logical_source'])
        self.assertEqual('p_path', comp['page_path_real_source'])
        self.assertEqual('PAGE144', comp['page_logical'])
        self.assertEqual('PAGE1', comp['page_submodule_real'])
        self.assertTrue(comp['page_path_real_raw'].startswith('@gpu_2sw_board_lib.gpu_2sw_board'))

    def test_parse_pstxprt_preserves_split_symbol_sections(self):
        components = pstx_parser.parse_pstxprt(split_symbol_part_block())
        comp = components['U46']
        self.assertEqual('TRUE', comp['split_inst'])
        self.assertEqual('HQ11112042009', comp['hq_code'])
        self.assertEqual('LCMXO3LF-9400C-5BG484C', comp['value'])
        self.assertEqual(2, comp['section_count'])
        self.assertEqual('section_path', comp['page_path_logical_source'])
        self.assertEqual('PAGE152', comp['page_logical'])
        self.assertEqual(['1', '2'], [section['section_number'] for section in comp['sections']])
        self.assertEqual(['HQ11112042009', 'HQ11112042009'], [section['hq_code'] for section in comp['sections']])
        self.assertEqual(['PAGE152', 'PAGE151'], [section['page_logical'] for section in comp['sections']])
        self.assertEqual('PAGE130', pstx_pages.extract_top_level_page(comp['sections'][1]['page_path_real_raw']))

    def test_parse_pstxprt_infers_split_symbol_hq_code_from_part_metadata(self):
        content = (
            "PART_NAME\n"
            "U46 'LCMXO3LF_9400C_HDL-HQ11112042009,LCMXO3LF-9400C-5BG484C':;\n"
            "SECTION_NUMBER 1\n"
            " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE152_I2"
            "@HQ_IC.LCMXO3LF_9400C_HDL(CHIPS)':\n"
            " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page152_i2"
            "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
            " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page131_i2"
            "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
            " SPLIT_INST='TRUE',\n"
            " LOCATION='U46',\n"
            " CDS_PART_NAME='LCMXO3LF_9400C_HDL-HQ11112042009,LCMXO3LF-9400C-5BG484C';\n"
            "SECTION_NUMBER 2\n"
            " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE151_I98"
            "@HQ_IC.LCMXO3LF_9400C_HDL(CHIPS)':\n"
            " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page151_i98"
            "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
            " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page130_i98"
            "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
            " SPLIT_INST='TRUE',\n"
            " LOCATION='U46';\n"
        )

        comp = pstx_parser.parse_pstxprt(content)['U46']

        self.assertEqual('HQ11112042009', comp['hq_code'])
        self.assertEqual('LCMXO3LF-9400C-5BG484C', comp['value'])
        self.assertEqual('HQ11112042009', comp['sections'][0]['hq_code'])

    def test_parse_all_links_component_pin_nets(self):
        components, nets, comp_nets = pstx_parser.parse_all(
            "PART_NAME\nR1 'RES_0402'\nVALUE='10K'\n",
            "NET_NAME\n'NET_A'\nNODE_NAME R1 1\n'1':\n",
        )
        self.assertIn('R1', components)
        self.assertEqual('NET_A', components['R1']['nets']['1'])
        self.assertEqual('NET_A', comp_nets['R1']['1'])
        self.assertEqual('R1', nets['NET_A'][0]['refdes'])


if __name__ == '__main__':
    unittest.main()
