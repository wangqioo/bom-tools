# -*- coding: utf-8 -*-
"""Tests for PSTX rule-layer analysis helpers."""

import tempfile
import unittest
from pathlib import Path, PureWindowsPath

from pstx_core import page_resolution
from pstx_core import pstx_parser
from pstx_rules.common import _build_analysis_scope, _infer_project_root_from_data_paths
from pstx_rules.derating import analyze_derating
from pstx_rules.drc import check_drc
from pstx_rules.network import analyze_networks
from pstx_rules.project_analysis import analyze_project_contents
from pstx_rules.resistor_bias import (
    _extract_pin_submodule_info,
    _extract_refdes_suffix_group,
    _net_is_gnd,
    _parse_ohms,
    analyze_resistors,
)
from tests.pstx_test_fixtures import make_cap, make_ic, make_res, sample_part_block


class RuleTests(unittest.TestCase):

    def test_infer_project_root_from_packaged_data_paths(self):
        prt_path = r'E:\demo\GPU_2SW_BOARD\packaged\pstxprt.dat'
        net_path = r'E:\demo\GPU_2SW_BOARD\packaged\pstxnet.dat'
        project_root = _infer_project_root_from_data_paths(prt_path, net_path)
        self.assertEqual(str(PureWindowsPath(r'E:\demo\GPU_2SW_BOARD')), project_root)

    def test_analyze_project_contents_adds_module_review_scope(self):
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

            bundle = analyze_project_contents(
                sample_part_block(),
                "NET_NAME\n'P1V8_AON'\nNODE_NAME C1A104 1\n'1':\n",
                project_name='gpu_2sw_board',
                project_root=str(root),
            )

        module_review = bundle['module_review']
        summary = module_review['summary']
        self.assertEqual(1, summary['submodule_count'])
        module_rows = {row['模块名']: row for row in module_review['module_rows']}
        submodule = module_rows['i2c_repeater_9617_cbb_v3']
        self.assertEqual('子模块', submodule['模块类型'])
        self.assertEqual('PAGE144', submodule['父级Symbol页码'])
        self.assertEqual('I70', submodule['父级Symbol实例'])
        self.assertEqual('PAGE177', submodule['起始页码'])
        self.assertEqual('PAGE210', submodule['结束页码'])
        component_row = next(row for row in module_review['component_rows'] if row['位号'] == 'C1A104')
        self.assertEqual('子模块', component_row['模块类型'])
        self.assertEqual('i2c_repeater_9617_cbb_v3', component_row['模块名'])
        self.assertEqual('PAGE177', component_row['页码'])
        self.assertEqual('PAGE144', component_row['主模块页'])
        self.assertEqual('PAGE1', component_row['子模块本地页'])

    def test_check_drc_bom_option_components_show_user_visible_page(self):
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
            page_resolution.resolve_component_pages(components, str(root))

        result = check_drc(components, {}, option_components_source=components)
        row = result['bom_option_components'][0]
        self.assertEqual('PAGE177', row['页面'])
        self.assertEqual('PAGE177', row['页码'])

    def test_build_analysis_scope_excludes_depop_components_by_default(self):
        components = {
            'U1': make_ic('U1', {'1': 'GPIO1'}),
            'R1': make_res('R1', 'P3V3', 'GPIO1', bom_option='DEPOP'),
            'R2': make_res('R2', 'P3V3', 'GPIO1'),
        }
        nets = {
            'GPIO1': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO1'},
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
                {'refdes': 'R2', 'pin': '1', 'pin_name': '1'},
            ],
        }
        active_components, active_nets, depop_refdes, excluded_refdes = _build_analysis_scope(
            components,
            nets,
            include_depop=False,
        )
        self.assertEqual(['R1'], depop_refdes)
        self.assertEqual(['R1'], excluded_refdes)
        self.assertNotIn('R1', active_components)
        self.assertEqual({'U1', 'R2'}, {node['refdes'] for node in active_nets['GPIO1']})

    def test_check_drc_keeps_bom_option_component_list_from_raw_source(self):
        analysis_components = {'U1': make_ic('U1')}
        raw_components = {
            'U1': make_ic('U1'),
            'R1': make_res('R1', 'P3V3', 'GPIO1', bom_option='DEPOP', page_real='PAGE518'),
        }
        result = check_drc(analysis_components, {}, option_components_source=raw_components)
        row_map = {row['位号']: row for row in result['bom_option_components']}
        self.assertIn('R1', row_map)
        self.assertEqual('是', row_map['R1']['是否DEPOP'])
        self.assertEqual('PAGE518', row_map['R1']['页面'])

    def test_check_drc_unnamed_nets_respect_depop_scope_for_multi_node_unnamed(self):
        full_components = {
            'R1': make_res('R1', 'UNNAMED_1', 'P3V3', bom_option='DEPOP', page_real='PAGE101'),
            'R2': make_res('R2', 'UNNAMED_1', 'P5V', bom_option='DNP', page_real='PAGE102'),
        }
        full_nets = {
            'UNNAMED_1': [
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
                {'refdes': 'R2', 'pin': '1', 'pin_name': '1'},
            ],
        }
        active_components, active_nets, _, _ = _build_analysis_scope(
            full_components,
            full_nets,
            include_depop=False,
        )

        result = check_drc(
            active_components,
            active_nets,
            single_pin_components=full_components,
            single_pin_nets=full_nets,
        )

        self.assertEqual([], result['unnamed_nets'])

    def test_check_drc_bom_option_components_use_top_level_logical_page_mapping(self):
        components = {
            'R1': {
                'refdes': 'R1',
                'part_name': 'RES_0402',
                'hq_code': '',
                'value': '10k',
                'package': '0402',
                'material': '',
                'tolerance': '',
                'voltage': '',
                'current': '',
                'power': '',
                'bom_option': 'DEPOP',
                'bom_cost': '',
                'room': '',
                'drawing': (
                    '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
                    '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
                ),
                'page': 'PAGE242 / PAGE1',
                'comp_type': 'RES',
                'nets': {'1': 'P3V3', '2': 'GPIO1'},
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
            page_resolution.resolve_component_pages(components, str(root))
        rows = check_drc({}, {}, option_components_source=components)['bom_option_components']
        self.assertEqual('PAGE518', rows[0]['页面'])

    def test_extract_refdes_suffix_group_prefers_trailing_letter_digit_cluster(self):
        self.assertEqual('A1', _extract_refdes_suffix_group('PU1A1'))
        self.assertEqual('A1', _extract_refdes_suffix_group('R1A1'))
        self.assertEqual('', _extract_refdes_suffix_group('U1'))

    def test_extract_pin_submodule_info_uses_parent_hierarchy_before_leaf_symbol(self):
        pin_name = (
            '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
            '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
            '@HQ_IC.TPS259261DRCR_11P_HDL(CHIPS)'
        )
        submodule, submodule_path = _extract_pin_submodule_info(pin_name)
        self.assertEqual('HQPWR_EFUSE_TPS259260_12VIN_4A', submodule)
        self.assertEqual('GPU_2SW_BOARD / HQPWR_EFUSE_TPS259260_12VIN_4A', submodule_path)

    def test_parse_ohms_supports_embedded_notation(self):
        self.assertEqual(4.7, _parse_ohms('4R7'))
        self.assertEqual(1500.0, _parse_ohms('1K5'))

    def test_parse_ohms_supports_ohm_word_suffixes(self):
        self.assertEqual(10, _parse_ohms('10OHM'))
        self.assertEqual(10, _parse_ohms('10OHMS'))
        self.assertEqual(10000, _parse_ohms('10KOHM'))
        self.assertEqual(4700, _parse_ohms('4.7KΩ'))

    def test_derating_does_not_infer_signal_like_pg_p1v8(self):
        components = {'C1': make_cap()}
        nets = {
            'PG_P1V8': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_derating(components, nets)
        self.assertIn('无法推断', rows[0]['状态'])
        self.assertEqual('', rows[0]['推断工作电压(V)'])

    def test_derating_requires_ground_and_single_known_positive_rail(self):
        components = {'C1': make_cap(rated='16V')}
        nets = {
            'P5V': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'P3V3': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_derating(components, nets)
        self.assertIn('未连接地', rows[0]['状态'])

    def test_custom_voltage_map_matches_prefix_boundary_only(self):
        components = {'C1': make_cap()}
        nets = {
            'PG_P1V8': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_derating(components, nets, custom_volt_map={'P1V8': 1.8})
        self.assertIn('无法推断', rows[0]['状态'])

    def test_exact_custom_voltage_map_can_override_signal_net_when_user_declares_it(self):
        components = {'C1': make_cap()}
        nets = {
            'PG_P1V8': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_derating(components, nets, custom_volt_map={'PG_P1V8': 1.8})
        self.assertEqual('1.8', rows[0]['推断工作电压(V)'])

    def test_derating_token_inference_is_candidate_not_confirmed(self):
        components = {'C1': make_cap()}
        nets = {
            'P1V8_AON': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertEqual('候选判断', row['结论类型'])
        self.assertEqual('网络首 token', row['推断来源类型'])
        self.assertEqual('single_positive_rail_token', row['原因代码'])

    def test_derating_custom_map_is_confirmed(self):
        components = {'C1': make_cap()}
        nets = {
            'VDD_SENSE': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_derating(components, nets, custom_volt_map={'VDD_SENSE': 1.2})[0]
        self.assertEqual('确定结论', row['结论类型'])
        self.assertEqual('自定义映射', row['推断来源类型'])
        self.assertEqual('custom_voltage_map', row['原因代码'])

    def test_derating_passes_high_rated_caps_when_global_max_voltage_is_not_above_12v(self):
        components = {'C1': make_cap(rated='50V')}
        nets = {
            'SIG_ALERT': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
            'P12V_AUX': [{'refdes': 'U1', 'pin': '1', 'pin_name': 'VIN'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertTrue(row['状态'].startswith('✅'))
        self.assertEqual('12.0', row['推断工作电压(V)'])
        self.assertEqual('P12V_AUX', row['推断来源网络'])
        self.assertEqual('global_max_voltage_under_12v_high_rated_cap', row['原因代码'])

    def test_derating_does_not_apply_50v_override_when_global_max_voltage_exceeds_12v(self):
        components = {'C1': make_cap(rated='50V')}
        nets = {
            'SIG_ALERT': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
            'P20V_SYS': [{'refdes': 'U1', 'pin': '1', 'pin_name': 'VIN'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertIn('无法推断', row['状态'])
        self.assertEqual('no_positive_voltage_evidence', row['原因代码'])

    def test_derating_marks_mirrored_diff_caps_as_low_risk_ac_coupling(self):
        components = {
            'C1': make_cap(refdes='C1'),
            'C2': make_cap(refdes='C2'),
        }
        nets = {
            'PCIE_TXA_P': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'PCIE_TXB_P': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
            'PCIE_TXA_N': [{'refdes': 'C2', 'pin': '1', 'pin_name': '1'}],
            'PCIE_TXB_N': [{'refdes': 'C2', 'pin': '2', 'pin_name': '2'}],
        }
        rows = {row['位号']: row for row in analyze_derating(components, nets)}
        self.assertTrue(rows['C1']['状态'].startswith('✅'))
        self.assertEqual('确定结论', rows['C1']['结论类型'])
        self.assertEqual('差分同极性 AC 耦合', rows['C1']['推断来源类型'])
        self.assertEqual('ac_coupling_same_polarity_diff_pair', rows['C1']['原因代码'])
        self.assertTrue(rows['C2']['状态'].startswith('✅'))

    def test_derating_marks_single_same_polarity_diff_cap_as_low_risk_ac_coupling(self):
        components = {'C1': make_cap(refdes='C1')}
        nets = {
            'PCIE_TXA_P': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'PCIE_TXB_P': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
            'PCIE_TXA_N': [{'refdes': 'U1', 'pin': '1', 'pin_name': 'RXN'}],
            'PCIE_TXB_N': [{'refdes': 'U2', 'pin': '1', 'pin_name': 'TXN'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertTrue(row['状态'].startswith('✅'))
        self.assertEqual('PCIE_TXA_P ↔ PCIE_TXB_P', row['推断来源网络'])
        self.assertEqual('ac_coupling_same_polarity_diff_pair', row['原因代码'])

    def test_derating_does_not_treat_lone_negative_suffix_as_ac_coupling(self):
        components = {'C1': make_cap()}
        nets = {
            'PCIE_TXA_N': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'PCIE_TXB_N': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertEqual('no_ground_reference', row['原因代码'])
        self.assertNotIn('AC 耦合', row['状态'])

    def test_analyze_derating_prefers_real_page_for_page_column(self):
        components = {'C1': make_cap(page='PAGE242', page_real='PAGE518')}
        nets = {
            'P1V8_AON': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertEqual('PAGE518', row['页面'])

    def test_analyze_derating_does_not_fallback_to_logical_page_when_real_page_missing(self):
        components = {'C1': make_cap(page='PAGE242', page_real='')}
        nets = {
            'P1V8_AON': [{'refdes': 'C1', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'C1', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_derating(components, nets)[0]
        self.assertEqual('', row['页面'])

    def test_analyze_resistors_uses_conservative_power_detection(self):
        components = {
            'R1': {
                'refdes': 'R1',
                'part_name': 'RES_0402',
                'hq_code': '',
                'value': '10k',
                'package': '0402',
                'material': '',
                'tolerance': '',
                'voltage': '',
                'current': '',
                'power': '',
                'bom_option': '',
                'bom_cost': '',
                'room': '',
                'drawing': 'SCH_PAGE1',
                'page': 'PAGE1',
                'comp_type': 'RES',
                'nets': {'1': 'DDR_VDDQ_EN', '2': 'ALERT_N'},
            },
        }
        nets = {
            'DDR_VDDQ_EN': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}],
            'ALERT_N': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}],
        }
        result = analyze_resistors(components, nets)
        self.assertEqual({}, result['pullups'])

    def test_analyze_resistors_duplicate_pullup_is_candidate(self):
        components = {
            'R1': {'refdes': 'R1', 'part_name': 'RES_0402', 'hq_code': '', 'value': '10k', 'package': '0402',
                   'material': '', 'tolerance': '', 'voltage': '', 'current': '', 'power': '',
                   'bom_option': '', 'bom_cost': '', 'room': '', 'drawing': '', 'page': 'PAGE1',
                   'comp_type': 'RES', 'nets': {'1': 'P3V3', '2': 'ALERT_N'}},
            'R2': {'refdes': 'R2', 'part_name': 'RES_0402', 'hq_code': '', 'value': '4.7k', 'package': '0402',
                   'material': '', 'tolerance': '', 'voltage': '', 'current': '', 'power': '',
                   'bom_option': '', 'bom_cost': '', 'room': '', 'drawing': '', 'page': 'PAGE2',
                   'comp_type': 'RES', 'nets': {'1': 'P3V3', '2': 'ALERT_N'}},
        }
        nets = {
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'ALERT_N': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}, {'refdes': 'R2', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_resistors(components, nets)['dup_pullups'][0]
        self.assertEqual('候选判断', row['结论类型'])
        self.assertEqual('multiple_pullup_paths', row['原因代码'])
        self.assertNotIn('单板场景', row)

    def test_analyze_resistors_duplicate_pullup_keeps_bom_options_without_single_board_inference(self):
        components = {
            'R1': make_res('R1', 'P3V3', 'ALERT_N', bom_option='MAIN'),
            'R2': make_res('R2', 'P3V3', 'ALERT_N', bom_option='ALT'),
        }
        nets = {
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'ALERT_N': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}, {'refdes': 'R2', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_resistors(components, nets)['dup_pullups'][0]
        self.assertIn('MAIN', row['BOM_OPTION'])
        self.assertIn('ALT', row['BOM_OPTION'])
        self.assertEqual(2, row['上拉数量'])
        self.assertEqual('低', row['严重级别'])
        self.assertEqual('multiple_pullup_paths_bom_option_variant', row['原因代码'])
        self.assertIn('可能是互斥装配', row['装配选项判断'])

    def test_analyze_resistors_duplicate_pulldown_is_candidate(self):
        components = {
            'R1': make_res('R1', 'ALERT_N', 'GND', value='100k'),
            'R2': make_res('R2', 'ALERT_N', 'AGND', value='47k'),
        }
        nets = {
            'ALERT_N': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}],
            'AGND': [{'refdes': 'R2', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_resistors(components, nets)['dup_pulldowns'][0]
        self.assertEqual('ALERT_N', row['信号网络'])
        self.assertEqual(2, row['下拉数量'])
        self.assertEqual('R1, R2', row['位号'])
        self.assertEqual('multiple_pulldown_paths', row['原因代码'])

    def test_analyze_resistors_reports_remote_duplicate_pullups_across_series(self):
        components = {
            'U1': make_ic('U1', {'1': 'SIG'}),
            'R1': make_res('R1', 'SIG', 'REMOTE', value='22R'),
            'R2': make_res('R2', 'P3V3', 'REMOTE', value='10k'),
            'R3': make_res('R3', 'P1V8', 'REMOTE', value='47k'),
        }
        nets = {
            'SIG': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
            ],
            'REMOTE': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R3', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'P1V8': [{'refdes': 'R3', 'pin': '1', 'pin_name': '1'}],
        }
        result = analyze_resistors(components, nets)
        row = next(item for item in result['dup_pullups'] if item['信号网络'] == 'SIG')
        self.assertEqual('SIG', row['信号网络'])
        self.assertEqual(2, row['上拉数量'])
        self.assertEqual('R2, R3', row['位号'])
        self.assertEqual('隔串阻', row['连接方式'])
        self.assertEqual('R1', row['隔串阻链'])
        self.assertEqual('multiple_pullup_paths_with_series', row['原因代码'])

    def test_analyze_resistors_marks_zero_ohm_series_usage(self):
        components = {
            'U1': make_ic('U1', {'1': 'SIG'}),
            'R1': make_res('R1', 'SIG', 'REMOTE', value='0R'),
            'R2': make_res('R2', 'P3V3', 'REMOTE', value='10k'),
        }
        nets = {
            'SIG': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
            ],
            'REMOTE': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
        }
        result = analyze_resistors(components, nets)
        divider_row = next(row for row in result['divider_risks'] if row['受影响网络'] == 'SIG')
        self.assertEqual('0R/跳线', divider_row['串阻类型候选'])
        chip_row = next(row for row in result['chip_pin_rows'] if row['芯片位号'] == 'U1')
        self.assertEqual('0R/跳线', chip_row['串阻类型候选'])

    def test_analyze_resistors_finds_series_bias_on_both_sides(self):
        components = {
            'R1': make_res('R1', 'NET_A', 'NET_B', value='100R'),
            'R2': make_res('R2', 'P3V3', 'NET_A', value='1k'),
            'R3': make_res('R3', 'NET_B', 'GND', value='2k'),
        }
        nets = {
            'NET_A': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '2', 'pin_name': '2'}],
            'NET_B': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}, {'refdes': 'R3', 'pin': '1', 'pin_name': '1'}],
            'P3V3': [{'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'R3', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_resistors(components, nets)['divider_risks']
        self.assertEqual(2, len(rows))
        row_map = {(row['偏置类型'], row['受影响网络']): row for row in rows}
        self.assertIn(('上拉', 'NET_B'), row_map)
        self.assertIn(('下拉', 'NET_A'), row_map)
        self.assertEqual('NET_A', row_map[('上拉', 'NET_B')]['偏置所在网络'])
        self.assertEqual('GND', row_map[('下拉', 'NET_A')]['偏置参考网络'])

    def test_analyze_resistors_reports_chip_pin_status_for_u_pu_xu(self):
        components = {
            'U1': make_ic('U1', {'1': 'GPIO1'}, page='PAGE242', page_real='PAGE518'),
            'PU2A1': make_ic('PU2A1', {'A1': 'GPIO1'}, page='PAGE242', page_real='PAGE518'),
            'XU3': make_ic('XU3', {'B2': 'NET_SER'}, page='PAGE300', page_real='PAGE612'),
            'R1': make_res('R1', 'GPIO1', 'NET_SER', value='22R', page='PAGE242', page_real='PAGE518'),
            'R2': make_res('R2', 'P3V3', 'GPIO1', value='10k', page='PAGE242', page_real='PAGE518'),
            'R3': make_res('R3', 'GPIO1', 'GND', value='100k', page='PAGE242', page_real='PAGE518'),
        }
        nets = {
            'GPIO1': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO1'},
                {
                    'refdes': 'PU2A1',
                    'pin': 'A1',
                    'pin_name': (
                        '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'
                        '@GPU_2SW_BOARD_LIB.HQPWR_EFUSE_TPS259260_12VIN_4A(SCH_1):PAGE1_I14'
                        '@HQ_IC.TPS259261DRCR_11P_HDL(CHIPS)'
                    ),
                },
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R3', 'pin': '1', 'pin_name': '1'},
            ],
            'NET_SER': [
                {'refdes': 'XU3', 'pin': 'B2', 'pin_name': 'DATA'},
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'R3', 'pin': '2', 'pin_name': '2'}],
        }
        rows = analyze_resistors(components, nets)['chip_pin_rows']
        row_map = {(row['芯片位号'], row['引脚']): row for row in rows}
        self.assertEqual('是', row_map[('U1', '1')]['有串阻'])
        self.assertEqual('是', row_map[('U1', '1')]['有上拉'])
        self.assertEqual('是', row_map[('U1', '1')]['有下拉'])
        self.assertEqual('A1', row_map[('PU2A1', 'A1')]['后缀组'])
        self.assertEqual('HQPWR_EFUSE_TPS259260_12VIN_4A', row_map[('PU2A1', 'A1')]['子模块'])
        self.assertEqual(
            'GPU_2SW_BOARD / HQPWR_EFUSE_TPS259260_12VIN_4A',
            row_map[('PU2A1', 'A1')]['子模块路径'],
        )
        self.assertEqual('PAGE518', row_map[('PU2A1', 'A1')]['页面'])
        self.assertNotIn('逻辑页', row_map[('PU2A1', 'A1')])
        self.assertEqual('是', row_map[('XU3', 'B2')]['有串阻'])
        self.assertEqual('否', row_map[('XU3', 'B2')]['有上拉'])
        self.assertEqual(1, row_map[('XU3', 'B2')]['隔串阻上拉数量'])
        self.assertEqual('R2', row_map[('XU3', 'B2')]['隔串阻上拉位号'])
        self.assertEqual('P3V3', row_map[('XU3', 'B2')]['隔串阻上拉电源'])
        self.assertEqual('R1', row_map[('XU3', 'B2')]['隔串阻上拉串阻链'])
        self.assertEqual(1, row_map[('XU3', 'B2')]['隔串阻下拉数量'])
        self.assertEqual('R1', row_map[('XU3', 'B2')]['隔串阻下拉串阻链'])

    def test_ground_detection_treats_analog_ground_variants_as_ground(self):
        for net_name in ['AGND', 'AGND1', 'AGND_ADC', 'GNDA', 'VSSA', 'AVSS', '0V', '0']:
            self.assertTrue(_net_is_gnd(net_name), net_name)

    def test_analyze_resistors_treats_agnd_variants_as_ground_not_series_path(self):
        components = {
            'U1': make_ic('U1', {'1': 'SIG'}),
            'R1': make_res('R1', 'SIG', 'AGND1', value='22R'),
            'R2': make_res('R2', 'P3V3', 'AGND1', value='10k'),
        }
        nets = {
            'SIG': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
            ],
            'AGND1': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
        }

        result = analyze_resistors(components, nets)
        row = {
            (item['芯片位号'], item['引脚']): item
            for item in result['chip_pin_rows']
        }[('U1', '1')]

        self.assertEqual('否', row['有串阻'])
        self.assertEqual('是', row['有下拉'])
        self.assertEqual('否', row['有上拉'])
        self.assertEqual([], result['indirect_pullups'].get('SIG', []))
        self.assertEqual([], result['divider_risks'])

    def test_analyze_resistors_finds_bias_across_multiple_series_resistors(self):
        components = {
            'XU1': make_ic('XU1', {'B2': 'NET_IN'}, page='PAGE10', page_real='PAGE510'),
            'R1': make_res('R1', 'NET_IN', 'NET_MID', value='22R', page='PAGE10', page_real='PAGE510'),
            'R2': make_res('R2', 'NET_MID', 'NET_BIAS', value='33R', page='PAGE11', page_real='PAGE511'),
            'R3': make_res('R3', 'P3V3', 'NET_BIAS', value='10k', page='PAGE12', page_real='PAGE512'),
            'R4': make_res('R4', 'NET_BIAS', 'GND', value='100k', page='PAGE12', page_real='PAGE512'),
        }
        nets = {
            'NET_IN': [
                {'refdes': 'XU1', 'pin': 'B2', 'pin_name': 'GPIO_CHAIN'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
            ],
            'NET_MID': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '1', 'pin_name': '1'},
            ],
            'NET_BIAS': [
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R3', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R4', 'pin': '1', 'pin_name': '1'},
            ],
            'P3V3': [{'refdes': 'R3', 'pin': '1', 'pin_name': '1'}],
            'GND': [{'refdes': 'R4', 'pin': '2', 'pin_name': '2'}],
        }
        result = analyze_resistors(components, nets)
        row = {
            (item['芯片位号'], item['引脚']): item
            for item in result['chip_pin_rows']
        }[('XU1', 'B2')]
        self.assertEqual(1, row['隔串阻上拉数量'])
        self.assertEqual('R3', row['隔串阻上拉位号'])
        self.assertEqual('NET_BIAS', row['隔串阻上拉来源网络'])
        self.assertEqual('P3V3', row['隔串阻上拉电源'])
        self.assertEqual('R1 -> R2', row['隔串阻上拉串阻链'])
        self.assertEqual(1, row['隔串阻下拉数量'])
        self.assertEqual('R4', row['隔串阻下拉位号'])
        self.assertEqual('R1 -> R2', row['隔串阻下拉串阻链'])

        net_in_rows = [item for item in result['divider_risks'] if item['受影响网络'] == 'NET_IN']
        self.assertEqual(2, len(net_in_rows))
        risk_map = {(item['偏置类型'], item['偏置位号']): item for item in net_in_rows}
        self.assertEqual('R1 -> R2', risk_map[('上拉', 'R3')]['串阻位号'])
        self.assertEqual('NET_IN -> NET_MID -> NET_BIAS', risk_map[('上拉', 'R3')]['串阻经过网络'])
        self.assertEqual(2, risk_map[('上拉', 'R3')]['串阻跳数'])
        self.assertEqual('P3V3', risk_map[('上拉', 'R3')]['偏置参考网络'])
        self.assertEqual('R1 -> R2', risk_map[('下拉', 'R4')]['串阻位号'])

    def test_analyze_resistors_keeps_parallel_series_paths_to_same_remote_pullup(self):
        components = {
            'U1': make_ic('U1', {'1': 'SIG'}),
            'R1': make_res('R1', 'SIG', 'MID_A', value='22R'),
            'R2': make_res('R2', 'MID_A', 'REMOTE', value='22R'),
            'R3': make_res('R3', 'SIG', 'MID_B', value='33R'),
            'R4': make_res('R4', 'MID_B', 'REMOTE', value='33R'),
            'R5': make_res('R5', 'P3V3', 'REMOTE', value='4.7k'),
        }
        nets = {
            'SIG': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'GPIO'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
                {'refdes': 'R3', 'pin': '1', 'pin_name': '1'},
            ],
            'MID_A': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '1', 'pin_name': '1'},
            ],
            'MID_B': [
                {'refdes': 'R3', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R4', 'pin': '1', 'pin_name': '1'},
            ],
            'REMOTE': [
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R4', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R5', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R5', 'pin': '1', 'pin_name': '1'}],
        }

        result = analyze_resistors(components, nets)

        sig_pullups = result['indirect_pullups']['SIG']
        self.assertEqual(
            {'R1 -> R2', 'R3 -> R4'},
            {item['via_refdes_chain'] for item in sig_pullups},
        )
        sig_dividers = [
            row for row in result['divider_risks']
            if row['受影响网络'] == 'SIG' and row['偏置位号'] == 'R5'
        ]
        self.assertEqual(2, len(sig_dividers))

    def test_analyze_resistors_does_not_emit_od_oc_missing_pullup_table(self):
        components = {
            'U1': make_ic('U1', {'1': 'SMBALERT_N'}),
        }
        nets = {
            'SMBALERT_N': [{'refdes': 'U1', 'pin': '1', 'pin_name': 'SMBALERT_N'}],
        }
        result = analyze_resistors(components, nets)
        self.assertNotIn('od_missing', result)

    def test_analyze_resistors_prefers_real_page_for_duplicate_and_divider_rows(self):
        components = {
            'R1': make_res('R1', 'P3V3', 'ALERT_N', page='PAGE1', page_real='PAGE501'),
            'R2': make_res('R2', 'P3V3', 'ALERT_N', value='4.7k', page='PAGE2', page_real='PAGE502'),
            'R3': make_res('R3', 'ALERT_N', 'GPIO_BUF', value='22R', page='PAGE3', page_real='PAGE503'),
        }
        nets = {
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'ALERT_N': [
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R2', 'pin': '2', 'pin_name': '2'},
                {'refdes': 'R3', 'pin': '1', 'pin_name': '1'},
            ],
            'GPIO_BUF': [{'refdes': 'R3', 'pin': '2', 'pin_name': '2'}],
        }
        result = analyze_resistors(components, nets)
        self.assertEqual('PAGE501, PAGE502', result['dup_pullups'][0]['页面'])
        divider_pages = {row['页面'] for row in result['divider_risks']}
        self.assertEqual({'PAGE503, PAGE501', 'PAGE503, PAGE502'}, divider_pages)

    def test_analyze_resistors_does_not_fallback_to_logical_page_when_real_page_missing(self):
        components = {
            'R1': make_res('R1', 'P3V3', 'ALERT_N', page='PAGE1', page_real=''),
            'R2': make_res('R2', 'P3V3', 'ALERT_N', value='4.7k', page='PAGE2', page_real=''),
        }
        nets = {
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}, {'refdes': 'R2', 'pin': '1', 'pin_name': '1'}],
            'ALERT_N': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}, {'refdes': 'R2', 'pin': '2', 'pin_name': '2'}],
        }
        row = analyze_resistors(components, nets)['dup_pullups'][0]
        self.assertEqual('', row['页面'])

    def test_analyze_resistors_ignores_depop_bias_resistors(self):
        components = {
            'U1': make_ic('U1', {'1': 'SMBALERT_N'}),
            'R1': make_res('R1', 'P3V3', 'SMBALERT_N', value='4.7k', bom_option='DEPOP'),
        }
        nets = {
            'SMBALERT_N': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'SMBALERT_N'},
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}],
        }
        result = analyze_resistors(components, nets)
        self.assertEqual({}, result['pullups'])
        self.assertNotIn('od_missing', result)

    def test_analyze_resistors_can_include_depop_when_requested(self):
        components = {
            'U1': make_ic('U1', {'1': 'SMBALERT_N'}),
            'R1': make_res('R1', 'P3V3', 'SMBALERT_N', value='4.7k', bom_option='DEPOP'),
        }
        nets = {
            'SMBALERT_N': [
                {'refdes': 'U1', 'pin': '1', 'pin_name': 'SMBALERT_N'},
                {'refdes': 'R1', 'pin': '2', 'pin_name': '2'},
            ],
            'P3V3': [{'refdes': 'R1', 'pin': '1', 'pin_name': '1'}],
        }
        result = analyze_resistors(components, nets, exclude_depop=False)
        self.assertIn('SMBALERT_N', result['pullups'])
        self.assertNotIn('od_missing', result)

    def test_analyze_networks_detects_diff_pairs_case_insensitively(self):
        nets = {
            'pcie_tx_p': [{'refdes': 'U1', 'pin': '1', 'pin_name': '1'}],
            'pcie_tx_n': [{'refdes': 'U1', 'pin': '2', 'pin_name': '2'}],
        }
        components = {'U1': {'page': 'PAGE1'}}
        result = analyze_networks(nets, components)
        self.assertIn('pcie_tx', result['diff_pairs'])
        self.assertEqual('pcie_tx_p', result['diff_pairs']['pcie_tx']['P'])
        self.assertEqual('pcie_tx_n', result['diff_pairs']['pcie_tx']['N'])
        self.assertEqual('候选判断', result['diff_pair_rows'][0]['结论类型'])

    def test_analyze_networks_does_not_treat_lone_n_suffix_as_diff_pair(self):
        nets = {
            'pcie_tx_n': [{'refdes': 'U1', 'pin': '1', 'pin_name': '1'}],
        }
        components = {'U1': {'page': 'PAGE1'}}
        result = analyze_networks(nets, components)
        self.assertEqual({}, result['diff_pairs'])

    def test_analyze_networks_uses_unknown_for_blank_page(self):
        nets = {}
        components = {
            'U1': {'page': ''},
            'U2': {'page': 'PAGE1'},
        }
        result = analyze_networks(nets, components)
        self.assertEqual(1, result['page_counter']['UNKNOWN'])
        self.assertEqual(1, result['page_counter']['PAGE1'])

    def test_analyze_networks_prefers_real_page_for_page_rows(self):
        nets = {}
        components = {
            'U1': {'page': 'PAGE242', 'page_real': 'PAGE518'},
            'U2': {'page': 'PAGE242', 'page_real': 'PAGE518'},
            'U3': {'page': 'PAGE300', 'page_real': ''},
        }
        result = analyze_networks(nets, components)
        self.assertEqual(2, result['page_counter']['PAGE518'])
        self.assertEqual(1, result['page_counter']['PAGE300'])
        self.assertEqual(['PAGE300', 'PAGE518'], [row['页码'] for row in result['page_rows']])

    def test_analyze_networks_normalizes_and_naturally_sorts_page_rows(self):
        nets = {}
        components = {
            'U1': {'page': 'PAGE01'},
            'U2': {'page': 'PAGE1'},
            'U3': {'page': 'PAGE10'},
            'U4': {'page': 'PAGE2'},
            'U5': {'page': ''},
        }
        result = analyze_networks(nets, components)
        self.assertEqual(2, result['page_counter']['PAGE1'])
        self.assertEqual(['PAGE1', 'PAGE2', 'PAGE10', 'UNKNOWN'],
                         [row['页码'] for row in result['page_rows']])

    def test_analyze_networks_prefers_real_pages_over_hierarchical_logic_pages(self):
        nets = {}
        components = {
            'U1': {'page': 'PAGE1', 'page_real': 'PAGE1'},
            'U2': {'page': 'PAGE242 / PAGE1', 'page_real': 'PAGE601'},
            'U3': {'page': 'PAGE242 / PAGE2', 'page_real': 'PAGE602'},
        }
        result = analyze_networks(nets, components)
        self.assertEqual(1, result['page_counter']['PAGE1'])
        self.assertEqual(1, result['page_counter']['PAGE601'])
        self.assertEqual(1, result['page_counter']['PAGE602'])

    def test_analyze_networks_uses_full_topology_to_avoid_depop_single_node_false_positive(self):
        active_nets = {
            'GPIO': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}],
        }
        full_nets = {
            'GPIO': [
                {'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'},
                {'refdes': 'R1', 'pin': '1', 'pin_name': '1'},
            ],
            'P3V3': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}],
        }
        components = {'U1': make_ic('U1', {'A1': 'GPIO'}, page='PAGE1', page_real='PAGE101')}

        result = analyze_networks(
            active_nets,
            components,
            single_node_topology_nets=full_nets,
        )

        self.assertEqual({}, result['single_node'])
        self.assertEqual([], result['single_node_rows'])

    def test_check_drc_does_not_report_testpoint_or_unnamed_net_as_single_pin_issue(self):
        components = {
            'TP1': {
                'refdes': 'TP1', 'part_name': 'testpoint', 'hq_code': '', 'value': '',
                'package': 'TP', 'material': '', 'tolerance': '', 'voltage': '',
                'current': '', 'power': '', 'bom_option': '', 'bom_cost': '',
                'room': '', 'drawing': 'SCH_PAGE1', 'page': 'PAGE1',
                'comp_type': 'TESTPOINT', 'nets': {},
            },
            'U1': {
                'refdes': 'U1', 'part_name': 'IC_CPU', 'hq_code': 'PN1', 'value': 'CPU',
                'package': 'BGA', 'material': '', 'tolerance': '', 'voltage': '',
                'current': '', 'power': '', 'bom_option': '', 'bom_cost': '',
                'room': '', 'drawing': 'SCH_PAGE1', 'page': 'PAGE1',
                'comp_type': 'IC', 'nets': {},
            },
        }
        nets = {
            'NET_ALONE': [{'refdes': 'TP1', 'pin': '1', 'pin_name': 'TP'}],
            'UNNAMED_1': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}],
        }
        result = check_drc(components, nets)
        self.assertEqual([], result['single_pin_nets'])
        self.assertEqual('候选判断', result['unnamed_nets'][0]['结论类型'])
        self.assertEqual('unnamed_net', result['unnamed_nets'][0]['原因代码'])
        self.assertEqual('PAGE1', result['unnamed_nets'][0]['页面'])
        self.assertEqual('PAGE1', result['unnamed_nets'][0]['页码'])
        self.assertEqual('U1.GPIO', result['unnamed_nets'][0]['连接点'])

    def test_check_drc_single_pin_uses_full_topology_with_depop_resistor(self):
        full_components = {
            'U1': make_ic('U1', {'A1': 'GPIO'}, page='PAGE1', page_real='PAGE101'),
            'R1': make_res('R1', 'GPIO', 'P3V3', bom_option='DEPOP', page='PAGE2', page_real='PAGE102'),
        }
        full_nets = {
            'GPIO': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}, {'refdes': 'R1', 'pin': '1', 'pin_name': '1'}],
            'P3V3': [{'refdes': 'R1', 'pin': '2', 'pin_name': '2'}],
        }
        active_components = {'U1': full_components['U1']}
        active_nets = {'GPIO': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}]}

        result = check_drc(
            active_components,
            active_nets,
            single_pin_components=full_components,
            single_pin_nets=full_nets,
        )

        self.assertEqual([], result['single_pin_nets'])

    def test_check_drc_bom_option_suggests_main(self):
        components = {
            'U1': {
                'refdes': 'U1', 'part_name': 'IC_CPU', 'hq_code': 'PN1', 'value': 'CPU',
                'package': 'BGA', 'material': '', 'tolerance': '', 'voltage': '',
                'current': '', 'power': '', 'bom_option': 'MIAN', 'bom_cost': '',
                'room': '', 'drawing': 'SCH_PAGE1', 'page': 'PAGE1',
                'comp_type': 'IC', 'nets': {},
            },
        }
        result = check_drc(components, {})
        self.assertEqual('MAIN', result['bom_option_typos'][0]['疑似应为'])

    def test_check_drc_missing_value_is_confirmed_high_confidence(self):
        components = {
            'U1': {
                'refdes': 'U1', 'part_name': 'IC_CPU', 'hq_code': 'PN1', 'value': '',
                'package': 'BGA', 'material': '', 'tolerance': '', 'voltage': '',
                'current': '', 'power': '', 'bom_option': '', 'bom_cost': '',
                'room': '', 'drawing': 'SCH_PAGE1', 'page': 'PAGE1',
                'comp_type': 'IC', 'nets': {},
            },
        }
        row = check_drc(components, {})['missing_value'][0]
        self.assertEqual('确定结论', row['结论类型'])
        self.assertEqual('高', row['严重级别'])
        self.assertEqual('高', row['置信度'])
        self.assertEqual('missing_value', row['原因代码'])

    def test_check_drc_lists_all_bom_option_components_with_depop_flag(self):
        components = {
            'U1': make_ic('U1', bom_option='DEPOP', page='PAGE1', page_real='PAGE518'),
            'R1': make_res('R1', 'P3V3', 'GPIO1', bom_option='ALT'),
            'C1': make_cap(),
        }
        rows = check_drc(components, {})['bom_option_components']
        row_map = {row['位号']: row for row in rows}
        self.assertEqual({'U1', 'R1'}, set(row_map))
        self.assertEqual('是', row_map['U1']['是否DEPOP'])
        self.assertEqual('否', row_map['R1']['是否DEPOP'])
        self.assertEqual('PAGE518', row_map['U1']['页面'])

    def test_check_drc_prefers_real_page_for_issue_rows(self):
        components = {'U1': make_ic('U1', page='PAGE242', page_real='PAGE518')}
        components['U1']['hq_code'] = ''
        nets = {
            'GPIO_ALONE': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}],
        }
        result = check_drc(components, nets)
        self.assertEqual('PAGE518', result['missing_hq_code'][0]['页面'])
        self.assertEqual('PAGE518', result['single_pin_nets'][0]['页面'])

    def test_check_drc_does_not_fallback_to_logical_page_when_real_page_missing(self):
        components = {'U1': make_ic('U1', page='PAGE242', page_real='')}
        components['U1']['hq_code'] = ''
        nets = {
            'GPIO_ALONE': [{'refdes': 'U1', 'pin': 'A1', 'pin_name': 'GPIO'}],
        }
        result = check_drc(components, nets)
        self.assertEqual('', result['missing_hq_code'][0]['页面'])
        self.assertEqual('', result['single_pin_nets'][0]['页面'])


if __name__ == '__main__':
    unittest.main()
