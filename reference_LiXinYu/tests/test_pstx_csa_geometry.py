# -*- coding: utf-8 -*-
"""Tests for CSA geometry and BOM_OPTION circle coverage."""

import tempfile
import unittest
from unittest import mock
from pathlib import Path

from pstx_core.cadence import csa_geometry
from pstx_core.cadence.csa_connectivity_overlay import build_csa_connectivity_overlay
from pstx_rules.bom_option_circle import check_bom_option_circle_coverage
from pstx_rules.project_analysis import analyze_project_contents
from tests.pstx_test_fixtures import (
    CSA_PAGE_DOT_CROSS,
    CSA_PAGE_DOTLESS_CROSS,
    CSA_PAGE_SPLIT_CROSS_WITH_ARC,
    CSA_PAGE_T_WITH_DOT,
    make_cap,
    make_res,
)


class CsaGeometryTests(unittest.TestCase):

    def test_core_csa_geometry_entrypoint_exports_public_api(self):
        self.assertTrue(callable(csa_geometry.analyze_csa_geometry))
        self.assertTrue(callable(csa_geometry.parse_csa_text))
        self.assertFalse(hasattr(csa_geometry, "_line_bbox"))
        self.assertFalse(Path("csa_geometry.py").exists())

    def test_csa_geometry_matches_reference_demo_rules(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir()
            (sch_dir / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            (sch_dir / 'page2.csa').write_text(CSA_PAGE_DOTLESS_CROSS, encoding='utf-8')
            (sch_dir / 'page3.csa').write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            (sch_dir / 'page4.csa').write_text(CSA_PAGE_SPLIT_CROSS_WITH_ARC, encoding='utf-8')

            result = csa_geometry.analyze_csa_geometry(root)

        self.assertTrue(result['enabled'])
        self.assertEqual(4, result['page_count'])
        self.assertEqual(2, result['cross_count'])
        self.assertEqual(3, result['circle_count'])
        self.assertEqual(['PAGE3', 'PAGE4'], [row['页面'] for row in result['dot_cross_rows']])
        self.assertEqual(['(450,0)', '(650,0)'], [row['坐标'] for row in result['dot_cross_rows']])
        self.assertIn('DOT 1 (450 0);', result['dot_cross_rows'][0]['DOT原始行'])
        self.assertIn('WIRE 16 -1 (400 0)(500 0);', result['dot_cross_rows'][0]['证据上下文'])
        self.assertIn('CIRCLE 16 -1 (1000 1000)(1100 1000);', result['circle_rows'][0]['证据上下文'])
        self.assertEqual(0, next(row for row in result['summary_rows'] if row['页面'] == 'PAGE1')['DOT四向十字数'])
        self.assertEqual(0, next(row for row in result['summary_rows'] if row['页面'] == 'PAGE2')['DOT四向十字数'])

    def test_package_scan_supports_file_recursive_missing_export_and_arc_defaults(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir()
            (sch_dir / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            (sch_dir / 'page3.csa').write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            nested = root / 'nested'
            nested.mkdir()
            (nested / 'page4.csa').write_text(CSA_PAGE_SPLIT_CROSS_WITH_ARC, encoding='utf-8')

            single = csa_geometry.collect_page_files(sch_dir / 'page3.csa', strict=True)
            self.assertEqual([sch_dir / 'page3.csa'], single)

            results, geometry = csa_geometry.scan_csa_geometry(
                root,
                check_missing=True,
                executor_kind='serial',
                include_arcs=False,
            )
            self.assertEqual([2], geometry['missing_pages'])
            self.assertEqual(2, geometry['page_count'])
            self.assertEqual(1, geometry['cross_count'])
            self.assertEqual(1, geometry['circle_count'])

            recursive_results, recursive_geometry = csa_geometry.scan_csa_geometry(
                root,
                recursive=True,
                executor_kind='thread',
                workers=2,
                include_arcs=False,
            )
            self.assertEqual(3, recursive_geometry['page_count'])
            self.assertEqual(2, recursive_geometry['cross_count'])
            self.assertEqual(1, recursive_geometry['circle_count'])

            _, recursive_with_arcs = csa_geometry.scan_csa_geometry(
                root,
                recursive=True,
                executor_kind='process',
                workers=2,
                include_arcs=True,
            )
            self.assertEqual(2, recursive_with_arcs['circle_count'])

            report_dir = root / 'report'
            written = csa_geometry.write_csa_geometry_reports(results, report_dir, json_report=True, html_report=True)
            self.assertTrue(Path(written['summary_csv']).is_file())
            self.assertTrue(Path(written['cross_detail_csv']).is_file())
            self.assertTrue(Path(written['circle_detail_csv']).is_file())
            self.assertTrue(Path(written['json_report']).is_file())
            self.assertTrue(Path(written['html_report']).is_file())
            self.assertIn('cross_positions', Path(written['summary_csv']).read_text(encoding='utf-8-sig').splitlines()[0])
            self.assertIn('source_context', Path(written['cross_detail_csv']).read_text(encoding='utf-8-sig').splitlines()[0])
            self.assertIn('source_context', Path(written['circle_detail_csv']).read_text(encoding='utf-8-sig').splitlines()[0])
            self.assertIn('cross_positions', Path(written['json_report']).read_text(encoding='utf-8'))
            html = Path(written['html_report']).read_text(encoding='utf-8')
            self.assertIn('CSA Geometry Report', html)
            self.assertIn('DOT Four-Way Cross Findings', html)
            self.assertIn('WIRE 16 -1 (400 0)(500 0);', html)

    def test_csa_geometry_payload_modes_and_bbox_circle_mode(self):
        bbox = csa_geometry.parse_circle_line('CIRCLE 16 -1 (0 0)(100 100);', 1, 'bbox')
        self.assertIsNotNone(bbox)
        self.assertEqual(50, bbox.center_x)
        self.assertEqual(50, bbox.radius)

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            (root / 'sch_1' / 'page3.csa').write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            _, geometry = csa_geometry.scan_csa_geometry(root, include_arcs=False)

        summary = csa_geometry.build_csa_geometry_payload(geometry, stdout='summary')
        self.assertEqual('pstx-csa-geometry.v1', summary['schema_version'])
        self.assertEqual([], summary['summary_rows'])

        hits = csa_geometry.build_csa_geometry_payload(geometry, stdout='hits', limit=1)
        self.assertEqual(1, len(hits['summary_rows']))
        self.assertEqual(1, len(hits['dot_cross_rows']))
        self.assertTrue(hits['truncated'])
        self.assertEqual('PAGE1', hits['circle_rows'][0]['页面'])

    def test_csa_connectivity_overlay_binds_dot_cross_to_page_semantics(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(
                "\n".join([
                    "FILE_TYPE = MACRO_DRAWING;",
                    "SET PAGE_NUMBER P1;",
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME NET_A",
                    "WIRE 16 -1 (50 -50)(50 50);",
                    "FORCEPROP 2 LAST SIG_NAME NET_A",
                    "DOT 1 (50 0);",
                    "NET_LABEL 1 (25 0) NET_A;",
                    "PORT 1 (100 0) NET_A INPUT;",
                    "OFFPAGE 1 (0 0) NET_A_REMOTE;",
                    "BUS 1 (75 0) NET_A[0..1];",
                    "NO_CONNECT 1 (50 0);",
                    "CIRCLE 16 -1 (0 -10)(120 10);",
                ]),
                encoding='utf-8',
            )
            _, geometry = csa_geometry.scan_csa_geometry(root, include_arcs=False)

            overlay = build_csa_connectivity_overlay(
                geometry,
                source_root=str(root),
                stdout='full',
            )

        self.assertEqual('pstx-csa-connectivity-overlay.v1', overlay['schema_version'])
        self.assertEqual(1, overlay['digest']['dot_cross_matched_count'])
        dot_row = overlay['dot_cross_overlay_rows'][0]
        self.assertEqual('matched', dot_row['binding_status'])
        self.assertEqual([50, 0], dot_row['coordinate'])
        self.assertIn('NET_A', dot_row['signal_names'])
        self.assertIn('NET_A', dot_row['labels'])
        self.assertIn('NET_A', dot_row['ports'])
        self.assertIn('NET_A_REMOTE', dot_row['offpage_connectors'])
        self.assertIn('NET_A[0..1]', dot_row['bus_names'])
        self.assertEqual([[50, 0]], dot_row['no_connect_points'])
        circle_row = overlay['circle_overlay_rows'][0]
        self.assertGreaterEqual(circle_row['contained_semantic_count'], 4)
        self.assertGreaterEqual(circle_row['intersecting_component_count'], 1)
        self.assertFalse(circle_row['connection_claim'])

    def test_csa_connectivity_overlay_keeps_unmatched_and_ambiguous_conservative(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME NET_WIDE",
                    "WIRE 16 -1 (25 0)(75 0);",
                    "FORCEPROP 2 LAST SIG_NAME NET_INNER",
                ]),
                encoding='utf-8',
            )
            geometry = {
                'enabled': True,
                'root': str(root / 'sch_1'),
                'page_count': 1,
                'cross_count': 2,
                'circle_count': 0,
                'error_count': 0,
                'missing_pages': [],
                'warnings': [],
                'summary_rows': [],
                'dot_cross_rows': [
                    {'页面': 'PAGE1', '文件': 'page1.csa', '序号': 1, '坐标': '(50,0)', 'X': 50, 'Y': 0},
                    {'页面': 'PAGE1', '文件': 'page1.csa', '序号': 2, '坐标': '(500,0)', 'X': 500, 'Y': 0},
                ],
                'circle_rows': [],
            }

            overlay = build_csa_connectivity_overlay(
                geometry,
                source_root=str(root),
                stdout='details',
            )

        statuses = [row['binding_status'] for row in overlay['dot_cross_overlay_rows']]
        self.assertEqual(['ambiguous', 'unmatched'], statuses)
        self.assertEqual(1, overlay['digest']['dot_cross_ambiguous_count'])
        self.assertEqual(1, overlay['digest']['dot_cross_unmatched_count'])
        self.assertNotIn('short', overlay['dot_cross_overlay_rows'][0]['note'].lower())

    def test_csa_geometry_page_filter_limits_scan_and_overlay(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            (root / 'sch_1' / 'page3.csa').write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            _, geometry = csa_geometry.scan_csa_geometry(root, include_arcs=False, page=3)

            payload = csa_geometry.build_csa_geometry_payload(geometry, stdout='full', page=3)
            overlay = build_csa_connectivity_overlay(
                geometry,
                source_root=str(root),
                page=3,
                stdout='full',
            )

        self.assertEqual(1, geometry['page_count'])
        self.assertEqual(1, payload['digest']['page_count'])
        self.assertEqual(1, len(payload['dot_cross_rows']))
        self.assertEqual(1, overlay['digest']['dot_cross_count'])

    def test_analyze_project_contents_includes_csa_geometry_when_sch1_has_csa(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page3.csa').write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            bundle = analyze_project_contents(
                "PART_NAME\nU1 'IC_CPU'\nHQ_CODE='PN'\nVALUE='CPU'\nPACKAGE='BGA'\n",
                "NET_NAME\n'N1'\nNODE_NAME U1 1\n'PIN1':\n",
                project_name='demo',
                project_root=str(root),
            )

        self.assertEqual(1, bundle['csa_geometry']['page_count'])
        self.assertEqual(1, bundle['csa_geometry']['cross_count'])
        self.assertEqual('PAGE3', bundle['csa_geometry']['dot_cross_rows'][0]['页面'])
        self.assertEqual('pstx-analysis-timings.v1', bundle['analysis_timings']['schema_version'])
        self.assertIn('csa_geometry', bundle['analysis_timings']['cache'])

    def test_analyze_project_contents_caches_and_invalidates_heavy_cadence_results(self):
        with tempfile.TemporaryDirectory() as temp_dir, tempfile.TemporaryDirectory() as cache_dir:
            root = Path(temp_dir)
            sch_dir = root / 'sch_1'
            sch_dir.mkdir()
            page_path = sch_dir / 'page3.csa'
            page_path.write_text(CSA_PAGE_DOT_CROSS, encoding='utf-8')
            env = {
                'PSTX_ANALYSIS_CACHE_DIR': cache_dir,
                'PSTX_DISABLE_ANALYSIS_CACHE': '',
            }

            with mock.patch.dict('os.environ', env, clear=False):
                first = analyze_project_contents(
                    "PART_NAME\nU1 'IC_CPU'\nHQ_CODE='PN'\nVALUE='CPU'\nPACKAGE='BGA'\n",
                    "NET_NAME\n'N1'\nNODE_NAME U1 1\n'PIN1':\n",
                    project_name='demo',
                    project_root=str(root),
                )
                second = analyze_project_contents(
                    "PART_NAME\nU1 'IC_CPU'\nHQ_CODE='PN'\nVALUE='CPU'\nPACKAGE='BGA'\n",
                    "NET_NAME\n'N1'\nNODE_NAME U1 1\n'PIN1':\n",
                    project_name='demo',
                    project_root=str(root),
                )
                page_path.write_text(CSA_PAGE_DOTLESS_CROSS, encoding='utf-8')
                third = analyze_project_contents(
                    "PART_NAME\nU1 'IC_CPU'\nHQ_CODE='PN'\nVALUE='CPU'\nPACKAGE='BGA'\n",
                    "NET_NAME\n'N1'\nNODE_NAME U1 1\n'PIN1':\n",
                    project_name='demo',
                    project_root=str(root),
                )

            self.assertEqual('miss', first['analysis_timings']['cache']['csa_geometry']['status'])
            self.assertEqual('hit', second['analysis_timings']['cache']['csa_geometry']['status'])
            self.assertEqual('hit', second['analysis_timings']['cache']['cadence_page_semantics']['status'])
            self.assertEqual('miss', third['analysis_timings']['cache']['csa_geometry']['status'])
            self.assertEqual(0, third['csa_geometry']['cross_count'])

            with mock.patch.dict('os.environ', {
                'PSTX_ANALYSIS_CACHE_DIR': cache_dir,
                'PSTX_DISABLE_ANALYSIS_CACHE': '1',
            }, clear=False):
                disabled = analyze_project_contents(
                    "PART_NAME\nU1 'IC_CPU'\nHQ_CODE='PN'\nVALUE='CPU'\nPACKAGE='BGA'\n",
                    "NET_NAME\n'N1'\nNODE_NAME U1 1\n'PIN1':\n",
                    project_name='demo',
                    project_root=str(root),
                )
            self.assertEqual('disabled', disabled['analysis_timings']['cache']['csa_geometry']['status'])

    def test_bom_option_circle_coverage_uses_component_center_overlap(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            csa = csa_geometry.analyze_csa_geometry(root)

        components = {
            'R1': make_res('R1', 'P3V3', 'GPIO1', bom_option='DEPOP', page_real='PAGE1'),
            'R2': make_res('R2', 'P3V3', 'GPIO2', bom_option='ALT', page_real='PAGE1'),
            'R3': make_res('R3', 'P3V3', 'GPIO3', bom_option='OPTION', page_real='PAGE1'),
        }
        components['R1'].update({'xy': '(1050,1000)', 'xy_x': 1050.0, 'xy_y': 1000.0})
        components['R2'].update({'xy': '(1300,1000)', 'xy_x': 1300.0, 'xy_y': 1000.0})
        components['R3'].update({'xy': '(1115,1000)', 'xy_x': 1115.0, 'xy_y': 1000.0})

        result = check_bom_option_circle_coverage(components, csa)

        rows = {row['位号']: row for row in result['coverage_rows']}
        self.assertEqual('已打圈', rows['R1']['覆盖状态'])
        self.assertEqual('50%', rows['R1']['中心重合度'])
        self.assertEqual('0.50', rows['R1']['距离/半径'])
        self.assertEqual('疑似打圈', rows['R3']['覆盖状态'])
        self.assertEqual('未打圈', rows['R2']['覆盖状态'])
        self.assertEqual(['R2'], [row['位号'] for row in result['issue_rows']])

    def test_bom_option_circle_coverage_checks_submodule_mapped_page(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page2.csa').write_text(CSA_PAGE_DOTLESS_CROSS, encoding='utf-8')
            csa = csa_geometry.analyze_csa_geometry(root)

        components = {
            'C1A101': make_cap('C1A101', bom_option='DEPOP', page_real='PAGE24'),
        }
        components['C1A101'].update({
            'xy': '(2000,2000)',
            'xy_x': 2000.0,
            'xy_y': 2000.0,
            'page_submodule_mapped': 'PAGE2',
        })

        result = check_bom_option_circle_coverage(components, csa)

        row = result['coverage_rows'][0]
        self.assertEqual('已打圈', row['覆盖状态'])
        self.assertEqual('PAGE2', row['最近画圈页'])
        self.assertEqual('页码', row['匹配来源'])
        self.assertEqual([], result['issue_rows'])

    def test_bom_option_circle_coverage_does_not_fall_back_to_real_page_for_submodule(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page24.csa').write_text(
                "FILE_TYPE = MACRO_DRAWING;\n"
                "SET PAGE_NUMBER P24;\n"
                "CIRCLE 16 -1 (2000 2000) 150;\n",
                encoding='utf-8',
            )
            csa = csa_geometry.analyze_csa_geometry(root)

        components = {
            'C1A101': make_cap('C1A101', bom_option='DEPOP', page_real='PAGE24'),
        }
        components['C1A101'].update({
            'xy': '(2000,2000)',
            'xy_x': 2000.0,
            'xy_y': 2000.0,
            'page_submodule_real': 'PAGE1',
            'page_submodule_mapped': 'PAGE177',
        })

        result = check_bom_option_circle_coverage(components, csa)

        row = result['coverage_rows'][0]
        self.assertEqual('PAGE177', row['候选检查页'])
        self.assertEqual('未打圈', row['覆盖状态'])
        self.assertEqual('', row['最近画圈页'])
        self.assertEqual('PAGE177', row['页码'])
        self.assertEqual('bom_option_circle_missing_no_circle_on_page', result['issue_rows'][0]['原因代码'])

    def test_bom_option_circle_coverage_reads_submodule_csa_page(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            worklib = Path(temp_dir) / 'worklib'
            main_root = worklib / 'gpu_2sw_board'
            sub_root = worklib / 'i2c_repeater_9617_cbb_v3'
            (main_root / 'sch_1').mkdir(parents=True)
            (sub_root / 'sch_1').mkdir(parents=True)
            (sub_root / 'sch_1' / 'page4.csa').write_text(CSA_PAGE_DOTLESS_CROSS, encoding='utf-8')
            csa = csa_geometry.analyze_csa_geometry(main_root)

            components = {
                'C2A106': make_cap('C2A106', bom_option='DEPOP', page_real='PAGE145'),
            }
            components['C2A106'].update({
                'xy': '(2000,2000)',
                'xy_x': 2000.0,
                'xy_y': 2000.0,
                'page_submodule_real': 'PAGE4',
                'page_submodule_mapped': 'PAGE307',
                'module_order_key': (
                    '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page145_i48'
                    '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1)'
                ),
            })

            result = check_bom_option_circle_coverage(
                components,
                csa,
                project_root=str(main_root),
            )

        row = result['coverage_rows'][0]
        self.assertEqual('PAGE307', row['候选检查页'])
        self.assertEqual('PAGE307', row['最近画圈页'])
        self.assertEqual('PAGE4', row['检查CSA页'])
        self.assertIn('i2c_repeater_9617_cbb_v3', row['检查CSA文件'])
        self.assertEqual('子模块CSA:i2c_repeater_9617_cbb_v3:PAGE4', row['匹配来源'])
        self.assertEqual('已打圈', row['覆盖状态'])
        self.assertEqual([], result['issue_rows'])

    def test_bom_option_circle_coverage_checks_each_split_symbol_section(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(
                "FILE_TYPE = MACRO_DRAWING;\n"
                "SET PAGE_NUMBER P1;\n"
                "CIRCLE 16 -1 (100 100) 80;\n",
                encoding='utf-8',
            )
            (root / 'sch_1' / 'page2.csa').write_text(
                "FILE_TYPE = MACRO_DRAWING;\n"
                "SET PAGE_NUMBER P2;\n",
                encoding='utf-8',
            )
            csa = csa_geometry.analyze_csa_geometry(root)

        components = {
            'U46': {
                'refdes': 'U46',
                'comp_type': 'IC',
                'bom_option': 'DEPOP',
                'sections': [
                    {
                        'section_number': '1',
                        'xy': '(100,100)',
                        'xy_x': 100.0,
                        'xy_y': 100.0,
                        'page_real': 'PAGE1',
                    },
                    {
                        'section_number': '2',
                        'xy': '(300,300)',
                        'xy_x': 300.0,
                        'xy_y': 300.0,
                        'page_real': 'PAGE2',
                    },
                ],
            },
        }

        result = check_bom_option_circle_coverage(components, csa)
        rows = {(row['位号'], row['SECTION_NUMBER']): row for row in result['coverage_rows']}

        self.assertEqual('已打圈', rows[('U46', '1')]['覆盖状态'])
        self.assertEqual('未打圈', rows[('U46', '2')]['覆盖状态'])
        self.assertEqual(['2'], [row['SECTION_NUMBER'] for row in result['issue_rows']])

    def test_bom_option_circle_coverage_reports_missing_xy_as_unknown(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / 'sch_1').mkdir()
            (root / 'sch_1' / 'page1.csa').write_text(CSA_PAGE_T_WITH_DOT, encoding='utf-8')
            csa = csa_geometry.analyze_csa_geometry(root)

        components = {
            'R1': make_res('R1', 'P3V3', 'GPIO1', bom_option='DEPOP', page_real='PAGE1'),
        }

        result = check_bom_option_circle_coverage(components, csa)

        self.assertEqual('无法判断', result['coverage_rows'][0]['覆盖状态'])
        self.assertEqual('无法判断', result['issue_rows'][0]['结论类型'])
        self.assertEqual('bom_option_circle_unknown_no_xy', result['issue_rows'][0]['原因代码'])


if __name__ == '__main__':
    unittest.main()
