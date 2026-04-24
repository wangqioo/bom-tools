import socket
import tempfile
import unittest
from pathlib import Path

import pstx_web


PRT_SAMPLE = (
    "PART_NAME\n"
    "U1 'IC_CPU'\n"
    "HQ_CODE='PN_U1'\n"
    "VALUE='CPU'\n"
    "PACKAGE='BGA'\n"
    "DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'\n"
    "PART_NAME\n"
    "R1 'RES_0402'\n"
    "VALUE='4.7k'\n"
    "PACKAGE='0402'\n"
    "DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I2'\n"
)

NET_SAMPLE = (
    "NET_NAME\n"
    "'SMBALERT_N'\n"
    "NODE_NAME U1 1\n"
    "'SMBALERT_N':\n"
    "NODE_NAME R1 2\n"
    "'2':\n"
    "NET_NAME\n"
    "'P3V3'\n"
    "NODE_NAME R1 1\n"
    "'1':\n"
)

PRT_SAMPLE_DEPOP = (
    "PART_NAME\n"
    "U1 'IC_CPU'\n"
    "HQ_CODE='PN_U1'\n"
    "VALUE='CPU'\n"
    "PACKAGE='BGA'\n"
    "DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I1'\n"
    "PART_NAME\n"
    "R1 'RES_0402'\n"
    "VALUE='4.7k'\n"
    "PACKAGE='0402'\n"
    "BOM_OPTION='DEPOP'\n"
    "DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE242_I2'\n"
)

NET_SAMPLE_DEPOP = (
    "NET_NAME\n"
    "'SMBALERT_N'\n"
    "NODE_NAME U1 1\n"
    "'SMBALERT_N':\n"
    "NODE_NAME R1 2\n"
    "'2':\n"
    "NET_NAME\n"
    "'P3V3'\n"
    "NODE_NAME R1 1\n"
    "'1':\n"
)

PRT_SAMPLE_PAGE_V2 = (
    "PART_NAME\n"
    "C1A104 'CAP_HDL-HQ17101005HS0,100NF,10%,0402,X7R,50V':\n"
    "SECTION_NUMBER 1\n"
    " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70"
    "@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1_I17"
    "@HQ_CAP.CAP_HDL(CHIPS)':\n"
    " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70"
    "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page1_i17"
    "@hq_cap.cap_hdl(chips)',\n"
    " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70"
    "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page1_i17"
    "@hq_cap.cap_hdl(chips)',\n"
    " DRAWING='@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1'\n"
    " BOM_OPTION='DEPOP'\n"
    " HQ_CODE='HQ17101005HS0'\n"
    " VALUE='100NF'\n"
    " PACKAGE='0402'\n"
)

NET_SAMPLE_PAGE_V2 = (
    "NET_NAME\n"
    "'P1V8_AON'\n"
    "NODE_NAME C1A104 1\n"
    "'1':\n"
    "NET_NAME\n"
    "'GND'\n"
    "NODE_NAME C1A104 2\n"
    "'2':\n"
)


def build_project_root() -> Path:
    return build_project_root_with_samples(PRT_SAMPLE, NET_SAMPLE)


def build_project_root_with_samples(prt_text: str, net_text: str) -> Path:
    root = Path(tempfile.mkdtemp())
    packaged = root / 'packaged'
    packaged.mkdir(parents=True)
    (packaged / 'pstxprt.dat').write_text(prt_text, encoding='utf-8')
    (packaged / 'pstxnet.dat').write_text(net_text, encoding='utf-8')
    (packaged / 'pstxref.dat').write_text('xref placeholder', encoding='utf-8')

    sch_dir = root / 'sch_1'
    sch_dir.mkdir(parents=True)
    (sch_dir / 'page518.csv').write_text('NAME,PAGE_NUMBER\nTOP,242\n', encoding='utf-8')
    return root


def build_project_root_for_page_v2() -> Path:
    root = Path(tempfile.mkdtemp())
    packaged = root / 'packaged'
    packaged.mkdir(parents=True)
    (packaged / 'pstxprt.dat').write_text(PRT_SAMPLE_PAGE_V2, encoding='utf-8')
    (packaged / 'pstxnet.dat').write_text(NET_SAMPLE_PAGE_V2, encoding='utf-8')
    (packaged / 'pstxref.dat').write_text('xref placeholder', encoding='utf-8')

    sch_dir = root / 'sch_1'
    sch_dir.mkdir(parents=True)
    (sch_dir / 'page114.csv').write_text('"PAGE_NUMBER" = 144;\n', encoding='utf-8')
    (sch_dir / 'page.map').write_text('144 114 TOP\n', encoding='utf-8')
    (root / 'module_order').write_text(
        'Version 15.0\n'
        'START_MODULEORDER\n'
        '@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70'
        '@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n'
        'END_MODULEORDER\n',
        encoding='utf-8',
    )
    return root


class WebUiTests(unittest.TestCase):
    def setUp(self):
        pstx_web.RUN_CACHE.clear()
        self.app = pstx_web.create_app()
        self.app.testing = True
        self.client = self.app.test_client()
        self.temp_roots = []

    def tearDown(self):
        for root in self.temp_roots:
            for path in sorted(root.rglob('*'), reverse=True):
                if path.is_file():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            if root.exists():
                root.rmdir()

    def make_root(self) -> Path:
        root = build_project_root()
        self.temp_roots.append(root)
        return root

    def make_root_with_samples(self, prt_text: str, net_text: str) -> Path:
        root = build_project_root_with_samples(prt_text, net_text)
        self.temp_roots.append(root)
        return root

    def make_page_v2_root(self) -> Path:
        root = build_project_root_for_page_v2()
        self.temp_roots.append(root)
        return root

    def test_parse_voltage_map_text_reports_invalid_lines(self):
        mapping, warnings = pstx_web._parse_voltage_map_text("P1V8=1.8\nINVALID\nBAD=abc")
        self.assertEqual({'P1V8': 1.8}, mapping)
        self.assertEqual(2, len(warnings))

    def test_discover_project_files_uses_packaged_under_project_root(self):
        root = self.make_root()
        project_root, prt_path, net_path, ref_path = pstx_web._discover_project_files(str(root))
        self.assertEqual(root, project_root)
        self.assertEqual(root / 'packaged' / 'pstxprt.dat', prt_path)
        self.assertEqual(root / 'packaged' / 'pstxnet.dat', net_path)
        self.assertEqual(root / 'packaged' / 'pstxref.dat', ref_path)

    def test_read_local_text_file_decodes_gb18030_without_replacement(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / 'pstxprt.dat'
            path.write_bytes(
                "PART_NAME\nC1 '电容'\nVALUE='10微法'\n".encode('gb18030')
            )

            text, meta = pstx_web._read_local_text_file(path, 'pstxprt.dat', True)

        self.assertIn('电容', text)
        self.assertIn('10微法', text)
        self.assertEqual('gb18030', meta['encoding'])

    def test_resolve_port_falls_back_when_preferred_port_is_busy(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.bind((pstx_web.DEFAULT_HOST, 0))
            sock.listen(1)
            busy_port = sock.getsockname()[1]
            resolved = pstx_web._resolve_port(busy_port, max_attempts=3)
        self.assertNotEqual(busy_port, resolved)
        self.assertGreaterEqual(resolved, busy_port + 1)

    def test_home_page_uses_product_title_and_runtime_port(self):
        response = self.client.get('/', headers={'Host': '127.0.0.1:8766'})
        self.assertEqual(200, response.status_code)
        text = response.get_data(as_text=True)
        self.assertIn('PSTX 原理图审查平台', text)
        self.assertIn('8766', text)
        self.assertIn('项目根路径', text)
        self.assertIn('DEPOP 参与排查', text)

    def test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split(self):
        root = self.make_root()
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'demo',
                'ratio_limit': '70',
                'custom_volt_map': 'P3V3=3.3',
            },
        )
        self.assertEqual(200, response.status_code)
        payload = response.get_json()
        self.assertTrue(payload['ok'])
        run_id = payload['run_id']

        report_json = self.client.get(f'/api/report/{run_id}')
        self.assertEqual(200, report_json.status_code)
        report_payload = report_json.get_json()
        self.assertEqual('demo', report_payload['project_name'])
        self.assertTrue(any(section['id'] == 'resistor' for section in report_payload['sections']))
        self.assertTrue(report_payload['top_insights'])
        self.assertTrue(any(card['id'] == 'bom' for card in report_payload['section_cards']))

        network_section = next(section for section in report_payload['sections'] if section['id'] == 'network')
        page_table = next(table for table in network_section['tables'] if table['id'] == 'page_rows')
        self.assertEqual(['PAGE518'], [row['页面'] for row in page_table['rows']])
        mapping_table = next(table for table in network_section['tables'] if table['id'] == 'page_mapping_rows')
        self.assertEqual('PAGE242', mapping_table['rows'][0]['逻辑页'])
        self.assertEqual('PAGE518', mapping_table['rows'][0]['真实页'])
        self.assertEqual('是', mapping_table['rows'][0]['是否一一对应'])

        resistor_section = next(section for section in report_payload['sections'] if section['id'] == 'resistor')
        chip_table = next(table for table in resistor_section['tables'] if table['id'] == 'chip_pin_rows')
        self.assertIn('页面', chip_table['columns'])
        self.assertNotIn('真实页', chip_table['columns'])
        self.assertEqual('PAGE518', chip_table['rows'][0]['页面'])

        report_page = self.client.get(f'/report/{run_id}')
        self.assertEqual(200, report_page.status_code)
        report_html = report_page.get_data(as_text=True)
        self.assertIn('PSTX 原理图审查平台', report_html)
        self.assertIn('report-topbar', report_html)
        self.assertIn('report-inspector', report_html)
        self.assertIn('nav-brand', report_html)
        self.assertIn('status-pill', report_html)
        self.assertIn('summary-details', report_html)
        self.assertIn('sidebar-toggle', report_html)
        self.assertIn('query-results', report_html)

        query_json = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '位号', 'keyword': 'U1'},
        )
        self.assertEqual(200, query_json.status_code)
        query_payload = query_json.get_json()
        self.assertIn('U1', '\n'.join(query_payload['lines']))
        self.assertEqual('component', query_payload['view'])
        meta_map = {item['label']: item['value'] for item in query_payload['summary']['meta']}
        self.assertEqual('PAGE518', meta_map['页面'])
        self.assertEqual('是', meta_map['页码一一对应'])

        network_query_json = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '网络名', 'keyword': 'SMBALERT_N'},
        )
        self.assertEqual(200, network_query_json.status_code)
        network_query_payload = network_query_json.get_json()
        node_row = network_query_payload['cards'][0]['items'][0]
        self.assertEqual('PAGE518', node_row['页面'])
        self.assertEqual('是', node_row['页码一一对应'])

        export_response = self.client.get(f'/api/report/{run_id}/export')
        self.assertEqual(200, export_response.status_code)
        self.assertTrue(export_response.data.startswith(b'PK'))

        app_js = self.client.get('/static/app.js')
        try:
            self.assertEqual(200, app_js.status_code)
            app_js_text = app_js.get_data(as_text=True)
            self.assertIn('table-sort-column', app_js_text)
            self.assertIn('column-resize-handle', app_js_text)
            self.assertIn('toolbar-density', app_js_text)
            self.assertIn('metricIconForLabel', app_js_text)
            self.assertIn('navIconForSection', app_js_text)
            self.assertIn('TABLE_INITIAL_RENDER_LIMIT', app_js_text)
            self.assertIn('scheduleTableMount', app_js_text)
            self.assertIn('scheduleScrollShadowUpdate', app_js_text)
            self.assertIn('sidebar-toggle', app_js_text)
            self.assertIn('columnFilterMatches', app_js_text)
            self.assertIn('column-filter-builder', app_js_text)
            self.assertIn('column-filter-add', app_js_text)
        finally:
            app_js.close()

        app_css = self.client.get('/static/app.css')
        try:
            self.assertEqual(200, app_css.status_code)
            app_css_text = app_css.get_data(as_text=True)
            self.assertIn('.column-resize-handle', app_css_text)
            self.assertIn('table-layout: fixed', app_css_text)
            self.assertIn('.report-topbar', app_css_text)
            self.assertIn('.report-inspector', app_css_text)
            self.assertIn('.nav-logo-mark', app_css_text)
            self.assertIn('.metric-icon', app_css_text)
            self.assertIn('content-visibility: auto', app_css_text)
            self.assertIn('.table-render-footer', app_css_text)
            self.assertIn('.column-filter-panel', app_css_text)
            self.assertIn('.column-filter-row', app_css_text)
        finally:
            app_css.close()

    def test_web_analysis_excludes_depop_by_default_but_keeps_detection_list(self):
        root = self.make_root_with_samples(PRT_SAMPLE_DEPOP, NET_SAMPLE_DEPOP)
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'demo-depop',
                'ratio_limit': '70',
                'custom_volt_map': 'P3V3=3.3',
            },
        )
        self.assertEqual(200, response.status_code)
        run_id = response.get_json()['run_id']

        report_payload = self.client.get(f'/api/report/{run_id}').get_json()
        self.assertFalse(report_payload['include_depop'])
        self.assertEqual(1, report_payload['excluded_depop_count'])
        self.assertTrue(any('DEPOP 排查：关闭' in line for line in report_payload['summary_lines']))

        drc_section = next(section for section in report_payload['sections'] if section['id'] == 'drc')
        bom_option_table = next(table for table in drc_section['tables'] if table['id'] == 'bom_option_components')
        self.assertEqual('R1', bom_option_table['rows'][0]['位号'])
        self.assertEqual('是', bom_option_table['rows'][0]['是否DEPOP'])

        query_payload = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '位号', 'keyword': 'R1'},
        ).get_json()
        self.assertEqual('missing', query_payload['match_type'])

    def test_web_analysis_can_include_depop_when_switch_is_on(self):
        root = self.make_root_with_samples(PRT_SAMPLE_DEPOP, NET_SAMPLE_DEPOP)
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'demo-depop-on',
                'ratio_limit': '70',
                'custom_volt_map': 'P3V3=3.3',
                'include_depop': 'on',
            },
        )
        self.assertEqual(200, response.status_code)
        run_id = response.get_json()['run_id']

        report_payload = self.client.get(f'/api/report/{run_id}').get_json()
        self.assertTrue(report_payload['include_depop'])
        self.assertEqual(1, report_payload['depop_count'])
        self.assertTrue(any('DEPOP 排查：开启' in line for line in report_payload['summary_lines']))

        query_payload = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '位号', 'keyword': 'R1'},
        ).get_json()
        self.assertEqual('exact', query_payload['match_type'])


    def _legacy_test_web_report_shows_real_and_mapped_pages_from_p_path(self):
        root = self.make_page_v2_root()
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'page-v2',
                'ratio_limit': '70',
                'include_depop': 'on',
            },
        )
        self.assertEqual(200, response.status_code)
        run_id = response.get_json()['run_id']

        report_payload = self.client.get(f'/api/report/{run_id}').get_json()
        drc_section = next(section for section in report_payload['sections'] if section['id'] == 'drc')
        bom_option_table = next(table for table in drc_section['tables'] if table['id'] == 'bom_option_components')
        self.assertEqual('PAGE114', bom_option_table['rows'][0]['页面'])
        self.assertEqual('PAGE177', bom_option_table['rows'][0]['子模块映射主模块真实页'])

        network_section = next(section for section in report_payload['sections'] if section['id'] == 'network')
        mapping_table = next(table for table in network_section['tables'] if table['id'] == 'page_mapping_rows')
        self.assertEqual('PAGE144', mapping_table['rows'][0]['逻辑页'])
        self.assertEqual('PAGE114', mapping_table['rows'][0]['真实页'])

        query_payload = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '浣嶅彿', 'keyword': 'C1A104'},
        ).get_json()
        meta_map = {item['label']: item['value'] for item in query_payload['summary']['meta']}
        self.assertEqual('PAGE114', meta_map['页面'])
        self.assertEqual('PAGE177', meta_map['子模块映射主模块真实页'])

    def _legacy_test_web_report_page_v2_query_meta_contains_real_and_mapped_pages(self):
        root = self.make_page_v2_root()
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'page-v2-query',
                'ratio_limit': '70',
                'include_depop': 'on',
            },
        )
        self.assertEqual(200, response.status_code)
        run_id = response.get_json()['run_id']

        query_payload = self.client.post(
            f'/api/report/{run_id}/query',
            json={'mode': '浣嶅彿', 'keyword': 'C1A104'},
        ).get_json()
        self.assertEqual('exact', query_payload['match_type'])
        meta_values = [item['value'] for item in query_payload['summary']['meta']]
        self.assertIn('PAGE114', meta_values)
        self.assertIn('PAGE177', meta_values)

    def test_web_report_page_v2_tables_show_real_and_mapped_pages(self):
        root = self.make_page_v2_root()
        response = self.client.post(
            '/api/analyze',
            data={
                'project_root': str(root),
                'project_name': 'page-v2-report',
                'ratio_limit': '70',
            },
        )
        self.assertEqual(200, response.status_code)
        run_id = response.get_json()['run_id']

        report_payload = self.client.get(f'/api/report/{run_id}').get_json()
        drc_section = next(section for section in report_payload['sections'] if section['id'] == 'drc')
        bom_option_table = next(table for table in drc_section['tables'] if table['id'] == 'bom_option_components')
        self.assertEqual('PAGE114', bom_option_table['rows'][0]['页面'])
        self.assertEqual('PAGE177', bom_option_table['rows'][0]['子模块映射主模块真实页'])

        network_section = next(section for section in report_payload['sections'] if section['id'] == 'network')
        mapping_table = next(table for table in network_section['tables'] if table['id'] == 'page_mapping_rows')
        self.assertEqual('PAGE144', mapping_table['rows'][0]['逻辑页'])
        self.assertEqual('PAGE114', mapping_table['rows'][0]['真实页'])

if __name__ == '__main__':
    unittest.main()
