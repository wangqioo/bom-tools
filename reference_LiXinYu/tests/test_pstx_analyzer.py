# -*- coding: utf-8 -*-
"""Smoke tests for the historical pstx_analyzer compatibility entrypoint."""

import subprocess
import sys
import textwrap
import types
import unittest
from unittest import mock


class AnalyzerCompatSmokeTests(unittest.TestCase):
    def test_import_does_not_load_openpyxl(self):
        script = textwrap.dedent(
            """
            import sys
            import pstx_analyzer
            assert 'openpyxl' not in sys.modules, sorted(k for k in sys.modules if k.startswith('openpyxl'))[:5]
            print('ok')
            """
        )
        output = subprocess.check_output([sys.executable, '-c', script], text=True)
        self.assertEqual('ok', output.strip())

    def test_compat_reexports_core_entrypoints(self):
        import pstx_analyzer

        for name in [
            'parse_all',
            'parse_pstxnet',
            'parse_pstxprt',
            'resolve_component_pages',
            'analyze_project_contents',
            'build_bom',
            'check_drc',
            'analyze_resistors',
        ]:
            self.assertTrue(callable(getattr(pstx_analyzer, name)))

    def test_export_to_excel_delegates_lazily(self):
        import pstx_analyzer

        fake_pkg = types.ModuleType('pstx_exports')
        fake_excel = types.ModuleType('pstx_exports.excel')
        calls = []

        def fake_export(data, out_path):
            calls.append((data, out_path))
            return 'delegated.xlsx'

        fake_excel.export_to_excel = fake_export
        fake_pkg.excel = fake_excel
        with mock.patch.dict(sys.modules, {'pstx_exports': fake_pkg, 'pstx_exports.excel': fake_excel}):
            result = pstx_analyzer.export_to_excel({'project_name': 'demo'}, 'out.xlsx')

        self.assertEqual('delegated.xlsx', result)
        self.assertEqual([({'project_name': 'demo'}, 'out.xlsx')], calls)

    def test_main_delegates_to_local_ui_entrypoint(self):
        import pstx_analyzer

        fake_local_ui = types.ModuleType('pstx_apps.local_ui')
        calls = []

        def fake_main():
            calls.append('main')
            return 17

        fake_local_ui.main = fake_main
        with mock.patch.dict(sys.modules, {'pstx_apps.local_ui': fake_local_ui}):
            result = pstx_analyzer.main()

        self.assertEqual(17, result)
        self.assertEqual(['main'], calls)


if __name__ == '__main__':
    unittest.main()
