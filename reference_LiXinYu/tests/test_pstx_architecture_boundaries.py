import subprocess
import sys
import textwrap
import unittest
import ast
from pathlib import Path

from pstx_core import pstx_parser
from pstx_harness.compare_tools import build_compare_tool_registry
from pstx_harness.report_tools import build_default_harness_registry
from pstx_harness.tool_core import HarnessTool, HarnessToolRegistry


class ArchitectureBoundaryTests(unittest.TestCase):
    def test_parser_public_module_parses_component_and_net(self):
        part_text = """
PART_NAME
 R1 'RES_0402-HQ12345678,10K':
 SECTION_NUMBER 1
 '@LIB.BOARD(SCH_1):PAGE2_I1@HQ_RES.RES_0402(CHIPS)':
  LOCATION='R1',
  HQ_CODE='HQ12345678',
  VALUE='10K',
  P_PATH='@lib.board(sch_1):page2_i1@hq_res.res_0402(chips)';
"""
        net_text = """
NET_NAME
'NET_A'
NODE_NAME R1 1
'1':
"""

        components, nets, comp_nets = pstx_parser.parse_all(part_text, net_text)

        self.assertIn("R1", components)
        self.assertEqual("HQ12345678", components["R1"]["hq_code"])
        self.assertIn("NET_A", nets)
        self.assertEqual("NET_A", comp_nets["R1"]["1"])

    def test_runtime_imports_do_not_load_openpyxl(self):
        script = textwrap.dedent(
            """
            import sys
            import pstx_analyzer
            import pstx_exports
            import pstx_webapp.app_factory
            assert 'openpyxl' not in sys.modules, sorted(k for k in sys.modules if k.startswith('openpyxl'))[:5]
            print('ok')
            """
        )
        output = subprocess.check_output([sys.executable, "-c", script], text=True)
        self.assertEqual("ok", output.strip())

    def test_production_code_does_not_import_analyzer_compat_shim(self):
        repo_root = Path(__file__).resolve().parents[1]
        scan_roots = [
            repo_root / "pstx_core",
            repo_root / "pstx_rules",
            repo_root / "pstx_queries",
            repo_root / "pstx_knowledge",
            repo_root / "pstx_integrations",
            repo_root / "pstx_harness",
            repo_root / "pstx_agent_runtime",
            repo_root / "pstx_webapp",
            repo_root / "pstx_apps",
        ]
        offenders = []
        for root in scan_roots:
            for path in root.rglob("*.py"):
                text = path.read_text(encoding="utf-8")
                for line_no, line in enumerate(text.splitlines(), 1):
                    stripped = line.strip()
                    if stripped.startswith("import pstx_analyzer") or stripped.startswith("from pstx_analyzer import"):
                        offenders.append(f"{path.relative_to(repo_root)}:{line_no}:{stripped}")

        self.assertEqual([], offenders)

    def test_agent_runtime_does_not_import_business_layers(self):
        repo_root = Path(__file__).resolve().parents[1]
        runtime_root = repo_root / "pstx_agent_runtime"
        forbidden_prefixes = {
            "pstx_analyzer",
            "pstx_apps",
            "pstx_core",
            "pstx_exports",
            "pstx_harness",
            "pstx_integrations",
            "pstx_knowledge",
            "pstx_queries",
            "pstx_rules",
            "pstx_webapp",
        }
        offenders = []
        for path in sorted(runtime_root.rglob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                imported = []
                if isinstance(node, ast.Import):
                    imported = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    imported = [node.module]
                for module in imported:
                    root_name = module.split(".", 1)[0]
                    if root_name in forbidden_prefixes:
                        offenders.append(f"{path.relative_to(repo_root)}:{node.lineno}:{module}")

        self.assertEqual([], offenders)

    def test_harness_registries_expose_read_only_tools_only(self):
        registries = {
            "report": build_default_harness_registry(),
            "compare": build_compare_tool_registry(),
        }
        allowed_approval_scopes = {"none", "read_project_file"}
        offenders = []
        for registry_name, registry in registries.items():
            for tool in registry.list_tools():
                if tool.get("readonly") is not True:
                    offenders.append(f"{registry_name}:{tool.get('name')}:readonly={tool.get('readonly')}")
                if tool.get("file_access") and tool.get("readonly") is not True:
                    offenders.append(f"{registry_name}:{tool.get('name')}:file_access_not_readonly")
                if tool.get("mutating"):
                    offenders.append(f"{registry_name}:{tool.get('name')}:mutating={tool.get('mutating')}")
                if tool.get("approval_scope") not in allowed_approval_scopes:
                    offenders.append(f"{registry_name}:{tool.get('name')}:approval_scope={tool.get('approval_scope')}")
                if not tool.get("evidence_kind"):
                    offenders.append(f"{registry_name}:{tool.get('name')}:missing_evidence_kind")
                if "supports_parallel" not in tool:
                    offenders.append(f"{registry_name}:{tool.get('name')}:missing_supports_parallel")

        self.assertEqual([], offenders)

    def test_harness_registry_rejects_mutating_tools(self):
        def handler(_context, _args):
            return {"ok": True}

        with self.assertRaises(ValueError):
            HarnessToolRegistry().register(HarnessTool(
                "write_file",
                "写文件",
                "不允许注册会修改项目的工具。",
                "file",
                handler,
                readonly=False,
            ))
        with self.assertRaises(ValueError):
            HarnessToolRegistry().register(HarnessTool(
                "patch_project",
                "改项目",
                "不允许注册 mutating 工具。",
                "file",
                handler,
                mutating=True,
            ))


if __name__ == "__main__":
    unittest.main()
