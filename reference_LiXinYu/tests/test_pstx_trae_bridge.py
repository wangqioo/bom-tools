import json
from pathlib import Path
import threading
import unittest
from urllib import request

from pstx_apps import trae_bridge
from pstx_webapp import run_store as webapp_run_store
from pstx_webapp import state as webapp_state


class PstxTraeBridgeTests(unittest.TestCase):
    def tearDown(self):
        webapp_state.clear_web_session_state()

    def _remember_demo_web_run(self, run_id="run-web", *, project_name="web_demo", net_name="P3V3"):
        bundle = {
            "project_name": project_name,
            "project_root": f"/tmp/{project_name}",
            "components": {
                "U1": {
                    "CDS_PART_NAME": "GPU_CORE",
                    "page_submodule_mapped": "12",
                },
                "U2": {
                    "CDS_PART_NAME": "PCIE_RETIMER",
                    "page_submodule_mapped": "13",
                },
            },
            "nets": {
                net_name: [{"refdes": "U1", "pin": "VDD"}, {"refdes": "U2", "pin": "VCC"}],
                "GND": [{"refdes": "U1", "pin": "GND"}, {"refdes": "U2", "pin": "GND"}],
            },
        }
        webapp_run_store.remember_run(run_id, {
            "bundle": bundle,
            "report": {
                "project_name": project_name,
                "generated_at": "2026-05-06 10:00:00",
            },
        })
        return bundle

    def test_default_bridge_port_is_high_and_project_specific(self):
        self.assertEqual(48765, trae_bridge.DEFAULT_BRIDGE_PORT)
        self.assertGreaterEqual(trae_bridge.DEFAULT_BRIDGE_PORT, 40000)

    def test_build_cli_argv_from_json_args(self):
        argv = trae_bridge.build_cli_argv(
            "evidence-pack",
            {
                "bundle_cache_in": "out/bundle-cache.json",
                "refdes": ["U46", "PU12"],
                "net": "P3V3",
                "table_id": ["chip_pin_rows"],
                "include_depop": True,
                "pretty": False,
            },
        )
        self.assertEqual("evidence-pack", argv[0])
        self.assertIn("--bundle-cache-in", argv)
        self.assertIn("out/bundle-cache.json", argv)
        self.assertIn("--refdes", argv)
        self.assertIn("U46", argv)
        self.assertIn("PU12", argv)
        self.assertIn("--include-depop", argv)
        self.assertNotIn("--pretty", argv)

    def test_bridge_only_run_id_is_not_forwarded_as_cli_flag(self):
        argv = trae_bridge.build_cli_argv(
            "net-catalog",
            {
                "run_id": "latest",
                "query": "P3V3",
                "latest_run": True,
            },
        )
        self.assertEqual(["net-catalog", "--query", "P3V3"], argv)
        self.assertNotIn("--run-id", argv)
        self.assertNotIn("--latest-run", argv)

    def test_bridge_payload_runs_schema_without_shell(self):
        payload = trae_bridge.run_bridge_payload({
            "command": "schema",
            "args": {"schema_command": "net-catalog"},
        })
        self.assertTrue(payload["ok"])
        self.assertEqual("schema", payload["command"])
        self.assertEqual(["net-catalog"], payload["commands"])
        self.assertEqual("pstx-trae-bridge.v1", payload["bridge"]["schema_version"])
        self.assertEqual(["schema", "net-catalog"], payload["bridge"]["cli_argv"])

    def test_bridge_can_read_harness_skill_body(self):
        payload = trae_bridge.run_bridge_payload({
            "command": "harness-skills",
            "args": {
                "skill_id": "datasheet-key-info",
                "include_body": True,
            },
        })

        self.assertTrue(payload["ok"])
        self.assertEqual("harness-skills", payload["command"])
        self.assertEqual(["harness-skills", "datasheet-key-info", "--include-body"], payload["bridge"]["cli_argv"])
        self.assertEqual("datasheet-key-info", payload["harness_skills"]["skills"][0]["id"])
        self.assertIn("已确认的 datasheet 事实", payload["harness_skills"]["skills"][0]["body"])

    def test_bridge_payload_reuses_web_run_bundle_by_run_id(self):
        self._remember_demo_web_run("run-web", project_name="web_demo", net_name="P3V3")

        payload = trae_bridge.run_bridge_payload({
            "command": "net-catalog",
            "args": {
                "run_id": "run-web",
                "query": "P3V3",
                "include_nodes": True,
            },
        })

        self.assertTrue(payload["ok"])
        self.assertEqual("net-catalog", payload["command"])
        self.assertEqual("web_run", payload["bridge"]["bundle_source"])
        self.assertEqual("run-web", payload["bridge"]["run_id"])
        self.assertEqual("web_demo", payload["bridge"]["project_name"])
        self.assertIn("--bundle-cache-in", payload["bridge"]["cli_argv"])
        self.assertNotIn("--run-id", payload["bridge"]["cli_argv"])
        self.assertEqual(1, payload["net_catalog"]["matched_count"])
        self.assertEqual("P3V3", payload["net_catalog"]["items"][0]["net_name"])

    def test_bridge_payload_defaults_to_latest_web_run_when_source_missing(self):
        self._remember_demo_web_run("run-old", project_name="old_demo", net_name="P1V8")
        self._remember_demo_web_run("run-new", project_name="new_demo", net_name="P5V")

        payload = trae_bridge.run_bridge_payload({
            "command": "net-catalog",
            "args": {
                "limit": 10,
            },
        })

        self.assertTrue(payload["ok"])
        self.assertEqual("web_run_latest", payload["bridge"]["bundle_source"])
        self.assertEqual("run-new", payload["bridge"]["run_id"])
        self.assertEqual("new_demo", payload["bridge"]["project_name"])
        self.assertEqual({"P5V", "GND"}, {item["net_name"] for item in payload["net_catalog"]["items"]})

    def test_bridge_payload_rejects_unknown_web_run(self):
        with self.assertRaises(trae_bridge.BridgeArgumentError):
            trae_bridge.run_bridge_payload({
                "command": "net-catalog",
                "args": {
                    "run_id": "missing-run",
                },
            })

    def test_batch_items_accept_json_arrays(self):
        argv = trae_bridge.build_cli_argv(
            "batch-query",
            {
                "bundle_cache_in": "out/bundle-cache.json",
                "mode": "位号",
                "items": ["U1", "U2", "U46"],
            },
        )
        self.assertIn("--items", argv)
        self.assertIn("U1,U2,U46", argv)

    def test_csa_geometry_bridge_args_map_to_safe_cli_flags(self):
        argv = trae_bridge.build_cli_argv(
            "csa-geometry",
            {
                "project_root": "/tmp/project",
                "recursive": True,
                "include_arcs": True,
                "include_connectivity": True,
                "circle_two_point_mode": "bbox",
                "executor": "serial",
                "page": 3,
                "html": True,
                "html_name": "review.html",
            },
        )
        self.assertEqual("csa-geometry", argv[0])
        self.assertEqual("/tmp/project", argv[1])
        self.assertIn("--recursive", argv)
        self.assertIn("--include-arcs", argv)
        self.assertIn("--include-connectivity", argv)
        self.assertIn("--page", argv)
        self.assertIn("3", argv)
        self.assertIn("--circle-two-point-mode", argv)
        self.assertIn("bbox", argv)
        self.assertIn("--html", argv)
        self.assertIn("--html-name", argv)
        self.assertIn("review.html", argv)

    def test_cadence_index_bridge_args_map_to_safe_cli_flags(self):
        argv = trae_bridge.build_cli_argv(
            "cadence-index",
            {
                "project_root": "/tmp/project",
                "stdout": "full",
                "kind": "offpage",
                "query": "P1V8",
                "page": 7,
            },
        )
        self.assertEqual("cadence-index", argv[0])
        self.assertEqual("/tmp/project", argv[1])
        self.assertIn("--stdout", argv)
        self.assertIn("full", argv)
        self.assertIn("--kind", argv)
        self.assertIn("offpage", argv)
        self.assertIn("--query", argv)
        self.assertIn("P1V8", argv)
        self.assertIn("--page", argv)
        self.assertIn("7", argv)

    def test_schematic_pdf_bridge_args_map_to_safe_cli_flags(self):
        argv = trae_bridge.build_cli_argv(
            "schematic-pdf-annotate",
            {
                "pdf": "/tmp/schematic.pdf",
                "project_root": "/tmp/project",
                "refdes": ["U1", "R1"],
                "pdf_page_map_json": "{\"PAGE1\": 1}",
                "allow_page_number_fallback": True,
                "stdout": "full",
            },
        )
        self.assertEqual("schematic-pdf-annotate", argv[0])
        self.assertEqual("/tmp/schematic.pdf", argv[1])
        self.assertEqual("/tmp/project", argv[2])
        self.assertIn("--refdes", argv)
        self.assertIn("U1", argv)
        self.assertIn("R1", argv)
        self.assertIn("--pdf-page-map-json", argv)
        self.assertIn("--allow-page-number-fallback", argv)
        self.assertIn("{\"PAGE1\": 1}", argv)
        self.assertIn("--stdout", argv)
        self.assertIn("full", argv)

    def test_agent_run_bridge_args_map_to_positional_id(self):
        self.assertEqual(
            ["agent-run-status", "report_abc"],
            trae_bridge.build_cli_argv("agent-run-status", {"agent_run_id": "report_abc"}),
        )
        self.assertEqual(
            ["agent-run-artifacts", "report_abc"],
            trae_bridge.build_cli_argv("agent-run-artifacts", {"agent_run_id": "report_abc"}),
        )
        self.assertEqual(
            ["agent-run-trace", "report_abc"],
            trae_bridge.build_cli_argv("agent-run-trace", {"agent_run_id": "report_abc"}),
        )
        self.assertNotIn("--agent-run-id", trae_bridge.build_cli_argv("agent-run-status", {"agent_run_id": "report_abc"}))

    def test_offline_migration_bridge_args_map_to_subcommand(self):
        argv = trae_bridge.build_cli_argv(
            "offline-migration",
            {
                "offline_action": "prepare",
                "out_dir": "out/offline",
                "name": "computer-a",
                "asset_cache_dir": "out/offline/_asset_cache",
                "python_mirror": "tuna",
                "python_version": "3.10.11",
                "mineru_venv": ".venv-mineru",
                "mineru_model_dir": "C:\\mineru\\models",
                "mineru_config": "C:\\Users\\me\\.mineru\\mineru.json",
                "download_mineru_models": True,
                "mineru_model_source": "huggingface",
                "mineru_model_type": "pipeline",
                "huggingface_endpoint": "https://hf-mirror.com",
                "target_profile": "windows-rtx4060-cuda",
                "download_wheels": True,
                "include_mineru_wheels": True,
                "mineru_wheel_spec": "mineru[pipeline]",
                "strict_mineru_wheels": True,
                "no_reuse_assets": True,
                "no_zip": True,
            },
        )
        self.assertEqual("offline-migration", argv[0])
        self.assertEqual("prepare", argv[1])
        self.assertIn("--python-mirror", argv)
        self.assertIn("tuna", argv)
        self.assertIn("--target-profile", argv)
        self.assertIn("windows-rtx4060-cuda", argv)
        self.assertIn("--asset-cache-dir", argv)
        self.assertIn("out/offline/_asset_cache", argv)
        self.assertIn("--mineru-model-dir", argv)
        self.assertIn("C:\\mineru\\models", argv)
        self.assertIn("--mineru-config", argv)
        self.assertIn("--download-mineru-models", argv)
        self.assertIn("--mineru-model-source", argv)
        self.assertIn("huggingface", argv)
        self.assertIn("--huggingface-endpoint", argv)
        self.assertIn("https://hf-mirror.com", argv)
        self.assertIn("--download-wheels", argv)
        self.assertIn("--include-mineru-wheels", argv)
        self.assertIn("--mineru-wheel-spec", argv)
        self.assertIn("mineru[pipeline]", argv)
        self.assertIn("--strict-mineru-wheels", argv)
        self.assertIn("--no-reuse-assets", argv)
        self.assertIn("--no-zip", argv)

    def test_trae_skill_contains_datasheet_mineru_and_64144_playbook(self):
        skill_path = Path(__file__).resolve().parents[1] / "trae_skill" / "pstx-cli-analysis" / "SKILL.md"
        content = skill_path.read_text(encoding="utf-8")
        self.assertIn("MinerU by default", content)
        self.assertIn("prior 64144 datasheet review", content)
        self.assertIn("harness_skills/datasheet-key-info/SKILL.md", content)
        self.assertIn("datasheet-key-info", content)
        self.assertIn("power_rail_voltage", content)
        self.assertIn("get_datasheet_parameter", content)

    def test_bridge_rejects_unknown_command(self):
        with self.assertRaises(trae_bridge.BridgeArgumentError):
            trae_bridge.build_cli_argv("rm", {"rf": "/"})

    def test_http_health_and_run(self):
        server = trae_bridge.TraeBridgeServer(("127.0.0.1", 0), trae_bridge.TraeBridgeHandler)
        server.token = "secret"
        server.cors_origin = "*"
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            port = server.server_address[1]
            req = request.Request(
                f"http://127.0.0.1:{port}/v1/health",
                headers={"X-PSTX-Bridge-Token": "secret"},
            )
            with request.urlopen(req, timeout=5) as resp:
                health = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(health["ok"])
            self.assertEqual("pstx-trae-bridge.v1", health["schema_version"])
            self.assertIn("GET /v1/projects", " ".join(health["notes"]))

            body = json.dumps({
                "command": "schema",
                "args": {"schema_command": "business-dictionary"},
            }).encode("utf-8")
            req = request.Request(
                f"http://127.0.0.1:{port}/v1/run",
                data=body,
                headers={
                    "Content-Type": "application/json",
                    "X-PSTX-Bridge-Token": "secret",
                },
            )
            with request.urlopen(req, timeout=5) as resp:
                result = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(result["ok"])
            self.assertEqual(["business-dictionary"], result["commands"])
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)

    def test_http_projects_endpoint_lists_web_runs(self):
        self._remember_demo_web_run("run-http", project_name="http_demo", net_name="P3V3")
        server = trae_bridge.TraeBridgeServer(("127.0.0.1", 0), trae_bridge.TraeBridgeHandler)
        server.token = "secret"
        server.cors_origin = "*"
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            port = server.server_address[1]
            req = request.Request(
                f"http://127.0.0.1:{port}/v1/projects",
                headers={"X-PSTX-Bridge-Token": "secret"},
            )
            with request.urlopen(req, timeout=5) as resp:
                projects = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(projects["ok"])
            self.assertEqual(1, projects["count"])
            self.assertEqual("run-http", projects["latest_run_id"])
            self.assertEqual("http_demo", projects["projects"][0]["project_name"])

            req = request.Request(
                f"http://127.0.0.1:{port}/v1/projects/run-http",
                headers={"X-PSTX-Bridge-Token": "secret"},
            )
            with request.urlopen(req, timeout=5) as resp:
                project = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(project["ok"])
            self.assertEqual("run-http", project["project"]["run_id"])
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)


if __name__ == "__main__":
    unittest.main()
