import contextlib
import io
import json
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from pstx_apps import cli as pstx_cli


PRT_SAMPLE = (
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
    " DRAWING='@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1',\n"
    " BOM_OPTION='DEPOP',\n"
    " HQ_CODE='HQ17101005HS0',\n"
    " VALUE='100NF',\n"
    " PACKAGE='0402'\n"
)

NET_SAMPLE = (
    "NET_NAME\n"
    "'P1V8_AON'\n"
    "NODE_NAME C1A104 1\n"
    "'1':\n"
    "NET_NAME\n"
    "'GND'\n"
    "NODE_NAME C1A104 2\n"
    "'2':\n"
)


def make_project_root() -> Path:
    root = Path(tempfile.mkdtemp())
    packaged = root / "packaged"
    packaged.mkdir(parents=True)
    (packaged / "pstxprt.dat").write_text(PRT_SAMPLE, encoding="utf-8")
    (packaged / "pstxnet.dat").write_text(NET_SAMPLE, encoding="utf-8")
    (packaged / "pstxref.dat").write_text("xref placeholder", encoding="utf-8")
    sch_dir = root / "sch_1"
    sch_dir.mkdir(parents=True)
    (sch_dir / "page114.csv").write_text('"PAGE_NUMBER" = 144;\n', encoding="utf-8")
    (sch_dir / "page.map").write_text("144 114 TOP\n", encoding="utf-8")
    (root / "module_order").write_text(
        "Version 15.0\n"
        "START_MODULEORDER\n"
        "@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70"
        "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1) 0 1 177 34 0\n"
        "END_MODULEORDER\n",
        encoding="utf-8",
    )
    return root


def write_minimal_pdf(path: Path, *, page_count: int = 1, width: int = 600, height: int = 800) -> None:
    objects = [
        "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n",
        "2 0 obj << /Type /Pages /Kids [" + " ".join(f"{idx + 3} 0 R" for idx in range(page_count)) + f"] /Count {page_count} >> endobj\n",
    ]
    for idx in range(page_count):
        obj_id = idx + 3
        objects.append(
            f"{obj_id} 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << >> >> endobj\n"
        )
    path.write_bytes(("%PDF-1.4\n" + "".join(objects) + "trailer << /Root 1 0 R >>\n%%EOF\n").encode("ascii"))


def run_cli(argv):
    stream = io.StringIO()
    with contextlib.redirect_stdout(stream):
        code = pstx_cli.main(argv)
    return code, json.loads(stream.getvalue())


class PstxCliTests(unittest.TestCase):
    def test_capabilities_lists_machine_interfaces(self):
        code, payload = run_cli(["capabilities", "--pretty"])
        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("pstx-cli", payload["interface"])
        self.assertEqual("pstx-cli.v1", payload["schema_version"])
        self.assertIn("analyze", {item["id"] for item in payload["capabilities"]})
        self.assertIn("schema", {item["id"] for item in payload["capabilities"]})
        self.assertIn("inspect", {item["id"] for item in payload["capabilities"]})
        self.assertIn("batch-query", {item["id"] for item in payload["capabilities"]})
        self.assertIn("evidence-pack", {item["id"] for item in payload["capabilities"]})
        self.assertIn("compare", {item["id"] for item in payload["capabilities"]})
        self.assertIn("net-catalog", {item["id"] for item in payload["capabilities"]})
        self.assertIn("topology-netlist", {item["id"] for item in payload["capabilities"]})
        self.assertIn("cadence-page", {item["id"] for item in payload["capabilities"]})
        self.assertIn("cadence-index", {item["id"] for item in payload["capabilities"]})
        self.assertIn("csa-geometry", {item["id"] for item in payload["capabilities"]})
        self.assertIn("schematic-pdf-annotate", {item["id"] for item in payload["capabilities"]})
        self.assertIn("business-dictionary", {item["id"] for item in payload["capabilities"]})
        self.assertIn("harness-skills", {item["id"] for item in payload["capabilities"]})
        self.assertIn("datasheet-template", {item["id"] for item in payload["capabilities"]})
        self.assertIn("datasheet-status", {item["id"] for item in payload["capabilities"]})
        self.assertIn("datasheet-search", {item["id"] for item in payload["capabilities"]})
        self.assertIn("datasheet-parameters", {item["id"] for item in payload["capabilities"]})
        self.assertIn("offline-migration", {item["id"] for item in payload["capabilities"]})
        self.assertIn("harness_runtime", payload)
        self.assertIn("skills", payload["harness_runtime"])
        self.assertIn("playbooks", payload["harness_runtime"])
        self.assertIn("completeness", payload["harness_runtime"]["tool_contract_fields"])

    def test_schema_describes_all_or_one_command(self):
        code, payload = run_cli(["schema"])
        self.assertEqual(0, code)
        self.assertEqual("schema", payload["command"])
        self.assertIn("batch-query", payload["commands"])
        self.assertIn("report-aggregate", payload["schema"])
        self.assertIn("evidence-pack", payload["schema"])
        self.assertIn("net-catalog", payload["schema"])
        self.assertIn("topology-netlist", payload["schema"])
        self.assertIn("cadence-page", payload["schema"])
        self.assertIn("cadence-index", payload["schema"])
        self.assertIn("csa-geometry", payload["schema"])
        self.assertIn("schematic-pdf-annotate", payload["schema"])
        self.assertIn("business-dictionary", payload["schema"])
        self.assertIn("harness-skills", payload["schema"])
        self.assertIn("datasheet-template", payload["schema"])
        self.assertIn("offline-migration", payload["schema"])
        self.assertIn("error_codes", payload)

        code, single_payload = run_cli(["schema", "batch-query"])
        self.assertEqual(0, code)
        self.assertEqual(["batch-query"], single_payload["commands"])
        self.assertEqual("reads --bundle-cache-in", single_payload["schema"]["batch-query"]["cache"])

        code, cadence_schema = run_cli(["schema", "cadence-page"])
        self.assertEqual(0, code)
        self.assertEqual(["cadence-page"], cadence_schema["commands"])
        self.assertIn("--page", cadence_schema["schema"]["cadence-page"]["inputs"])

        code, cadence_index_schema = run_cli(["schema", "cadence-index"])
        self.assertEqual(0, code)
        self.assertEqual(["cadence-index"], cadence_index_schema["commands"])
        self.assertIn("--kind all|net|port|offpage|bus|no_connect|unbound", cadence_index_schema["schema"]["cadence-index"]["inputs"])

        code, csa_schema = run_cli(["schema", "csa-geometry"])
        self.assertEqual(0, code)
        self.assertEqual(["csa-geometry"], csa_schema["commands"])
        self.assertIn("--demo", csa_schema["schema"]["csa-geometry"]["inputs"])
        self.assertIn("--include-connectivity", csa_schema["schema"]["csa-geometry"]["inputs"])
        self.assertIn("--page?", csa_schema["schema"]["csa-geometry"]["inputs"])
        self.assertIn("--html", csa_schema["schema"]["csa-geometry"]["inputs"])

        code, pdf_schema = run_cli(["schema", "schematic-pdf-annotate"])
        self.assertEqual(0, code)
        self.assertEqual(["schematic-pdf-annotate"], pdf_schema["commands"])
        self.assertIn("--pdf-page-map-json?", pdf_schema["schema"]["schematic-pdf-annotate"]["inputs"])
        self.assertIn("--calibrations-json?", pdf_schema["schema"]["schematic-pdf-annotate"]["inputs"])
        self.assertIn("--allow-page-number-fallback", pdf_schema["schema"]["schematic-pdf-annotate"]["inputs"])

        code, skill_schema = run_cli(["schema", "harness-skills"])
        self.assertEqual(0, code)
        self.assertEqual(["harness-skills"], skill_schema["commands"])
        self.assertIn("--include-body", skill_schema["schema"]["harness-skills"]["inputs"])

        code, compare_schema = run_cli(["schema", "compare"])
        self.assertEqual(0, code)
        self.assertEqual(["compare"], compare_schema["commands"])
        self.assertIn("left_project_root|left_project_container|left_archive", compare_schema["schema"]["compare"]["inputs"])
        self.assertIn("right_project_root|right_project_container|right_archive", compare_schema["schema"]["compare"]["inputs"])

        code, offline_schema = run_cli(["schema", "offline-migration"])
        self.assertEqual(0, code)
        self.assertEqual(["offline-migration"], offline_schema["commands"])
        self.assertIn("build-python-url|prepare|verify", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--allow-system-python-on-b", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--target-profile?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--mineru-model-dir?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--mineru-config?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--download-mineru-models", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--mineru-model-source huggingface|modelscope", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--mineru-model-type pipeline|vlm|all", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--huggingface-endpoint?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--mineru-wheel-spec?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--strict-mineru-wheels", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--asset-cache-dir?", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--no-reuse-assets", offline_schema["schema"]["offline-migration"]["inputs"])
        self.assertIn("--skip-runtime-probe", offline_schema["schema"]["offline-migration"]["inputs"])

    def test_parse_errors_are_json_enveloped(self):
        code, payload = run_cli(["batch-query", "--mode", "不支持"])
        self.assertEqual(2, code)
        self.assertFalse(payload["ok"])
        self.assertEqual("invalid_request", payload["error_code"])
        self.assertEqual("pstx-cli.v1", payload["schema_version"])

    def test_knowledge_and_datasheet_template_cli_are_llm_readable(self):
        code, dictionary = run_cli(["business-dictionary", "--pretty"])
        self.assertEqual(0, code)
        self.assertTrue(dictionary["ok"])
        self.assertIn("PCE", dictionary["business_dictionary"]["interface_aliases"]["pcie"])

        code, templates = run_cli(["datasheet-template", "--without-questions"])
        self.assertEqual(0, code)
        self.assertTrue(templates["ok"])
        self.assertEqual("pstx-datasheet-review-template.v1", templates["datasheet_templates"]["schema_version"])
        template_ids = {item["template_id"] for item in templates["datasheet_templates"]["templates"]}
        self.assertIn("complex_chip", template_ids)

        code, template = run_cli(["datasheet-template", "complex_chip"])
        self.assertEqual(0, code)
        self.assertTrue(template["ok"])
        self.assertEqual("complex_chip", template["datasheet_template"]["template_id"])
        self.assertIn("required_evidence", template["datasheet_template"])
        self.assertIn("extraction_sections", template["datasheet_template"])
        self.assertIn("power_budget_current", json.dumps(template["datasheet_template"], ensure_ascii=False))
        self.assertIn("power down sequence timing", json.dumps(template["datasheet_template"], ensure_ascii=False))

        code, skills = run_cli(["harness-skills", "datasheet-key-info", "--include-body"])
        self.assertEqual(0, code)
        self.assertEqual("pstx-harness-skills.v1", skills["harness_skills"]["schema_version"])
        self.assertEqual("single", skills["harness_skills"]["mode"])
        self.assertEqual("datasheet-key-info", skills["harness_skills"]["skills"][0]["id"])
        self.assertIn("已确认的 datasheet 事实", skills["harness_skills"]["skills"][0]["body"])

        code, selected = run_cli([
            "harness-skills",
            "--query",
            "MinerU 读取 64144 datasheet 关键电源参数",
            "--capability-profile",
            "datasheet_qa",
            "--include-body",
        ])
        self.assertEqual(0, code)
        self.assertEqual("select", selected["harness_skills"]["mode"])
        self.assertIn("datasheet-key-info", {item["id"] for item in selected["harness_skills"]["skills"]})

    def test_datasheet_status_cli_is_structured_without_configuration(self):
        code, payload = run_cli(["datasheet-status"])
        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertIn("configured", payload["datasheet_status"])
        self.assertIn("extractor", payload["datasheet_status"])

    def test_offline_migration_prepare_verify_and_python_mirror_url(self):
        code, url_payload = run_cli([
            "offline-migration",
            "build-python-url",
            "--python-version",
            "3.10.11",
            "--python-mirror",
            "tuna",
        ])
        self.assertEqual(0, code)
        self.assertIn("mirrors.tuna.tsinghua.edu.cn", url_payload["offline_migration"]["python_url"])
        self.assertIn("python-3.10.11-embed-amd64.zip", url_payload["offline_migration"]["python_url"])

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            fake_mineru = temp / "mineru_venv" / "bin"
            fake_mineru.mkdir(parents=True)
            (fake_mineru / "mineru").write_text("# fake mineru\n", encoding="utf-8")
            fake_models = temp / "mineru_models"
            fake_models.mkdir()
            (fake_models / "layout.pt").write_text("fake model", encoding="utf-8")
            fake_config = temp / "mineru.json"
            fake_config.write_text(
                json.dumps({"models_dir": str(fake_models), "device": "__PSTX_MINERU_MODELS_DIR__"}, ensure_ascii=False),
                encoding="utf-8",
            )
            out_dir = temp / "offline_out"

            code, prepared = run_cli([
                "offline-migration",
                "prepare",
                "--project-root",
                str(Path(__file__).resolve().parents[1]),
                "--out-dir",
                str(out_dir),
                "--name",
                "smoke",
                "--python-dir",
                str(fake_python),
                "--mineru-venv",
                str(temp / "mineru_venv"),
                "--mineru-model-dir",
                str(fake_models),
                "--mineru-config",
                str(fake_config),
                "--no-zip",
            ])
            self.assertEqual(0, code)
            bundle_root = Path(prepared["written"]["bundle_root"])
            self.assertTrue((bundle_root / "offline_manifest.json").is_file())
            self.assertTrue((bundle_root / "VERIFY_OFFLINE_B.py").is_file())
            self.assertTrue((bundle_root / "RUN_VERIFY_B.sh").is_file())
            self.assertTrue((bundle_root / "RUN_VERIFY_B.ps1").is_file())
            self.assertTrue((bundle_root / "RUN_SETUP_B.bat").is_file())
            self.assertTrue((bundle_root / "RUN_SETUP_B.ps1").is_file())
            self.assertTrue((bundle_root / "CONFIGURE_B.py").is_file())
            self.assertTrue((bundle_root / "RUN_INSTALL_WHEELHOUSE_B.sh").is_file())
            self.assertTrue((bundle_root / "RUN_INSTALL_WHEELHOUSE_B.ps1").is_file())
            self.assertTrue((bundle_root / "runtime" / "mineru_models" / "layout.pt").is_file())
            self.assertTrue((bundle_root / "runtime" / "mineru_config" / "mineru.template.json").is_file())
            self.assertEqual("", prepared["written"]["zip_path"] or "")
            self.assertEqual("windows-rtx4060-cuda", prepared["offline_migration"]["target_profile"])
            self.assertIn("RUN_SETUP_B", " ".join(prepared["computer_b_command"]))

            code, verified = run_cli(["offline-migration", "verify", str(bundle_root), "--skip-runtime-probe"])
            self.assertEqual(0, code)
            verification = verified["verification"]
            self.assertTrue(verification["ok"])
            self.assertEqual("windows-rtx4060-cuda", verification["target_profile"])
            self.assertIn("runtime/python/python", verification["python"]["candidates"])
            self.assertIn("runtime/mineru_venv/bin/mineru", verification["mineru"]["candidates"])
            self.assertIsNone(verification["dependency_probe"]["ok"])

            import subprocess

            standalone = subprocess.run(
                [
                    sys.executable,
                    str(bundle_root / "VERIFY_OFFLINE_B.py"),
                    str(bundle_root),
                    "--skip-runtime-probe",
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(0, standalone.returncode, standalone.stderr or standalone.stdout)
            standalone_payload = json.loads(standalone.stdout)
            self.assertTrue(standalone_payload["ok"])

            configure = subprocess.run(
                [
                    sys.executable,
                    str(bundle_root / "CONFIGURE_B.py"),
                    str(bundle_root),
                    "--write-env",
                    "--pretty",
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(0, configure.returncode, configure.stderr or configure.stdout)
            self.assertTrue((bundle_root / "RUN_ENV_B.bat").is_file())
            self.assertTrue((bundle_root / "START_WEB_B.bat").is_file())
            generated_config = (bundle_root / "runtime" / "mineru_config" / "mineru.json").read_text(encoding="utf-8")
            self.assertIn(str(bundle_root / "runtime" / "mineru_models"), generated_config)

            (bundle_root / "project" / "pstx_cli.py").unlink()
            code, failed_verify = run_cli(["offline-migration", "verify", str(bundle_root)])
            self.assertEqual(1, code)
            self.assertFalse(failed_verify["verification"]["ok"])
            self.assertTrue(failed_verify["verification"]["issues"])

    def test_offline_migration_prepare_can_download_mineru_models_from_hf_mirror(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            fake_mineru = temp / "mineru_venv" / "bin"
            fake_mineru.mkdir(parents=True)
            (fake_mineru / "mineru").write_text("# fake mineru\n", encoding="utf-8")
            fake_downloader = fake_mineru / "mineru-models-download"
            fake_downloader.write_text("# fake downloader\n", encoding="utf-8")
            os.chmod(fake_downloader, 0o755)
            downloaded_models = temp / "downloaded_models"
            downloaded_models.mkdir()
            (downloaded_models / "layout.pt").write_text("fake model", encoding="utf-8")

            def fake_download(cmd, capture_output, text, check, env=None):
                self.assertEqual(str(fake_downloader), cmd[0])
                self.assertEqual(["-s", "huggingface", "-m", "pipeline"], cmd[1:])
                self.assertEqual("https://hf-mirror.com", env.get("HF_ENDPOINT"))
                config_path = Path(env["MINERU_TOOLS_CONFIG_JSON"])
                config_path.parent.mkdir(parents=True, exist_ok=True)
                config_path.write_text(
                    json.dumps({"models-dir": {"pipeline": str(downloaded_models)}}, ensure_ascii=False),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(cmd, 0, stdout="models ok", stderr="")

            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_download):
                code, prepared = run_cli([
                    "offline-migration",
                    "prepare",
                    "--project-root",
                    str(project),
                    "--out-dir",
                    str(temp / "offline_out"),
                    "--name",
                    "model-download",
                    "--python-dir",
                    str(fake_python),
                    "--mineru-venv",
                    str(temp / "mineru_venv"),
                    "--download-mineru-models",
                    "--mineru-model-source",
                    "huggingface",
                    "--mineru-model-type",
                    "pipeline",
                    "--huggingface-endpoint",
                    "https://hf-mirror.com",
                    "--no-zip",
                ])

            self.assertEqual(0, code)
            bundle_root = Path(prepared["written"]["bundle_root"])
            self.assertTrue((bundle_root / "runtime" / "mineru_models" / "layout.pt").is_file())
            self.assertTrue((bundle_root / "runtime" / "mineru_config" / "mineru.template.json").is_file())
            model_info = prepared["offline_migration"]["mineru"]["assets"]["models"]
            self.assertTrue(model_info["provided"])
            self.assertEqual(1, model_info["file_count"])
            self.assertTrue(model_info["download"]["ok"])
            self.assertEqual("huggingface", model_info["download"]["source"])
            self.assertEqual("pipeline", model_info["download"]["model_type"])
            self.assertEqual("https://hf-mirror.com", model_info["download"]["huggingface_endpoint"])

    def test_offline_migration_prepare_reuses_python_archive_asset_cache(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            calls = {"download": 0}

            def fake_download(url, target):
                calls["download"] += 1
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(b"fake portable python archive")

            base_args = [
                "offline-migration",
                "prepare",
                "--project-root",
                str(project),
                "--out-dir",
                str(temp / "offline_out"),
                "--name",
                "python-cache",
                "--python-url",
                "https://mirror.example/python-3.10.11-embed-amd64.zip",
                "--no-extract-python",
                "--allow-system-python-on-b",
                "--no-zip",
            ]
            with mock.patch("pstx_apps.offline_migration._download_file", side_effect=fake_download):
                code, first = run_cli(base_args)
                self.assertEqual(0, code)
                code, second = run_cli(base_args)
                self.assertEqual(0, code)
            self.assertEqual(1, calls["download"])
            self.assertFalse(first["offline_migration"]["asset_cache"]["python_archive"]["hit"])
            self.assertTrue(second["offline_migration"]["asset_cache"]["python_archive"]["hit"])
            self.assertTrue((Path(second["written"]["bundle_root"]) / "runtime" / "python_archive" / "python-3.10.11-embed-amd64.zip").is_file())

    def test_offline_migration_prepare_reuses_mineru_model_asset_cache(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            fake_mineru = temp / "mineru_venv" / "bin"
            fake_mineru.mkdir(parents=True)
            (fake_mineru / "mineru").write_text("# fake mineru\n", encoding="utf-8")
            fake_downloader = fake_mineru / "mineru-models-download"
            fake_downloader.write_text("# fake downloader\n", encoding="utf-8")
            os.chmod(fake_downloader, 0o755)
            downloaded_models = temp / "downloaded_models"
            downloaded_models.mkdir()
            (downloaded_models / "layout.pt").write_text("fake model", encoding="utf-8")
            calls = {"download": 0}

            def fake_download(cmd, capture_output, text, check, env=None):
                calls["download"] += 1
                config_path = Path(env["MINERU_TOOLS_CONFIG_JSON"])
                config_path.parent.mkdir(parents=True, exist_ok=True)
                config_path.write_text(
                    json.dumps({"models-dir": {"pipeline": str(downloaded_models)}}, ensure_ascii=False),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(cmd, 0, stdout="models ok", stderr="")

            base_args = [
                "offline-migration",
                "prepare",
                "--project-root",
                str(project),
                "--out-dir",
                str(temp / "offline_out"),
                "--name",
                "model-cache",
                "--python-dir",
                str(fake_python),
                "--mineru-venv",
                str(temp / "mineru_venv"),
                "--download-mineru-models",
                "--mineru-model-source",
                "huggingface",
                "--mineru-model-type",
                "pipeline",
                "--huggingface-endpoint",
                "https://hf-mirror.com",
                "--no-zip",
            ]
            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_download):
                code, first = run_cli(base_args)
                self.assertEqual(0, code)
                code, second = run_cli(base_args)
                self.assertEqual(0, code)
            self.assertEqual(1, calls["download"])
            self.assertFalse(first["offline_migration"]["asset_cache"]["mineru_models"]["hit"])
            self.assertTrue(second["offline_migration"]["asset_cache"]["mineru_models"]["hit"])
            self.assertTrue((Path(second["written"]["bundle_root"]) / "runtime" / "mineru_models" / "layout.pt").is_file())

    def test_offline_migration_prepare_bootstraps_mineru_venv_for_model_download(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            downloaded_models = temp / "downloaded_models"
            downloaded_models.mkdir()
            (downloaded_models / "layout.pt").write_text("fake model", encoding="utf-8")

            def fake_run(cmd, capture_output, text, check, env=None, input=None):
                command = [str(part) for part in cmd]
                if command[:3] == [sys.executable, "-m", "venv"]:
                    venv_root = Path(command[3])
                    bin_dir = venv_root / "bin"
                    bin_dir.mkdir(parents=True)
                    (bin_dir / "python").write_text("# fake python\n", encoding="utf-8")
                    os.chmod(bin_dir / "python", 0o755)
                    return subprocess.CompletedProcess(cmd, 0, stdout="created venv", stderr="")
                if command[1:4] == ["-m", "pip", "install"]:
                    bin_dir = Path(command[0]).parent
                    (bin_dir / "mineru").write_text("# fake mineru\n", encoding="utf-8")
                    (bin_dir / "mineru-models-download").write_text("# fake downloader\n", encoding="utf-8")
                    os.chmod(bin_dir / "mineru", 0o755)
                    os.chmod(bin_dir / "mineru-models-download", 0o755)
                    return subprocess.CompletedProcess(cmd, 0, stdout="installed mineru", stderr="")
                if command[0].endswith("mineru-models-download"):
                    self.assertEqual(["-s", "huggingface", "-m", "pipeline"], command[1:])
                    config_path = Path(env["MINERU_TOOLS_CONFIG_JSON"])
                    config_path.parent.mkdir(parents=True, exist_ok=True)
                    config_path.write_text(
                        json.dumps({"models-dir": {"pipeline": str(downloaded_models)}}, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    return subprocess.CompletedProcess(cmd, 0, stdout="models ok", stderr="")
                return subprocess.CompletedProcess(cmd, 1, stdout="", stderr=f"unexpected command: {command}")

            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_run):
                code, prepared = run_cli([
                    "offline-migration",
                    "prepare",
                    "--project-root",
                    str(project),
                    "--out-dir",
                    str(temp / "offline_out"),
                    "--name",
                    "bootstrap-mineru",
                    "--python-dir",
                    str(fake_python),
                    "--download-mineru-models",
                    "--mineru-model-source",
                    "huggingface",
                    "--mineru-model-type",
                    "pipeline",
                    "--huggingface-endpoint",
                    "https://hf-mirror.com",
                    "--no-zip",
                ])

            self.assertEqual(0, code)
            bundle_root = Path(prepared["written"]["bundle_root"])
            self.assertTrue((bundle_root / "runtime" / "mineru_venv" / "bin" / "mineru").is_file())
            self.assertTrue((bundle_root / "runtime" / "mineru_models" / "layout.pt").is_file())
            mineru_info = prepared["offline_migration"]["mineru"]
            self.assertTrue(mineru_info["provided"])
            self.assertTrue(mineru_info["source"].endswith(".venv-mineru"))
            bootstrap = mineru_info["venv_bootstrap"]
            self.assertTrue(bootstrap["requested"])
            self.assertTrue(bootstrap["created"])
            self.assertTrue(bootstrap["installed"])

    def test_offline_migration_prepare_can_use_mineru_module_downloader(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            fake_mineru = temp / "mineru_venv" / "bin"
            fake_mineru.mkdir(parents=True)
            (fake_mineru / "python").write_text("# fake python\n", encoding="utf-8")
            (fake_mineru / "mineru").write_text("# fake mineru\n", encoding="utf-8")
            os.chmod(fake_mineru / "python", 0o755)
            os.chmod(fake_mineru / "mineru", 0o755)
            downloaded_models = temp / "downloaded_models"
            downloaded_models.mkdir()
            (downloaded_models / "layout.pt").write_text("fake model", encoding="utf-8")

            def fake_download(cmd, capture_output, text, check, env=None, input=None):
                self.assertEqual(
                    [str(fake_mineru / "python"), "-m", "mineru.cli.models_download", "-s", "huggingface", "-m", "pipeline"],
                    [str(part) for part in cmd],
                )
                config_path = Path(env["MINERU_TOOLS_CONFIG_JSON"])
                config_path.parent.mkdir(parents=True, exist_ok=True)
                config_path.write_text(
                    json.dumps({"models-dir": {"pipeline": str(downloaded_models)}}, ensure_ascii=False),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess(cmd, 0, stdout="models ok", stderr="")

            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_download):
                code, prepared = run_cli([
                    "offline-migration",
                    "prepare",
                    "--project-root",
                    str(project),
                    "--out-dir",
                    str(temp / "offline_out"),
                    "--name",
                    "module-downloader",
                    "--python-dir",
                    str(fake_python),
                    "--mineru-venv",
                    str(temp / "mineru_venv"),
                    "--download-mineru-models",
                    "--mineru-model-source",
                    "huggingface",
                    "--mineru-model-type",
                    "pipeline",
                    "--no-zip",
                ])

            self.assertEqual(0, code)
            download = prepared["offline_migration"]["mineru"]["assets"]["models"]["download"]
            self.assertEqual([str(fake_mineru / "python"), "-m", "mineru.cli.models_download", "-s", "huggingface", "-m", "pipeline"], download["command"])

    def test_offline_migration_prepare_treats_mineru_wheels_as_best_effort_with_venv(self):
        import subprocess

        def fake_pip_download(cmd, capture_output, text, check):
            self.assertIn("pip", cmd)
            self.assertIn("download", cmd)
            if "mineru[pipeline]" in cmd:
                return subprocess.CompletedProcess(cmd, 1, stdout="", stderr="ERROR: ResolutionImpossible")
            return subprocess.CompletedProcess(cmd, 0, stdout="project wheels ok", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            fake_mineru = temp / "mineru_venv" / "bin"
            fake_mineru.mkdir(parents=True)
            (fake_mineru / "mineru").write_text("# fake mineru\n", encoding="utf-8")

            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_pip_download):
                code, prepared = run_cli([
                    "offline-migration",
                    "prepare",
                    "--project-root",
                    str(project),
                    "--out-dir",
                    str(temp / "offline_out"),
                    "--name",
                    "mineru-best-effort",
                    "--python-dir",
                    str(fake_python),
                    "--mineru-venv",
                    str(temp / "mineru_venv"),
                    "--download-wheels",
                    "--include-mineru-wheels",
                    "--no-zip",
                ])

        self.assertEqual(0, code)
        wheelhouse = prepared["offline_migration"]["wheelhouse"]
        self.assertTrue(wheelhouse["ok"])
        self.assertTrue(wheelhouse["partial"])
        self.assertFalse(wheelhouse["mineru"]["ok"])
        self.assertEqual("mineru[pipeline]", wheelhouse["mineru"]["spec"])
        self.assertTrue(wheelhouse["warnings"])
        self.assertIn("ResolutionImpossible", wheelhouse["mineru"]["stderr"])

    def test_offline_migration_prepare_reuses_wheelhouse_asset_cache(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            calls = {"pip": 0}

            def fake_pip_download(cmd, capture_output, text, check):
                calls["pip"] += 1
                wheelhouse = Path(cmd[cmd.index("-d") + 1])
                wheelhouse.mkdir(parents=True, exist_ok=True)
                (wheelhouse / "Flask-3.0.0-py3-none-any.whl").write_text("fake wheel", encoding="utf-8")
                return subprocess.CompletedProcess(cmd, 0, stdout="project wheels ok", stderr="")

            base_args = [
                "offline-migration",
                "prepare",
                "--project-root",
                str(project),
                "--out-dir",
                str(temp / "offline_out"),
                "--name",
                "wheel-cache",
                "--python-dir",
                str(fake_python),
                "--download-wheels",
                "--no-zip",
            ]
            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_pip_download):
                code, first = run_cli(base_args)
                self.assertEqual(0, code)
                code, second = run_cli(base_args)
                self.assertEqual(0, code)
            self.assertEqual(1, calls["pip"])
            self.assertFalse(first["offline_migration"]["asset_cache"]["wheelhouse"]["hit"])
            self.assertTrue(second["offline_migration"]["asset_cache"]["wheelhouse"]["hit"])
            self.assertTrue((Path(second["written"]["bundle_root"]) / "wheelhouse" / "Flask-3.0.0-py3-none-any.whl").is_file())

    def test_offline_migration_prepare_seeds_wheelhouse_cache_when_requirements_change(self):
        import subprocess

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            requirements = project / "requirements.txt"
            requirements.write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)
            calls = {"pip": 0, "seeded_on_second": False}

            def fake_pip_download(cmd, capture_output, text, check):
                calls["pip"] += 1
                wheelhouse = Path(cmd[cmd.index("-d") + 1])
                wheelhouse.mkdir(parents=True, exist_ok=True)
                if calls["pip"] == 2:
                    calls["seeded_on_second"] = (wheelhouse / "Flask-3.0.0-py3-none-any.whl").is_file()
                    self.assertIn("--find-links", cmd)
                if "requests" in requirements.read_text(encoding="utf-8"):
                    (wheelhouse / "requests-2.32.0-py3-none-any.whl").write_text("fake requests", encoding="utf-8")
                else:
                    (wheelhouse / "Flask-3.0.0-py3-none-any.whl").write_text("fake flask", encoding="utf-8")
                return subprocess.CompletedProcess(cmd, 0, stdout="project wheels ok", stderr="")

            base_args = [
                "offline-migration",
                "prepare",
                "--project-root",
                str(project),
                "--out-dir",
                str(temp / "offline_out"),
                "--name",
                "wheel-incremental",
                "--python-dir",
                str(fake_python),
                "--download-wheels",
                "--no-zip",
            ]
            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_pip_download):
                code, first = run_cli(base_args)
                self.assertEqual(0, code)
                requirements.write_text("Flask>=3.0\nrequests>=2.32\n", encoding="utf-8")
                code, second = run_cli(base_args)
                self.assertEqual(0, code)
            self.assertEqual(2, calls["pip"])
            self.assertTrue(calls["seeded_on_second"])
            self.assertFalse(first["offline_migration"]["asset_cache"]["wheelhouse"]["hit"])
            cache_info = second["offline_migration"]["asset_cache"]["wheelhouse"]
            self.assertFalse(cache_info["hit"])
            self.assertGreaterEqual(cache_info["seeded_file_count"], 1)
            bundle_wheelhouse = Path(second["written"]["bundle_root"]) / "wheelhouse"
            self.assertTrue((bundle_wheelhouse / "Flask-3.0.0-py3-none-any.whl").is_file())
            self.assertTrue((bundle_wheelhouse / "requests-2.32.0-py3-none-any.whl").is_file())

    def test_offline_migration_prepare_uses_configurable_pipeline_mineru_spec(self):
        import subprocess

        seen_mineru_commands = []

        def fake_pip_download(cmd, capture_output, text, check):
            self.assertIn("pip", cmd)
            self.assertIn("download", cmd)
            if "mineru[pipeline]" in cmd:
                seen_mineru_commands.append(list(cmd))
                return subprocess.CompletedProcess(cmd, 0, stdout="mineru wheels ok", stderr="")
            return subprocess.CompletedProcess(cmd, 0, stdout="project wheels ok", stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            project = temp / "project"
            project.mkdir()
            (project / "requirements.txt").write_text("Flask>=3.0\n", encoding="utf-8")
            fake_python = temp / "python_runtime"
            fake_python.mkdir()
            (fake_python / "python").write_text(
                "#!/usr/bin/env sh\n"
                f"exec {json.dumps(sys.executable)} \"$@\"\n",
                encoding="utf-8",
            )
            os.chmod(fake_python / "python", 0o755)

            with mock.patch("pstx_apps.offline_migration.subprocess.run", side_effect=fake_pip_download):
                code, prepared = run_cli([
                    "offline-migration",
                    "prepare",
                    "--project-root",
                    str(project),
                    "--out-dir",
                    str(temp / "offline_out"),
                    "--name",
                    "mineru-pipeline",
                    "--python-dir",
                    str(fake_python),
                    "--download-wheels",
                    "--include-mineru-wheels",
                    "--mineru-wheel-spec",
                    "mineru[pipeline]",
                    "--no-zip",
                ])

            self.assertEqual(0, code)
            self.assertTrue(seen_mineru_commands)
            wheelhouse = prepared["offline_migration"]["wheelhouse"]
            self.assertEqual("mineru[pipeline]", wheelhouse["mineru"]["spec"])
            bundle_root = Path(prepared["written"]["bundle_root"])
            configure_text = (bundle_root / "CONFIGURE_B.py").read_text(encoding="utf-8")
            self.assertIn("should_install_mineru", configure_text)
            self.assertIn("mineru_spec", configure_text)

    def test_offline_migration_prepare_requires_portable_python_by_default(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            code, payload = run_cli([
                "offline-migration",
                "prepare",
                "--project-root",
                str(Path(__file__).resolve().parents[1]),
                "--out-dir",
                str(Path(temp_dir) / "offline_out"),
                "--name",
                "no-python",
                "--no-zip",
            ])
        self.assertEqual(2, code)
        self.assertFalse(payload["ok"])
        self.assertEqual("invalid_request", payload["error_code"])
        self.assertIn("computer B may not have Python", payload["error_message"])

    def test_prepare_migration_a_script_defaults_to_folder_bundle_and_asset_cache(self):
        root = Path(__file__).resolve().parents[1]
        ps1 = (root / "scripts" / "PREPARE_MIGRATION_A.ps1").read_text(encoding="utf-8")
        cmd = (root / "scripts" / "PREPARE_MIGRATION_A.cmd").read_text(encoding="utf-8")
        self.assertIn("[string]$AssetCacheDir", ps1)
        self.assertIn("[switch]$NoReuseAssets", ps1)
        self.assertIn("[switch]$MakeZip", ps1)
        self.assertIn("Join-Path $OutDir '_asset_cache'", ps1)
        self.assertIn("'--asset-cache-dir', $AssetCacheDir", ps1)
        self.assertIn("if (-not $MakeZip -or $NoZip)", ps1)
        self.assertIn("$argsList += '--no-zip'", ps1)
        self.assertIn("PREPARE_MIGRATION_A.ps1", cmd)

    def test_agent_run_status_and_artifacts_cli_read_workspace(self):
        from pstx_agent_runtime import AgentDurableRunStore, write_workspace_scratch_files

        old_workspace = os.environ.get("PSTX_AGENT_WORKSPACE_DIR")
        with tempfile.TemporaryDirectory() as tmp:
            os.environ["PSTX_AGENT_WORKSPACE_DIR"] = tmp
            try:
                store = AgentDurableRunStore()
                store.create_run(
                    scope_id="cli-run",
                    kind="report",
                    request={"profile": "quick_scan", "question": "CLI status"},
                    agent_run_id="cli-agent-run",
                )
                store.finish_record("cli-agent-run", {
                    "ok": True,
                    "agent_run_id": "cli-agent-run",
                    "status": "completed",
                    "answer": "CLI artifact answer",
                })
                write_workspace_scratch_files(
                    "cli-run",
                    "cli-agent-run",
                    [{"filename": "scratch-note.md", "content": "temporary note"}],
                )

                code, status = run_cli(["agent-run-status", "cli-agent-run"])
                self.assertEqual(0, code)
                self.assertTrue(status["ok"])
                self.assertEqual("completed", status["agent_run_status"]["status"])
                self.assertEqual("CLI artifact answer", status["agent_run_status"]["agent_run"]["answer"])

                code, artifacts = run_cli(["agent-run-artifacts", "cli-agent-run"])
                self.assertEqual(0, code)
                self.assertTrue(artifacts["ok"])
                names = {item["name"] for item in artifacts["agent_run_artifacts"]["artifacts"]}
                self.assertIn("result.json", names)
                self.assertIn("answer.md", names)
                self.assertIn("trace.json", names)
                self.assertIn("scratch-note.md", names)
                answer_artifact = next(item for item in artifacts["agent_run_artifacts"]["artifacts"] if item["name"] == "answer.md")
                self.assertIn("CLI artifact answer", answer_artifact["content_preview"])
                self.assertFalse(answer_artifact["content_truncated"])

                code, trace = run_cli(["agent-run-trace", "cli-agent-run"])
                self.assertEqual(0, code)
                self.assertTrue(trace["ok"])
                self.assertEqual("completed", trace["agent_run_trace"]["status"])
                self.assertEqual("cli-agent-run", trace["agent_run_trace"]["agent_run_id"])
            finally:
                if old_workspace is None:
                    os.environ.pop("PSTX_AGENT_WORKSPACE_DIR", None)
                else:
                    os.environ["PSTX_AGENT_WORKSPACE_DIR"] = old_workspace

    def test_topology_netlist_exports_llm_schema_from_bundle_cache(self):
        out_dir = Path(tempfile.mkdtemp())
        cache_path = out_dir / "bundle-cache.json"
        topology_path = out_dir / "llm-topology.json"
        cache_path.write_text(
            json.dumps({
                "bundle": {
                    "project_name": "topology_demo",
                    "components": {
                        "U1": {
                            "CDS_PART_NAME": "GPU_FPGA_CORE",
                            "page_submodule_mapped": "12",
                            "nets": {"A1": "I2C_SCL", "A2": "I2C_SDA", "VDD": "P3V3"},
                        },
                        "U2": {
                            "CDS_PART_NAME": "TXS0108_LEVEL_TRANSLATOR",
                            "page_submodule_mapped": "14",
                            "nets": {"A1": "I2C_SCL", "A2": "I2C_SDA", "VCCA": "P1V8", "VCCB": "P3V3"},
                        },
                    },
                    "nets": {
                        "I2C_SCL": [{"refdes": "U1", "pin": "A1"}, {"refdes": "U2", "pin": "A1"}],
                        "I2C_SDA": [{"refdes": "U1", "pin": "A2"}, {"refdes": "U2", "pin": "A2"}],
                        "P3V3": [{"refdes": "U1", "pin": "VDD"}, {"refdes": "U2", "pin": "VCCB"}],
                        "P1V8": [{"refdes": "U2", "pin": "VCCA"}],
                    },
                },
            }, ensure_ascii=False),
            encoding="utf-8",
        )

        code, payload = run_cli([
            "topology-netlist",
            "--bundle-cache-in",
            str(cache_path),
            "--out",
            str(topology_path),
            "--stdout",
            "full",
        ])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("topology-netlist", payload["command"])
        topology = payload["topology_netlist"]
        self.assertEqual("llm-topology.v1", topology["schema_version"])
        self.assertEqual(2, topology["node_count"])
        self.assertEqual(1, topology["edge_count"])
        self.assertIn("evidence_cards", topology)
        self.assertEqual("llm-topology-business-view.v1", topology["business_view"]["schema_version"])
        self.assertEqual(topology["counts"], payload["topology_summary"]["counts"])
        self.assertEqual(topology["business_view"], payload["topology_business_view"])
        self.assertTrue(topology_path.is_file())
        written = json.loads(topology_path.read_text(encoding="utf-8"))
        self.assertEqual("llm-topology.v1", written["schema_version"])
        self.assertIn("business_view", written)

        code, summary_payload = run_cli([
            "topology-netlist",
            "--bundle-cache-in",
            str(cache_path),
            "--stdout",
            "summary",
        ])
        self.assertEqual(0, code)
        self.assertIn("counts", summary_payload["topology_summary"])
        self.assertEqual("llm-topology-business-view.v1", summary_payload["topology_business_view"]["schema_version"])

    def test_cadence_page_cli_returns_summary_objects_and_detail(self):
        root = make_project_root()
        (root / "sch_1" / "page114.csa").write_text(
            "\n".join([
                "WIRE 16 -1 (0 0)(100 0);",
                "FORCEPROP 2 LAST SIG_NAME I2C_SCL;",
                "NET_LABEL 1 (50 0) I2C_SCL;",
                "PORT 1 (100 0) I2C_SCL OUTPUT;",
            ]),
            encoding="utf-8",
        )

        code, summary = run_cli(["cadence-page", str(root), "--page", "114"])
        self.assertEqual(0, code)
        self.assertTrue(summary["ok"])
        self.assertEqual("cadence-page", summary["command"])
        self.assertEqual("pstx-cadence-page.v1", summary["cadence_page"]["schema_version"])
        self.assertEqual(1, summary["cadence_page"]["connectivity_summary"]["semantic_counts"]["NET_LABEL"])

        code, objects = run_cli(["cadence-page", str(root), "--page", "114", "--stdout", "objects"])
        self.assertEqual(0, code)
        label_id = next(item["object_id"] for item in objects["cadence_page"]["objects"] if item["type"] == "NET_LABEL")

        code, detail = run_cli(["cadence-page", str(root), "--page", "114", "--object-id", label_id])
        self.assertEqual(0, code)
        self.assertEqual(label_id, detail["cadence_page"]["object"]["object_id"])

        code, missing = run_cli(["cadence-page", str(root), "--page", "114", "--object-id", "NO_SUCH"])
        self.assertNotEqual(0, code)
        self.assertFalse(missing["ok"])
        self.assertEqual("invalid_request", missing["error_code"])

    def test_cadence_and_csa_commands_use_project_input_snapshot(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            container = temp / "smb" / "board"
            project_root = container / "worklib" / "MAIN_MOD"
            packaged = project_root / "packaged"
            sch_dir = project_root / "sch_1"
            packaged.mkdir(parents=True)
            sch_dir.mkdir(parents=True)
            (container / "MAIN_MOD.cpm").write_text("placeholder", encoding="utf-8")
            (packaged / "pstxprt.dat").write_text(PRT_SAMPLE, encoding="utf-8")
            (packaged / "pstxnet.dat").write_text(NET_SAMPLE, encoding="utf-8")
            (sch_dir / "page114.csa").write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME I2C_SCL;",
                    "NET_LABEL 1 (50 0) I2C_SCL;",
                    "WIRE 16 -1 (50 -50)(50 50);",
                    "FORCEPROP 2 LAST SIG_NAME I2C_SCL;",
                    "DOT 1 (50 0);",
                ]),
                encoding="utf-8",
            )
            archive_path = container / "MAIN_MOD_project.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                for path in project_root.rglob("*"):
                    if path.is_file():
                        archive.write(path, path.relative_to(container))

            with mock.patch.dict(os.environ, {"PSTX_PROJECT_SNAPSHOT_DIR": str(temp / "snapshots")}):
                code, page_payload = run_cli(["cadence-page", str(container), "--page", "114"])
                self.assertEqual(0, code)
                self.assertTrue(page_payload["summary"]["project_input_snapshot"]["enabled"])
                self.assertIn("snapshots", page_payload["summary"]["project_root"])

                code, index_payload = run_cli(["cadence-index", str(container), "--stdout", "nets"])
                self.assertEqual(0, code)
                self.assertTrue(index_payload["summary"]["project_input_snapshot"]["enabled"])
                self.assertIn("I2C_SCL", {row["name"] for row in index_payload["cadence_index"]["net_rows"]})

                code, csa_payload = run_cli(["csa-geometry", str(project_root), "--stdout", "hits"])
                self.assertEqual(0, code)
                self.assertTrue(csa_payload["summary"]["project_input_snapshot"]["enabled"])
                self.assertIn("snapshots", csa_payload["summary"]["project_root"])
                self.assertEqual(1, csa_payload["csa_geometry"]["digest"]["cross_count"])

    def test_cadence_index_cli_returns_project_semantic_catalog(self):
        root = make_project_root()
        (root / "sch_1" / "page1.csa").write_text(
            "\n".join([
                "WIRE 16 -1 (0 0)(100 0);",
                "FORCEPROP 2 LAST SIG_NAME P1V8_AON;",
                "NET_LABEL 1 (50 0) P1V8_AON;",
                "PORT 1 (100 0) P1V8_AON INPUT;",
                "OFFPAGE 1 (0 0) P1V8_AON_REMOTE;",
                "NET_LABEL 1 (300 300) FLOATING_LABEL;",
            ]),
            encoding="utf-8",
        )
        (root / "sch_1" / "page2.csa").write_text(
            "WIRE 16 -1 (0 0)(100 0);\nOFFPAGE 1 (100 0) P1V8_AON_REMOTE;\n",
            encoding="utf-8",
        )
        cache_path = Path(tempfile.mkdtemp()) / "bundle-cache.json"
        cache_path.write_text(
            json.dumps({
                "bundle": {
                    "project_name": "cadence-index-demo",
                    "project_root": str(root),
                    "components": {},
                    "nets": {"P1V8_AON": [], "GND": []},
                },
            }, ensure_ascii=False),
            encoding="utf-8",
        )

        code, full = run_cli([
            "cadence-index",
            "--bundle-cache-in",
            str(cache_path),
            "--stdout",
            "full",
            "--query",
            "P1V8",
        ])
        self.assertEqual(0, code)
        self.assertEqual("cadence-index", full["command"])
        index = full["cadence_index"]
        self.assertEqual("pstx-cadence-index.v1", index["schema_version"])
        self.assertEqual("cadence-index-demo", full["summary"]["project_name"])
        self.assertEqual(1, len(index["net_rows"]))
        self.assertTrue(index["net_rows"][0]["pstx_net_match"])
        self.assertEqual(1, len(index["port_rows"]))
        self.assertEqual("same_name_multi_page_evidence", index["offpage_link_rows"][0]["link_status"])

        code, nets = run_cli(["cadence-index", str(root), "--stdout", "nets", "--limit", "1"])
        self.assertEqual(0, code)
        self.assertEqual(1, len(nets["cadence_index"]["net_rows"]))
        self.assertEqual([], nets["cadence_index"]["port_rows"])

    def test_csa_geometry_cli_demo_exports_fail_flags_and_bundle_cache(self):
        code, demo = run_cli(["csa-geometry", "--demo", "--stdout", "full"])
        self.assertEqual(0, code)
        self.assertTrue(demo["ok"])
        self.assertEqual("DEMO_OK", demo["demo"]["status"])
        self.assertEqual("pstx-csa-geometry.v1", demo["csa_geometry"]["schema_version"])
        self.assertEqual("builtin-demo", demo["summary"]["source"])
        self.assertEqual(2, demo["csa_geometry"]["digest"]["cross_count"])
        self.assertEqual(3, demo["csa_geometry"]["digest"]["circle_count"])
        self.assertIn("DOT 1", demo["csa_geometry"]["dot_cross_rows"][0]["证据上下文"])

        root = make_project_root()
        (root / "sch_1" / "page3.csa").write_text(
            "FILE_TYPE = MACRO_DRAWING;\n"
            "SET PAGE_NUMBER P3;\n"
            "WIRE 16 -1 (400 0)(500 0);\n"
            "FORCEPROP 2 LAST SIG_NAME CROSS_DOT_H\n"
            "WIRE 16 -1 (450 -50)(450 50);\n"
            "FORCEPROP 2 LAST SIG_NAME CROSS_DOT_V\n"
            "DOT 1 (450 0);\n"
            "ARC 16 -1 (3000 3000)(3100 3000)(3050 3050);\n",
            encoding="utf-8",
        )
        code, no_arcs = run_cli(["csa-geometry", str(root), "--stdout", "full"])
        self.assertEqual(0, code)
        self.assertEqual(1, no_arcs["csa_geometry"]["digest"]["cross_count"])
        self.assertEqual(0, no_arcs["csa_geometry"]["digest"]["circle_count"])

        out_dir = Path(tempfile.mkdtemp()) / "csa-report"
        code, with_arcs = run_cli([
            "csa-geometry",
            str(root),
            "--include-arcs",
            "--stdout",
            "full",
            "--out-dir",
            str(out_dir),
            "--json",
            "--html",
        ])
        self.assertEqual(0, code)
        self.assertEqual(1, with_arcs["csa_geometry"]["digest"]["circle_count"])
        self.assertTrue(Path(with_arcs["written"]["summary_csv"]).is_file())
        self.assertTrue(Path(with_arcs["written"]["json_report"]).is_file())
        self.assertTrue(Path(with_arcs["written"]["html_report"]).is_file())
        self.assertIn("CSA Geometry Report", Path(with_arcs["written"]["html_report"]).read_text(encoding="utf-8"))

        code, failed = run_cli(["csa-geometry", str(root), "--fail-on-findings"])
        self.assertEqual(1, code)
        self.assertTrue(failed["ok"])

        cache_path = Path(tempfile.mkdtemp()) / "bundle-cache.json"
        cache_path.write_text(
            json.dumps({
                "bundle": {
                    "project_name": "csa-cache-demo",
                    "project_root": str(root),
                    "components": {},
                    "nets": {},
                },
            }, ensure_ascii=False),
            encoding="utf-8",
        )
        code, cached = run_cli(["csa-geometry", "--bundle-cache-in", str(cache_path), "--stdout", "details"])
        self.assertEqual(0, code)
        self.assertEqual("csa-cache-demo", cached["summary"]["project_name"])
        self.assertEqual(1, len(cached["csa_geometry"]["dot_cross_rows"]))

        code, overlay = run_cli([
            "csa-geometry",
            "--bundle-cache-in",
            str(cache_path),
            "--include-connectivity",
            "--page",
            "3",
            "--stdout",
            "full",
        ])
        self.assertEqual(0, code)
        self.assertTrue(overlay["summary"]["include_connectivity"])
        self.assertEqual(3, overlay["summary"]["page"])
        semantic_overlay = overlay["csa_geometry"]["semantic_overlay"]
        self.assertEqual("pstx-csa-connectivity-overlay.v1", semantic_overlay["schema_version"])
        self.assertEqual(1, semantic_overlay["digest"]["dot_cross_matched_count"])
        self.assertEqual("matched", semantic_overlay["dot_cross_overlay_rows"][0]["binding_status"])

    def test_schematic_pdf_annotate_cli_returns_overlay_schema(self):
        root = make_project_root()
        pdf_path = Path(tempfile.mkdtemp()) / "schematic.pdf"
        write_minimal_pdf(pdf_path, page_count=1)
        target_json = json.dumps({
            "kind": "coordinate",
            "page": "PAGE1",
            "label": "BOM warning",
            "severity": "warning",
            "pdf_page_number": 1,
            "pdf_bbox": [10, 20, 40, 60],
        })

        code, payload = run_cli([
            "schematic-pdf-annotate",
            str(pdf_path),
            str(root),
            "--target-json",
            target_json,
            "--stdout",
            "full",
        ])

        self.assertEqual(0, code)
        self.assertEqual("schematic-pdf-annotate", payload["command"])
        annotation = payload["schematic_pdf_annotation"]
        self.assertEqual("pstx-schematic-pdf-annotation.v1", annotation["schema_version"])
        self.assertEqual(1, annotation["summary"]["matched_count"])
        self.assertEqual("explicit_pdf_bbox", annotation["annotations"][0]["confidence"])
        self.assertEqual([10.0, 20.0, 40.0, 60.0], annotation["annotations"][0]["pdf_bbox"])

    def test_net_catalog_lists_and_filters_net_labels_from_bundle_cache(self):
        cache_path = Path(tempfile.mkdtemp()) / "bundle-cache.json"
        cache_path.write_text(
            json.dumps({
                "bundle": {
                    "project_name": "net_catalog_demo",
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
                        "P3V3": [{"refdes": "U1", "pin": "VDD"}, {"refdes": "U2", "pin": "VCC"}],
                        "GND": [{"refdes": "U1", "pin": "GND"}, {"refdes": "U2", "pin": "GND"}],
                        "PCIE_TX0_P": [{"refdes": "U1", "pin": "A1"}, {"refdes": "U2", "pin": "B1"}],
                        "P5E_RX0_N": [{"refdes": "U1", "pin": "A2"}, {"refdes": "U2", "pin": "B2"}],
                        "I2C_SCL": [{"refdes": "U1", "pin": "C1"}, {"refdes": "U2", "pin": "D1"}],
                    },
                },
            }, ensure_ascii=False),
            encoding="utf-8",
        )

        code, payload = run_cli([
            "net-catalog",
            "--bundle-cache-in",
            str(cache_path),
            "--query",
            "PCE",
            "--kind",
            "differential",
            "--include-nodes",
            "--pretty",
        ])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("net-catalog", payload["command"])
        catalog = payload["net_catalog"]
        self.assertEqual("pstx-net-catalog.v1", catalog["schema_version"])
        self.assertEqual(2, catalog["matched_count"])
        self.assertEqual({"P5E_RX0_N", "PCIE_TX0_P"}, {item["net_name"] for item in catalog["items"]})
        self.assertEqual("differential", catalog["items"][0]["kind"])
        self.assertIn("PCE", catalog["filters"]["expanded_query_terms"])
        self.assertIn("nodes", catalog["items"][0])
        self.assertIn("evidence-pack", catalog["items"][0]["detail_command"])
        self.assertEqual(1, catalog["kind_counts"]["power"])
        self.assertEqual(1, catalog["kind_counts"]["ground"])

    def test_inspect_reports_project_files_and_next_commands(self):
        root = make_project_root()

        code, payload = run_cli(["inspect", str(root)])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("inspect", payload["command"])
        self.assertTrue(payload["project"]["is_directory"])
        self.assertTrue(any(item["label"] == "pstxprt.dat" and item["exists"] for item in payload["files"]))
        self.assertTrue(payload["page_sources"]["module_order_available"])
        self.assertTrue(any("bundle-cache-out" in item for item in payload["suggested_workflow"]))

    def test_inspect_reports_missing_project_files_without_failing(self):
        root = Path(tempfile.mkdtemp())

        code, payload = run_cli(["inspect", str(root)])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertFalse(payload["project"]["packaged_exists"])
        required = {item["label"]: item for item in payload["files"] if item["required"]}
        self.assertFalse(required["pstxprt.dat"]["exists"])
        self.assertFalse(required["pstxnet.dat"]["exists"])

    def test_inspect_reports_ambiguous_cpm_as_invalid_request(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "A.cpm").write_text("a", encoding="utf-8")
            (root / "B.cpm").write_text("b", encoding="utf-8")

            code, payload = run_cli(["inspect", str(root)])

        self.assertEqual(2, code)
        self.assertFalse(payload["ok"])
        self.assertEqual("invalid_request", payload["error_code"])
        self.assertIn("多个 .cpm", payload["error_message"])

    def test_analyze_writes_bundle_and_report_json(self):
        root = make_project_root()
        out_dir = Path(tempfile.mkdtemp())
        bundle_path = out_dir / "bundle.json"
        report_path = out_dir / "report.json"
        cache_path = out_dir / "bundle-cache.json"

        code, payload = run_cli([
            "analyze",
            str(root),
            "--include-total-bom",
            "--include-depop",
            "--json-out",
            str(bundle_path),
            "--bundle-cache-out",
            str(cache_path),
            "--report-json-out",
            str(report_path),
        ])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("analyze", payload["command"])
        self.assertEqual("pstx-cli.v1", payload["schema_version"])
        self.assertEqual(2, payload["module_scope"]["module_count"])
        self.assertEqual("pstx-analysis-timings.v1", payload["summary"]["analysis_timings"]["schema_version"])
        self.assertTrue(bundle_path.is_file())
        self.assertTrue(cache_path.is_file())
        self.assertTrue(report_path.is_file())
        bundle = json.loads(bundle_path.read_text(encoding="utf-8"))
        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertIn("module_review", bundle)
        self.assertIn("analysis_timings", report)
        self.assertTrue(any(row["stage"] == "report_payload" for row in report["analysis_timings"]["stages"]))
        self.assertTrue(any(section["id"] == "module" for section in report["sections"]))

        code, cached_query = run_cli([
            "query",
            "--bundle-cache-in",
            str(cache_path),
            "--mode",
            "位号",
            "--keyword",
            "C1A104",
        ])
        self.assertEqual(0, code)
        self.assertTrue(cached_query["summary"]["bundle_cache"]["loaded"])
        self.assertEqual("component", cached_query["query"]["entity_type"])

    def test_cli_analyze_uses_archive_snapshot_from_cpm_container(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            container = Path(temp_dir) / "smb" / "board"
            project_root = container / "worklib" / "MAIN_MOD"
            packaged = project_root / "packaged"
            packaged.mkdir(parents=True)
            (container / "MAIN_MOD.cpm").write_text("placeholder", encoding="utf-8")
            (packaged / "pstxprt.dat").write_text(PRT_SAMPLE, encoding="utf-8")
            (packaged / "pstxnet.dat").write_text(NET_SAMPLE, encoding="utf-8")
            (packaged / "pstxref.dat").write_text("xref", encoding="utf-8")
            archive_path = container / "MAIN_MOD_project.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                for path in project_root.rglob("*"):
                    if path.is_file():
                        archive.write(path, path.relative_to(container))

            with mock.patch.dict(os.environ, {"PSTX_PROJECT_SNAPSHOT_DIR": str(Path(temp_dir) / "snapshots")}):
                code, payload = run_cli(["analyze", str(container)])

            self.assertEqual(0, code)
            snapshot = payload["summary"]["project_input_snapshot"]
            self.assertTrue(snapshot["enabled"])
            self.assertEqual(str(archive_path), snapshot["source_archive"])
            self.assertTrue(Path(snapshot["local_archive"]).is_file())
            self.assertIn("snapshots", payload["summary"]["project_root"])

            code, inspect_payload = run_cli(["inspect", str(container)])

            self.assertEqual(0, code)
            inspect_snapshot = inspect_payload["project"]["snapshot"]
            self.assertTrue(inspect_snapshot["enabled"])
            self.assertEqual(str(archive_path), inspect_snapshot["source_archive"])
            self.assertIn("snapshots", inspect_payload["project"]["root"])

    def test_module_review_and_query_are_json_friendly(self):
        root = make_project_root()

        code, module_payload = run_cli(["module-review", str(root)])
        self.assertEqual(0, code)
        self.assertEqual(1, module_payload["module_review"]["summary"]["submodule_count"])

        code, filtered_payload = run_cli([
            "module-review",
            str(root),
            "--module-name",
            "i2c_repeater_9617_cbb_v3",
        ])
        self.assertEqual(0, code)
        self.assertTrue(filtered_payload["module_review"]["summary"]["filtered"])
        self.assertEqual(1, filtered_payload["module_review"]["summary"]["module_count"])
        self.assertEqual(["C1A104"], [row["位号"] for row in filtered_payload["module_review"]["component_rows"]])

        code, query_payload = run_cli(["query", str(root), "--include-depop", "--mode", "位号", "--keyword", "C1A104"])
        self.assertEqual(0, code)
        self.assertEqual("component", query_payload["query"]["entity_type"])
        self.assertTrue(any(item["label"] == "页码" and item["value"] == "PAGE177" for item in query_payload["query"]["summary"]["meta"]))

    def test_report_table_lists_and_pages_rows(self):
        root = make_project_root()

        code, catalog_payload = run_cli(["report-table", str(root)])
        self.assertEqual(0, code)
        self.assertEqual("report-table", catalog_payload["command"])
        self.assertIn("module_component_rows", {item["table_id"] for item in catalog_payload["tables"]})

        code, table_payload = run_cli([
            "report-table",
            str(root),
            "--table-id",
            "module_component_rows",
            "--module-name",
            "i2c_repeater_9617_cbb_v3",
            "--limit",
            "5",
        ])
        self.assertEqual(0, code)
        self.assertEqual("module_component_rows", table_payload["table"]["table_id"])
        self.assertEqual(1, table_payload["table"]["returned_count"])
        self.assertEqual("C1A104", table_payload["table"]["rows"][0]["位号"])

        code, scope_payload = run_cli([
            "report-table",
            str(root),
            "--table-id",
            "module_scope_rows",
            "--module-type",
            "子模块",
        ])
        self.assertEqual(0, code)
        self.assertEqual(1, scope_payload["table"]["returned_count"])
        self.assertEqual("子模块", scope_payload["table"]["rows"][0]["模块类型"])

    def test_report_aggregate_counts_columns_from_cache(self):
        root = make_project_root()
        cache_path = Path(tempfile.mkdtemp()) / "bundle-cache.json"
        code, _payload = run_cli([
            "analyze",
            str(root),
            "--include-depop",
            "--bundle-cache-out",
            str(cache_path),
        ])
        self.assertEqual(0, code)

        code, aggregate_payload = run_cli([
            "report-aggregate",
            "--bundle-cache-in",
            str(cache_path),
            "--table-id",
            "module_scope_rows",
            "--column",
            "模块类型",
            "--operation",
            "count",
        ])
        self.assertEqual(0, code)
        self.assertEqual("report-aggregate", aggregate_payload["command"])
        self.assertEqual(2, aggregate_payload["aggregation"]["unique_count"])
        self.assertEqual({"主模块", "子模块"}, {item["value"] for item in aggregate_payload["aggregation"]["items"]})

    def test_evidence_pack_collects_mixed_targets_from_cache(self):
        root = make_project_root()
        cache_path = Path(tempfile.mkdtemp()) / "bundle-cache.json"
        code, _payload = run_cli([
            "analyze",
            str(root),
            "--include-depop",
            "--bundle-cache-out",
            str(cache_path),
        ])
        self.assertEqual(0, code)

        code, payload = run_cli([
            "evidence-pack",
            "--bundle-cache-in",
            str(cache_path),
            "--refdes",
            "C1A104,NOPE",
            "--net",
            "GND",
            "--hq",
            "HQ17101005HS0",
            "--page",
            "177",
            "--table-id",
            "module_component_rows",
            "--limit-per-target",
            "3",
            "--table-limit",
            "2",
        ])

        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("evidence-pack", payload["command"])
        pack = payload["evidence_pack"]
        self.assertEqual(5, len(pack["items"]))
        self.assertEqual(1, len(pack["tables"]))
        self.assertEqual(4, pack["target_summary"]["found_count"])
        self.assertEqual(1, pack["target_summary"]["missing_count"])
        self.assertTrue(any(item["kind"] == "page" and item["result"]["normalized_query"] == "PAGE177" for item in pack["items"]))

    def test_batch_query_supports_multiple_entity_modes(self):
        root = make_project_root()

        code, refdes_payload = run_cli([
            "batch-query",
            str(root),
            "--include-depop",
            "--mode",
            "位号",
            "--items",
            "C1A104,NOPE",
        ])
        self.assertEqual(0, code)
        self.assertEqual("batch-query", refdes_payload["command"])
        self.assertEqual(2, refdes_payload["requested_count"])
        self.assertEqual(1, refdes_payload["found_count"])
        self.assertEqual("found", refdes_payload["results"][0]["status"])
        self.assertEqual("missing", refdes_payload["results"][1]["status"])

        code, hq_payload = run_cli([
            "batch-query",
            str(root),
            "--include-depop",
            "--mode",
            "HQ料号",
            "--items",
            "HQ17101005HS0",
            "--module-name",
            "i2c_repeater_9617_cbb_v3",
        ])
        self.assertEqual(0, code)
        self.assertTrue(hq_payload["module_scope"]["filtered"])
        self.assertEqual(1, hq_payload["results"][0]["result_count"])
        self.assertEqual("C1A104", hq_payload["results"][0]["items"][0]["位号"])

        code, page_payload = run_cli([
            "batch-query",
            str(root),
            "--include-depop",
            "--mode",
            "页码",
            "--items",
            "177",
        ])
        self.assertEqual(0, code)
        self.assertEqual("PAGE177", page_payload["results"][0]["normalized_query"])
        self.assertEqual(1, page_payload["results"][0]["result_count"])

    def test_compare_outputs_diff_totals(self):
        root = make_project_root()
        code, payload = run_cli(["compare", str(root), str(root)])
        self.assertEqual(0, code)
        self.assertTrue(payload["ok"])
        self.assertEqual("compare", payload["command"])
        self.assertEqual("pstx-cli.v1", payload["schema_version"])
        self.assertEqual(0, payload["diff_totals"]["components"])

    def test_errors_are_structured_for_machine_callers(self):
        code, payload = run_cli(["analyze", "/definitely/not/a/pstx/project"])

        self.assertEqual(2, code)
        self.assertFalse(payload["ok"])
        self.assertEqual("pstx-cli.v1", payload["schema_version"])
        self.assertIn(payload["error_code"], {"file_not_found", "invalid_request", "internal_error"})
        self.assertIn("code", payload["error"])


if __name__ == "__main__":
    unittest.main()
