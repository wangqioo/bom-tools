import os
import sqlite3
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pstx_knowledge.datasheets import (
    batch_search_datasheet_chunks,
    build_datasheet_status,
    datasheet_db_path,
    get_datasheet_chunk,
    get_datasheet_excerpt,
    get_datasheet_parameter,
    list_datasheet_documents,
    match_component_datasheets,
    reindex_datasheets,
    search_datasheet_parameters,
    search_datasheet_chunks,
    search_datasheets,
    summarize_datasheet_coverage,
)


class DatasheetIndexTests(unittest.TestCase):
    def setUp(self):
        self._old_dir = os.environ.get("PSTX_DATASHEET_DIR")
        self._old_data_dir = os.environ.get("PSTX_DATASHEET_DATA_DIR")
        self._old_pdf_extractor = os.environ.get("PSTX_PDF_EXTRACTOR")
        self._old_mineru_bin = os.environ.get("PSTX_MINERU_BIN")
        self._old_mineru_backend = os.environ.get("PSTX_MINERU_BACKEND")
        self._old_mineru_device = os.environ.get("PSTX_MINERU_DEVICE")
        self._old_mineru_method = os.environ.get("PSTX_MINERU_METHOD")
        self._old_mineru_model_source = os.environ.get("PSTX_MINERU_MODEL_SOURCE")
        self._old_mineru_timeout = os.environ.get("PSTX_MINERU_TIMEOUT_SECONDS")
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / "pdfs"
        self.source.mkdir()
        self.data_dir = self.root / "data"
        os.environ["PSTX_DATASHEET_DATA_DIR"] = str(self.data_dir)

    def tearDown(self):
        if self._old_dir is None:
            os.environ.pop("PSTX_DATASHEET_DIR", None)
        else:
            os.environ["PSTX_DATASHEET_DIR"] = self._old_dir
        if self._old_data_dir is None:
            os.environ.pop("PSTX_DATASHEET_DATA_DIR", None)
        else:
            os.environ["PSTX_DATASHEET_DATA_DIR"] = self._old_data_dir
        if self._old_pdf_extractor is None:
            os.environ.pop("PSTX_PDF_EXTRACTOR", None)
        else:
            os.environ["PSTX_PDF_EXTRACTOR"] = self._old_pdf_extractor
        if self._old_mineru_bin is None:
            os.environ.pop("PSTX_MINERU_BIN", None)
        else:
            os.environ["PSTX_MINERU_BIN"] = self._old_mineru_bin
        if self._old_mineru_backend is None:
            os.environ.pop("PSTX_MINERU_BACKEND", None)
        else:
            os.environ["PSTX_MINERU_BACKEND"] = self._old_mineru_backend
        if self._old_mineru_device is None:
            os.environ.pop("PSTX_MINERU_DEVICE", None)
        else:
            os.environ["PSTX_MINERU_DEVICE"] = self._old_mineru_device
        if self._old_mineru_method is None:
            os.environ.pop("PSTX_MINERU_METHOD", None)
        else:
            os.environ["PSTX_MINERU_METHOD"] = self._old_mineru_method
        if self._old_mineru_model_source is None:
            os.environ.pop("PSTX_MINERU_MODEL_SOURCE", None)
        else:
            os.environ["PSTX_MINERU_MODEL_SOURCE"] = self._old_mineru_model_source
        if self._old_mineru_timeout is None:
            os.environ.pop("PSTX_MINERU_TIMEOUT_SECONDS", None)
        else:
            os.environ["PSTX_MINERU_TIMEOUT_SECONDS"] = self._old_mineru_timeout
        self.tmp.cleanup()

    def test_unconfigured_status_is_explicit(self):
        os.environ.pop("PSTX_DATASHEET_DIR", None)

        status = build_datasheet_status()
        reindex = reindex_datasheets()

        self.assertTrue(status["ok"])
        self.assertFalse(status["configured"])
        self.assertIn("未配置", status["summary"])
        self.assertFalse(reindex["configured"])

    def test_index_search_excerpt_and_component_match(self):
        pdf = self.source / "HQ100_GPU_CORE_TEST_IC.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQ100 GPU_CORE_TEST_IC absolute maximum ratings and I2C pins"], "fake", ""),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["indexed_count"])
        status = build_datasheet_status()
        self.assertEqual(1, status["document_count"])
        self.assertGreaterEqual(status["chunk_count"], 1)
        self.assertEqual(0, status["parameter_count"])

        documents = list_datasheet_documents()
        self.assertEqual(1, documents["total_documents"])
        self.assertEqual(1, documents["documents"][0]["chunk_count"])
        self.assertEqual(0, documents["documents"][0]["parameter_count"])

        search = search_datasheets("HQ100 GPU_CORE_TEST_IC")
        self.assertEqual(1, search["total_matches"])
        self.assertIn("HQ100", search["matches"][0]["snippet"])

        chunk_search = search_datasheet_chunks("HQ100 absolute maximum")
        self.assertEqual(1, chunk_search["total_matches"])
        self.assertEqual("p1-c1", chunk_search["matches"][0]["chunk_id"])

        chunk = get_datasheet_chunk(chunk_search["matches"][0]["doc_id"], "p1-c1", max_chars=32)
        self.assertTrue(chunk["ok"])
        self.assertEqual("p1-c1", chunk["chunk_id"])
        self.assertTrue(chunk["truncated"])
        self.assertLessEqual(len(chunk["content"]), 32)

        batch = batch_search_datasheet_chunks(["HQ100", "NO_MATCH"], limit_per_query=2)
        self.assertEqual(2, batch["query_count"])
        self.assertEqual("found", batch["items"][0]["status"])
        self.assertEqual("missing", batch["items"][1]["status"])

        excerpt = get_datasheet_excerpt(search["matches"][0]["doc_id"], 1, max_chars=20)
        self.assertTrue(excerpt["ok"])
        self.assertTrue(excerpt["truncated"])
        self.assertLessEqual(len(excerpt["content"]), 20)

        match = match_component_datasheets({
            "refdes": "U1",
            "hq_no": "HQ100",
            "spec": "GPU_CORE_TEST_IC",
            "candidate_chip_type": "gpu",
        })
        self.assertEqual("U1", match["refdes"])
        self.assertTrue(match["matches"])

    def test_reindex_extracts_deterministic_datasheet_parameters(self):
        pdf = self.source / "AMS1117_LDO.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        page_text = """
        AMS1117 1A LOW DROPOUT VOLTAGE REGULATOR
        FEATURES
        Three Terminal Adjustable or Fixed Voltages 1.5V, 1.8V, 2.5V, 2.85V, 3.3V and 5.0V
        Output Current of 1A
        Operates Down to 1V Dropout
        Line Regulation: 0.2% Max.
        Load Regulation: 0.4% Max.
        SOT-223, TO-252 and SO-8 package available.
        The dropout voltage of the device is guaranteed maximum 1.3V.

        ABSOLUTE MAXIMUM RATINGS
        Input Voltage 15V
        Lead Temperature (25 sec) 265°C
        Thermal Resistance SO-8 package Φ_JA = 1 6 0 ^ ° C / W

        APPLICATION HINTS
        The addition of 22uF solid tantalum on the output will ensure stability for all operating conditions.
        """
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", [page_text], "mineru", ""),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["indexed_count"])
        status = build_datasheet_status()
        self.assertGreaterEqual(status["parameter_count"], 7)
        documents = list_datasheet_documents()["documents"]
        self.assertGreaterEqual(documents[0]["parameter_count"], 7)

        output_current = search_datasheet_parameters(parameter_key="output_current")
        self.assertEqual(1, output_current["total_matches"])
        current_card = output_current["parameters"][0]
        self.assertEqual("Output Current", current_card["parameter_name"])
        self.assertEqual(1.0, current_card["value_typ"])
        self.assertEqual("A", current_card["unit"])

        input_voltage = search_datasheet_parameters(query="Input Voltage 15V", parameter_key="absolute_max_input_voltage")
        self.assertEqual(1, input_voltage["total_matches"])
        voltage_card = input_voltage["parameters"][0]
        self.assertEqual(15.0, voltage_card["value_max"])
        self.assertEqual("V", voltage_card["unit"])
        self.assertIn("Absolute Maximum", voltage_card["condition"])

        dropout = search_datasheet_parameters(parameter_key="dropout_voltage_max")
        self.assertEqual(1, dropout["total_matches"])
        self.assertEqual(1.3, dropout["parameters"][0]["value_max"])

        line_regulation = search_datasheet_parameters(parameter_key="line_regulation")
        self.assertEqual(1, line_regulation["total_matches"])
        self.assertEqual("%", line_regulation["parameters"][0]["unit"])
        self.assertEqual(0.2, line_regulation["parameters"][0]["value_max"])

        packages = search_datasheet_parameters(parameter_key="packages")
        self.assertEqual(1, packages["total_matches"])
        self.assertIn("SOT-223", packages["parameters"][0]["value_text"])
        self.assertIn("TO-252", packages["parameters"][0]["value_text"])

        capacitor = search_datasheet_parameters(query="22uF output stability", parameter_key="output_capacitor")
        self.assertEqual(1, capacitor["total_matches"])
        self.assertEqual(22.0, capacitor["parameters"][0]["value_min"])
        self.assertEqual("uF", capacitor["parameters"][0]["unit"])

        thermal = search_datasheet_parameters(parameter_key="thermal_resistance_ja")
        self.assertEqual(1, thermal["total_matches"])
        self.assertEqual(160.0, thermal["parameters"][0]["value_typ"])
        self.assertEqual("°C/W", thermal["parameters"][0]["unit"])
        self.assertIn("SO-8", thermal["parameters"][0]["condition"])

        lead_temperature = search_datasheet_parameters(parameter_key="lead_temperature")
        self.assertEqual(1, lead_temperature["total_matches"])
        self.assertEqual(265.0, lead_temperature["parameters"][0]["value_max"])
        self.assertEqual("°C", lead_temperature["parameters"][0]["unit"])

        fixed_voltages = search_datasheet_parameters(parameter_key="fixed_output_voltages")
        self.assertEqual(1, fixed_voltages["total_matches"])
        self.assertIn("3.3V", fixed_voltages["parameters"][0]["value_text"])

        detail = get_datasheet_parameter(current_card["parameter_id"], max_chars=400)
        self.assertTrue(detail["ok"])
        self.assertIn("Output Current", detail["summary"])
        self.assertIn("Output Current of 1A", detail["source_text"])

    def test_reindex_extracts_complex_chip_parameter_tables(self):
        pdf = self.source / "C82_PSW_64144.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        page_text = """
        3.1 工作条件及范围
        工作环境温度为0℃-70℃，相对湿度20%\\~80%。
        储存环境温度为-40℃\\~125℃，存储相对湿度：5%\\~95%。
        3.2 供电特性
        <table><tr><td>电源名称</td><td>说明</td><td>最小值</td><td>典型值</td><td>最大值</td><td>AC噪声</td><td>参考地</td></tr>
        <tr><td>VDD</td><td>数字CORE 电源</td><td>0.7125</td><td>0.75</td><td>0.7875</td><td>6%</td><td>VSS</td></tr>
        <tr><td>VDDA12_VPH_G[0-8]</td><td>SerDes 模拟电源 VPH</td><td></td><td>1.2</td><td>1.26</td><td>3%</td><td>VSS</td></tr>
        <tr><td>VDDI0</td><td>I0电源</td><td>1.71DA</td><td>1.8</td><td>1.89</td><td>6%</td><td>VSS</td></tr></table>
        3.3 功耗
        <table><tr><td>场景</td><td>电源管脚</td><td>电压/V</td><td>电流/A</td><td>功耗/W</td><td>总功耗/W</td></tr>
        <tr><td>最大功耗</td><td>VDD</td><td>0.75</td><td>122</td><td>91.5</td><td>138.84</td></tr></table>
        3.4 上下电时序
        <table><tr><td>参数</td><td>最小值</td><td>最大值</td><td>说明</td></tr>
        <tr><td>T4/T13</td><td>-5ms</td><td>5ms</td><td>VDD 和VDDA两个电源到达 50%的间隔不超过5ms。</td></tr>
        <tr><td>T8</td><td>6ms</td><td></td><td>所有电源上电完成到PWR_ON_RST_N拉高。</td></tr></table>
        5.5 热特性
        <table><tr><td>符号</td><td>参数</td><td>最小值</td><td>标准值</td><td>最大值</td><td>单位</td></tr>
        <tr><td>0JC</td><td>结壳热阻</td><td>1</td><td></td><td></td><td>C/W</td></tr></table>
        注意：如果芯片结温超过105℃限制，芯片运行或可靠性将无法保证。
        """
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", [page_text], "mineru", ""),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["indexed_count"])
        self.assertGreaterEqual(build_datasheet_status()["parameter_count"], 12)

        rails = search_datasheet_parameters(parameter_key="power_rail_voltage", limit=20)
        self.assertGreaterEqual(rails["total_matches"], 3)
        by_name = {item["parameter_name"]: item for item in rails["parameters"]}
        self.assertEqual(0.7125, by_name["VDD Voltage Range"]["value_min"])
        self.assertEqual(1.26, by_name["VDDA12_VPH_G[0-8] Voltage Range"]["value_max"])
        self.assertEqual(1.71, by_name["VDDIO Voltage Range"]["value_min"])

        timing = search_datasheet_parameters(query="T4", parameter_key="power_sequence_timing")
        self.assertEqual(1, timing["total_matches"])
        self.assertEqual("ms", timing["parameters"][0]["unit"])
        self.assertEqual(-5.0, timing["parameters"][0]["value_min"])
        self.assertEqual(5.0, timing["parameters"][0]["value_max"])

        thermal = search_datasheet_parameters(parameter_key="thermal_characteristic")
        self.assertEqual(1, thermal["total_matches"])
        self.assertEqual("θJC", thermal["parameters"][0]["condition"])
        self.assertEqual("°C/W", thermal["parameters"][0]["unit"])

        junction = search_datasheet_parameters(parameter_key="junction_temperature_limit")
        self.assertEqual(1, junction["total_matches"])
        self.assertEqual(105.0, junction["parameters"][0]["value_max"])

    def test_reindex_backfills_chunks_for_existing_page_index(self):
        pdf = self.source / "HQ200_POWER_IC.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQ200 POWER_IC recommended operating conditions"], "fake", ""),
        ):
            reindex_datasheets(force=True)
        with sqlite3.connect(datasheet_db_path()) as conn:
            conn.execute("DELETE FROM chunks")
            conn.commit()
        self.assertEqual(0, build_datasheet_status()["chunk_count"])
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQ200 POWER_IC recommended operating conditions"], "fake", ""),
        ) as extractor:
            reindex = reindex_datasheets(force=False)

        self.assertEqual(1, reindex["indexed_count"])
        self.assertEqual(0, reindex["skipped_count"])
        extractor.assert_called_once()
        self.assertGreaterEqual(build_datasheet_status()["chunk_count"], 1)

    def test_search_ignores_stale_documents_outside_current_datasheet_dir(self):
        first_source = self.source
        second_source = self.root / "other_pdfs"
        second_source.mkdir()
        first_pdf = first_source / "HQSTALE_OLD.pdf"
        second_pdf = second_source / "HQNEW_ACTIVE.pdf"
        first_pdf.write_bytes(b"%PDF old")
        second_pdf.write_bytes(b"%PDF new")
        os.environ["PSTX_DATASHEET_DIR"] = str(first_source)
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQSTALE stale document"], "fake", ""),
        ):
            reindex_datasheets(force=True)

        os.environ["PSTX_DATASHEET_DIR"] = str(second_source)
        self.assertEqual(0, search_datasheet_chunks("HQSTALE")["total_matches"])
        with mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages",
            return_value=("indexed", ["HQNEW active document"], "fake", ""),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["removed_count"])
        self.assertEqual(0, search_datasheet_chunks("HQSTALE")["total_matches"])
        self.assertEqual(1, search_datasheet_chunks("HQNEW")["total_matches"])
        self.assertEqual(1, list_datasheet_documents()["total_documents"])

    def test_coverage_marks_datasheet_gaps_without_crashing(self):
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        summary = summarize_datasheet_coverage([
            {"refdes": "U1", "category": "chip", "hq_no": "HQ404", "spec": "NO_MATCH"},
            {"refdes": "R1", "category": "passive", "hq_no": "HQR", "spec": "RES"},
        ])

        self.assertEqual(1, summary["total_key_components"])
        self.assertEqual(0, summary["matched_count"])
        self.assertEqual(1, summary["gap_count"])
        self.assertEqual("U1", summary["gap_cards"][0]["refdes"])

    def test_default_extractor_is_mineru_and_missing_cli_marks_manual_review(self):
        pdf = self.source / "HQ250_DEFAULT_MINERU.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        os.environ.pop("PSTX_PDF_EXTRACTOR", None)

        with mock.patch("pstx_knowledge.datasheet_extractors.shutil.which", return_value=None):
            status_before = build_datasheet_status()
            reindex = reindex_datasheets(force=True)

        self.assertEqual("mineru", status_before["extractor"]["mode"])
        self.assertFalse(status_before["extractor"]["mineru"]["available"])
        self.assertIn("默认 PDF 抽取需要 MinerU", status_before["extractor"]["mineru"]["error"])
        self.assertEqual(1, reindex["failed_count"])
        documents = list_datasheet_documents()["documents"]
        self.assertEqual("mineru", documents[0]["extractor"])
        self.assertEqual("needs_manual_review", documents[0]["status"])
        self.assertIn("MinerU CLI 不可用", documents[0]["error"])

    def test_auto_without_mineru_falls_back_to_pypdf(self):
        pdf = self.source / "HQ300_FALLBACK.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        os.environ["PSTX_PDF_EXTRACTOR"] = "auto"
        with mock.patch("pstx_knowledge.datasheet_extractors.shutil.which", return_value=None), mock.patch(
            "pstx_knowledge.datasheets._extract_pdf_pages_with_pypdf",
            return_value=("indexed", ["HQ300 fallback pypdf text"], "pypdf", ""),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["indexed_count"])
        documents = list_datasheet_documents()["documents"]
        self.assertEqual("pypdf", documents[0]["extractor"])
        self.assertEqual(1, search_datasheet_chunks("HQ300")["total_matches"])
        status = build_datasheet_status()
        self.assertFalse(status["extractor"]["mineru"]["available"])

    def test_mineru_mock_cli_indexes_structured_chunks(self):
        pdf = self.source / "HQ400_MINERU.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        os.environ["PSTX_PDF_EXTRACTOR"] = "mineru"
        os.environ["PSTX_MINERU_BIN"] = "/opt/mineru/bin/mineru"
        os.environ["PSTX_MINERU_DEVICE"] = "mps"
        os.environ["PSTX_MINERU_METHOD"] = "txt"
        os.environ["PSTX_MINERU_MODEL_SOURCE"] = "modelscope"

        def fake_run(cmd, **kwargs):
            if "--version" in cmd:
                return subprocess.CompletedProcess(cmd, 0, stdout="mineru 3.0.9\n", stderr="")
            self.assertEqual("mps", kwargs.get("env", {}).get("MINERU_DEVICE_MODE"))
            self.assertEqual("modelscope", kwargs.get("env", {}).get("MINERU_MODEL_SOURCE"))
            self.assertEqual("txt", cmd[cmd.index("-m") + 1])
            output_dir = Path(cmd[cmd.index("-o") + 1])
            (output_dir / "content_list.json").write_text(
                """
                [
                  {"type":"text","page_idx":0,"text":"HQ400 MinerU absolute maximum ratings"},
                  {"type":"table","page_idx":1,"table_body":"<table><tr><td>VCC</td><td>3.3V</td></tr></table>"}
                ]
                """,
                encoding="utf-8",
            )
            return subprocess.CompletedProcess(cmd, 0, stdout="ok", stderr="")

        with mock.patch("pstx_knowledge.datasheet_extractors.subprocess.run", side_effect=fake_run):
            reindex = reindex_datasheets(force=True)
            status = build_datasheet_status()

        self.assertEqual(1, reindex["indexed_count"])
        documents = list_datasheet_documents()["documents"]
        self.assertEqual("mineru", documents[0]["extractor"])
        self.assertEqual("mineru 3.0.9", status["extractor"]["mineru"]["version"])
        self.assertEqual("mps", status["extractor"]["mineru"]["device"])
        self.assertEqual("txt", status["extractor"]["mineru"]["method"])
        self.assertEqual("modelscope", status["extractor"]["mineru"]["model_source"])
        self.assertEqual(1, search_datasheet_chunks("absolute maximum")["total_matches"])
        self.assertEqual(1, search_datasheet_chunks("VCC 3.3V")["total_matches"])

    def test_forced_mineru_failure_marks_document_for_review(self):
        pdf = self.source / "HQ500_MINERU_FAIL.pdf"
        pdf.write_bytes(b"%PDF fake")
        os.environ["PSTX_DATASHEET_DIR"] = str(self.source)
        os.environ["PSTX_PDF_EXTRACTOR"] = "mineru"
        os.environ["PSTX_MINERU_BIN"] = "/opt/mineru/bin/mineru"

        with mock.patch(
            "pstx_knowledge.datasheet_extractors.subprocess.run",
            return_value=subprocess.CompletedProcess(["mineru"], 2, stdout="", stderr="boom"),
        ):
            reindex = reindex_datasheets(force=True)

        self.assertEqual(1, reindex["failed_count"])
        documents = list_datasheet_documents()["documents"]
        self.assertEqual("mineru", documents[0]["extractor"])
        self.assertEqual("needs_manual_review", documents[0]["status"])
        self.assertIn("MinerU 返回 2", documents[0]["error"])


if __name__ == "__main__":
    unittest.main()
