import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pstx_knowledge.reference_library import (
    agent_checklist_ref_dir,
    agent_ref_dir,
    build_agent_ref_status,
    build_review_checklist_status,
    get_agent_ref_excerpt,
    get_review_checklist_excerpt,
    reindex_agent_ref,
    reindex_review_checklists,
    search_agent_ref,
    search_review_checklists,
)


class AgentRefIndexTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.ref_dir = self.root / "ref"
        self.ref_dir.mkdir()
        self.checklist_dir = self.root / "ref_checklist"
        self.checklist_dir.mkdir()
        self.data_dir = self.root / "data"
        self.checklist_data_dir = self.root / "checklist_data"
        self.old_ref_dir = os.environ.get("PSTX_AGENT_REF_DIR")
        self.old_data_dir = os.environ.get("PSTX_AGENT_REF_DATA_DIR")
        self.old_checklist_dir = os.environ.get("PSTX_AGENT_CHECKLIST_REF_DIR")
        self.old_checklist_data_dir = os.environ.get("PSTX_AGENT_CHECKLIST_DATA_DIR")
        os.environ["PSTX_AGENT_REF_DIR"] = str(self.ref_dir)
        os.environ["PSTX_AGENT_REF_DATA_DIR"] = str(self.data_dir)
        os.environ["PSTX_AGENT_CHECKLIST_REF_DIR"] = str(self.checklist_dir)
        os.environ["PSTX_AGENT_CHECKLIST_DATA_DIR"] = str(self.checklist_data_dir)

    def tearDown(self):
        if self.old_ref_dir is None:
            os.environ.pop("PSTX_AGENT_REF_DIR", None)
        else:
            os.environ["PSTX_AGENT_REF_DIR"] = self.old_ref_dir
        if self.old_data_dir is None:
            os.environ.pop("PSTX_AGENT_REF_DATA_DIR", None)
        else:
            os.environ["PSTX_AGENT_REF_DATA_DIR"] = self.old_data_dir
        if self.old_checklist_dir is None:
            os.environ.pop("PSTX_AGENT_CHECKLIST_REF_DIR", None)
        else:
            os.environ["PSTX_AGENT_CHECKLIST_REF_DIR"] = self.old_checklist_dir
        if self.old_checklist_data_dir is None:
            os.environ.pop("PSTX_AGENT_CHECKLIST_DATA_DIR", None)
        else:
            os.environ["PSTX_AGENT_CHECKLIST_DATA_DIR"] = self.old_checklist_data_dir
        self.temp_dir.cleanup()

    def test_reindex_search_and_excerpt_ref_pdf(self):
        (self.ref_dir / "agent_capability_manual.pdf").write_bytes(b"%PDF fake")
        with mock.patch(
            "pstx_knowledge.reference_library._extract_pdf_pages",
            return_value=("indexed", ["Agent Lab can search ref PDF evidence and cite page one."], "fake", ""),
        ):
            result = reindex_agent_ref(force=True)

        self.assertTrue(result["ok"])
        self.assertEqual(1, result["indexed_count"])
        status = build_agent_ref_status()
        self.assertEqual(1, status["indexed_count"])
        self.assertEqual(1, status["page_count"])

        search = search_agent_ref("Agent Lab evidence", limit=5)
        self.assertTrue(search["ok"])
        self.assertEqual(1, search["total_matches"])
        self.assertEqual(1, search["matches"][0]["doc_id"])
        self.assertEqual(1, search["matches"][0]["page"])

        excerpt = get_agent_ref_excerpt(search["matches"][0]["doc_id"], 1, max_chars=20)
        self.assertTrue(excerpt["ok"])
        self.assertTrue(excerpt["truncated"])
        self.assertIn("Agent Lab", excerpt["content"])

    def test_ref_search_and_excerpt_ignore_stale_documents_outside_current_root(self):
        (self.ref_dir / "old_manual.pdf").write_bytes(b"%PDF fake")
        with mock.patch(
            "pstx_knowledge.reference_library._extract_pdf_pages",
            return_value=("indexed", ["STALE_REF_TOKEN should disappear after root changes."], "fake", ""),
        ):
            reindex_agent_ref(force=True)
        old_search = search_agent_ref("STALE_REF_TOKEN", limit=5)
        self.assertEqual(1, old_search["total_matches"])
        old_doc_id = old_search["matches"][0]["doc_id"]

        new_ref_dir = self.root / "new_ref"
        new_ref_dir.mkdir()
        os.environ["PSTX_AGENT_REF_DIR"] = str(new_ref_dir)

        self.assertEqual(0, search_agent_ref("STALE_REF_TOKEN", limit=5)["total_matches"])
        stale_excerpt = get_agent_ref_excerpt(old_doc_id, 1)
        self.assertFalse(stale_excerpt["ok"])
        self.assertIn("不属于当前 ref 目录", stale_excerpt["error"])

    def test_source_dir_helpers_do_not_create_reference_roots(self):
        missing_ref = self.root / "missing_ref"
        missing_checklist = self.root / "missing_checklist"
        os.environ["PSTX_AGENT_REF_DIR"] = str(missing_ref)
        os.environ["PSTX_AGENT_CHECKLIST_REF_DIR"] = str(missing_checklist)

        self.assertEqual(missing_ref, agent_ref_dir())
        self.assertEqual(missing_checklist, agent_checklist_ref_dir())
        self.assertFalse(missing_ref.exists())
        self.assertFalse(missing_checklist.exists())

    def test_review_checklist_indexes_markdown_csv_and_excel(self):
        (self.checklist_dir / "gpu_review_notes.md").write_text(
            "# GPU Review\n- DDR VREF 上下拉重复，需要检查 U46 和 R120。\n- BOM_OPTION 未画圈要复核。",
            encoding="utf-8",
        )
        (self.checklist_dir / "review_changelist.csv").write_text(
            "问题,位号,建议\n未命名网络,U46,展示页码\n",
            encoding="utf-8",
        )

        try:
            from openpyxl import Workbook  # type: ignore
        except Exception:
            Workbook = None
        if Workbook is not None:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Checklist"
            sheet.append(["场景", "问题", "处理"])
            sheet.append(["电容降额", "差分同极性耦合电容低风险", "标记低风险通过"])
            workbook.save(self.checklist_dir / "review_cases.xlsx")

        result = reindex_review_checklists(force=True)

        self.assertTrue(result["ok"])
        self.assertGreaterEqual(result["indexed_count"], 2)
        status = build_review_checklist_status()
        self.assertGreaterEqual(status["indexed_count"], 2)

        search = search_review_checklists("U46 页码", limit=5)
        self.assertTrue(search["ok"])
        self.assertGreaterEqual(search["total_matches"], 1)

        excerpt = get_review_checklist_excerpt(search["matches"][0]["doc_id"], search["matches"][0]["page"], max_chars=80)
        self.assertTrue(excerpt["ok"])
        self.assertTrue(excerpt["content"])

    def test_review_checklist_marks_legacy_xls_manual_review(self):
        (self.checklist_dir / "legacy.xls").write_bytes(b"fake xls")

        result = reindex_review_checklists(force=True)

        self.assertTrue(result["ok"])
        self.assertEqual(1, result["failed_count"])
        status = build_review_checklist_status()
        self.assertEqual(1, status["failed_count"])
        self.assertIn("转换为 .xlsx", status["documents"][0]["error"])


if __name__ == "__main__":
    unittest.main()
