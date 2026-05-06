import unittest

from pstx_harness.eval import AgentEvalError, build_agent_eval_status, run_agent_eval


class AgentEvalTests(unittest.TestCase):
    def test_eval_status_lists_cases(self):
        status = build_agent_eval_status()

        self.assertTrue(status["ok"])
        self.assertGreaterEqual(status["case_count"], 6)
        case_ids = {item["case_id"] for item in status["cases"]}
        self.assertIn("mock_quick_scan", case_ids)
        self.assertIn("profile_blocks_file_read", case_ids)

    def test_run_all_eval_cases_passes(self):
        payload = run_agent_eval()

        self.assertTrue(payload["ok"])
        self.assertEqual(payload["case_count"], payload["passed_count"])
        self.assertEqual(0, payload["failed_count"])
        self.assertEqual(100.0, payload["score"])

    def test_run_selected_case(self):
        payload = run_agent_eval(["invalid_citation_flagged"])

        self.assertTrue(payload["ok"])
        self.assertEqual(1, payload["case_count"])
        self.assertEqual("invalid_citation_flagged", payload["cases"][0]["case_id"])
        self.assertGreaterEqual(payload["cases"][0]["metrics"]["invalid_citation_count"], 1)

    def test_unknown_case_is_rejected(self):
        with self.assertRaises(AgentEvalError):
            run_agent_eval(["missing-case"])


if __name__ == "__main__":
    unittest.main()
