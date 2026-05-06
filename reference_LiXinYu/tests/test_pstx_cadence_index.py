# -*- coding: utf-8 -*-
"""Tests for project-level Cadence semantic index."""

import tempfile
import unittest
from pathlib import Path

from pstx_core.cadence.semantic_index import build_cadence_index_payload


class CadenceSemanticIndexTests(unittest.TestCase):

    def test_cadence_index_groups_page_semantics_and_matches_pstx_nets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "sch_1").mkdir()
            (root / "sch_1" / "page1.csa").write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME SMB_ALERT;",
                    "NET_LABEL 1 (50 0) SMB_ALERT;",
                    "PORT 1 (100 0) SMB_ALERT INPUT;",
                    "OFFPAGE 1 (0 0) SMB_LINK;",
                    "BUS 1 (75 0) SMBUS[0..1];",
                    "NO_CONNECT 1 (100 0);",
                    "NET_LABEL 1 (300 300) FLOATING_LABEL;",
                ]),
                encoding="utf-8",
            )
            (root / "sch_1" / "page2.csa").write_text(
                "\n".join([
                    "WIRE 16 -1 (0 0)(100 0);",
                    "FORCEPROP 2 LAST SIG_NAME SMB_ALERT;",
                    "NET_LABEL 1 (50 0) SMB_ALERT;",
                    "OFFPAGE 1 (100 0) SMB_LINK;",
                ]),
                encoding="utf-8",
            )

            payload = build_cadence_index_payload(
                root,
                pstx_nets={"SMB_ALERT": [], "OTHER": []},
                stdout="full",
            )

        self.assertEqual("pstx-cadence-index.v1", payload["schema_version"])
        self.assertEqual(2, payload["digest"]["page_count"])
        net_row = next(row for row in payload["net_rows"] if row["normalized_name"] == "SMB_ALERT")
        self.assertTrue(net_row["pstx_net_match"])
        self.assertEqual([1, 2], net_row["pages"])
        self.assertIn("SIG_NAME", net_row["source_counts"])
        self.assertIn("NET_LABEL", net_row["source_counts"])

        port_row = payload["port_rows"][0]
        self.assertEqual("SMB_ALERT", port_row["name"])
        self.assertEqual(["INPUT"], port_row["directions"])
        self.assertEqual([1], port_row["pages"])

        link_row = payload["offpage_link_rows"][0]
        self.assertEqual("SMB_LINK", link_row["name"])
        self.assertEqual("same_name_multi_page_evidence", link_row["link_status"])
        self.assertFalse(link_row["connection_claim"])

        self.assertEqual("SMBUS[0..1]", payload["bus_rows"][0]["name"])
        self.assertEqual("NO_CONNECT", payload["no_connect_rows"][0]["name"])
        self.assertEqual("FLOATING_LABEL", payload["unbound_semantic_rows"][0]["name"])

    def test_cadence_index_filters_query_kind_page_and_limit(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "sch_1").mkdir()
            (root / "sch_1" / "page1.csa").write_text(
                "WIRE 16 -1 (0 0)(100 0);\nNET_LABEL 1 (50 0) NET_A;\n",
                encoding="utf-8",
            )
            (root / "sch_1" / "page2.csa").write_text(
                "WIRE 16 -1 (0 0)(100 0);\nNET_LABEL 1 (50 0) NET_B;\n",
                encoding="utf-8",
            )

            nets = build_cadence_index_payload(root, stdout="nets", query="net_", limit=1)
            page_filtered = build_cadence_index_payload(root, stdout="full", kind="net", page=2)

        self.assertEqual(2, nets["digest"]["net_count"])
        self.assertEqual(1, len(nets["net_rows"]))
        self.assertTrue(nets["truncated"])
        self.assertEqual([], nets["port_rows"])
        self.assertEqual(["NET_B"], [row["name"] for row in page_filtered["net_rows"]])
        self.assertEqual([], page_filtered["offpage_link_rows"])

    def test_cadence_index_empty_project_returns_warning(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            payload = build_cadence_index_payload(temp_dir, stdout="summary")

        self.assertFalse(payload["digest"]["enabled"])
        self.assertEqual(0, payload["digest"]["page_count"])
        self.assertTrue(payload["warnings"])
        self.assertEqual([], payload["net_rows"])


if __name__ == "__main__":
    unittest.main()
