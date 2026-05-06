import tempfile
import unittest
from pathlib import Path

from pstx_core.schematic_pdf_annotation import (
    build_schematic_pdf_annotation_payload,
    read_pdf_metadata,
)


def write_minimal_pdf(path: Path, *, page_count: int = 2, width: int = 600, height: int = 800) -> None:
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


class SchematicPdfAnnotationTests(unittest.TestCase):
    def test_reads_pdf_metadata_with_fallback(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=3, width=640, height=480)

            metadata = read_pdf_metadata(pdf_path)

            self.assertEqual(3, metadata["page_count"])
            self.assertEqual(640.0, metadata["pages"][0]["width"])
            self.assertEqual(480.0, metadata["pages"][0]["height"])
            self.assertTrue(metadata["sha256"])

    def test_refdes_uses_pdf_text_bbox_when_available(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=1)
            bundle = {
                "components": {
                    "U1": {
                        "refdes": "U1",
                        "part_name": "IC_CPU",
                        "value": "CPU",
                        "hq_code": "PN_U1",
                        "comp_type": "IC",
                        "page": "PAGE242",
                        "page_real": "PAGE242",
                        "page_logical": "PAGE242",
                    }
                },
                "nets": {},
            }

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "U1", "severity": "warning"}],
                pdf_page_map={"PAGE242": 1},
                text_boxes=[{"text": "U1", "pdf_page_number": 1, "bbox": [100, 120, 130, 140], "source": "test"}],
                stdout="full",
            )

            annotation = payload["annotations"][0]
            self.assertEqual("pstx-schematic-pdf-annotation.v1", payload["schema_version"])
            self.assertEqual("matched", annotation["status"])
            self.assertEqual("pdf_text_match", annotation["confidence"])
            self.assertEqual([100.0, 120.0, 130.0, 140.0], annotation["pdf_bbox"])
            self.assertEqual([100.0 / 600.0, 120.0 / 800.0, 130.0 / 600.0, 140.0 / 800.0], annotation["normalized_bbox"])
            self.assertEqual("rect", annotation["overlay"]["shape"])

    def test_schematic_xy_requires_calibration_before_pdf_bbox(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=1)
            bundle = {
                "components": {
                    "R1": {
                        "refdes": "R1",
                        "part_name": "RES_0402",
                        "value": "4.7k",
                        "comp_type": "RES",
                        "page": "PAGE242",
                        "page_real": "PAGE242",
                        "page_logical": "PAGE242",
                        "xy_x": 1000,
                        "xy_y": 1000,
                        "xy": "(1000,1000)",
                    }
                },
                "nets": {},
            }

            without_calibration = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "R1"}],
                pdf_page_map={"PAGE242": 1},
                text_boxes=[],
                stdout="full",
            )
            self.assertEqual("page_only", without_calibration["annotations"][0]["confidence"])
            self.assertEqual([], without_calibration["annotations"][0]["pdf_bbox"])

            calibrated = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "R1"}],
                pdf_page_map={"PAGE242": 1},
                page_calibrations=[{
                    "project_page": "PAGE242",
                    "schematic_bbox": [0, 0, 2000, 2000],
                    "pdf_bbox": [0, 0, 600, 800],
                    "invert_y": True,
                }],
                text_boxes=[],
                stdout="full",
            )

            annotation = calibrated["annotations"][0]
            self.assertEqual("calibrated_xy", annotation["confidence"])
            self.assertEqual([291.0, 391.0, 309.0, 409.0], annotation["pdf_bbox"])
            self.assertEqual("rect", annotation["overlay"]["shape"])

    def test_net_target_expands_to_component_annotations(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=2)
            bundle = {
                "components": {
                    "U1": {"refdes": "U1", "page": "PAGE1", "page_real": "PAGE1", "page_logical": "PAGE1"},
                    "R1": {"refdes": "R1", "page": "PAGE2", "page_real": "PAGE2", "page_logical": "PAGE2"},
                },
                "nets": {"P3V3": [{"refdes": "U1", "pin": "1"}, {"refdes": "R1", "pin": "1"}]},
            }

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "net", "net": "P3V3"}],
                text_boxes=[],
                stdout="full",
            )

            self.assertEqual(2, payload["summary"]["annotation_count"])
            self.assertEqual(["U1", "R1"], [row["refdes"] for row in payload["annotations"]])
            self.assertEqual(["unmatched", "unmatched"], [row["confidence"] for row in payload["annotations"]])
            self.assertEqual([0, 0], [row["pdf_page_number"] for row in payload["annotations"]])

    def test_submodule_mapped_page_is_project_page_and_requires_reliable_pdf_map(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=300)
            bundle = {
                "components": {
                    "U9": {
                        "refdes": "U9",
                        "page": "PAGE1",
                        "page_real": "PAGE1",
                        "page_logical": "PAGE1",
                        "page_submodule_real": "PAGE1",
                        "page_submodule_mapped": "PAGE242",
                        "module_order_key": "@LIB.SUB(SCH_1):PAGE1_I1",
                    }
                },
                "nets": {},
            }

            strict_payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "U9"}],
                text_boxes=[],
                stdout="full",
            )
            strict_annotation = strict_payload["annotations"][0]
            self.assertEqual("PAGE242", strict_annotation["project_page"])
            self.assertEqual(0, strict_annotation["pdf_page_number"])
            self.assertEqual("unmatched", strict_annotation["confidence"])
            self.assertEqual("strict", strict_payload["digest"]["page_mapping_policy"])

            mapped_payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "U9"}],
                pdf_page_map={"PAGE242": 1},
                text_boxes=[],
                stdout="full",
            )
            mapped_annotation = mapped_payload["annotations"][0]
            self.assertEqual("PAGE242", mapped_annotation["project_page"])
            self.assertEqual(1, mapped_annotation["pdf_page_number"])
            self.assertEqual("pdf_page_map", mapped_annotation["pdf_page_source"])
            self.assertEqual("page_only", mapped_annotation["confidence"])

    def test_pdf_text_page_label_maps_reordered_pdf_without_number_fallback(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=3)
            bundle = {"components": {}, "nets": {}}

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "page", "page": "PAGE242"}],
                text_boxes=[{
                    "text": "PAGE242",
                    "pdf_page_number": 1,
                    "bbox": [500, 760, 560, 780],
                    "source": "test.page_label",
                }],
                stdout="full",
            )

            annotation = payload["annotations"][0]
            self.assertEqual("PAGE242", annotation["project_page"])
            self.assertEqual(1, annotation["pdf_page_number"])
            self.assertEqual("pdf_text_page_label", annotation["pdf_page_source"])
            self.assertEqual("page_only", annotation["confidence"])
            self.assertEqual({"PAGE242": 1}, payload["inputs"]["pdf_text_page_map"])

    def test_page_number_fallback_is_opt_in_and_marked_weak(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=5)
            bundle = {"components": {}, "nets": {}}

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "page", "page": "PAGE2"}],
                text_boxes=[],
                stdout="full",
                allow_page_number_fallback=True,
            )

            annotation = payload["annotations"][0]
            self.assertEqual(2, annotation["pdf_page_number"])
            self.assertEqual("page_label_number_weak", annotation["pdf_page_source"])
            self.assertEqual("page_only", annotation["confidence"])
            self.assertEqual("weak_page_number_fallback", payload["digest"]["page_mapping_policy"])

    def test_ambiguous_pdf_text_page_label_is_not_used(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=3)
            bundle = {"components": {}, "nets": {}}

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "page", "page": "PAGE242"}],
                text_boxes=[
                    {"text": "PAGE242", "pdf_page_number": 1, "bbox": [10, 10, 60, 25]},
                    {"text": "PAGE242", "pdf_page_number": 3, "bbox": [10, 10, 60, 25]},
                ],
                stdout="full",
            )

            annotation = payload["annotations"][0]
            self.assertEqual(0, annotation["pdf_page_number"])
            self.assertEqual("unmatched", annotation["confidence"])
            self.assertTrue(any("命中多个 PDF 页" in warning for warning in payload["warnings"]))

    def test_pdf_text_match_does_not_fake_project_page_when_component_missing(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=2)
            bundle = {"components": {}, "nets": {}}

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "refdes", "refdes": "U404"}],
                text_boxes=[{
                    "text": "U404",
                    "pdf_page_number": 2,
                    "bbox": [100, 120, 140, 140],
                    "source": "test.refdes_text",
                }],
                stdout="full",
            )

            annotation = payload["annotations"][0]
            self.assertEqual("", annotation["project_page"])
            self.assertEqual(2, annotation["pdf_page_number"])
            self.assertEqual("pdf_text_match", annotation["confidence"])

    def test_pdf_page_map_with_mismatched_pdf_hash_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf_path = Path(temp) / "schematic.pdf"
            write_minimal_pdf(pdf_path, page_count=2)
            bundle = {"components": {}, "nets": {}}

            payload = build_schematic_pdf_annotation_payload(
                pdf_path,
                bundle,
                [{"kind": "page", "page": "PAGE1"}],
                pdf_page_map={
                    "pdf_sha256": "0" * 64,
                    "pages": {"PAGE1": 1},
                },
                text_boxes=[],
                stdout="full",
            )

            annotation = payload["annotations"][0]
            self.assertEqual(0, annotation["pdf_page_number"])
            self.assertEqual("unmatched", annotation["confidence"])
            self.assertFalse(payload["inputs"]["pdf_page_map_meta"]["sha256_match"])
            self.assertTrue(any("sha256" in warning for warning in payload["warnings"]))


if __name__ == "__main__":
    unittest.main()
