import os
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from pstx_knowledge import dfmea_workbench


class DfmeaWorkbenchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_data_dir = os.environ.get(dfmea_workbench.DFMEA_DATA_DIR_ENV)
        os.environ[dfmea_workbench.DFMEA_DATA_DIR_ENV] = self.tmp.name
        self.report = {"project_name": "demo"}
        self.bundle = {
            "project_name": "demo",
            "all_components": {
                "U2": {
                    "refdes": "U2",
                    "page_submodule_mapped": "PAGE10",
                    "page_user_visible_pages": "PAGE10, PAGE11",
                    "page_logical_pages": "PAGE3, PAGE4",
                    "HQ_CODE": "HQU2",
                    "VALUE": "MCU",
                    "PACKAGE": "BGA",
                },
                "R1": {
                    "refdes": "R1",
                    "page_submodule_mapped": "PAGE2",
                    "VALUE": "10k",
                    "PACKAGE": "0402",
                },
                "C3": {
                    "refdes": "C3",
                    "page_submodule_mapped": "PAGE2",
                    "VALUE": "1uF",
                    "PACKAGE": "0402",
                    "BOM_OPTION": "DEPOP",
                },
            },
        }

    def tearDown(self):
        if self.old_data_dir is None:
            os.environ.pop(dfmea_workbench.DFMEA_DATA_DIR_ENV, None)
        else:
            os.environ[dfmea_workbench.DFMEA_DATA_DIR_ENV] = self.old_data_dir
        self.tmp.cleanup()

    def test_workbench_groups_pending_components_and_exports_excel(self):
        payload = dfmea_workbench.get_dfmea_workbench("run1", self.report, self.bundle)
        self.assertEqual(["R1", "U2"], [row["refdes"] for row in payload["pending_components"]])
        self.assertNotIn("C3", [row["refdes"] for row in payload["pending_components"]])

        with_depop = dfmea_workbench.get_dfmea_workbench("run1", self.report, self.bundle, include_depop=True)
        self.assertEqual(["R1", "C3", "U2"], [row["refdes"] for row in with_depop["pending_components"]])
        exclude_rc = dfmea_workbench.get_dfmea_workbench(
            "run1",
            self.report,
            self.bundle,
            include_depop=True,
            exclude_rc=True,
        )
        self.assertEqual(["U2"], [row["refdes"] for row in exclude_rc["pending_components"]])
        self.assertTrue(exclude_rc["exclude_rc"])

        created = dfmea_workbench.create_dfmea_group("run1", {
            "refdes": ["U2", "R1"],
            "function_requirement": "控制与分压",
            "failure_mode": "开路",
            "failure_effect": "功能异常",
            "failure_cause": "焊接异常",
            "prevention_detection": "ICT/FCT",
        })
        self.assertTrue(created["group_id"])

        after = dfmea_workbench.get_dfmea_workbench("run1", self.report, self.bundle, include_depop=True)
        self.assertEqual(["C3"], [row["refdes"] for row in after["pending_components"]])
        self.assertEqual(1, len(after["groups"]))
        self.assertEqual("R1, U2", after["groups"][0]["refdes_text"])
        self.assertEqual("PAGE2, PAGE10, PAGE11", after["groups"][0]["pages_text"])
        self.assertEqual(["R1", "U2"], [row["refdes"] for row in after["groups"][0]["components"]])
        u2 = next(row for row in after["groups"][0]["components"] if row["refdes"] == "U2")
        self.assertEqual("PAGE3, PAGE4", u2["summary"]["main_module_page"])

        data = dfmea_workbench.export_dfmea_workbook("run1")
        target = Path(self.tmp.name) / "dfmea.xlsx"
        target.write_bytes(data)
        workbook = load_workbook(target)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(dfmea_workbench.EXPORT_HEADERS, headers)
        self.assertEqual("R1, U2", sheet["B2"].value)
        self.assertEqual("PAGE2, PAGE10, PAGE11", sheet["C2"].value)
        self.assertEqual("控制与分压", sheet["D2"].value)

        deleted = dfmea_workbench.delete_dfmea_group("run1", created["group_id"])
        self.assertTrue(deleted["ok"])
        restored = dfmea_workbench.get_dfmea_workbench("run1", self.report, self.bundle)
        self.assertEqual(["R1", "U2"], [row["refdes"] for row in restored["pending_components"]])


if __name__ == "__main__":
    unittest.main()
