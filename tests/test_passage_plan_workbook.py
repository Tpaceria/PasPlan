import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
import unittest


REPO = Path(__file__).resolve().parents[1]
WORKBOOKS = [
    REPO / "NEW_FINAL_LOGO_EXACT_5_05x1_00cm.xlsm",
    REPO / "PasPlan_NORDLAND.xlsm",
]
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def find(root, path):
    return root.find(path, NS)


class PassagePlanWorkbookTests(unittest.TestCase):
    def test_initial_state_is_single_plan_page(self):
        for workbook in WORKBOOKS:
            with self.subTest(workbook=workbook.name), zipfile.ZipFile(workbook) as archive:
                workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
                print_area = [
                    node.text
                    for node in workbook_xml.findall("x:definedNames/x:definedName", NS)
                    if node.attrib.get("name") == "_xlnm.Print_Area"
                ]
                self.assertEqual(print_area, ["'Plan list'!$A$1:$BN$46"])

                plan_xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                self.assertEqual(find(plan_xml, "x:dimension").attrib["ref"], "A1:BN46")
                self.assertIsNone(find(plan_xml, "x:rowBreaks"))

                rows = [int(row.attrib["r"]) for row in find(plan_xml, "x:sheetData").findall("x:row", NS)]
                self.assertEqual(max(rows), 46)

    def test_initial_data_template_is_empty_but_keeps_helper_formulas(self):
        for workbook in WORKBOOKS:
            with self.subTest(workbook=workbook.name), zipfile.ZipFile(workbook) as archive:
                data_xml = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
                self.assertEqual(find(data_xml, "x:dimension").attrib["ref"], "A1:G600")

                expected_rows = set(range(2, 601))
                a_rows = set()
                f_rows = set()

                for row in find(data_xml, "x:sheetData").findall("x:row", NS):
                    row_number = int(row.attrib["r"])
                    if row_number not in expected_rows:
                        continue

                    refs = [cell.attrib["r"] for cell in row.findall("x:c", NS)]
                    self.assertEqual(refs, [f"A{row_number}", f"F{row_number}"])

                    formulas = [cell.find("x:f", NS).text for cell in row.findall("x:c", NS)]
                    self.assertEqual(
                        formulas,
                        [
                            f'IF(B{row_number}="","",CHAR(160))',
                            f'IF(D{row_number}="","",SUM($D$2:D{row_number}))',
                        ],
                    )

                    a_rows.add(row_number)
                    f_rows.add(row_number)

                self.assertEqual(a_rows, expected_rows)
                self.assertEqual(f_rows, expected_rows)

    def test_vba_exposes_rebuild_entrypoints_without_fixed_five_page_bootstrap(self):
        source = (REPO / "excel_sources/PassagePlanModule_bas.txt").read_text(encoding="utf-8")
        self.assertRegex(source, r"\bSub\s+Auto_Open\(\)")
        self.assertRegex(source, r"\bPublic Sub\s+RouteDataChanged\b")
        self.assertRegex(source, r"\bPublic Sub\s+UpdatePassagePlan\b")
        self.assertIn("ROUTE COLUMN B", source)
        self.assertIn("DATA G -> Plan list X", source)
        self.assertNotIn("pagesNeeded = 5", source)

    def test_data_edits_always_trigger_rebuild_across_editable_range(self):
        sheet_source = (REPO / "excel_sources/Лист2_cls.txt").read_text(encoding="utf-8")
        module_source = (REPO / "excel_sources/PassagePlanModule_bas.txt").read_text(encoding="utf-8")
        handler = re.search(
            r"Private Sub Worksheet_Change\(ByVal Target As Range\)\s*(.*?)\s*End Sub",
            sheet_source,
            re.S,
        )
        self.assertIsNotNone(handler)
        self.assertEqual(handler.group(1).strip(), "RouteDataChanged Target")

        delegate = re.search(
            r"Public Sub RouteDataChanged\(ByVal Target As Range\)\s*(.*?)\s*End Sub",
            module_source,
            re.S,
        )
        self.assertIsNotNone(delegate)
        delegate_body = delegate.group(1)
        self.assertRegex(delegate_body, r"Target\.CountLarge\s*>\s*500")
        self.assertIn("A2:G600", delegate_body)
        self.assertRegex(delegate_body, r"\bUpdatePassagePlan\s+True\b")


if __name__ == "__main__":
    unittest.main()
