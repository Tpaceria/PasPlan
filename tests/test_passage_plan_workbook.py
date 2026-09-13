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

                for row_number in (2, 3, 35):
                    row = find(data_xml, f"x:sheetData/x:row[@r='{row_number}']")
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

    def test_vba_uses_dynamic_column_b_page_count_without_fixed_minimum(self):
        source = (REPO / "excel_sources/PassagePlanModule_bas.txt").read_text(encoding="utf-8")
        self.assertRegex(source, r"\bSub\s+Auto_Open\(\)")
        self.assertRegex(source, r"\bCountDataRowsByColumnB\b")
        self.assertRegex(source, r"pagesNeeded\s*=\s*\(dataCount\s*\+\s*DATA_ROWS\s*-\s*1\)\s*\\\s*DATA_ROWS")
        self.assertRegex(source, r"If\s+dataCount\s*=\s*0\s+Then\s+[\s\S]*?pagesNeeded\s*=\s*1")
        self.assertRegex(source, r'Range\("X"\s*&\s*destRow\)\.Formula')
        self.assertRegex(source, r'Rows\(\(TEMPLATE_LAST_ROW\s*\+\s*1\)\s*&\s*":"\s*&\s*oldLastRow\)\.Delete')
        self.assertNotIn("pagesNeeded = 5", source)

    def test_data_edits_always_trigger_rebuild_across_editable_range(self):
        sheet_source = (REPO / "excel_sources/Лист2_cls.txt").read_text(encoding="utf-8")
        module_source = (REPO / "excel_sources/PassagePlanModule_bas.txt").read_text(encoding="utf-8")
        self.assertRegex(sheet_source, r"\bRouteDataChanged\s+Target\b")
        self.assertRegex(module_source, r"Target\.CountLarge\s*>\s*500")
        self.assertRegex(module_source, r'Target\.Worksheet\.Range\("A2:G600"\)')
        self.assertRegex(module_source, r"\bUpdatePassagePlan\s+True\b")


if __name__ == "__main__":
    unittest.main()
