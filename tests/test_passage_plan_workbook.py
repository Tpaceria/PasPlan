import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


REPO = Path("/home/runner/work/PasPlan/PasPlan")
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
        self.assertIn("Sub Auto_Open()", source)
        self.assertIn("dataCount = CountDataRowsByColumnB(wsD)", source)
        self.assertIn("pagesNeeded = (dataCount + DATA_ROWS - 1) \\ DATA_ROWS", source)
        self.assertIn("If dataCount = 0 Then", source)
        self.assertIn('wsP.Range("X" & destRow).Formula', source)
        self.assertIn('wsP.Rows((TEMPLATE_LAST_ROW + 1) & ":" & oldLastRow).Delete', source)
        self.assertNotIn("pagesNeeded = 5", source)

    def test_data_edits_always_trigger_rebuild_across_editable_range(self):
        source = (REPO / "excel_sources/Лист2_cls.txt").read_text(encoding="utf-8")
        self.assertIn('Me.Range("A2:G600")', source)
        self.assertIn("UpdatePassagePlan True", source)


if __name__ == "__main__":
    unittest.main()
