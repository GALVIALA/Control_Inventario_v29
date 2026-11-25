import hashlib
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
ORIGINAL = ROOT / "planilla_vacia.xlsx"
TEMPLATE = ROOT / "Planilla_vacía.xlsx"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class EmptyTemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.archive = zipfile.ZipFile(TEMPLATE)
        cls.names = set(cls.archive.namelist())
        cls.workbook_xml = cls.archive.read("xl/workbook.xml")
        cls.workbook = ET.fromstring(cls.workbook_xml)

    @classmethod
    def tearDownClass(cls):
        cls.archive.close()

    def test_original_file_remains_unchanged(self):
        digest = hashlib.sha256(ORIGINAL.read_bytes()).hexdigest()
        self.assertEqual("0db9cb1d7948110b770cf6f2bd5d906c430d7ddf609f8c463bc690ddc5d61795", digest)

    def test_clean_copy_retains_requested_sheets_in_order(self):
        sheets = self.workbook.find(f"{{{MAIN_NS}}}sheets")
        self.assertEqual(["Inventario", "Series", "MB52"], [sheet.attrib["name"] for sheet in sheets])

    def test_real_inventory_residue_and_external_features_are_absent(self):
        forbidden_parts = (
            "comments", "externalLinks", "connections", "vbaProject", "calcChain",
            "sheet4.xml", "sheet5.xml", "sheet6.xml", "table4.xml", "table5.xml",
            "vmlDrawing", "LabelInfo.xml",
        )
        for part in self.names:
            self.assertFalse(any(token in part for token in forbidden_parts), part)
        xml_text = "\n".join(
            self.archive.read(name).decode("utf-8", errors="ignore")
            for name in self.names
            if name.endswith((".xml", ".rels"))
        )
        for residue in (
            "/Users/", "TIE,CABLE,205 mm", "WASHER,FLAT,ISO 7089", "#REF!",
            "Planilla Oficial", "VL10G", "CIM 4529",
        ):
            self.assertNotIn(residue, xml_text)

    def test_native_excel_structure_and_formulas_are_preserved(self):
        inventory_sheet = self.archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        series_sheet = self.archive.read("xl/worksheets/sheet2.xml").decode("utf-8")
        mb52_sheet = self.archive.read("xl/worksheets/sheet3.xml").decode("utf-8")
        inventory_table = self.archive.read("xl/tables/table1.xml").decode("utf-8")
        mb52_table = self.archive.read("xl/tables/table3.xml").decode("utf-8")
        self.assertIn("<cols>", inventory_sheet)
        self.assertIn("<conditionalFormatting", inventory_sheet)
        self.assertIn("<tableParts", inventory_sheet)
        self.assertIn("<tableParts", series_sheet)
        self.assertIn("<tableParts", mb52_sheet)
        self.assertIn("<autoFilter", inventory_table)
        self.assertIn("<autoFilter", mb52_table)
        self.assertIn("VLOOKUP(Table1[[#This Row],[MATERIAL]],TABLA1[],9,FALSE)", inventory_table)
        self.assertIn("Value Unrestricted", mb52_table)
        self.assertIn("<calculatedColumnFormula>", mb52_table)

    def test_data_rows_are_empty_but_formula_cells_remain(self):
        inventory_sheet = self.archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        series_sheet = self.archive.read("xl/worksheets/sheet2.xml").decode("utf-8")
        mb52_sheet = self.archive.read("xl/worksheets/sheet3.xml").decode("utf-8")
        self.assertIn('<c r="O2" s="6"><f>', inventory_sheet)
        self.assertIn('<c r="I2" s="20"><f>', mb52_sheet)
        for sheet in (inventory_sheet, series_sheet, mb52_sheet):
            row = re.search(r'<row\b[^>]*\br="2"[^>]*>(.*?)</row>', sheet, re.DOTALL).group(1)
            values_without_formula_cache = re.sub(r'<c\b[^>]*>\s*<f>.*?</f>\s*</c>', "", row, flags=re.DOTALL)
            self.assertNotIn("<v>", values_without_formula_cache)

    def test_metadata_is_generic(self):
        core = self.archive.read("docProps/core.xml").decode("utf-8")
        self.assertIn("Control Inventario V29", core)
        self.assertNotIn("/Users/", core)


if __name__ == "__main__":
    unittest.main()
