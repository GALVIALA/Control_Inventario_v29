import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class ConsistencyPresentationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = (ROOT / "app.js").read_text(encoding="utf-8")
        cls.index = (ROOT / "index.html").read_text(encoding="utf-8")

    def test_categories_are_collapsed_native_details(self):
        markup = re.search(r"function consistencyReportMarkup.*?\n  }", self.app, re.S).group(0)
        self.assertEqual(3, markup.count('<details class="consistency-group">'))
        self.assertNotIn('<details open', markup)
        self.assertIn("if(data.seriesOnlyMaterials.length)", markup)
        self.assertIn("if(data.quantityDifferences.length)", markup)
        self.assertIn("if(data.locationDifferences.length)", markup)

    def test_each_header_reports_affected_materials(self):
        self.assertGreaterEqual(self.app.count("material${"), 3)

    def test_shared_text_contains_every_category_and_detail(self):
        report = re.search(r"function consistencyReportText.*?\n  }", self.app, re.S).group(0)
        for title in (
            "Materiales presentes en Series y ausentes de Inventario",
            "Diferencias entre cantidades SAP y registros de series",
            "Diferencias por St Loc",
        ):
            self.assertIn(title, report)
        self.assertIn("item.sapQuantity", report)
        self.assertIn("item.seriesCount", report)
        self.assertIn("location.sapQuantity", report)
        self.assertIn("location.seriesCount", report)

    def test_share_uses_native_api_with_copy_fallback(self):
        self.assertIn("navigator.share({title:'Discrepancias del inventario',text})", self.app)
        self.assertIn("await copyPlainText(text)", self.app)
        self.assertIn('id="shareConsistencyButton"', self.index)
        self.assertIn('class="consistency-share-button hidden"', self.index)

    def test_existing_dialog_actions_remain_dialog_buttons(self):
        self.assertRegex(self.index, r'id="continueConsistency"[^>]+value="continue"')
        self.assertRegex(self.index, r'id="replaceConsistencyFile"[^>]+value="replace"')

    def test_reimport_label_follows_consistency_result(self):
        self.assertIn("report.discrepancyCount?'Corregir y volver a importar':'Volver a importar'", self.app)


if __name__ == "__main__":
    unittest.main()
