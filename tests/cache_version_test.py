import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class CacheVersionTests(unittest.TestCase):
    def test_visual_version_is_v29(self):
        index = (ROOT / "index.html").read_text(encoding="utf-8")
        self.assertIn("Control Inventario · V29", index)
        self.assertNotIn("Inventario · By ", index)

    def test_first_party_assets_share_one_cache_buster(self):
        index = (ROOT / "index.html").read_text(encoding="utf-8")
        versions = re.findall(
            r'(?:styles\.css|manifest\.webmanifest|inventory-(?:model|storage|xlsx|lifecycle)\.js|app\.js|sw\.js)\?v=(\d+)',
            index,
        )
        self.assertTrue(versions)
        self.assertEqual({"51"}, set(versions))

    def test_service_worker_uses_new_consistent_cache(self):
        service_worker = (ROOT / "sw.js").read_text(encoding="utf-8")
        self.assertIn("const CACHE='inventario-offline-v51'", service_worker)
        for asset in (
            "styles.css",
            "manifest.webmanifest",
            "inventory-model.js",
            "inventory-storage.js",
            "inventory-xlsx.js",
            "inventory-lifecycle.js",
            "app.js",
        ):
            self.assertIn(f"{asset}?v=51", service_worker)

    def test_empty_template_is_cached_for_offline_download(self):
        service_worker = (ROOT / "sw.js").read_text(encoding="utf-8")
        self.assertIn("./Planilla_vacía.xlsx", service_worker)


if __name__ == "__main__":
    unittest.main()
