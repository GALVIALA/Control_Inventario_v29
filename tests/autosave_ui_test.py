import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class AutosaveUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = (ROOT / "app.js").read_text(encoding="utf-8")
        cls.index = (ROOT / "index.html").read_text(encoding="utf-8")

    def test_initial_read_error_is_visible_and_retryable(self):
        check = re.search(r"async function checkSavedSession\(\).*?\n  }", self.app, re.S).group(0)
        self.assertIn("catch(error)", check)
        self.assertIn("recoveryLookupError.classList.remove('hidden')", check)
        self.assertIn("retryRecoveryLookup.disabled=false", check)
        self.assertIn("ui.retryRecoveryLookup.addEventListener('click',checkSavedSession)", self.app)
        self.assertIn('id="retryRecoveryLookup"', self.index)

    def test_failed_save_cannot_assign_a_new_timestamp(self):
        persist = re.search(r"async function persist\(bytes\).*?\n  }", self.app, re.S).group(0)
        self.assertIn("state.lastSavedAt=await Storage.confirmSave", persist)
        self.assertNotRegex(persist, r"state\.lastSavedAt\s*=\s*Date\.now")

    def test_pending_and_error_states_have_distinct_labels(self):
        self.assertIn("status==='saving'?'Guardando…'", self.app)
        self.assertIn("status==='error'?'Error al guardar'", self.app)
        self.assertIn("status==='saved'?'Guardado'", self.app)


if __name__ == "__main__":
    unittest.main()
