import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class V29UiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = (ROOT / "index.html").read_text(encoding="utf-8")
        cls.app = (ROOT / "app.js").read_text(encoding="utf-8")
        cls.styles = (ROOT / "styles.css").read_text(encoding="utf-8")

    def test_home_exposes_primary_inventory_actions(self):
        for text in ("Continuar inventario", "Iniciar nuevo inventario", "Consultar inventarios anteriores"):
            self.assertIn(text, self.index)
        self.assertNotIn("<summary>Administración</summary>", self.index)
        self.assertIn('id="newInventoryPanel" class="new-inventory-panel hidden"', self.index)
        self.assertIn('id="cancelNewInventory"', self.index)
        self.assertRegex(self.index, r'id="restoreButton"[^>]+disabled')

    def test_finish_is_a_real_view_and_not_a_dialog(self):
        self.assertIn('id="finishView"', self.index)
        self.assertNotIn('id="finishDialog"', self.index)
        self.assertIn('id="closeInventoryButton"', self.index)
        self.assertIn("showFinishView", self.app)

    def test_closed_inventory_is_read_only(self):
        self.assertIn("state.readOnly", self.app)
        self.assertIn("El inventario histórico es de solo lectura", self.app)
        self.assertIn("ui.finishButton.disabled=readOnly", self.app)

    def test_lifecycle_module_is_cached_and_loaded(self):
        worker = (ROOT / "sw.js").read_text(encoding="utf-8")
        self.assertIn("inventory-lifecycle.js?v=53", self.index)
        self.assertIn("inventory-lifecycle.js?v=53", worker)

    def test_v53_material_card_labels_and_layout(self):
        for text in ("Mantener cantidad y ubicación", "Mantener ubicación", "Cantidad SAP", "Ubicación SAP"):
            self.assertIn(text, self.app)
        for old_text in ("Mantener cantidad y ubicación SAP", "Mantener solo ubicación", "SAP qty", "SAP location"):
            self.assertNotIn(old_text, self.app)
        for class_name in ("material-um", "material-status", "material-unit-value", "loc-tab-code", "loc-tab-suffix"):
            self.assertIn(class_name, self.app)
        self.assertRegex(self.styles, r'\.loc-tab,\.loc-tab\.batch-tab\{[^}]*flex:0 0 105px[^}]*height:48px')
        self.assertRegex(self.styles, r'\.loc-tab-suffix\{[^}]*font-size:70%[^}]*font-weight:400')
        self.assertRegex(self.styles, r'\.material-meta\{[^}]*display:flex[^}]*flex-wrap:wrap')

    def test_ios_status_bar_uses_the_header_surface(self):
        self.assertIn('viewport-fit=cover', self.index)
        self.assertNotIn('apple-mobile-web-app-status-bar-style', self.index)
        self.assertIn('padding:calc(9px + env(safe-area-inset-top))', self.styles)
        self.assertIn('html{background:#063b31}', self.styles)
        self.assertRegex(self.styles, r'\.app-header\{[^}]*background:#063b31')
        self.assertNotIn('body::before', self.styles)

    def test_home_downloads_clean_empty_template_directly(self):
        self.assertIn('id="emptyTemplateButton"', self.index)
        self.assertIn('href="./Planilla_vacía.xlsx"', self.index)
        self.assertIn('download="Planilla_vacía.xlsx"', self.index)
        self.assertIn('>Exportar planilla vacía</a>', self.index)

    def test_closed_inventory_deletion_requires_explicit_confirmation(self):
        self.assertIn('id="deleteInventoryDialog"', self.index)
        self.assertIn('id="deleteInventoryName"', self.index)
        self.assertRegex(self.index, r'id="confirmDeleteInventory"[^>]+disabled')
        self.assertIn("deleteClosedInventory", self.app)
        self.assertIn("await showHistory()", self.app)

    def test_cancel_does_not_delete_and_empty_history_is_supported(self):
        dialog = self.index.split('id="deleteInventoryDialog"', 1)[1].split('</dialog>', 1)[0]
        self.assertIn('value="cancel"', dialog)
        self.assertNotIn('deleteClosedInventory', dialog)
        self.assertIn("historyEmpty.classList.toggle('hidden',visible.length>0)", self.app)

    def test_delete_action_is_limited_to_closed_history_items(self):
        history = self.app.split('async function showHistory()', 1)[1].split('function openDeleteInventoryDialog', 1)[0]
        self.assertIn("if(closed&&state.historyAdminMode)", history)
        self.assertIn("deleteButton.textContent='Eliminar inventario'", history)

    def test_history_admin_mode_is_temporary(self):
        self.assertIn('id="manageHistory"', self.index)
        self.assertIn("state.historyAdminMode?'Terminar':'Administrar'", self.app)
        self.assertIn("state.historyAdminMode=!state.historyAdminMode", self.app)
        self.assertIn("$('#closeHistory').addEventListener('click',()=>{state.historyAdminMode=false;showHome();})", self.app)

    def test_delete_action_has_iphone_layout(self):
        self.assertIn('@media(max-width:430px)', self.styles)
        self.assertRegex(self.styles, r'@media\(max-width:430px\).*?\.history-item\{grid-template-columns:1fr\}.*?\.history-delete-button')

    def test_status_percentage_has_two_decimals(self):
        self.assertIn("minimumFractionDigits:2,maximumFractionDigits:2", self.app)

    def test_demo_exit_is_global_and_isolated(self):
        self.assertIn('id="demoExitButton"', self.index)
        self.assertIn('Salir de demo', self.index)
        self.assertIn("ui.demoExitButton.addEventListener('click',showHome)", self.app)

    def test_update_preview_lists_real_details_and_no_change_message(self):
        self.assertIn('id="updatePreviewDetails"', self.index)
        for text in ('detail.material', 'detail.label', 'detail.before', 'detail.after', 'No se detectaron cambios'):
            self.assertIn(text, self.app)

    def test_history_administration_offers_neutral_export_and_red_delete(self):
        self.assertIn("exportButton.textContent='Exportar Excel'", self.app)
        self.assertIn("deleteButton.textContent='Eliminar inventario'", self.app)
        self.assertIn('.history-export-button', self.styles)
        self.assertIn('color:var(--danger)', self.styles)

    def test_toast_moves_inside_open_dialog(self):
        self.assertIn("const modal=$('dialog[open]')", self.app)
        self.assertIn("(modal||document.body).appendChild(ui.toast)", self.app)


if __name__ == "__main__":
    unittest.main()
