load('inventory-model.js');
load('inventory-lifecycle.js');

let passed = 0;
function assert(condition, message) { if (!condition) throw new Error(message); }
function equal(actual, expected, message) { assert(actual === expected, `${message}: ${actual} !== ${expected}`); }
function test(name, callback) { callback(); passed += 1; print(`OK ${name}`); }

function record(overrides = {}) {
  return { rowNumber: 2, material: 'MAT-001', stloc: '1140', batch: '', stockType: 'Libre', barcode: '', sapQty: '2', sapLocations: 'A1', serialized: true, serials: [], alternateSerials: [], count: '1', currentLocation: 'A1', comments: 'ANTES', ...overrides };
}
function workbook(records) { InventoryModel.assignStableRecordKeys(records); return { fileName: 'inventario.xlsx', bytes: new Uint8Array([1]), records, invalidPhysicalCountRows: [] }; }

test('crea inventarios independientes con nombre permanente', () => {
  const wb = workbook([record()]);
  const first = InventoryLifecycle.createInventory({ name: 'End Year 2026', workbook: wb, sourceFingerprint: 'source-a', now: 10 });
  const second = InventoryLifecycle.createInventory({ name: 'Mid Year 2027', workbook: wb, sourceFingerprint: 'source-a', now: 20 });
  equal(first.name, 'End Year 2026', 'nombre');
  assert(first.id !== second.id, 'identificadores independientes');
  equal(first.status, 'open', 'estado inicial');
});

test('actualiza cantidad, borra comentario y cambia series conocidas', () => {
  const oldRecord = record({ serials: [{ serial: 'S1', selectionKey: 'S1' }, { serial: 'S2', selectionKey: 'S2' }] });
  const newRecord = record({ count: '3', comments: '', serials: [{ serial: 'S1', selectionKey: 'S1', styleState: 'unselected' }, { serial: 'S2', selectionKey: 'S2', styleState: 'selected' }] });
  const current = workbook([oldRecord]), candidate = workbook([newRecord]);
  const inventory = { id: 'i1', status: 'open', sourceFingerprint: 'same', changes: {}, selectedSeries: { [oldRecord.key]: ['S1'] }, bytes: new Uint8Array([1]) };
  const plan = InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: current, candidateWorkbook: candidate, candidateFingerprint: 'same', model: InventoryModel });
  equal(plan.counts.quantities, 1, 'cantidad cambiada');
  equal(plan.counts.comments, 1, 'comentario eliminado');
  equal(plan.changes[oldRecord.key].comments, '', 'vacío intencional');
  equal(plan.selectedSeries[oldRecord.key][0], 'S2', 'serie actualizada');
});

test('un estilo de serie desconocido conserva el estado anterior', () => {
  const oldRecord = record({ serials: [{ serial: 'S1', selectionKey: 'S1' }] });
  const newRecord = record({ serials: [{ serial: 'S1', selectionKey: 'S1', styleState: 'unknown', originalStyle: 99 }] });
  const current = workbook([oldRecord]), candidate = workbook([newRecord]);
  const inventory = { id: 'i1', status: 'open', sourceFingerprint: 'same', changes: {}, selectedSeries: { [oldRecord.key]: ['S1'] } };
  const plan = InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: current, candidateWorkbook: candidate, candidateFingerprint: 'same', model: InventoryModel });
  equal(plan.selectedSeries[oldRecord.key][0], 'S1', 'estado preservado');
  equal(plan.unknownSeriesStyles.length, 1, 'advertencia registrada');
});

test('una serie alternativa seleccionada no se duplica como cambio', () => {
  const first = record({ material: 'MAT-ALT', stloc: 'A1' });
  const second = record({ rowNumber: 3, material: 'MAT-ALT', stloc: 'B1' });
  const current = workbook([first, second]);
  const selected = { serial: 'S1', selectionKey: 'S1', styleState: 'selected', defaultRecordKey: first.key };
  current.records[0].serials = [selected];
  current.records[1].alternateSerials = [selected];
  const candidate = workbook([record({ material: 'MAT-ALT', stloc: 'A1' }), record({ rowNumber: 3, material: 'MAT-ALT', stloc: 'B1' })]);
  const candidateSelected = { ...selected, defaultRecordKey: candidate.records[0].key };
  candidate.records[0].serials = [candidateSelected];
  candidate.records[1].alternateSerials = [candidateSelected];
  const inventory = { id: 'i1', status: 'open', sourceFingerprint: 'same', changes: {}, selectedSeries: { [first.key]: ['S1'] } };
  const plan = InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: current, candidateWorkbook: candidate, candidateFingerprint: 'same', model: InventoryModel });
  equal(plan.counts.series, 0, 'sin falso cambio de series');
  equal(plan.details.length, 0, 'sin diferencias reales');
});

test('genera revisiones independientes y recupera metadatos cerrados válidos', () => {
  const inventory = { id: 'i1', name: 'Inventario Mid Year 2026', status: 'open', exportRevision: 2 };
  const metadata = InventoryLifecycle.nextExportMetadata(inventory, new Date(2026, 10, 24, 12).getTime());
  equal(metadata.fileName, 'Inventario_Mid_Year_2026_REV03_2026-11-24.xlsx', 'nombre de revisión');
  const closed = InventoryLifecycle.closeInventory(inventory, { backupFileName: metadata.fileName, exportRevision: metadata.revision, exportDate: metadata.date, finalBackupBytes: new Uint8Array([1]), now: 30 });
  equal(closed.finalExportRevision, 3, 'revisión congelada');
  equal(InventoryLifecycle.closedExportMetadata(closed).fileName, metadata.fileName, 'nombre congelado recuperable');
  equal(InventoryLifecycle.closedExportMetadata({ status: 'closed', finalBackupFileName: 'legado.xlsx' }), null, 'no inventa datos históricos');
});

test('rechaza un Excel de otro inventario', () => {
  const current = workbook([record()]), candidate = workbook([record()]), inventory = { id: 'i1', status: 'open', sourceFingerprint: 'correcta', changes: {}, selectedSeries: {} };
  let error = null; try { InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: current, candidateWorkbook: candidate, candidateFingerprint: 'otra', model: InventoryModel }); } catch (caught) { error = caught; }
  assert(error && /no corresponde/.test(error.message), 'archivo incorrecto bloqueado');
});

test('rechaza registros ausentes y duplicados', () => {
  const current = workbook([record(), record({ rowNumber: 3, material: 'MAT-002' })]);
  const missing = workbook([record()]);
  const inventory = { id: 'i1', status: 'open', sourceFingerprint: 'same', changes: {}, selectedSeries: {} };
  let error = null; try { InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: current, candidateWorkbook: missing, candidateFingerprint: 'same', model: InventoryModel }); } catch (caught) { error = caught; }
  assert(error && /ausente/.test(error.message), 'registro ausente bloqueado');
  const duplicates = workbook([record(), record({ rowNumber: 3 })]);
  error = null; try { InventoryLifecycle.buildExcelUpdatePlan({ inventory, currentWorkbook: duplicates, candidateWorkbook: duplicates, candidateFingerprint: 'same', model: InventoryModel }); } catch (caught) { error = caught; }
  assert(error && /duplicados/.test(error.message), 'duplicados bloqueados');
});

test('aplicar una actualización no muta el estado anterior', () => {
  const inventory = { id: 'i1', status: 'open', fileName: 'old.xlsx', bytes: new Uint8Array([1]), changes: { a: { count: '1', currentLocation: 'A1' } }, selectedSeries: {} };
  const plan = { inventoryId: 'i1', candidateFileName: 'new.xlsx', candidateBytes: new Uint8Array([2]), changes: { a: { count: '2' } }, selectedSeries: {} };
  const updated = InventoryLifecycle.applyExcelUpdate(inventory, plan, 50);
  equal(updated.changes.a.count, '2', 'nuevo conteo');
  equal(updated.changes.a.currentLocation, 'A1', 'campo no actualizado preservado');
  equal(inventory.changes.a.count, '1', 'estado anterior intacto');
  const backup = InventoryLifecycle.createBackup(inventory, 'prueba', 40);
  equal(backup.snapshot.changes.a.count, '1', 'respaldo recuperable');
});

test('el cierre exige respaldo y bloquea ediciones futuras', () => {
  const inventory = { id: 'i1', name: 'End Year 2026', status: 'open', changes: {}, selectedSeries: {} };
  let error = null; try { InventoryLifecycle.closeInventory(inventory); } catch (caught) { error = caught; }
  assert(error && /respaldo/.test(error.message), 'respaldo obligatorio');
  const closed = InventoryLifecycle.closeInventory(inventory, { backupFileName: 'final_REV01_2026-11-24.xlsx', exportRevision: 1, exportDate: '2026-11-24', now: 30 });
  equal(closed.status, 'closed', 'cerrado');
  try { InventoryLifecycle.assertOpen(closed); } catch (caught) { error = caught; }
  assert(error && /solo puede consultarse/.test(error.message), 'solo lectura');
});

print(`${passed} pruebas aprobadas`);
