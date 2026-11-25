load('inventory-model.js');

let passed = 0;
function assert(condition, message) { if (!condition) throw new Error(message); }
function equal(actual, expected, message) { assert(actual === expected, `${message}: se esperaba ${expected}, se obtuvo ${actual}`); }
function test(name, callback) { callback(); passed += 1; print(`OK ${name}`); }

function record(overrides = {}) {
  return {
    rowNumber: 2, plant: '5810', material: 'MAT-001', stloc: '1140', sapQty: 1,
    serialized: true, serialControlRequired: true, serials: [], ...overrides,
  };
}

function series(overrides = {}) {
  return {
    rowNumber: 2, serialCol: 'E', plant: '5810', location: '1140', material: 'MAT-001',
    description: 'Prueba', serial: 'SER-001', available: true, selected: false, ...overrides,
  };
}

test('lista una sola vez materiales de Series ausentes de Inventario y conserva Plant', () => {
  const report = InventoryModel.seriesConsistencyReport([record()], [
    series({ rowNumber: 10, material: 'AUSENTE', serial: 'A-1' }),
    series({ rowNumber: 11, material: 'AUSENTE', serial: 'A-2' }),
    series({ rowNumber: 12, plant: '9999', material: 'AUSENTE', serial: 'A-3' }),
  ]);
  equal(report.seriesOnlyMaterials.length, 1, 'código único');
  equal(report.seriesOnlyMaterials[0].material, 'AUSENTE', 'material ausente');
  equal(report.seriesOnlyMaterials[0].plants.length, 2, 'Plants afectados');
});

test('detecta 29006499DEF con tres unidades SAP y dos registros de series', () => {
  const report = InventoryModel.seriesConsistencyReport([
    record({ material: '29006499DEF', stloc: '1140', sapQty: 1 }),
    record({ rowNumber: 3, material: '29006499DEF', stloc: '6025', sapQty: 1 }),
    record({ rowNumber: 4, material: '29006499DEF', stloc: '6066', sapQty: 1 }),
  ], [
    series({ material: '29006499DEF', location: '1140', serial: '0008NXMNB' }),
    series({ rowNumber: 3, material: '29006499DEF', location: '6066', serial: 'JAORI-20261018-1' }),
  ]);
  equal(report.quantityDifferences.length, 1, 'diferencia total');
  equal(report.quantityDifferences[0].sapQuantity, 3, 'cantidad SAP');
  equal(report.quantityDifferences[0].seriesCount, 2, 'registros de series');
  equal(report.quantityDifferences[0].difference, -1, 'falta una serie');
});

test('detecta más registros de series que cantidad SAP', () => {
  const report = InventoryModel.seriesConsistencyReport([record({ sapQty: 1 })], [series(), series({ rowNumber: 3, serial: 'SER-002' })]);
  equal(report.quantityDifferences[0].difference, 1, 'sobra una serie');
});

test('detecta distribución incorrecta aunque los totales coincidan', () => {
  const report = InventoryModel.seriesConsistencyReport([
    record({ stloc: '1140', sapQty: 5 }), record({ rowNumber: 3, stloc: '6066', sapQty: 5 }),
  ], [
    ...Array.from({ length: 6 }, (_, i) => series({ rowNumber: 10 + i, location: '1140', serial: `A-${i}` })),
    ...Array.from({ length: 4 }, (_, i) => series({ rowNumber: 20 + i, location: '6066', serial: `B-${i}` })),
  ]);
  equal(report.quantityDifferences.length, 0, 'totales coincidentes');
  equal(report.locationDifferences.length, 1, 'diferencia por ubicación');
  equal(report.locationDifferences[0].locations.length, 2, 'ubicaciones afectadas');
});

test('cuenta cada fila DEF aunque repita un número de serie', () => {
  const report = InventoryModel.seriesConsistencyReport([record({ material: 'MAT-DEF', sapQty: 2 })], [
    series({ material: 'MAT-DEF', serial: 'TEMP' }), series({ rowNumber: 3, material: 'MAT-DEF', serial: 'TEMP' }),
  ]);
  equal(report.quantityDifferences.length, 0, 'dos filas DEF contadas');
});

test('cuenta independientemente varias series 0 de un material DEF', () => {
  const report = InventoryModel.seriesConsistencyReport([record({ material: 'MAT-DEF', sapQty: 2 })], [
    series({ material: 'MAT-DEF', serial: '0' }), series({ rowNumber: 3, material: 'MAT-DEF', serial: '0' }),
  ]);
  equal(report.discrepancyCount, 0, 'dos unidades 0 coherentes');
});

test('omite materiales no serializados y aprueba materiales correctos', () => {
  const report = InventoryModel.seriesConsistencyReport([
    record({ material: 'CORRECTO', sapQty: 1 }),
    record({ rowNumber: 3, material: 'SIN-CONTROL', sapQty: 9, serialized: false, serialControlRequired: false }),
  ], [series({ material: 'CORRECTO' })]);
  equal(report.discrepancyCount, 0, 'planilla consistente');
});

test('una nueva importación corregida vuelve a ejecutar la validación', () => {
  const records = [record({ sapQty: 2 })];
  const first = InventoryModel.seriesConsistencyReport(records, [series()]);
  const corrected = InventoryModel.seriesConsistencyReport(records, [series(), series({ rowNumber: 3, serial: 'SER-002' })]);
  equal(first.quantityDifferences.length, 1, 'primera importación');
  equal(corrected.discrepancyCount, 0, 'archivo corregido');
});

test('la validación no modifica el conteo ni impide seleccionar series', () => {
  const target = record({ sapQty: 1, count: '' });
  const items = [series()];
  InventoryModel.seriesConsistencyReport([target], items);
  InventoryModel.assignStableRecordKeys([target]);
  InventoryModel.assignSeriesOwnership([target], items);
  const selected = InventoryModel.toggleSeriesSelection([target], {}, target, InventoryModel.seriesSelectionKey(target.serials[0]));
  equal(selected[target.key].length, 1, 'conteo normal disponible');
  equal(target.count, '', 'conteo original intacto');
});

print(`${passed} pruebas aprobadas`);
