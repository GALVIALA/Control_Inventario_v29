load('inventory-model.js');

let passed = 0;

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function equal(actual, expected, message) {
  assert(actual === expected, `${message}: se esperaba ${expected}, se obtuvo ${actual}`);
}

function test(name, callback) {
  callback();
  passed += 1;
  print(`OK ${name}`);
}

function record(overrides = {}) {
  return {
    rowNumber: 2,
    plant: '5810',
    material: 'MAT-001',
    stloc: '1140',
    batch: '',
    stockType: '(Un)Restricted-use',
    barcode: '',
    sapQty: '1',
    sapLocations: '',
    serialized: true,
    serials: [],
    count: '',
    currentLocation: '',
    comments: '',
    ...overrides,
  };
}

function series(overrides = {}) {
  return {
    rowNumber: 2,
    serialCol: 'E',
    plant: '5810',
    location: '1140',
    material: 'MAT-001',
    description: 'Material de prueba',
    serial: 'SER-001',
    available: true,
    selected: false,
    ...overrides,
  };
}

function assign(records, items) {
  InventoryModel.assignStableRecordKeys(records);
  return InventoryModel.assignSeriesOwnership(records, items);
}

test('asocia solo una coincidencia exacta de Plant, Material y Location', () => {
  const exact = record();
  const otherPlant = record({ rowNumber: 3, plant: '9999' });
  const result = assign([exact, otherPlant], [series()]);
  equal(exact.serials.length, 1, 'serie exacta');
  equal(otherPlant.serials.length, 0, 'otra planta');
  equal(result.unresolved.length, 0, 'series pendientes');
});

test('permite seleccionar LM0012 conscientemente desde 6066 para el St Loc 6S21', () => {
  const item = record({ rowNumber: 1200, material: '29080784', stloc: '6S21', count: '1' });
  const result = assign([item], [series({ rowNumber: 7, material: '29080784', location: '6066', serial: 'LM0012', selected: true })]);
  equal(item.serials.length, 0, 'series exactas');
  equal(item.alternateSerials.length, 1, 'series disponibles de otra ubicación');
  equal(item.count, '1', 'conteo físico conservado');
  equal(result.unresolved.length, 1, 'serie incompatible pendiente');
  assert(result.issues.some(issue => issue.type === 'incompatible-location' && issue.severity === 'warning'), 'debe advertir ambas ubicaciones');
  const selected = InventoryModel.toggleSeriesSelection([item], {}, item, InventoryModel.seriesSelectionKey(item.alternateSerials[0]));
  equal(selected[item.key].length, 1, 'conteo seleccionable sin confirmación adicional');
});

test('elimina la heurística de fila restante para 106498DEF', () => {
  const rows = [
    record({ rowNumber: 1112, material: '106498DEF', stloc: '1140', sapQty: '3' }),
    record({ rowNumber: 1171, material: '106498DEF', stloc: '6025' }),
    record({ rowNumber: 1187, material: '106498DEF', stloc: '6066' }),
  ];
  const items = [
    series({ rowNumber: 145, material: '106498DEF', location: '1140', serial: 'EXACTA-1140' }),
    series({ rowNumber: 146, material: '106498DEF', location: '6066', serial: 'EXACTA-6066' }),
    ...['6049', '6052', '6050', '6053', '6054', '6053'].map((location, index) => series({ rowNumber: 200 + index, material: '106498DEF', location, serial: `OTRA-${index + 1}` })),
  ];
  const result = assign(rows, items);
  equal(rows[0].serials.length, 1, 'series 1140');
  equal(rows[1].serials.length, 0, 'series 6025');
  equal(rows[2].serials.length, 1, 'series 6066');
  equal(rows[1].alternateSerials.length, 8, 'series del mismo material visibles en 6025');
  equal(result.unresolved.length, 6, 'series de otras ubicaciones');
  equal(result.issues.filter(issue => issue.type === 'incompatible-location').length, 8, 'advertencias de ubicación');
});

test('detecta materiales serializados sin series por material o ubicación', () => {
  const rows = [
    record({ rowNumber: 1119, material: '29006499DEF', stloc: '1140' }),
    record({ rowNumber: 1172, material: '29006499DEF', stloc: '6025' }),
    record({ rowNumber: 1188, material: '29006499DEF', stloc: '6066' }),
    record({ rowNumber: 1193, material: '29060554', stloc: '6P22', count: '1' }),
  ];
  const result = assign(rows, [
    series({ rowNumber: 210, material: '29006499DEF', location: '1140', serial: 'YAW-1140' }),
    series({ rowNumber: 8, material: '29006499DEF', location: '6066', serial: 'YAW-6066' }),
  ]);
  equal(rows[1].serials.length, 0, 'series exactas de 29006499DEF en 6025');
  equal(rows[1].alternateSerials.length, 2, 'series de otras ubicaciones disponibles en 6025');
  equal(rows[3].serials.length, 0, '29060554 en 6P22');
  assert(rows[1].seriesIssues.some(issue => issue.type === 'incompatible-location'), 'falta advertencia para 29006499DEF');
  assert(rows[3].seriesIssues.some(issue => issue.type === 'serialized-without-series'), 'falta advertencia para 29060554');
  equal(rows[3].count, '1', 'conteo guardado de 29060554');
  equal(result.issues.filter(issue => issue.type === 'serialized-without-series').length, 1, 'materiales sin series disponibles');
});

test('mantiene fuera de la asociación las 25 series sin material de Inventario', () => {
  const valid = record({ material: 'VALIDO' });
  const items = Array.from({ length: 25 }, (_, index) => series({ rowNumber: 50 + index, material: `AUSENTE-${index}`, serial: `SIN-MATERIAL-${index}` }));
  const result = assign([valid], items);
  equal(valid.serials.length, 0, 'series del material válido');
  equal(result.unresolved.length, 25, 'series no asociadas');
  equal(result.issues.filter(issue => issue.type === 'series-without-material').length, 25, 'incidencias sin material');
});

test('duplica solo por Material y Serial Number', () => {
  const a1140 = record({ material: 'MAT-A', stloc: '1140' });
  const a6025 = record({ rowNumber: 3, material: 'MAT-A', stloc: '6025' });
  const b1140 = record({ rowNumber: 4, material: 'MAT-B', stloc: '1140' });
  const result = assign([a1140, a6025, b1140], [
    series({ rowNumber: 2, material: 'MAT-A', location: '1140', serial: 'GP28' }),
    series({ rowNumber: 3, material: 'MAT-A', location: '6025', serial: 'GP28' }),
    series({ rowNumber: 4, material: 'MAT-B', location: '1140', serial: 'GP28' }),
  ]);
  equal(a1140.serials.length, 0, 'duplicado MAT-A en 1140');
  equal(a6025.serials.length, 0, 'duplicado MAT-A en 6025');
  equal(b1140.serials.length, 1, 'GP28 legítima de MAT-B');
  equal(result.issues.filter(issue => issue.type === 'duplicate-series').length, 2, 'filas duplicadas de MAT-A');
});

test('rechaza asociaciones ambiguas sin detener materiales válidos', () => {
  const first = record({ material: 'AMB', stloc: '1140' });
  const second = record({ rowNumber: 3, material: 'AMB', stloc: '1140', batch: 'LOTE-B' });
  const valid = record({ rowNumber: 4, material: 'VALIDO', stloc: '6025' });
  const result = assign([first, second, valid], [
    series({ rowNumber: 2, material: 'AMB', location: '1140', serial: 'AMB-1' }),
    series({ rowNumber: 3, material: 'VALIDO', location: '6025', serial: 'OK-1' }),
  ]);
  equal(first.serials.length + second.serials.length, 0, 'series ambiguas asignadas');
  equal(valid.serials.length, 1, 'material válido');
  assert(result.issues.some(issue => issue.type === 'ambiguous-series'), 'incidencia ambigua');
});

test('detecta filas incompletas y advierte la serie 0 sin invalidarla', () => {
  const target = record({ material: '106498' });
  const result = assign([target], [
    series({ rowNumber: 10, material: '106498', serial: '0' }),
    series({ rowNumber: 11, material: '106498', serial: 'SIN-LOCATION', location: '' }),
    series({ rowNumber: 12, material: '106498', serial: 'SIN-DESCRIPCION', description: '' }),
  ]);
  equal(target.serials.length, 2, 'series asociables, incluida 0');
  assert(target.serials.some(item => item.serial === '0'), 'la serie 0 debe conservarse');
  equal(result.unresolved.length, 1, 'fila sin Location');
  assert(result.issues.some(issue => issue.type === 'suspicious-series' && issue.serial === '0'), 'advertencia de serie 0');
  assert(result.issues.some(issue => issue.type === 'incomplete-series-row' && issue.rowNumber === 11 && issue.severity === 'critical'), 'fila sin Location');
  assert(result.issues.some(issue => issue.type === 'incomplete-series-row' && issue.rowNumber === 12 && issue.severity === 'warning'), 'descripción faltante');
});

test('cuenta de forma independiente series 0 repetidas en materiales DEF', () => {
  const target = record({ material: '106498DEF' });
  const result = assign([target], [
    series({ rowNumber: 20, material: '106498DEF', serial: '0' }),
    series({ rowNumber: 21, material: '106498DEF', serial: '0' }),
  ]);
  equal(target.serials.length, 2, 'unidades DEF disponibles');
  assert(result.issues.every(issue => issue.type !== 'duplicate-series' && issue.type !== 'suspicious-series'), 'DEF no debe generar incidencias bloqueantes por 0');
  equal(result.issues.filter(issue => issue.type === 'duplicate-series-def').length, 2, 'advertencias DEF');
  const firstKey = InventoryModel.seriesSelectionKey(target.serials[0]);
  const secondKey = InventoryModel.seriesSelectionKey(target.serials[1]);
  assert(firstKey !== secondKey, 'cada unidad DEF debe tener identidad física propia');
  let selected = InventoryModel.toggleSeriesSelection([target], {}, target, firstKey);
  equal(selected[target.key].length, 1, 'primera unidad DEF');
  selected = InventoryModel.toggleSeriesSelection([target], selected, target, secondKey);
  equal(selected[target.key].length, 2, 'dos unidades DEF');
  selected = InventoryModel.toggleSeriesSelection([target], selected, target, firstKey);
  equal(selected[target.key].length, 1, 'desmarcar una unidad no afecta la otra');
  equal(selected[target.key][0], secondKey, 'segunda unidad conservada');
  InventoryModel.assertUniqueSeriesOwnership([target]);
});

test('la validación defensiva permite el mismo serial en materiales distintos', () => {
  const first = record({ material: 'MAT-A' });
  const second = record({ rowNumber: 3, material: 'MAT-B' });
  assign([first, second], [
    series({ rowNumber: 2, material: 'MAT-A', serial: 'GP28' }),
    series({ rowNumber: 3, material: 'MAT-B', serial: 'GP28' }),
  ]);
  InventoryModel.assertUniqueSeriesOwnership([first, second]);
});

print(`${passed} pruebas aprobadas`);
