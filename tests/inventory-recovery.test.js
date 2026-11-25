globalThis.TextEncoder ??= class TextEncoder {
  encode(value) {
    const encoded = unescape(encodeURIComponent(String(value)));
    return Uint8Array.from(encoded, character => character.charCodeAt(0));
  }
};

load('inventory-model.js');
load('inventory-storage.js');

let passed = 0;

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function equal(actual, expected, message) {
  assert(actual === expected, `${message}: se esperaba ${expected}, se obtuvo ${actual}`);
}

async function test(name, callback) {
  await callback();
  passed += 1;
  print(`OK ${name}`);
}

function record(overrides = {}) {
  return {
    rowNumber: 2,
    material: 'MAT-001',
    stloc: '1140',
    batch: 'LOTE-A',
    stockType: 'Unrestricted',
    barcode: '7800001',
    sapQty: '2',
    sapLocations: 'BIN-A1',
    serialized: true,
    serials: [{ serial: 'SER-001' }, { serial: 'SER-002' }],
    count: '',
    currentLocation: '',
    comments: '',
    ...overrides,
  };
}

function inventory() {
  return [
    record(),
    record({ rowNumber: 3, material: 'MAT-002', batch: 'LOTE-B', barcode: '7800002', sapQty: 5, sapLocations: 'BIN-B1', serialized: false, serials: [] }),
  ];
}

async function sessionId(records, unmatchedSeries = []) {
  return InventoryStorage.inventorySessionId(records, null, InventoryModel, unmatchedSeries);
}

async function run() {
  await test('prepara un único reinicio de las sesiones anteriores', () => {
    let cleared = 0;
    const database = {
      objectStoreNames: { contains: name => name === 'sessions' },
      createObjectStore: () => { throw new Error('no debe recrear la colección existente'); },
    };
    const transaction = { objectStore: () => ({ clear: () => { cleared += 1; } }) };
    equal(InventoryStorage.upgradeSessionStore(database, transaction, 1), 'cleared', 'migración desde V1');
    equal(cleared, 1, 'una limpieza durante la actualización');
    equal(InventoryStorage.upgradeSessionStore(database, transaction, 2), 'unchanged', 'la misma versión no repite el reinicio');
    equal(cleared, 1, 'sin segunda limpieza');
  });

  await test('una instalación nueva crea el almacén sin ejecutar una limpieza', () => {
    let created = 0;
    const database = {
      objectStoreNames: { contains: () => false },
      createObjectStore: (name, options) => { if (name === 'sessions' && options.keyPath === 'id') created += 1; },
    };
    const transaction = { objectStore: () => { throw new Error('no existe todavía'); } };
    equal(InventoryStorage.upgradeSessionStore(database, transaction, 0), 'created', 'instalación nueva');
    equal(created, 1, 'almacén creado');
  });

  await test('reconoce exactamente el mismo inventario', async () => {
    equal(await sessionId(inventory()), await sessionId(inventory()), 'huella exacta');
  });

  await test('ignora campos de conteo al identificar el inventario', async () => {
    const changed = inventory();
    changed[0].count = '7';
    changed[0].currentLocation = 'PATIO';
    changed[0].comments = 'REVISADO';
    equal(await sessionId(inventory()), await sessionId(changed), 'los datos físicos no cambian la huella SAP');
  });

  for (const [name, change] of [
    ['batch', item => { item.batch = 'LOTE-NUEVO'; }],
    ['tipo de stock', item => { item.stockType = 'Blocked'; }],
    ['cantidad SAP', item => { item.sapQty = 3; }],
    ['ubicación SAP', item => { item.sapLocations = 'BIN-Z9'; }],
    ['números de serie', item => { item.serials = [{ serial: 'SER-001' }, { serial: 'SER-003' }]; }],
  ]) {
    await test(`cambia la huella cuando cambia ${name}`, async () => {
      const changed = inventory();
      change(changed[0]);
      assert(await sessionId(inventory()) !== await sessionId(changed), `la huella debe incluir ${name}`);
    });
  }

  await test('un archivo reordenado requiere confirmación pero conserva coincidencias seguras', async () => {
    const original = inventory();
    const reordered = [original[1], original[0]];
    assert(await sessionId(original) !== await sessionId(reordered), 'el orden modificado debe producir otra sesión');
    const plan = InventoryModel.recoveryPlan(reordered, original);
    equal(plan.compatibleCount, 2, 'registros compatibles');
    equal(plan.omittedCount, 0, 'registros omitidos');
  });

  await test('omite registros cuya cantidad SAP cambió', () => {
    const original = inventory();
    const changed = inventory();
    changed[0].sapQty = 4;
    const plan = InventoryModel.recoveryPlan(changed, original);
    equal(plan.compatibleCount, 1, 'solo un registro compatible');
    equal(plan.omittedCount, 1, 'registro modificado omitido');
  });

  await test('mide solo progreso guardado explícitamente por la aplicación', () => {
    const previous = inventory();
    previous[0].count = '9';
    InventoryModel.assignStableRecordKeys(previous);
    equal(InventoryStorage.savedSessionProgress({ records: previous }, { changes: {} }, InventoryModel), 0, 'el conteo del Excel histórico no es progreso de la aplicación');
    const saved = { changes: { [previous[0].key]: { count: 0 } }, selectedSeries: {} };
    equal(InventoryStorage.savedSessionProgress({ records: previous }, saved, InventoryModel), 1, 'el cero explícito sí es progreso');
  });

  await test('una sesión exacta vacía no oculta una sesión histórica compatible', async () => {
    const current = inventory();
    const exactRecords = inventory();
    const historicalRecords = inventory();
    historicalRecords[0].sapQty = 1;
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(exactRecords);
    InventoryModel.assignStableRecordKeys(historicalRecords);
    const exact = { id: 'exacta-vacia', fileName: 'actual.xlsx', bytes: new Uint8Array([1]), changes: {}, selectedSeries: {}, workbook: { records: exactRecords, sessionId: 'current' } };
    const historical = { id: 'historica', fileName: 'anterior.xlsx', bytes: new Uint8Array([2]), changes: { [historicalRecords[1].key]: { count: '5' } }, selectedSeries: {}, workbook: { records: historicalRecords, sessionId: 'previous' } };
    const candidates = await InventoryStorage.evaluateSessionCandidates([exact, historical], { records: current }, 'current', {
      loadWorkbook: saved => saved.workbook,
      inventorySessionId: workbook => workbook.sessionId,
      recoveryPlan: InventoryModel.recoveryPlan,
      model: InventoryModel,
    });
    equal(candidates.exact.progress, 0, 'sesión exacta sin progreso');
    equal(candidates.compatible.recoverableCount, 1, 'conteo histórico recuperable');
    const choice = InventoryStorage.chooseRecoveryCandidate(candidates);
    equal(choice.automatic, null, 'la sesión exacta vacía no se aplica automáticamente');
    equal(choice.prompt.saved.id, 'historica', 'se ofrece la sesión histórica');
  });

  await test('una sesión exacta con progreso conserva prioridad', async () => {
    const current = inventory();
    const exactRecords = inventory();
    const historicalRecords = inventory();
    historicalRecords[0].sapQty = 1;
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(exactRecords);
    InventoryModel.assignStableRecordKeys(historicalRecords);
    const exact = { id: 'exacta', fileName: 'actual.xlsx', bytes: new Uint8Array([1]), changes: { [exactRecords[0].key]: { count: '2' } }, selectedSeries: {}, workbook: { records: exactRecords, sessionId: 'current' } };
    const historical = { id: 'historica', fileName: 'anterior.xlsx', bytes: new Uint8Array([2]), changes: { [historicalRecords[1].key]: { count: '5' } }, selectedSeries: {}, workbook: { records: historicalRecords, sessionId: 'previous' } };
    const candidates = await InventoryStorage.evaluateSessionCandidates([exact, historical], { records: current }, 'current', {
      loadWorkbook: saved => saved.workbook,
      inventorySessionId: workbook => workbook.sessionId,
      recoveryPlan: InventoryModel.recoveryPlan,
      model: InventoryModel,
    });
    const choice = InventoryStorage.chooseRecoveryCandidate(candidates);
    equal(choice.automatic.saved.id, 'exacta', 'sesión exacta elegida');
    equal(choice.prompt, null, 'sin aviso compatible cuando la exacta tiene progreso');
  });

  await test('no ofrece sesiones incompatibles aunque contengan conteos', async () => {
    const current = [record({ sapQty: 3 })];
    const previous = [record({ sapQty: 1 })];
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = { id: 'incompatible', fileName: 'anterior.xlsx', bytes: new Uint8Array([1]), changes: { [previous[0].key]: { count: '1' } }, selectedSeries: {}, workbook: { records: previous, sessionId: 'previous' } };
    const candidates = await InventoryStorage.evaluateSessionCandidates([saved], { records: current }, 'current', {
      loadWorkbook: item => item.workbook,
      inventorySessionId: workbook => workbook.sessionId,
      recoveryPlan: InventoryModel.recoveryPlan,
      model: InventoryModel,
    });
    equal(candidates.compatible, null, 'sesión incompatible descartada');
    equal(candidates.errors.length, 0, 'la incompatibilidad no es error de lectura');
  });

  await test('informa errores al leer sesiones históricas', async () => {
    const saved = { id: 'ilegible', fileName: 'sesion-antigua.xlsx', bytes: new Uint8Array([1]), changes: {}, selectedSeries: {} };
    const candidates = await InventoryStorage.evaluateSessionCandidates([saved], { records: inventory() }, 'current', {
      loadWorkbook: () => { throw new Error('Falta la pestaña histórica requerida.'); },
      inventorySessionId: () => 'unused',
      recoveryPlan: InventoryModel.recoveryPlan,
      model: InventoryModel,
    });
    equal(candidates.errors.length, 1, 'error conservado');
    equal(candidates.errors[0].fileName, 'sesion-antigua.xlsx', 'archivo informado');
    assert(/pestaña histórica/.test(candidates.errors[0].message), 'motivo comprensible');
  });

  await test('recupera solo filas existentes en ambos inventarios', () => {
    const current = [...inventory(), record({ rowNumber: 4, material: 'MAT-003', batch: 'LOTE-C', barcode: '7800003', serialized: false, serials: [] })];
    const plan = InventoryModel.recoveryPlan(current, inventory());
    equal(plan.compatibleCount, 2, 'registros compartidos');
    equal(plan.omittedCount, 1, 'registro nuevo omitido');
  });

  await test('no recupera registros duplicados o ambiguos', () => {
    const duplicateA = record({ rowNumber: 2 });
    const duplicateB = record({ rowNumber: 3 });
    const plan = InventoryModel.recoveryPlan([duplicateA, duplicateB], [record({ rowNumber: 8 }), record({ rowNumber: 9 })]);
    equal(plan.compatibleCount, 0, 'duplicados compatibles');
    equal(plan.omittedCount, 2, 'duplicados omitidos');
    equal(plan.ambiguousCount, 2, 'duplicados informados como ambiguos');
  });

  await test('recupera cambios y series al reabrir el mismo inventario', () => {
    const current = inventory();
    const previous = inventory();
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = {
      changes: { [previous[0].key]: { count: '2', currentLocation: 'BIN-A1', comments: 'CONTADO' } },
      selectedSeries: { [previous[0].key]: ['SER-002'] },
    };
    const plan = InventoryModel.recoveryPlan(current, previous);
    const result = InventoryModel.recoveryData(current, previous, saved, plan);
    equal(result.changes[current[0].key].count, '2', 'conteo recuperado');
    equal(result.changes[current[0].key].comments, 'CONTADO', 'comentario recuperado');
    equal(result.selectedSeries[current[0].key][0], 'SER-002', 'serie recuperada');
    equal(result.recoveredRows, 1, 'fila recuperada');
  });

  await test('el conteo guardado por la aplicación prevalece en recuperación compatible', () => {
    const current = inventory();
    const previous = inventory();
    current[0].count = '9';
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = { changes: { [previous[0].key]: { count: '2', comments: 'ANTERIOR' } } };
    const plan = InventoryModel.recoveryPlan(current, previous);
    const result = InventoryModel.recoveryData(current, previous, saved, plan, { mode: 'compatible' });
    equal(result.changes[current[0].key].count, '2', 'conteo de la aplicación recuperado');
    equal(result.changes[current[0].key].comments, 'ANTERIOR', 'campo vacío recuperado');
    equal(result.changes[current[0].key].sapQty, undefined, 'la recuperación no escribe cantidad SAP');
    equal(current[0].sapQty, previous[0].sapQty, 'la cantidad SAP actual permanece intacta');
    equal(result.skippedConflicts, 0, 'el conteo no se trata como conflicto');
  });

  await test('conserva un conteo explícito de cero sobre el Excel', () => {
    const current = inventory();
    const previous = inventory();
    current[0].count = '8';
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = { changes: { [previous[0].key]: { count: 0 } } };
    const result = InventoryModel.recoveryData(current, previous, saved, InventoryModel.recoveryPlan(current, previous), { mode: 'compatible' });
    equal(result.changes[current[0].key].count, 0, 'cero recuperado');
    equal(InventoryModel.physicalCountState(result.changes[current[0].key].count).status, 'valid', 'cero válido');
  });

  await test('usa el Excel nuevo cuando no existe conteo explícito en la aplicación', () => {
    const current = inventory();
    const previous = inventory();
    current[0].count = '7';
    previous[0].count = '4';
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const result = InventoryModel.recoveryData(current, previous, { changes: {} }, InventoryModel.recoveryPlan(current, previous), { includeSavedWorkbook: true });
    equal(result.changes[current[0].key]?.count, undefined, 'el Excel antiguo no aporta conteo');
    equal(current[0].count, '7', 'conteo del Excel nuevo conservado');
  });

  await test('rechaza conteos guardados inválidos sin convertirlos en cero', () => {
    const current = inventory();
    const previous = inventory();
    current[0].count = '3';
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = { changes: { [previous[0].key]: { count: 'NO CONTADO' } } };
    const result = InventoryModel.recoveryData(current, previous, saved, InventoryModel.recoveryPlan(current, previous));
    equal(result.changes[current[0].key]?.count, undefined, 'conteo inválido omitido');
    equal(current[0].count, '3', 'conteo válido del Excel conservado');
  });

  await test('distingue vacío, cero y valores físicos inválidos', () => {
    equal(InventoryModel.physicalCountState('').status, 'empty', 'celda vacía');
    equal(InventoryModel.physicalCountState('   ').status, 'empty', 'espacios vacíos');
    equal(InventoryModel.physicalCountState(0).status, 'valid', 'cero numérico');
    equal(InventoryModel.physicalCountState('0').status, 'valid', 'cero de texto');
    for(const value of ['texto', '-1', '1,5', Infinity, NaN])equal(InventoryModel.physicalCountState(value).status, 'invalid', `valor inválido ${value}`);
  });

  await test('filtra selecciones de series que ya no existen', () => {
    const selected = InventoryModel.compatibleSelectedSeries(record(), ['SER-002', 'SER-999', 'ser-002']);
    equal(selected.length, 1, 'series recuperables');
    equal(selected[0], 'SER-002', 'serie conservada');
  });

  await test('conserva el conteo guardado aunque una serie incompatible deje de asociarse', () => {
    const current = [record({ material: '29080784', stloc: '6S21', serials: [], count: '1' })];
    const previous = [record({ material: '29080784', stloc: '6S21', serials: [], count: '1' })];
    InventoryModel.assignStableRecordKeys(current);
    InventoryModel.assignStableRecordKeys(previous);
    const saved = { changes: { [previous[0].key]: { count: '1' } }, selectedSeries: { [previous[0].key]: ['LM0012'] } };
    const result = InventoryModel.recoveryData(current, previous, saved, InventoryModel.recoveryPlan(current, previous));
    equal(result.changes[current[0].key].count, '1', 'conteo físico recuperado');
    equal(result.selectedSeries[current[0].key].length, 0, 'serie incompatible descartada');
  });

  await test('incluye series no asociadas en la identidad global', async () => {
    const first = await sessionId(inventory(), [{ material: 'MAT-X', location: '1140', serial: 'SIN-ASIGNAR-1' }]);
    const second = await sessionId(inventory(), [{ material: 'MAT-X', location: '1140', serial: 'SIN-ASIGNAR-2' }]);
    assert(first !== second, 'las series no asociadas deben cambiar la huella');
  });

  await test('mantiene compatibilidad de lectura con claves guardadas antiguas', () => {
    const old = record();
    InventoryModel.assignStableRecordKeys([old]);
    const saved = { changes: { [old.legacyKey]: { count: '6' } }, selectedSeries: { [old.legacyKey]: ['SER-001'] } };
    equal(InventoryModel.savedChange(saved, old).count, '6', 'conteo legado');
    equal(InventoryModel.savedSeries(saved, old)[0], 'SER-001', 'series legadas');
  });

  await test('conserva datos existentes al actualizar una sesión exacta', () => {
    const existing = { changes: { ambiguo: { count: '4' }, actual: { comments: 'ANTERIOR' } }, selectedSeries: { ambiguo: ['SER-001'] } };
    const merged = InventoryStorage.mergeSessionValue(existing, { id: 'inventory-v3-test', changes: { actual: { comments: 'NUEVO' } }, selectedSeries: {} });
    equal(merged.changes.ambiguo.count, '4', 'conteo ambiguo preservado');
    equal(merged.changes.actual.comments, 'NUEVO', 'cambio actual actualizado');
    equal(merged.selectedSeries.ambiguo[0], 'SER-001', 'series ambiguas preservadas');
    equal(existing.changes.actual.comments, 'ANTERIOR', 'sesión fuente sin mutación');
  });

  await test('confirma la hora solo después de un guardado exitoso', async () => {
    let lastSavedAt = null;
    lastSavedAt = await InventoryStorage.confirmSave(async () => {}, 12345);
    equal(lastSavedAt, 12345, 'hora confirmada');
  });

  await test('un fallo de escritura no reemplaza la última hora confirmada', async () => {
    let lastSavedAt = 12345;
    let error = null;
    try {
      lastSavedAt = await InventoryStorage.confirmSave(async () => { throw new Error('sin espacio'); }, 67890);
    } catch (caught) {
      error = caught;
    }
    equal(lastSavedAt, 12345, 'hora anterior conservada');
    assert(error && /sin espacio/.test(error.message), 'error propagado');
  });

  await test('serializa tres guardados y solo confirma al terminar la cola', async () => {
    const events = [];
    const statuses = [];
    let operation = 0;
    const queue = InventoryStorage.createSerialSaveQueue(async () => {
      const current = ++operation;
      events.push(`inicio-${current}`);
      await new Promise(resolve => setTimeout(resolve, 3));
      events.push(`fin-${current}`);
    }, status => statuses.push(status));
    const first = queue.enqueue();
    const second = queue.enqueue();
    const third = queue.enqueue();
    await Promise.all([first, second, third]);
    equal(events.join(','), 'inicio-1,fin-1,inicio-2,fin-2,inicio-3,fin-3', 'orden de escritura');
    equal(statuses.filter(status => status === 'saved').length, 1, 'una confirmación final');
    equal(statuses.at(-1), 'saved', 'estado final');
  });

  await test('rechaza una transacción abortada una sola vez', async () => {
    let transaction;
    const fakeIndexedDB = {
      open() {
        const request = {};
        setTimeout(() => {
          request.result = {
            transaction() {
              transaction = {
                error: new Error('aborto simulado'),
                objectStore: () => ({ put: () => setTimeout(() => { transaction.onabort(); transaction.onerror(); }, 0) }),
              };
              return transaction;
            },
          };
          request.onsuccess();
        }, 0);
        return request;
      },
    };
    let error = null;
    try {
      await InventoryStorage.createSessionStore(fakeIndexedDB, 'prueba-aborto').put({ id: 'uno' });
    } catch (caught) {
      error = caught;
    }
    assert(error && /aborto simulado/.test(error.message), 'aborto propagado');
  });

  print(`${passed} pruebas aprobadas`);
}

run().catch(error => { throw error; });
