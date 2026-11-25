load('inventory-storage.js');

let passed = 0;
function assert(condition, message) { if (!condition) throw new Error(message); }
function equal(actual, expected, message) { assert(actual === expected, `${message}: ${actual} !== ${expected}`); }
function test(name, callback) { callback(); passed += 1; print(`OK ${name}`); }
async function asyncTest(name, callback) { await callback(); passed += 1; print(`OK ${name}`); }

function database(existing = ['sessions']) {
  const names = new Set(existing), created = [];
  return {
    objectStoreNames: { contains: name => names.has(name) },
    createObjectStore(name) { names.add(name); created.push(name); return { createIndex() {} }; },
    created,
  };
}

test('V2 a V3 crea almacenes sin repetir la limpieza', () => {
  const db = database(), transaction = { objectStore: () => ({ clear: () => { throw new Error('no debe limpiar'); } }) };
  equal(InventoryStorage.upgradeDatabase(db, transaction, 2), 'unchanged', 'sesiones preservadas');
  equal(db.created.sort().join(','), 'backups,inventories,settings', 'almacenes V3');
});

test('una instalación nueva crea todos los almacenes vacíos', () => {
  const db = database([]), transaction = { objectStore: () => ({ clear: () => { throw new Error('no debe limpiar'); } }) };
  equal(InventoryStorage.upgradeDatabase(db, transaction, 0), 'created', 'sessions nuevo');
  equal(db.created.sort().join(','), 'backups,inventories,sessions,settings', 'estructura completa');
});

test('la sesión V2 se convierte sin perder cambios ni series', () => {
  const session = { id: 'inventory-v3-test', fileName: 'End Year 2026.xlsx', changes: { a: { count: '7' } }, selectedSeries: { a: ['S1'] }, updatedAt: 100 };
  const migrated = InventoryStorage.legacyInventoryFromSession(session, { active: true });
  equal(migrated.name, 'End Year 2026', 'nombre derivado');
  equal(migrated.changes.a.count, '7', 'conteo');
  equal(migrated.selectedSeries.a[0], 'S1', 'serie');
  equal(migrated.status, 'open', 'activo');
});

function disposableIndexedDB(seed) {
  const data = {
    inventories: new Map(seed.inventories.map(value => [value.id, value])),
    backups: new Map(seed.backups.map(value => [value.id, value])),
    settings: new Map(seed.settings.map(value => [value.key, value])),
    sessions: new Map(seed.sessions.map(value => [value.id, value])),
  };
  const database = {
    onversionchange: null,
    close() {},
    transaction() {
      let pending = 0;
      const transaction = {
        error: null,
        aborted: false,
        objectStore(name) {
          const values = data[name];
          const request = operation => {
            const result = {};
            pending += 1;
            setTimeout(() => {
              if (transaction.aborted) return;
              try { result.result = operation(); result.onsuccess?.(); }
              catch (error) { result.error = error; transaction.error = error; result.onerror?.(); transaction.abort(); }
              pending -= 1;
              setTimeout(() => { if (!pending && !transaction.aborted) transaction.oncomplete?.(); }, 0);
            }, 0);
            return result;
          };
          return {
            get: key => request(() => values.get(key)),
            getAll: () => request(() => [...values.values()]),
            put: value => request(() => { const key = value.id ?? value.key; values.set(key, value); return key; }),
            delete: key => values.delete(key),
            index(indexName) {
              assert(name === 'backups' && indexName === 'inventoryId', 'índice solicitado');
              return { getAllKeys: inventoryId => request(() => [...values].filter(([, value]) => value.inventoryId === inventoryId).map(([key]) => key)) };
            },
          };
        },
        abort() {
          if (transaction.aborted) return;
          transaction.aborted = true;
          setTimeout(() => transaction.onabort?.(), 0);
        },
      };
      return transaction;
    },
  };
  return {
    data,
    open() {
      const request = {};
      setTimeout(() => { request.result = database; request.onsuccess?.(); }, 0);
      return request;
    },
  };
}

async function runDeletionTests() {
  await asyncTest('elimina solo el inventario cerrado seleccionado y sus registros asociados', async () => {
    const indexedDB = disposableIndexedDB({
      inventories: [
        { id: 'cerrado-a', name: 'Prueba desechable A', status: 'closed', legacySessionId: 'sesion-a', updatedAt: 3 },
        { id: 'cerrado-b', name: 'Prueba desechable B', status: 'closed', updatedAt: 2 },
        { id: 'abierto', name: 'Prueba abierta', status: 'open', updatedAt: 1 },
      ],
      backups: [
        { id: 'respaldo-a1', inventoryId: 'cerrado-a' },
        { id: 'respaldo-a2', inventoryId: 'cerrado-a' },
        { id: 'respaldo-b', inventoryId: 'cerrado-b' },
      ],
      settings: [{ key: 'activeInventoryId', value: 'abierto' }],
      sessions: [{ id: 'sesion-a' }, { id: 'sesion-b' }],
    });
    const firstConnection = InventoryStorage.createSessionStore(indexedDB, 'eliminacion-desechable');
    const result = await firstConnection.deleteClosedInventory('cerrado-a');
    equal(result.deletedBackups, 2, 'respaldos asociados eliminados');
    assert(!indexedDB.data.inventories.has('cerrado-a'), 'inventario seleccionado eliminado');
    assert(indexedDB.data.inventories.has('cerrado-b'), 'otro inventario cerrado conservado');
    assert(indexedDB.data.inventories.has('abierto'), 'inventario abierto conservado');
    assert(indexedDB.data.backups.has('respaldo-b'), 'respaldo ajeno conservado');
    assert(!indexedDB.data.sessions.has('sesion-a'), 'sesión heredada asociada eliminada');
    assert(indexedDB.data.sessions.has('sesion-b'), 'sesión ajena conservada');

    const reopened = InventoryStorage.createSessionStore(indexedDB, 'eliminacion-desechable');
    const remaining = await reopened.allInventories();
    equal(remaining.map(item => item.id).sort().join(','), 'abierto,cerrado-b', 'eliminación persistente al reabrir');
  });

  await asyncTest('rechaza eliminar un inventario abierto sin modificar datos', async () => {
    const indexedDB = disposableIndexedDB({ inventories: [{ id: 'abierto', status: 'open' }], backups: [], settings: [], sessions: [] });
    let error = null;
    try { await InventoryStorage.createSessionStore(indexedDB, 'rechazo-desechable').deleteClosedInventory('abierto'); }
    catch (caught) { error = caught; }
    assert(error && /cerrados/.test(error.message), 'rechazo comprensible');
    assert(indexedDB.data.inventories.has('abierto'), 'inventario abierto intacto');
  });

  await asyncTest('el historial puede quedar vacío y sigue vacío al reabrir', async () => {
    const indexedDB = disposableIndexedDB({
      inventories: [{ id: 'unico', name: 'Único inventario desechable', status: 'closed' }],
      backups: [], settings: [], sessions: [],
    });
    await InventoryStorage.createSessionStore(indexedDB, 'historial-vacio-desechable').deleteClosedInventory('unico');
    const remaining = await InventoryStorage.createSessionStore(indexedDB, 'historial-vacio-desechable').allInventories();
    equal(remaining.length, 0, 'historial vacío persistente');
  });

  await asyncTest('conserva la revisión de exportación después de reabrir', async () => {
    const indexedDB = disposableIndexedDB({ inventories: [], backups: [], settings: [], sessions: [] });
    const first = InventoryStorage.createSessionStore(indexedDB, 'revision-desechable');
    await first.saveInventory({ id: 'inventario-revision', status: 'open', exportRevision: 7, lastExportDate: '2026-11-24', lastExportFileName: 'Inventario_REV07_2026-11-24.xlsx' });
    const reopened = InventoryStorage.createSessionStore(indexedDB, 'revision-desechable');
    const saved = await reopened.getInventory('inventario-revision');
    equal(saved.exportRevision, 7, 'revisión persistente');
    equal(saved.lastExportDate, '2026-11-24', 'fecha persistente');
  });
}

runDeletionTests().then(() => print(`${passed} pruebas aprobadas`)).catch(error => { throw error; });
