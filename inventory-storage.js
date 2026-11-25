/* IndexedDB persistence and stable workbook-session identifiers. */
(() => {
  'use strict';

  const DATABASE_VERSION = 3;
  const ONE_TIME_RESET_VERSION = 2;
  const LEGACY_MIGRATION_KEY = 'legacy-v2-to-v3';

  function upgradeSessionStore(database, transaction, oldVersion, resetVersion = ONE_TIME_RESET_VERSION) {
    const hasSessions = database.objectStoreNames.contains('sessions');
    if (!hasSessions) {
      database.createObjectStore('sessions', { keyPath: 'id' });
      return 'created';
    }
    if (oldVersion < resetVersion) {
      transaction.objectStore('sessions').clear();
      return 'cleared';
    }
    return 'unchanged';
  }

  function upgradeDatabase(database, transaction, oldVersion) {
    const sessionResult = upgradeSessionStore(database, transaction, oldVersion);
    if (!database.objectStoreNames.contains('inventories')) database.createObjectStore('inventories', { keyPath: 'id' });
    if (!database.objectStoreNames.contains('settings')) database.createObjectStore('settings', { keyPath: 'key' });
    if (!database.objectStoreNames.contains('backups')) {
      const backups = database.createObjectStore('backups', { keyPath: 'id' });
      backups.createIndex?.('inventoryId', 'inventoryId', { unique: false });
    }
    return sessionResult;
  }

  const cloneValue = value => typeof structuredClone === 'function' ? structuredClone(value) : value;
  const requestPromise = request => new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('La operación de almacenamiento falló.'));
  });
  const transactionPromise = transaction => new Promise((resolve, reject) => {
    let settled = false;
    const finish = callback => { if (!settled) { settled = true; callback(); } };
    transaction.oncomplete = () => finish(resolve);
    transaction.onerror = () => finish(() => reject(transaction.error || new Error('La transacción de almacenamiento falló.')));
    transaction.onabort = () => finish(() => reject(transaction.error || new Error('La transacción de almacenamiento fue abortada.')));
  });

  function legacyInventoryFromSession(session, { active = false } = {}) {
    const updatedAt = Number(session?.updatedAt) || Number(session?.createdAt) || Date.now();
    const baseName = String(session?.fileName || 'Inventario migrado').replace(/\.xlsx$/i, '').replace(/_actualizado_.*$/i, '').trim() || 'Inventario migrado';
    return {
      ...cloneValue(session),
      id: `legacy-${session.id}`,
      name: baseName,
      status: 'open',
      createdAt: Number(session?.createdAt) || updatedAt,
      closedAt: null,
      sourceFingerprint: session?.inventoryId || session?.id || '',
      legacySessionId: session?.id || '',
      migratedFromVersion: 2,
      migratedAsHistorical: false,
      migratedAsActive: active,
      updatedAt,
    };
  }

  function createSessionStore(indexedDB, databaseName = 'vestas-inventario-offline') {
    return {
      open() {
        return new Promise((resolve, reject) => {
          const request = indexedDB.open(databaseName, DATABASE_VERSION);
          request.onupgradeneeded = event => upgradeDatabase(request.result, request.transaction, event.oldVersion);
          request.onsuccess = () => {
            request.result.onversionchange = () => request.result.close();
            resolve(request.result);
          };
          request.onerror = () => reject(request.error);
          request.onblocked = () => reject(new Error('Cierra las otras pestañas de Inventario para completar la actualización local.'));
        });
      },
      async put(value) {
        const database = await this.open();
        return new Promise((resolve, reject) => {
          const transaction = database.transaction('sessions', 'readwrite');
          let settled = false;
          const finish = callback => {
            if (settled) return;
            settled = true;
            callback();
          };
          transaction.objectStore('sessions').put(value);
          transaction.oncomplete = () => finish(resolve);
          transaction.onerror = () => finish(() => reject(transaction.error || new Error('No fue posible guardar la sesión.')));
          transaction.onabort = () => finish(() => reject(transaction.error || new Error('La transacción de guardado fue abortada.')));
        });
      },
      async get(id) {
        const database = await this.open();
        return new Promise((resolve, reject) => {
          const request = database.transaction('sessions').objectStore('sessions').get(id);
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
      },
      async all() {
        const database = await this.open();
        return new Promise((resolve, reject) => {
          const request = database.transaction('sessions').objectStore('sessions').getAll();
          request.onsuccess = () => resolve(request.result.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)));
          request.onerror = () => reject(request.error);
        });
      },
      async latest() {
        return (await this.all())[0];
      },
      async getInventory(id) {
        const database = await this.open();
        return requestPromise(database.transaction('inventories').objectStore('inventories').get(id));
      },
      async allInventories() {
        const database = await this.open();
        const values = await requestPromise(database.transaction('inventories').objectStore('inventories').getAll());
        return values.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
      },
      async getSetting(key) {
        const database = await this.open();
        return requestPromise(database.transaction('settings').objectStore('settings').get(key));
      },
      async putSetting(key, value) {
        const database = await this.open();
        const transaction = database.transaction('settings', 'readwrite');
        transaction.objectStore('settings').put({ key, value, updatedAt: Date.now() });
        await transactionPromise(transaction);
        return value;
      },
      async saveInventory(inventory, { backup = null, activate = false, clearActive = false } = {}) {
        const database = await this.open();
        const stores = backup ? ['inventories', 'backups', 'settings'] : ['inventories', 'settings'];
        const transaction = database.transaction(stores, 'readwrite');
        transaction.objectStore('inventories').put(cloneValue(inventory));
        if (backup) transaction.objectStore('backups').put(cloneValue(backup));
        if (activate) transaction.objectStore('settings').put({ key: 'activeInventoryId', value: inventory.id, updatedAt: Date.now() });
        else if (clearActive) transaction.objectStore('settings').put({ key: 'activeInventoryId', value: null, updatedAt: Date.now() });
        await transactionPromise(transaction);
        return inventory;
      },
      async backupsForInventory(inventoryId) {
        const database = await this.open();
        const store = database.transaction('backups').objectStore('backups');
        if (store.index) return requestPromise(store.index('inventoryId').getAll(inventoryId));
        return (await requestPromise(store.getAll())).filter(item => item.inventoryId === inventoryId);
      },
      async deleteClosedInventory(inventoryId) {
        const database = await this.open();
        return new Promise((resolve, reject) => {
          const transaction = database.transaction(['inventories', 'backups', 'settings', 'sessions'], 'readwrite');
          const inventories = transaction.objectStore('inventories');
          const backups = transaction.objectStore('backups');
          const settings = transaction.objectStore('settings');
          const sessions = transaction.objectStore('sessions');
          let failure = null;
          let result = null;
          const fail = message => {
            failure = new Error(message);
            transaction.abort();
          };
          transaction.oncomplete = () => resolve(result);
          transaction.onerror = () => reject(failure || transaction.error || new Error('No fue posible eliminar el inventario.'));
          transaction.onabort = () => reject(failure || transaction.error || new Error('La eliminación del inventario fue abortada.'));
          const inventoryRequest = inventories.get(inventoryId);
          inventoryRequest.onerror = () => { failure = inventoryRequest.error; };
          inventoryRequest.onsuccess = () => {
            const inventory = inventoryRequest.result;
            if (!inventory) { fail('El inventario seleccionado ya no existe.'); return; }
            if (inventory.status !== 'closed') { fail('Solo se pueden eliminar inventarios cerrados.'); return; }
            const activeRequest = settings.get('activeInventoryId');
            activeRequest.onerror = () => { failure = activeRequest.error; };
            activeRequest.onsuccess = () => {
              if (activeRequest.result?.value === inventoryId) { fail('No se puede eliminar el inventario activo.'); return; }
              const backupKeysRequest = backups.index('inventoryId').getAllKeys(inventoryId);
              backupKeysRequest.onerror = () => { failure = backupKeysRequest.error; };
              backupKeysRequest.onsuccess = () => {
                const backupKeys = backupKeysRequest.result || [];
                backupKeys.forEach(key => backups.delete(key));
                inventories.delete(inventoryId);
                if (inventory.legacySessionId) sessions.delete(inventory.legacySessionId);
                result = { inventoryId, deletedBackups: backupKeys.length, deletedLegacySession: Boolean(inventory.legacySessionId) };
              };
            };
          };
        });
      },
      async migrateLegacySessions() {
        const database = await this.open();
        return new Promise((resolve, reject) => {
          const transaction = database.transaction(['sessions', 'inventories', 'settings'], 'readwrite');
          const sessions = transaction.objectStore('sessions');
          const inventories = transaction.objectStore('inventories');
          const settings = transaction.objectStore('settings');
          let result = null;
          transaction.oncomplete = () => resolve(result);
          transaction.onerror = () => reject(transaction.error || new Error('No fue posible migrar los inventarios guardados.'));
          transaction.onabort = () => reject(transaction.error || new Error('La migración local fue abortada.'));
          const markerRequest = settings.get(LEGACY_MIGRATION_KEY);
          markerRequest.onerror = () => transaction.abort();
          markerRequest.onsuccess = () => {
            if (markerRequest.result) { result = markerRequest.result.value; return; }
            const inventoriesRequest = inventories.getAll();
            inventoriesRequest.onerror = () => transaction.abort();
            inventoriesRequest.onsuccess = () => {
              const sessionsRequest = sessions.getAll();
              sessionsRequest.onerror = () => transaction.abort();
              sessionsRequest.onsuccess = () => {
                const existingInventories = inventoriesRequest.result || [];
                const legacySessions = sessionsRequest.result || [];
                let migrated = 0;
                let activeInventoryId = null;
                if (!existingInventories.length && legacySessions.length) {
                  legacySessions.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
                  legacySessions.forEach((session, index) => {
                    const inventory = legacyInventoryFromSession(session, { active: index === 0 });
                    inventories.put(inventory);
                    if (index === 0) activeInventoryId = inventory.id;
                    migrated += 1;
                  });
                  if (activeInventoryId) settings.put({ key: 'activeInventoryId', value: activeInventoryId, updatedAt: Date.now() });
                }
                result = { migrated, activeInventoryId, completedAt: Date.now() };
                settings.put({ key: LEGACY_MIGRATION_KEY, value: result, updatedAt: result.completedAt });
              };
            };
          };
        });
      },
    };
  }

  async function digest(bytes, cryptoApi = globalThis.crypto) {
    if (cryptoApi?.subtle) {
      const hash = await cryptoApi.subtle.digest('SHA-256', bytes);
      return [...new Uint8Array(hash)].slice(0, 12).map(value => value.toString(16).padStart(2, '0')).join('');
    }
    let hash = 2166136261;
    for (const byte of new Uint8Array(bytes)) {
      hash ^= byte;
      hash = Math.imul(hash, 16777619);
    }
    return `${bytes.byteLength}-${(hash >>> 0).toString(16)}`;
  }

  function inventorySessionSignature(records, model = globalThis.InventoryModel, unmatchedSeries = []) {
    if (!model?.sourceRecordIdentity) throw new Error('No se pudo construir la identidad segura del inventario.');
    const recordSignature = records.map((record, index) => [
      index + 1,
      Number(record.rowNumber) || 0,
      model.sourceRecordIdentity(record),
    ].join('\u001f')).join('\n');
    const normalized = value => String(value ?? '').trim().toUpperCase().replace(/\s+/g, ' ');
    const unmatchedSignature = unmatchedSeries.map(item => [normalized(item.material), normalized(item.location), normalized(item.serial)].join('\u001f')).sort().join('\n');
    return `${recordSignature}\n\u001eUNMATCHED-SERIES\u001e\n${unmatchedSignature}`;
  }

  async function inventorySessionId(records, cryptoApi = globalThis.crypto, model = globalThis.InventoryModel, unmatchedSeries = []) {
    const signature = inventorySessionSignature(records, model, unmatchedSeries);
    return `inventory-v3-${await digest(new TextEncoder().encode(signature).buffer, cryptoApi)}`;
  }

  function recordHasExplicitProgress(record, saved, model = globalThis.InventoryModel) {
    if (!model?.savedChange || !model?.savedSeries || !model?.physicalCountState) throw new Error('No se pudo evaluar el progreso guardado.');
    const change = model.savedChange(saved, record);
    const hasCount = Object.prototype.hasOwnProperty.call(change, 'count') && model.physicalCountState(change.count).status === 'valid';
    const hasLocation = Object.prototype.hasOwnProperty.call(change, 'currentLocation') && String(change.currentLocation ?? '').trim() !== '';
    const hasComments = Object.prototype.hasOwnProperty.call(change, 'comments') && String(change.comments ?? '').trim() !== '';
    const selectedSeries = model.savedSeries(saved, record);
    return hasCount || hasLocation || hasComments || (Array.isArray(selectedSeries) && selectedSeries.length > 0);
  }

  function savedSessionProgress(workbook, saved, model = globalThis.InventoryModel) {
    return workbook.records.filter(record => recordHasExplicitProgress(record, saved, model)).length;
  }

  function recoverableSessionProgress(plan, saved, model = globalThis.InventoryModel) {
    return plan.matches.filter(match => recordHasExplicitProgress(match.saved, saved, model)).length;
  }

  async function evaluateSessionCandidates(savedSessions, currentWorkbook, currentSessionId, operations) {
    const { loadWorkbook, inventorySessionId: identify, recoveryPlan, model = globalThis.InventoryModel } = operations;
    const exact = [];
    const compatible = [];
    const errors = [];
    for (const saved of savedSessions) {
      try {
        if (!saved?.bytes) throw new Error('La sesión no contiene una copia del Excel de origen.');
        const workbook = await loadWorkbook(saved);
        const savedSessionId = await identify(workbook);
        const plan = recoveryPlan(currentWorkbook.records, workbook.records);
        const progress = savedSessionProgress(workbook, saved, model);
        const recoverableCount = recoverableSessionProgress(plan, saved, model);
        const candidate = { saved, workbook, plan, progress, recoverableCount, savedSessionId };
        if (savedSessionId === currentSessionId) exact.push(candidate);
        else if (recoverableCount > 0) compatible.push(candidate);
      } catch (error) {
        errors.push({
          sessionId: saved?.id || '',
          fileName: saved?.fileName || 'sesión sin nombre',
          message: error?.message || 'No fue posible leer la sesión.',
        });
      }
    }
    exact.sort((left, right) => right.progress - left.progress || (right.saved.updatedAt || 0) - (left.saved.updatedAt || 0));
    compatible.sort((left, right) => right.recoverableCount - left.recoverableCount || right.plan.compatibleCount - left.plan.compatibleCount || (right.saved.updatedAt || 0) - (left.saved.updatedAt || 0));
    return { exact: exact[0] || null, compatible: compatible[0] || null, errors };
  }

  function chooseRecoveryCandidate(candidates) {
    if (candidates.exact?.progress > 0) return { automatic: candidates.exact, prompt: null };
    if (candidates.compatible?.recoverableCount > 0) return { automatic: null, prompt: candidates.compatible };
    return { automatic: candidates.exact || null, prompt: null };
  }

  function mergeSessionValue(existing, value) {
    return {
      ...value,
      changes: { ...(existing?.changes || {}), ...(value.changes || {}) },
      selectedSeries: { ...(existing?.selectedSeries || {}), ...(value.selectedSeries || {}) },
    };
  }

  async function confirmSave(write, savedAt = Date.now()) {
    await write();
    return savedAt;
  }

  function createSerialSaveQueue(save, onStatus = () => {}) {
    let tail = Promise.resolve();
    let pending = 0;
    let lastError = null;
    return {
      enqueue() {
        pending += 1;
        onStatus('saving', { pending, error: null });
        const operation = tail.catch(() => {}).then(() => save());
        tail = operation.then(
          value => { lastError = null; return value; },
          error => { lastError = error; return undefined; },
        ).finally(() => {
          pending -= 1;
          if (!pending) onStatus(lastError ? 'error' : 'saved', { pending, error: lastError });
        });
        return tail;
      },
      get pending() { return pending; },
    };
  }

  globalThis.InventoryStorage = Object.freeze({
    DATABASE_VERSION,
    ONE_TIME_RESET_VERSION,
    LEGACY_MIGRATION_KEY,
    upgradeSessionStore,
    upgradeDatabase,
    legacyInventoryFromSession,
    createSessionStore,
    digest,
    inventorySessionSignature,
    inventorySessionId,
    recordHasExplicitProgress,
    savedSessionProgress,
    recoverableSessionProgress,
    evaluateSessionCandidates,
    chooseRecoveryCandidate,
    mergeSessionValue,
    confirmSave,
    createSerialSaveQueue,
  });
})();
