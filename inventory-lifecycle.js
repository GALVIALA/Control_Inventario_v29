/* Pure lifecycle helpers for independent inventories and safe Excel updates. */
(() => {
  'use strict';

  const text = value => String(value ?? '').trim();
  const clone = value => typeof structuredClone === 'function' ? structuredClone(value) : JSON.parse(JSON.stringify(value));
  const sameList = (left, right) => left.length === right.length && left.every((value, index) => value === right[index]);
  const makeId = (prefix = 'inventory') => `${prefix}-${globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`}`;

  function localDateStamp(now = Date.now()) {
    const date = new Date(now);
    return [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
  }

  function exportFileBase(name) {
    return assertInventoryName(name).replace(/\s+/g, '_').replace(/[\\/:*?"<>|]/g, '_');
  }

  function nextExportMetadata(inventory, now = Date.now()) {
    assertOpen(inventory);
    const revision = Math.max(0, Number(inventory.exportRevision) || 0) + 1;
    const date = localDateStamp(now);
    return { revision, date, fileName: `${exportFileBase(inventory.name)}_REV${String(revision).padStart(2, '0')}_${date}.xlsx` };
  }

  function closedExportMetadata(inventory) {
    if (!inventory || inventory.status !== 'closed') return null;
    const revision = Number(inventory.finalExportRevision);
    const date = String(inventory.finalExportDate || '');
    const storedName = String(inventory.finalBackupFileName || '');
    if (revision > 0 && /^\d{4}-\d{2}-\d{2}$/.test(date) && storedName) return { revision, date, fileName: storedName };
    const recovered = storedName.match(/_REV(\d+)_([0-9]{4}-[0-9]{2}-[0-9]{2})\.xlsx$/i);
    return recovered ? { revision: Number(recovered[1]), date: recovered[2], fileName: storedName } : null;
  }

  function assertInventoryName(name) {
    const clean = text(name).replace(/\s+/g, ' ');
    if (clean.length < 3) throw new Error('Ingresa un nombre de inventario de al menos tres caracteres.');
    if (clean.length > 80) throw new Error('El nombre del inventario no puede superar 80 caracteres.');
    return clean;
  }

  function createInventory({ name, workbook, sourceFingerprint, bytes, now = Date.now() }) {
    if (!workbook?.records?.length) throw new Error('El Excel no contiene materiales para crear el inventario.');
    return {
      id: makeId(),
      name: assertInventoryName(name),
      status: 'open',
      createdAt: now,
      closedAt: null,
      updatedAt: now,
      fileName: workbook.fileName,
      bytes: bytes || workbook.bytes,
      sourceFingerprint,
      identityVersion: 3,
      changes: {},
      selectedSeries: {},
      lastView: { view: 'inventory', search: '', scrollY: 0 },
    };
  }

  function createBackup(inventory, reason, now = Date.now()) {
    return {
      id: makeId('backup'),
      inventoryId: inventory.id,
      reason,
      createdAt: now,
      snapshot: clone(inventory),
    };
  }

  function assertOpen(inventory) {
    if (!inventory || inventory.status !== 'open') throw new Error('El inventario está cerrado y solo puede consultarse.');
  }

  function duplicateBases(records, model) {
    const counts = new Map();
    for (const record of records || []) {
      const base = model.recordIdentityBase(record);
      counts.set(base, (counts.get(base) || 0) + 1);
    }
    return [...counts].filter(([, count]) => count > 1).map(([base]) => base);
  }

  function effectiveValue(inventory, record, field, model) {
    const change = model.savedChange(inventory, record);
    return Object.prototype.hasOwnProperty.call(change, field) ? change[field] : record[field];
  }

  function buildExcelUpdatePlan({ inventory, currentWorkbook, candidateWorkbook, candidateFingerprint, model }) {
    assertOpen(inventory);
    if (!currentWorkbook?.records || !candidateWorkbook?.records) throw new Error('No fue posible comparar los inventarios.');
    if (candidateFingerprint !== inventory.sourceFingerprint) throw new Error('El archivo no corresponde al inventario activo o cambió información estructural protegida.');
    const currentDuplicates = duplicateBases(currentWorkbook.records, model);
    const candidateDuplicates = duplicateBases(candidateWorkbook.records, model);
    if (currentDuplicates.length || candidateDuplicates.length) throw new Error('La actualización contiene materiales duplicados o ambiguos y no puede aplicarse automáticamente.');
    if (candidateWorkbook.invalidPhysicalCountRows?.length) {
      throw new Error(`COUNT QTY contiene ${candidateWorkbook.invalidPhysicalCountRows.length} valor(es) que no pudieron interpretarse.`);
    }
    const current = new Map(currentWorkbook.records.map(record => [record.key, record]));
    const candidate = new Map(candidateWorkbook.records.map(record => [record.key, record]));
    const missing = [...current.keys()].filter(key => !candidate.has(key));
    const extra = [...candidate.keys()].filter(key => !current.has(key));
    if (missing.length || extra.length) throw new Error(`La estructura no coincide: ${missing.length} registro(s) ausente(s) y ${extra.length} registro(s) adicional(es).`);

    const changes = {};
    const selectedSeries = clone(inventory.selectedSeries || {});
    const details = [];
    const unknownSeriesStyles = [];
    const counts = { quantities: 0, locations: 0, comments: 0, series: 0 };
    for (const [key, oldRecord] of current) {
      const nextRecord = candidate.get(key);
      const patch = {};
      for (const [field, counter, label] of [
        ['count', 'quantities', 'COUNT QTY'],
        ['currentLocation', 'locations', 'COUNT BIN'],
        ['comments', 'comments', 'COMMENTS'],
      ]) {
        const before = String(effectiveValue(inventory, oldRecord, field, model) ?? '');
        const after = String(nextRecord[field] ?? '');
        if (before !== after) {
          patch[field] = nextRecord[field];
          counts[counter] += 1;
          details.push({ key, material: oldRecord.material, field, label, before, after });
        }
      }
      if (Object.keys(patch).length) changes[key] = patch;

      const beforeSeries = [...(inventory.selectedSeries?.[key] || [])];
      const nextSet = new Set(beforeSeries);
      for (const item of [...(nextRecord.serials || []), ...(nextRecord.alternateSerials || [])]) {
        const selectionKey = model.seriesSelectionKey(item);
        if (item.styleState === 'selected') {
          if (!item.defaultRecordKey || item.defaultRecordKey === nextRecord.key) nextSet.add(selectionKey);
          else nextSet.delete(selectionKey);
        }
        else if (item.styleState === 'unselected') nextSet.delete(selectionKey);
        else unknownSeriesStyles.push({ material: nextRecord.material, serial: item.serial, style: item.originalStyle });
      }
      const afterSeries = [...nextSet];
      if (!sameList([...beforeSeries].sort(), [...afterSeries].sort())) {
        selectedSeries[key] = afterSeries;
        counts.series += 1;
        details.push({ key, material: oldRecord.material, field: 'selectedSeries', label: 'SERIES', before: beforeSeries, after: afterSeries });
      }
    }
    return {
      inventoryId: inventory.id,
      candidateBytes: candidateWorkbook.bytes,
      candidateFileName: candidateWorkbook.fileName,
      changes,
      selectedSeries,
      counts,
      details,
      unknownSeriesStyles,
      hasChanges: details.length > 0,
    };
  }

  function applyExcelUpdate(inventory, plan, now = Date.now()) {
    assertOpen(inventory);
    if (plan.inventoryId !== inventory.id) throw new Error('La actualización preparada pertenece a otro inventario.');
    const changes = clone(inventory.changes || {});
    for (const [key, patch] of Object.entries(plan.changes || {})) changes[key] = { ...(changes[key] || {}), ...patch };
    return {
      ...clone(inventory),
      fileName: plan.candidateFileName || inventory.fileName,
      bytes: plan.candidateBytes || inventory.bytes,
      changes,
      selectedSeries: clone(plan.selectedSeries || inventory.selectedSeries || {}),
      updatedAt: now,
      lastExcelUpdateAt: now,
    };
  }

  function closeInventory(inventory, { backupFileName, exportRevision, exportDate, finalBackupBytes = null, now = Date.now() } = {}) {
    assertOpen(inventory);
    if (!backupFileName) throw new Error('Debes generar y guardar un respaldo Excel antes de cerrar el inventario.');
    if (!(Number(exportRevision) > 0) || !/^\d{4}-\d{2}-\d{2}$/.test(String(exportDate || ''))) throw new Error('El respaldo definitivo no tiene revisión y fecha válidas.');
    return { ...clone(inventory), status: 'closed', closedAt: now, updatedAt: now, finalBackupFileName: backupFileName, finalExportRevision: Number(exportRevision), finalExportDate: exportDate, finalBackupBytes: finalBackupBytes ? clone(finalBackupBytes) : null };
  }

  globalThis.InventoryLifecycle = Object.freeze({
    assertInventoryName,
    localDateStamp,
    exportFileBase,
    nextExportMetadata,
    closedExportMetadata,
    createInventory,
    createBackup,
    assertOpen,
    duplicateBases,
    effectiveValue,
    buildExcelUpdatePlan,
    applyExcelUpdate,
    closeInventory,
  });
})();
