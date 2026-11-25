/* Pure inventory model helpers. No DOM, network, or storage access. */
(() => {
  'use strict';

  const text = value => String(value ?? '').trim();
  const normalized = value => text(value).toUpperCase().replace(/\s+/g, ' ');
  const encoded = value => encodeURIComponent(normalized(value));
  const normalizedQuantity = value => {
    const raw = text(value);
    if (!raw) return '';
    const number = Number(raw);
    return Number.isFinite(number) ? String(number) : normalized(raw);
  };

  function legacyRecordKey(record) {
    return `${text(record.material)}::${text(record.stloc)}::${record.rowNumber}`;
  }

  function recordIdentityBase(record) {
    const fields = [
      record.material,
      record.stloc,
      record.batch,
      record.stockType,
      record.barcode,
    ];
    return `record-v3::${fields.map(encoded).join('::')}`;
  }

  function sourceRecordIdentity(record) {
    const serials = (record.serials || []).map(item => normalized(item.serial)).filter(Boolean).sort();
    const fields = [
      record.material,
      record.stloc,
      record.batch,
      record.stockType,
      record.barcode,
      normalizedQuantity(record.sapQty),
      record.sapLocations,
      record.serialized ? '1' : '0',
      serials.join('\u001d'),
    ];
    return `source-v1::${fields.map(encoded).join('::')}`;
  }

  function recoveryPlan(currentRecords, savedRecords) {
    const group = records => {
      const groups = new Map();
      for (const record of records) {
        const identity = sourceRecordIdentity(record);
        const list = groups.get(identity) || [];
        list.push(record);
        groups.set(identity, list);
      }
      return groups;
    };
    const currentGroups = group(currentRecords);
    const savedGroups = group(savedRecords);
    const matches = [];
    let ambiguousCount = 0;
    for (const [identity, currentGroup] of currentGroups) {
      const savedGroup = savedGroups.get(identity) || [];
      if (currentGroup.length === 1 && savedGroup.length === 1) {
        matches.push({ current: currentGroup[0], saved: savedGroup[0] });
      } else if (savedGroup.length && (currentGroup.length > 1 || savedGroup.length > 1)) {
        ambiguousCount += currentGroup.length;
      }
    }
    return {
      matches,
      compatibleCount: matches.length,
      omittedCount: currentRecords.length - matches.length,
      ambiguousCount,
    };
  }

  function compatibleSelectedSeries(record, selected) {
    const items = [...(record.serials || []), ...(record.alternateSerials || [])];
    const available = new Map(items.map(item => [normalized(seriesSelectionKey(item)), seriesSelectionKey(item)]));
    const legacy = new Map();
    for (const item of items) {
      const key = normalized(item.serial);
      const list = legacy.get(key) || [];
      list.push(seriesSelectionKey(item));
      legacy.set(key, list);
    }
    const result = [];
    for (const serial of selected || []) {
      const normalizedSelected = normalized(serial);
      const current = available.get(normalizedSelected) || (legacy.get(normalizedSelected) || []).find(key => !result.includes(key));
      if (current != null && !result.includes(current)) result.push(current);
    }
    return result;
  }

  function seriesSelectionKey(item) {
    return text(item?.selectionKey || item?.serial);
  }

  function toggleSeriesSelection(records, selectedSeries, targetRecord, selectionKey) {
    const next = { ...(selectedSeries || {}) };
    const targetSelected = new Set(next[targetRecord.key] || []);
    const removeOnly = targetSelected.has(selectionKey);
    for (const record of records) {
      if (normalized(record.material) !== normalized(targetRecord.material)) continue;
      const current = new Set(next[record.key] || []);
      current.delete(selectionKey);
      if (record.key === targetRecord.key && !removeOnly) current.add(selectionKey);
      next[record.key] = [...current];
    }
    return next;
  }

  function physicalCountState(value) {
    if (value == null) return { status: 'empty', value: '' };
    const raw = text(value);
    if (!raw) return { status: 'empty', value: '' };
    if (typeof value !== 'number' && !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw)) return { status: 'invalid', value: raw };
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0) return { status: 'invalid', value: raw };
    return { status: 'valid', value: typeof value === 'number' ? value : raw, number };
  }

  function numericQuantity(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    const raw = text(value);
    if (!raw) return 0;
    let cleaned = raw.replace(/[^\d,.-]/g, '');
    if (cleaned.includes(',') && cleaned.includes('.')) cleaned = cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.') ? cleaned.replace(/\./g, '').replace(',', '.') : cleaned.replace(/,/g, '');
    else if (cleaned.includes(',')) cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    const number = Number(cleaned);
    return Number.isFinite(number) ? number : 0;
  }

  function seriesConsistencyReport(records, seriesItems) {
    const ownerKey = (plant, material) => `${normalized(plant)}::${normalized(material)}`;
    const locationKey = value => normalized(value) || '(VACÍA)';
    const sortText = (a, b) => normalized(a).localeCompare(normalized(b), 'es', { numeric: true });
    const inventoryKeys = new Set();
    const inventoryGroups = new Map();
    for (const record of records || []) {
      if (!text(record.plant) || !text(record.material)) continue;
      const key = ownerKey(record.plant, record.material);
      inventoryKeys.add(key);
      const group = inventoryGroups.get(key) || {
        plant: text(record.plant), material: text(record.material), serialized: false,
        sapQuantity: 0, locations: new Map(),
      };
      group.serialized ||= record.serialControlRequired ?? !!record.serialized;
      const quantity = numericQuantity(record.sapQty);
      group.sapQuantity += quantity;
      const locKey = locationKey(record.stloc);
      const location = group.locations.get(locKey) || { location: text(record.stloc) || '(vacía)', sapQuantity: 0 };
      location.sapQuantity += quantity;
      group.locations.set(locKey, location);
      inventoryGroups.set(key, group);
    }

    const seriesGroups = new Map();
    for (const item of seriesItems || []) {
      if (!text(item.plant) || !text(item.material)) continue;
      const key = ownerKey(item.plant, item.material);
      const group = seriesGroups.get(key) || {
        plant: text(item.plant), material: text(item.material), count: 0, locations: new Map(),
      };
      if (text(item.serial)) {
        group.count += 1;
        const locKey = locationKey(item.location);
        const location = group.locations.get(locKey) || { location: text(item.location) || '(vacía)', count: 0 };
        location.count += 1;
        group.locations.set(locKey, location);
      }
      seriesGroups.set(key, group);
    }

    const absentByMaterial = new Map();
    for (const [key, group] of seriesGroups) {
      if (inventoryKeys.has(key)) continue;
      const materialKey = normalized(group.material);
      const absent = absentByMaterial.get(materialKey) || { material: group.material, plants: [] };
      if (!absent.plants.includes(group.plant)) absent.plants.push(group.plant);
      absentByMaterial.set(materialKey, absent);
    }
    const seriesOnlyMaterials = [...absentByMaterial.values()]
      .map(item => ({ ...item, plants: item.plants.sort(sortText) }))
      .sort((a, b) => sortText(a.material, b.material));

    const quantityDifferences = [];
    const locationDifferences = [];
    for (const [key, inventory] of inventoryGroups) {
      if (!inventory.serialized) continue;
      const series = seriesGroups.get(key) || { count: 0, locations: new Map() };
      const difference = series.count - inventory.sapQuantity;
      if (Math.abs(difference) > 1e-9) {
        quantityDifferences.push({
          plant: inventory.plant, material: inventory.material,
          sapQuantity: inventory.sapQuantity, seriesCount: series.count, difference,
        });
      }
      const locationKeys = new Set([...inventory.locations.keys(), ...series.locations.keys()]);
      const locations = [];
      for (const locKey of locationKeys) {
        const expected = inventory.locations.get(locKey);
        const actual = series.locations.get(locKey);
        const sapQuantity = expected?.sapQuantity || 0;
        const seriesCount = actual?.count || 0;
        if (Math.abs(seriesCount - sapQuantity) <= 1e-9) continue;
        locations.push({
          location: expected?.location || actual?.location || '(vacía)',
          sapQuantity, seriesCount, difference: seriesCount - sapQuantity,
        });
      }
      if (locations.length) {
        locations.sort((a, b) => sortText(a.location, b.location));
        locationDifferences.push({
          plant: inventory.plant, material: inventory.material,
          sapQuantity: inventory.sapQuantity, seriesCount: series.count, locations,
        });
      }
    }
    quantityDifferences.sort((a, b) => sortText(a.material, b.material) || sortText(a.plant, b.plant));
    locationDifferences.sort((a, b) => sortText(a.material, b.material) || sortText(a.plant, b.plant));
    return {
      seriesOnlyMaterials,
      quantityDifferences,
      locationDifferences,
      discrepancyCount: seriesOnlyMaterials.length + quantityDifferences.length + locationDifferences.length,
    };
  }

  function recoveryData(currentRecords, savedRecords, saved, plan, { includeSavedWorkbook = false, mode = 'exact' } = {}) {
    const fields = ['count', 'currentLocation', 'comments', 'keepSap', 'keepLocation'];
    const changes = {};
    const selectedSeries = {};
    for (const record of currentRecords) {
      const selected = [...(record.serials || []), ...(record.alternateSerials || [])].filter(item => item.selected && (!item.defaultRecordKey || item.defaultRecordKey === record.key)).map(seriesSelectionKey);
      if (selected.length) selectedSeries[record.key] = selected;
    }
    const savedWorkbookSeries = {};
    for (const record of savedRecords) {
      const selected = [...(record.serials || []), ...(record.alternateSerials || [])].filter(item => item.selected && (!item.defaultRecordKey || item.defaultRecordKey === record.key)).map(seriesSelectionKey);
      if (selected.length) savedWorkbookSeries[record.key] = selected;
    }
    const recoveredKeys = new Set();
    let skippedConflicts = 0;
    for (const match of plan.matches) {
      const record = match.current;
      const old = match.saved;
      const explicit = savedChange(saved, old);
      const next = {};
      for (const field of fields) {
        const hasExplicit = Object.prototype.hasOwnProperty.call(explicit, field);
        const value = hasExplicit ? explicit[field] : old[field];
        if (field === 'count') {
          if (!hasExplicit) continue;
          const count = physicalCountState(value);
          if (count.status === 'valid') next.count = count.value;
          continue;
        }
        if (!hasExplicit && (!includeSavedWorkbook || text(value) === '')) continue;
        const currentValue = record[field];
        if (mode === 'compatible' && text(currentValue) !== '' && String(currentValue ?? '') !== String(value ?? '')) {
          skippedConflicts += 1;
          continue;
        }
        next[field] = value;
      }
      if (Object.keys(next).length) {
        changes[record.key] = next;
        recoveredKeys.add(record.key);
      }
      const selected = savedSeries(saved, old) ?? savedWorkbookSeries[old.key];
      const compatibleSeries = compatibleSelectedSeries(record, selected);
      if (selected !== undefined) {
        if (mode !== 'compatible' || !(selectedSeries[record.key] || []).length) selectedSeries[record.key] = compatibleSeries;
        if (compatibleSeries.length) recoveredKeys.add(record.key);
      }
    }
    return { changes, selectedSeries, recoveredRows: recoveredKeys.size, skippedConflicts };
  }

  function assignStableRecordKeys(records) {
    const groups = new Map();
    for (const record of records) {
      record.legacyKey = record.legacyKey || legacyRecordKey(record);
      const base = recordIdentityBase(record);
      const list = groups.get(base) || [];
      list.push(record);
      groups.set(base, list);
    }
    for (const [base, group] of groups) {
      group.sort((a, b) => Number(a.rowNumber) - Number(b.rowNumber));
      for (let index = 0; index < group.length; index += 1) {
        group[index].key = group.length === 1 ? base : `${base}::duplicate-${index + 1}`;
      }
    }
    return records;
  }

  function assignSeriesOwnership(records, seriesItems) {
    const rowsByMaterial = new Map();
    const rowsByPlantMaterial = new Map();
    const issues = [];
    const unresolved = [];
    const duplicateItems = new Set();
    const itemIdentity = item => `${item.rowNumber}::${item.serialCol || ''}`;
    const materialKey = value => normalized(value);
    const ownerKey = (plant, material) => `${normalized(plant)}::${materialKey(material)}`;
    const addRecordIssue = (record, issue) => {
      if (!record) return;
      record.seriesIssues.push(issue);
    };
    const addIssue = (issue, affectedRecords = []) => {
      const complete = { severity: 'warning', ...issue, recordKeys: affectedRecords.map(record => record.key) };
      issues.push(complete);
      affectedRecords.forEach(record => addRecordIssue(record, complete));
      return complete;
    };
    for (const record of records) {
      record.serials = [];
      record.alternateSerials = [];
      record.seriesIssues = [];
      const byMaterial = rowsByMaterial.get(materialKey(record.material)) || [];
      byMaterial.push(record);
      rowsByMaterial.set(materialKey(record.material), byMaterial);
      const byOwner = rowsByPlantMaterial.get(ownerKey(record.plant, record.material)) || [];
      byOwner.push(record);
      rowsByPlantMaterial.set(ownerKey(record.plant, record.material), byOwner);
    }

    const duplicates = new Map();
    for (const item of seriesItems) {
      item.selectionKey = text(item.serial);
      if (!text(item.material) || !text(item.serial)) continue;
      const key = `${materialKey(item.material)}::${normalized(item.serial)}`;
      const items = duplicates.get(key) || [];
      items.push(item);
      duplicates.set(key, items);
    }
    for (const items of duplicates.values()) {
      if (items.length < 2) continue;
      const relatedRows = items.map(item => item.rowNumber);
      const defective = materialKey(items[0].material).endsWith('DEF');
      for (const item of items) {
        if (defective) item.selectionKey = `${text(item.serial)}@@${itemIdentity(item)}`;
        else duplicateItems.add(itemIdentity(item));
        const affected = rowsByMaterial.get(materialKey(item.material)) || [];
        addIssue({
          type: defective ? 'duplicate-series-def' : 'duplicate-series',
          severity: defective ? 'warning' : 'critical',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber, relatedRows,
        }, affected);
      }
    }

    for (const item of seriesItems) {
      const missingFields = ['plant', 'location', 'material', 'description', 'serial'].filter(field => !text(item[field]));
      const missingAssociationField = missingFields.some(field => ['plant', 'location', 'material', 'serial'].includes(field));
      const materialRows = rowsByMaterial.get(materialKey(item.material)) || [];
      const candidateRows = rowsByPlantMaterial.get(ownerKey(item.plant, item.material)) || [];
      if (missingFields.length) {
        addIssue({
          type: 'incomplete-series-row',
          severity: missingAssociationField ? 'critical' : 'warning',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber, missingFields,
        }, candidateRows.length ? candidateRows : materialRows);
      }
      if (missingAssociationField || duplicateItems.has(itemIdentity(item))) {
        unresolved.push(item);
        continue;
      }
      if (!candidateRows.length) {
        addIssue({
          type: materialRows.length ? 'incompatible-plant' : 'series-without-material',
          severity: item.selected ? 'critical' : 'warning',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber,
        }, materialRows);
        unresolved.push(item);
        continue;
      }
      const exact = candidateRows.filter(record => normalized(record.stloc) === normalized(item.location));
      if (exact.length !== 1) {
        if (!exact.length) {
          for (const record of candidateRows) record.alternateSerials.push(item);
          if (candidateRows.length === 1) item.defaultRecordKey = candidateRows[0].key;
        }
        addIssue({
          type: exact.length ? 'ambiguous-series' : 'incompatible-location',
          severity: exact.length ? 'critical' : 'warning',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber, inventoryLocations: candidateRows.map(record => text(record.stloc)),
        }, exact.length ? exact : candidateRows);
        unresolved.push(item);
        continue;
      }
      exact[0].serials.push(item);
      item.defaultRecordKey = exact[0].key;
      const alternateRows = candidateRows.filter(record => record !== exact[0]);
      if (alternateRows.length) {
        alternateRows.forEach(record => record.alternateSerials.push(item));
        addIssue({
          type: 'incompatible-location',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber, inventoryLocations: alternateRows.map(record => text(record.stloc)),
        }, alternateRows);
      }
      if (normalized(item.serial) === '0' && !materialKey(item.material).endsWith('DEF')) {
        addIssue({
          type: 'suspicious-series',
          material: text(item.material), plant: text(item.plant), location: text(item.location), serial: text(item.serial),
          rowNumber: item.rowNumber,
        }, exact);
      }
    }

    for (const record of records) {
      if (!record.serialized || record.serials.length || record.alternateSerials.length) continue;
      addIssue({
        type: 'serialized-without-series',
        severity: 'warning',
        material: text(record.material), plant: text(record.plant), location: text(record.stloc), serial: '',
        inventoryRowNumber: record.rowNumber,
      }, [record]);
    }
    return { records, unresolved, issues };
  }

  function savedChange(saved, record) {
    if (!saved?.changes || !record) return {};
    return saved.changes[record.key] || saved.changes[record.legacyKey] || {};
  }

  function savedSeries(saved, record) {
    if (!saved?.selectedSeries || !record) return undefined;
    return saved.selectedSeries[record.key] || saved.selectedSeries[record.legacyKey];
  }

  function recoveryRecordMap(records) {
    const map = new Map();
    for (const record of records) {
      map.set(record.key, record);
      if (record.legacyKey) map.set(record.legacyKey, record);
    }
    return map;
  }

  function assertUniqueSeriesOwnership(records) {
    const owners = new Map();
    const logicalOwners = new Map();
    for (const record of records) {
      for (const item of record.serials || []) {
        const identity = `${item.rowNumber}::${item.serialCol}`;
        const prior = owners.get(identity);
        if (prior && prior !== record.key) {
          throw new Error(`La serie ${item.serial} está asociada a más de una fila de inventario.`);
        }
        owners.set(identity, record.key);
        const logicalIdentity = `${normalized(item.material || record.material)}::${normalized(item.serial)}`;
        const logicalPrior = logicalOwners.get(logicalIdentity);
        if (logicalPrior && logicalPrior !== identity && !normalized(item.material || record.material).endsWith('DEF')) {
          throw new Error(`La serie ${item.serial} está duplicada para el material ${item.material || record.material}.`);
        }
        logicalOwners.set(logicalIdentity, identity);
      }
    }
    return owners;
  }

  globalThis.InventoryModel = Object.freeze({
    legacyRecordKey,
    recordIdentityBase,
    sourceRecordIdentity,
    recoveryPlan,
    compatibleSelectedSeries,
    seriesSelectionKey,
    toggleSeriesSelection,
    physicalCountState,
    numericQuantity,
    seriesConsistencyReport,
    recoveryData,
    assignStableRecordKeys,
    assignSeriesOwnership,
    savedChange,
    savedSeries,
    recoveryRecordMap,
    assertUniqueSeriesOwnership,
  });
})();
