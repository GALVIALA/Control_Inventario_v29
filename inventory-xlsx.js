/* XLSX parsing and export helpers. Corporate data stays in browser memory. */
(() => {
  'use strict';

  const escapeXml = value => String(value ?? '').replace(/[&<>'"]/g, character => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&apos;', '"':'&quot;' }[character]));
  const colFromRef = reference => (String(reference).match(/[A-Z]+/i) || [''])[0].toUpperCase();
  const colToNumber = column => [...String(column).toUpperCase()].reduce((number, character) => number * 26 + character.charCodeAt(0) - 64, 0);

  function parseXml(text) {
    const document = new DOMParser().parseFromString(text, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('XML inválido dentro del Excel.');
    return document;
  }

  function textOf(node) {
    return node ? [...node.querySelectorAll('t')].map(item => item.textContent).join('') : '';
  }

  function readSharedStrings(xml) {
    return xml ? [...parseXml(xml).querySelectorAll('si')].map(textOf) : [];
  }

  function cellValue(cell, shared) {
    if (!cell) return '';
    const type = cell.getAttribute('t');
    if (type === 'inlineStr') return textOf(cell.querySelector('is'));
    const raw = cell.querySelector('v')?.textContent ?? '';
    if (type === 's') return shared[Number(raw)] ?? '';
    if (type === 'b') return raw === '1';
    return raw;
  }

  function parseSheet(xml, shared) {
    const document = parseXml(xml);
    const rows = [];
    for (const row of document.querySelectorAll('sheetData > row')) {
      const values = {};
      const cells = {};
      for (const cell of row.querySelectorAll(':scope > c')) {
        const column = colFromRef(cell.getAttribute('r'));
        values[column] = cellValue(cell, shared);
        cells[column] = { formula: Boolean(cell.querySelector('f')), style: Number(cell.getAttribute('s') || 0) };
      }
      rows.push({ rowNumber: Number(row.getAttribute('r')), values, cells });
    }
    return rows;
  }

  function analyzeStyles(xml) {
    const result = { greenStyleIds: new Set(), baseByGreen: {}, greenByBase: {} };
    if (!xml) return result;
    const document = parseXml(xml);
    const fills = [...document.querySelectorAll('fills > fill')];
    const greenFills = new Set();
    fills.forEach((fill, index) => {
      const rgb = (fill.querySelector('fgColor')?.getAttribute('rgb') || '').toUpperCase();
      if (rgb === 'FF00B050' || rgb === 'FF20A66D') greenFills.add(index);
    });
    const xfs = [...document.querySelectorAll('cellXfs > xf')];
    const signature = xf => ['numFmtId','fontId','borderId','xfId','quotePrefix'].map(attribute => xf.getAttribute(attribute) || '').join('|') + `|${xf.innerHTML || ''}`;
    xfs.forEach((xf, index) => {
      if (greenFills.has(Number(xf.getAttribute('fillId') || 0))) result.greenStyleIds.add(index);
    });
    for (const green of result.greenStyleIds) {
      const greenSignature = signature(xfs[green]);
      const base = xfs.findIndex((xf, index) => index !== green && !greenFills.has(Number(xf.getAttribute('fillId') || 0)) && signature(xf) === greenSignature);
      if (base >= 0) {
        result.baseByGreen[green] = base;
        result.greenByBase[base] = green;
      }
    }
    return result;
  }

  function appendGreenStyle(stylesXml, baseStyleId) {
    if (!stylesXml) throw new Error('El libro no contiene xl/styles.xml.');
    const fills = stylesXml.match(/<fills\b([^>]*)>([\s\S]*?)<\/fills>/i);
    const cellXfs = stylesXml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/i);
    if (!fills || !cellXfs) throw new Error('La tabla de estilos del libro no es compatible.');

    const fillEntries = fills[2].match(/<fill\b[^>]*?\/>|<fill\b[^>]*?>[\s\S]*?<\/fill>/gi) || [];
    const xfEntries = cellXfs[2].match(/<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/gi) || [];
    const base = xfEntries[Number(baseStyleId)];
    if (!base) throw new Error(`No existe el estilo base ${baseStyleId} de la serie.`);

    const greenFillId = fillEntries.length;
    const greenStyleId = xfEntries.length;
    const greenFill = '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor indexed="64"/></patternFill></fill>';
    const greenXf = base.replace(/^<xf\b([^>]*?)(\/>|>)/i, (match, attributes, close) => {
      let updated = setXmlAttribute(attributes, 'fillId', greenFillId);
      updated = setXmlAttribute(updated, 'applyFill', 1);
      return `<xf${updated}${close}`;
    });

    const updatedFills = `<fills${setXmlAttribute(fills[1], 'count', greenFillId + 1)}>${fills[2]}${greenFill}</fills>`;
    const updatedCellXfs = `<cellXfs${setXmlAttribute(cellXfs[1], 'count', greenStyleId + 1)}>${cellXfs[2]}${greenXf}</cellXfs>`;
    return {
      xml: stylesXml.replace(fills[0], updatedFills).replace(cellXfs[0], updatedCellXfs),
      fillId: greenFillId,
      styleId: greenStyleId,
    };
  }

  function seriesStyleState(styleInfo, styleId) {
    if (styleInfo?.greenStyleIds?.has(styleId)) return 'selected';
    if (Object.prototype.hasOwnProperty.call(styleInfo?.greenByBase || {}, styleId)) return 'unselected';
    return 'unknown';
  }

  const REQUIRED_SHEET_NAMES = Object.freeze({ inventory: 'Inventario', series: 'Series' });
  const WORKBOOK_SCHEMA = Object.freeze({
    inventory: Object.freeze({
      expectedHeaders: Object.freeze([
        'INVENTORY DOC', 'DOC ITEM', 'PLANT', 'ST LOC', 'MATERIAL', 'MATERIAL DESCRIPTION', 'BATCH', 'UM', 'STOCK TYPE',
        'SAP QTY', 'COUNT QTY', 'COUNT BIN', 'SAP BIN', 'SERIALIZED', 'UNIT VALUE', 'CRCY', 'COMMENTS', 'QTY DIFF',
        'SAP VALUE', 'COUNT VALUE', 'DIFF VALUE',
      ]),
      fieldHeaders: Object.freeze({
        plant: 'PLANT', material: 'MATERIAL', description: 'MATERIAL DESCRIPTION', batch: 'BATCH', stloc: 'ST LOC',
        stockType: 'STOCK TYPE', sapQty: 'SAP QTY', sapLocations: 'SAP BIN', unitValue: 'UNIT VALUE', unitMeasure: 'UM',
        count: 'COUNT QTY', currentLocation: 'COUNT BIN', serialized: 'SERIALIZED', comments: 'COMMENTS',
      }),
    }),
    series: Object.freeze({
      expectedHeaders: Object.freeze(['PLANT', 'ST LOC', 'MATERIAL', 'MATERIAL DESCRIPTION', 'SERIAL NUMBER']),
      fieldHeaders: Object.freeze({
        seriesPlant: 'PLANT', seriesLocation: 'ST LOC', seriesMaterial: 'MATERIAL',
        seriesDescription: 'MATERIAL DESCRIPTION', serialNumber: 'SERIAL NUMBER',
      }),
    }),
  });
  const FIELD_LABELS = Object.freeze({
    material: 'Material',
    sapQty: 'SAP QTY',
    sapLocations: 'SAP BIN',
    count: 'COUNT QTY',
    currentLocation: 'COUNT BIN',
    comments: 'COMMENTS',
    seriesPlant: 'Plant de series',
    seriesMaterial: 'Material de series',
    seriesLocation: 'ubicación de series',
    seriesDescription: 'descripción de series',
    serialNumber: 'número de serie',
  });

  const fieldLabel = field => FIELD_LABELS[field] || field;

  function assertWritableColumns(map) {
    const writable = ['count', 'currentLocation', 'comments'];
    const missing = writable.filter(field => !map[field]);
    if (missing.length) throw new Error(`No se puede exportar: faltan las columnas ${missing.map(fieldLabel).join(', ')}.`);
    const seen = new Map();
    for (const field of writable) {
      const previous = seen.get(map[field]);
      if (previous) throw new Error(`No se puede exportar: la columna ${map[field]} está asignada tanto a ${fieldLabel(previous)} como a ${fieldLabel(field)}.`);
      seen.set(map[field], field);
    }
  }

  function selectRequiredSheets(sheets, requiredSheetNames = REQUIRED_SHEET_NAMES) {
    const selected = {};
    const missing = [];
    for (const [kind, configuredNames] of Object.entries(requiredSheetNames)) {
      const names = Array.isArray(configuredNames) ? configuredNames : [configuredNames];
      const matches = sheets.filter(sheet => names.includes(sheet.name));
      const description = names.map(name => `"${name}"`).join(' o ');
      if (!matches.length) missing.push(description);
      else if (matches.length > 1) throw new Error(`No se puede importar: existe más de una pestaña compatible con ${description}.`);
      else selected[kind] = matches[0];
    }
    if (missing.length) throw new Error(`No se puede importar: ${missing.length === 1 ? 'falta la pestaña obligatoria' : 'faltan las pestañas obligatorias'} ${missing.join(' y ')}.`);
    return { ...selected, sheets: [selected.inventory, selected.series] };
  }

  const normalizeSchemaHeader = value => String(value ?? '').trim().toUpperCase().replace(/\s+/g, ' ');

  function detectSchemaHeader(rows, schema, sheetName) {
    const expected = schema.expectedHeaders.map(normalizeSchemaHeader);
    const expectedSet = new Set(expected);
    let best = null;
    for (const row of rows.slice(0, 25)) {
      const entries = Object.entries(row.values).map(([column, value]) => [column, normalizeSchemaHeader(value)]).filter(([, value]) => value);
      const overlap = entries.filter(([, value]) => expectedSet.has(value)).length;
      if (!best || overlap > best.overlap) best = { rowNumber: row.rowNumber, entries, overlap };
    }
    if (!best || !best.overlap) throw new Error(`No se puede importar: no encontré los encabezados obligatorios en la pestaña "${sheetName}".`);
    const byValue = new Map();
    for (const [column, value] of best.entries) {
      const columns = byValue.get(value) || [];
      columns.push(column);
      byValue.set(value, columns);
    }
    const missing = expected.filter(value => !byValue.has(value));
    const unexpected = [...byValue.keys()].filter(value => !expectedSet.has(value));
    const duplicates = expected.filter(value => (byValue.get(value) || []).length > 1);
    if (missing.length || unexpected.length || duplicates.length || best.entries.length !== expected.length) {
      const details = [];
      if (missing.length) details.push(`faltan ${missing.join(', ')}`);
      if (unexpected.length) details.push(`sobran o cambiaron ${unexpected.join(', ')}`);
      if (duplicates.length) details.push(`están duplicados ${duplicates.join(', ')}`);
      throw new Error(`No se puede importar: los encabezados de la pestaña "${sheetName}" no corresponden a la plantilla actual (${details.join('; ')}).`);
    }
    const map = {};
    for (const [field, header] of Object.entries(schema.fieldHeaders)) map[field] = byValue.get(normalizeSchemaHeader(header))[0];
    return { rowNumber: best.rowNumber, map, headers: Object.fromEntries(best.entries), ambiguities: [], score: expected.length };
  }

  async function loadWorkbook(bytes, fileName, options) {
    const { JSZip, buildRecords, requiredSheetNames = REQUIRED_SHEET_NAMES, workbookSchema = WORKBOOK_SCHEMA } = options;
    const zip = await JSZip.loadAsync(bytes);
    const read = async path => zip.file(path) ? zip.file(path).async('text') : '';
    const workbookXml = await read('xl/workbook.xml');
    const relationshipsXml = await read('xl/_rels/workbook.xml.rels');
    if (!workbookXml || !relationshipsXml) throw new Error('El archivo no contiene una estructura XLSX válida.');
    const workbookDocument = parseXml(workbookXml);
    const relationshipsDocument = parseXml(relationshipsXml);
    const relationships = {};
    for (const relationship of relationshipsDocument.querySelectorAll('Relationship')) relationships[relationship.getAttribute('Id')] = relationship.getAttribute('Target');
    const shared = readSharedStrings(await read('xl/sharedStrings.xml'));
    const stylesXml = await read('xl/styles.xml');
    const styleInfo = analyzeStyles(stylesXml);
    const sheetDescriptors = [];
    for (const sheet of workbookDocument.querySelectorAll('sheets > sheet')) {
      const relationshipId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let target = relationships[relationshipId] || '';
      target = target.replace(/^\//, '').replace(/^xl\//, '');
      const path = `xl/${target}`.replace(/\/+/g, '/');
      sheetDescriptors.push({ name: sheet.getAttribute('name'), path });
    }
    const selected = selectRequiredSheets(sheetDescriptors, requiredSheetNames);
    const loadSelectedSheet = async descriptor => {
      const xml = await read(descriptor.path);
      if (!xml) throw new Error(`No se puede importar: no fue posible leer la pestaña "${descriptor.name}".`);
      return { ...descriptor, xml, rows: parseSheet(xml, shared) };
    };
    const inventorySheet = await loadSelectedSheet(selected.inventory);
    const seriesBase = await loadSelectedSheet(selected.series);
    const inventoryHeader = detectSchemaHeader(inventorySheet.rows, workbookSchema.inventory, selected.inventory.name);
    const inventory = { ...inventorySheet, header: inventoryHeader };
    const seriesHeader = detectSchemaHeader(seriesBase.rows, workbookSchema.series, selected.series.name);
    const seriesSheet = { ...seriesBase, header: seriesHeader };
    const sheets = [inventory, seriesSheet];
    const workbook = { zip, fileName, sheets, inventory, seriesSheet, bytes, stylesXml, styleInfo };
    workbook.records = buildRecords(workbook);
    return workbook;
  }

  function replaceCell(sheetXml, rowNumber, column, value, numeric = false) {
    const reference = `${column}${rowNumber}`;
    const rowExpression = new RegExp(`(<row\\b[^>]*\\br=["']${rowNumber}["'][^>]*>)([\\s\\S]*?)(</row>)`, 'i');
    const row = sheetXml.match(rowExpression);
    if (!row) throw new Error(`No existe la fila ${rowNumber} en la hoja.`);
    const cellExpression = new RegExp(`<c\\b([^>]*\\br=["']${reference}["'][^>]*)\\s*/>|<c\\b([^>]*\\br=["']${reference}["'][^>]*)>([\\s\\S]*?)</c>`, 'i');
    const cell = row[2].match(cellExpression);
    if (cell && /<f\b/i.test(cell[3] || '')) {
      const cached = (cell[3].match(/<v\b[^>]*>([\s\S]*?)<\/v>/i) || [])[1] ?? '';
      const target = String(value ?? '');
      const same = numeric && target !== '' && cached !== '' ? Number(target) === Number(cached) : target === cached;
      if (same) return sheetXml;
      throw new Error(`La celda ${reference} contiene una fórmula y no se modificó.`);
    }
    let attributes = (cell?.[1] || cell?.[2] || ` r="${reference}"`).replace(/\s*\/$/, '').replace(/\s+t=["'][^"']*["']/ig, '');
    const content = numeric ? `<v>${Number(value)}</v>` : `<is><t xml:space="preserve">${escapeXml(value)}</t></is>`;
    if (!numeric) attributes += ' t="inlineStr"';
    const replacement = `<c${attributes}>${content}</c>`;
    let body;
    if (cell) body = row[2].replace(cell[0], replacement);
    else {
      const targetColumn = colToNumber(column);
      const followingCell = [...row[2].matchAll(/<c\b[^>]*\br=["']([A-Z]+)\d+["'][^>]*(?:\s*\/>|>[\s\S]*?<\/c>)/gi)]
        .find(match => colToNumber(match[1]) > targetColumn);
      body = followingCell
        ? row[2].slice(0, followingCell.index) + replacement + row[2].slice(followingCell.index)
        : row[2] + replacement;
    }
    return sheetXml.replace(rowExpression, `${row[1]}${body}${row[3]}`);
  }

  function setCellStyle(sheetXml, rowNumber, column, styleId) {
    const reference = `${column}${rowNumber}`;
    const rowExpression = new RegExp(`(<row\\b[^>]*\\br=["']${rowNumber}["'][^>]*>)([\\s\\S]*?)(</row>)`, 'i');
    const row = sheetXml.match(rowExpression);
    if (!row) throw new Error(`No existe la fila ${rowNumber} en Series.`);
    const cellExpression = new RegExp(`<c\\b([^>]*\\br=["']${reference}["'][^>]*)\\s*/>|<c\\b([^>]*\\br=["']${reference}["'][^>]*)>([\\s\\S]*?)</c>`, 'i');
    const cell = row[2].match(cellExpression);
    if (!cell) throw new Error(`No existe la celda ${reference} en Series.`);
    const attributes = (cell[1] || cell[2]).replace(/\s*\/$/, '').replace(/\s+s=["'][^"']*["']/i, '') + ` s="${styleId}"`;
    const replacement = cell[1] ? `<c${attributes}/>` : `<c${attributes}>${cell[3] || ''}</c>`;
    return sheetXml.replace(rowExpression, `${row[1]}${row[2].replace(cell[0], replacement)}${row[3]}`);
  }

  function stripCalculationChainReference(xml, kind) {
    if (!xml) return xml;
    if (kind === 'relationships') {
      return xml.replace(/<Relationship\b(?=[^>]*\bType=["'][^"']*\/calcChain["'])[^>]*\/>/ig, '');
    }
    if (kind === 'contentTypes') {
      return xml.replace(/<Override\b(?=[^>]*\bPartName=["']\/xl\/calcChain\.xml["'])[^>]*\/>/ig, '');
    }
    return xml;
  }

  function setXmlAttribute(attributes, name, value) {
    const expression = new RegExp(`\\s${name}=["'][^"']*["']`, 'i');
    return expression.test(attributes)
      ? attributes.replace(expression, ` ${name}="${value}"`)
      : `${attributes} ${name}="${value}"`;
  }

  function markWorkbookForFullCalculation(xml) {
    if (!xml) return xml;
    const update = attributes => {
      let next = attributes.replace(/\s*\/$/, '');
      next = setXmlAttribute(next, 'calcMode', 'auto');
      next = setXmlAttribute(next, 'fullCalcOnLoad', '1');
      next = setXmlAttribute(next, 'forceFullCalc', '1');
      next = setXmlAttribute(next, 'calcId', '0');
      return next;
    };
    if (/<calcPr\b/i.test(xml)) {
      if (/<calcPr\b[^>]*\/>/i.test(xml)) {
        return xml.replace(/<calcPr\b([^>]*)\/>/i, (_, attributes) => `<calcPr${update(attributes)}/>`);
      }
      return xml.replace(/<calcPr\b([^>]*)>/i, (_, attributes) => `<calcPr${update(attributes)}>`);
    }
    return xml.replace(/<\/workbook>/i, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcId="0"/></workbook>');
  }

  async function prepareFullRecalculation(zip) {
    const relationshipsPath = 'xl/_rels/workbook.xml.rels';
    const contentTypesPath = '[Content_Types].xml';
    const workbookPath = 'xl/workbook.xml';
    const relationships = await zip.file(relationshipsPath)?.async('text');
    const contentTypes = await zip.file(contentTypesPath)?.async('text');
    const workbookXml = await zip.file(workbookPath)?.async('text');
    if (relationships != null) zip.file(relationshipsPath, stripCalculationChainReference(relationships, 'relationships'));
    if (contentTypes != null) zip.file(contentTypesPath, stripCalculationChainReference(contentTypes, 'contentTypes'));
    if (workbookXml != null) zip.file(workbookPath, markWorkbookForFullCalculation(workbookXml));
    if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml');
  }

  async function buildExport({ workbook, records, changes, selectedSeries, model, outputFileName = '', FileType = File }) {
    const map = workbook.inventory.header.map;
    assertWritableColumns(map);
    let inventoryXml = workbook.inventory.xml;
    let calculationInputsChanged = false;
    for (const record of records) {
      const change = changes[record.key];
      if (!change) continue;
      if (Object.prototype.hasOwnProperty.call(change, 'count')) {
        const value = change.count;
        const previous = inventoryXml;
        inventoryXml = replaceCell(inventoryXml, record.rowNumber, map.count, value, value !== '' && Number.isFinite(Number(value)));
        calculationInputsChanged ||= inventoryXml !== previous;
      }
      if (Object.prototype.hasOwnProperty.call(change, 'currentLocation')) {
        const previous = inventoryXml;
        inventoryXml = replaceCell(inventoryXml, record.rowNumber, map.currentLocation, change.currentLocation);
        calculationInputsChanged ||= inventoryXml !== previous;
      }
      if (Object.prototype.hasOwnProperty.call(change, 'comments')) {
        const previous = inventoryXml;
        inventoryXml = replaceCell(inventoryXml, record.rowNumber, map.comments, change.comments);
        calculationInputsChanged ||= inventoryXml !== previous;
      }
    }
    workbook.zip.file(workbook.inventory.path, inventoryXml);
    if (calculationInputsChanged) await prepareFullRecalculation(workbook.zip);
    if (workbook.seriesSheet) {
      model.assertUniqueSeriesOwnership(records);
      let seriesXml = workbook.seriesSheet.xml;
      let stylesXml = await workbook.zip.file('xl/styles.xml')?.async('text');
      const greenStyleIds = new Set(workbook.styleInfo.greenStyleIds);
      const greenByBase = { ...workbook.styleInfo.greenByBase };
      const baseByGreen = { ...workbook.styleInfo.baseByGreen };
      let stylesChanged = false;
      const normalized = value => String(value ?? '').trim().toUpperCase().replace(/\s+/g, ' ');
      const selected = new Set();
      for (const record of records) {
        for (const key of selectedSeries[record.key] || []) selected.add(`${normalized(record.material)}::${key}`);
      }
      const processed = new Set();
      for (const record of records) {
        for (const series of [...(record.serials || []), ...(record.alternateSerials || [])]) {
          const identity = `${series.rowNumber}::${series.serialCol}`;
          if (processed.has(identity)) continue;
          processed.add(identity);
          let style = series.originalStyle;
          if (selected.has(`${normalized(series.material || record.material)}::${model.seriesSelectionKey(series)}`)) {
            if (!greenStyleIds.has(style)) {
              let greenStyle = greenByBase[style];
              if (greenStyle == null) {
                const created = appendGreenStyle(stylesXml, style);
                stylesXml = created.xml;
                greenStyle = created.styleId;
                greenStyleIds.add(greenStyle);
                greenByBase[style] = greenStyle;
                baseByGreen[greenStyle] = style;
                stylesChanged = true;
              }
              style = greenStyle;
            }
          } else if (greenStyleIds.has(style)) style = baseByGreen[style];
          if (style != null && style !== series.originalStyle) seriesXml = setCellStyle(seriesXml, series.rowNumber, series.serialCol, style);
        }
      }
      if (stylesChanged) workbook.zip.file('xl/styles.xml', stylesXml);
      workbook.zip.file(workbook.seriesSheet.path, seriesXml);
    }
    const blob = await workbook.zip.generateAsync({ type:'blob', compression:'DEFLATE', compressionOptions:{ level:6 }, mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const fallbackBase = workbook.fileName.replace(/\.xlsx$/i, '').replace(/(?:_actualizado_\d{4}-\d{2}-\d{2}(?:-\d{2}-\d{2})?)+$/i, '').replace(/\bInvetario\b/gi, 'Inventario');
    return new FileType([blob], outputFileName || `${fallbackBase}_actualizado.xlsx`, { type: blob.type });
  }

  globalThis.InventoryXlsx = Object.freeze({
    WORKBOOK_SCHEMA,
    escapeXml,
    colFromRef,
    parseXml,
    parseSheet,
    analyzeStyles,
    appendGreenStyle,
    seriesStyleState,
    selectRequiredSheets,
    detectSchemaHeader,
    assertWritableColumns,
    loadWorkbook,
    replaceCell,
    setCellStyle,
    stripCalculationChainReference,
    markWorkbookForFullCalculation,
    prepareFullRecalculation,
    buildExport,
  });
})();
