load('inventory-xlsx.js');

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

function assertUnique(map) {
  const columns = Object.values(map);
  equal(new Set(columns).size, columns.length, 'cada columna debe asignarse una sola vez');
}

const columns = Array.from({ length: 21 }, (_, index) => {
  let number = index + 1;
  let column = '';
  while (number) {
    number -= 1;
    column = String.fromCharCode(65 + (number % 26)) + column;
    number = Math.floor(number / 26);
  }
  return column;
});

const inventoryValues = Object.fromEntries(InventoryXlsx.WORKBOOK_SCHEMA.inventory.expectedHeaders.map((header, index) => [columns[index], header]));
const seriesValues = { A: 'PLANT', B: 'ST LOC', C: 'MATERIAL', D: 'MATERIAL DESCRIPTION', E: 'SERIAL NUMBER' };

function detectInventory(values = inventoryValues) {
  return InventoryXlsx.detectSchemaHeader([{ rowNumber: 1, values }], InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario');
}

function expectSchemaError(values, schema, sheetName, message) {
  let error = null;
  try {
    InventoryXlsx.detectSchemaHeader([{ rowNumber: 1, values }], schema, sheetName);
  } catch (caught) {
    error = caught;
  }
  assert(error && /plantilla actual/.test(error.message), message);
}

test('acepta exclusivamente los 21 encabezados nuevos de Inventario', () => {
  const header = detectInventory();
  equal(Object.keys(header.headers).length, 21, 'cantidad exacta de encabezados');
  equal(header.map.plant, 'C', 'PLANT');
  equal(header.map.stloc, 'D', 'ST LOC');
  equal(header.map.material, 'E', 'MATERIAL');
  equal(header.map.sapQty, 'J', 'SAP QTY');
  equal(header.map.count, 'K', 'COUNT QTY');
  equal(header.map.currentLocation, 'L', 'COUNT BIN');
  equal(header.map.sapLocations, 'M', 'SAP BIN');
  equal(header.map.comments, 'Q', 'COMMENTS');
  assertUnique(header.map);
});

test('rechaza el esquema antiguo de 20 encabezados', () => {
  const old = ['Phys. Inve Doc', 'Item', 'Plant', 'St Loc', 'Material', 'Material Desc', 'Batch', 'BUN', 'Stock type', 'SAP', 'Quantity', 'Location', 'Ubicación SAP', 'Serie?', 'Unit Value', 'Crcy', 'Comment', 'Total Value', 'Total Count', 'Difference'];
  const values = Object.fromEntries(old.map((header, index) => [columns[index], header]));
  expectSchemaError(values, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'el formato anterior debe quedar fuera');
});

test('rechaza alias antiguos aunque los demás encabezados coincidan', () => {
  const values = { ...inventoryValues, J: 'SAP', K: 'Quantity', L: 'Location', Q: 'Comment' };
  expectSchemaError(values, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'no debe aceptar alias antiguos');
});

test('rechaza variaciones con guiones que no pertenecen a la plantilla', () => {
  expectSchemaError({ ...inventoryValues, J: 'SAP-QTY' }, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'SAP-QTY no es SAP QTY');
  expectSchemaError({ ...inventoryValues, K: 'COUNT_QTY' }, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'COUNT_QTY no es COUNT QTY');
});

test('rechaza encabezados faltantes, extra o duplicados', () => {
  const missing = { ...inventoryValues };
  delete missing.U;
  expectSchemaError(missing, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'falta DIFF VALUE');
  expectSchemaError({ ...inventoryValues, V: 'EXTRA' }, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'sobra EXTRA');
  expectSchemaError({ ...inventoryValues, U: 'SAP QTY' }, InventoryXlsx.WORKBOOK_SCHEMA.inventory, 'Inventario', 'SAP QTY duplicado');
});

test('selecciona exclusivamente Inventario y Series', () => {
  const sheets = [
    { name: 'Interna 1', path: 'sheet1.xml' },
    { name: 'Inventario', path: 'sheet2.xml' },
    { name: 'Interna 2', path: 'sheet3.xml' },
    { name: 'Series', path: 'sheet4.xml' },
    { name: 'Interna 3', path: 'sheet5.xml' },
    { name: 'Interna 4', path: 'sheet6.xml' },
  ];
  const selected = InventoryXlsx.selectRequiredSheets(sheets);
  equal(selected.inventory.name, 'Inventario', 'hoja principal');
  equal(selected.series.name, 'Series', 'hoja de series');
  equal(selected.sheets.length, 2, 'cantidad de hojas procesables');
  assert(selected.sheets.every(sheet => sheet.name === 'Inventario' || sheet.name === 'Series'), 'las hojas internas deben quedar excluidas');
  equal(sheets.length, 6, 'el libro original conserva todas sus hojas');
  equal(sheets[0].name, 'Interna 1', 'la selección no modifica las hojas internas');
});

test('exige los nombres exactos Inventario y Series', () => {
  for (const sheets of [
    [{ name: 'Inventario Parque Demo' }, { name: 'Series' }],
    [{ name: 'Inventario' }, { name: 'Seriales' }],
  ]) {
    let error = null;
    try {
      InventoryXlsx.selectRequiredSheets(sheets);
    } catch (caught) {
      error = caught;
    }
    assert(error && /pestaña(?:s)? obligatoria/.test(error.message), 'debe informar la pestaña obligatoria ausente');
  }
});

test('acepta exclusivamente los cinco encabezados nuevos de Series', () => {
  const header = InventoryXlsx.detectSchemaHeader([{ rowNumber: 1, values: seriesValues }], InventoryXlsx.WORKBOOK_SCHEMA.series, 'Series');
  equal(header.map.seriesPlant, 'A', 'Plant de series');
  equal(header.map.seriesLocation, 'B', 'Location de series');
  equal(header.map.seriesMaterial, 'C', 'Material de series');
  equal(header.map.seriesDescription, 'D', 'Description de series');
  equal(header.map.serialNumber, 'E', 'Serial Number');
  assertUnique(header.map);
});

test('rechaza alias y columnas extra en Series', () => {
  expectSchemaError({ ...seriesValues, E: 'Serial No' }, InventoryXlsx.WORKBOOK_SCHEMA.series, 'Series', 'alias antiguo de serie');
  expectSchemaError({ ...seriesValues, F: 'Comments' }, InventoryXlsx.WORKBOOK_SCHEMA.series, 'Series', 'columna extra en Series');
});

test('bloquea exportaciones con columnas de salida faltantes', () => {
  let error = null;
  try {
    InventoryXlsx.assertWritableColumns({ count: 'C', comments: 'F' });
  } catch (caught) {
    error = caught;
  }
  assert(error && /COUNT BIN/.test(error.message), 'debe nombrar claramente la columna faltante');
});

test('bloquea exportaciones si una columna se asigna a dos campos', () => {
  let error = null;
  try {
    InventoryXlsx.assertWritableColumns({ count: 'C', currentLocation: 'C', comments: 'F' });
  } catch (caught) {
    error = caught;
  }
  assert(error && /asignada tanto/.test(error.message), 'debe informar la asignación duplicada');
});

test('inserta celdas nuevas en orden sin alterar ni duplicar las existentes', () => {
  const material = '<c r="A2" s="4" t="inlineStr"><is><t>001234</t></is></c>';
  const formula = '<c r="C2" s="7"><f>1+1</f><v>2</v></c>';
  const later = '<c r="AA2" custom="intacto" t="inlineStr"><is><t>FINAL</t></is></c>';
  const xml = `<worksheet><sheetData><row r="2">${material}${formula}${later}</row></sheetData></worksheet>`;
  let updated = InventoryXlsx.replaceCell(xml, 2, 'D', '0', true, true);
  updated = InventoryXlsx.replaceCell(updated, 2, 'Z', '', false, true);
  updated = InventoryXlsx.replaceCell(updated, 2, 'D', '0', true, true);
  const references = [...updated.matchAll(/<c\b[^>]*\br="([A-Z]+)2"/g)].map(match => match[1]);
  equal(references.join(','), 'A,C,D,Z,AA', 'orden ascendente de columnas');
  equal((updated.match(/\br="D2"/g) || []).length, 1, 'celda sin duplicación tras exportaciones consecutivas');
  assert(updated.includes(material), 'material con ceros iniciales intacto');
  assert(updated.includes(formula), 'fórmula, estilo y valor existentes intactos');
  assert(updated.includes(later), 'atributos y valor de la celda posterior intactos');
  assert(updated.includes('<c r="D2"><v>0</v></c>'), 'valor cero conservado');
  assert(updated.includes('<c r="Z2" t="inlineStr"><is><t xml:space="preserve"></t></is></c>'), 'campo vacío conservado');
});

test('cambiar el color de una serie conserva su Location original', () => {
  const xml = '<worksheet><sheetData><row r="7"><c r="B7" t="inlineStr"><is><t>6066</t></is></c><c r="E7" s="3" t="inlineStr"><is><t>LM0012</t></is></c></row></sheetData></worksheet>';
  const updated = InventoryXlsx.setCellStyle(xml, 7, 'E', 38);
  assert(updated.includes('<c r="B7" t="inlineStr"><is><t>6066</t></is></c>'), 'Location debe permanecer intacta');
  assert(updated.includes('<c r="E7" t="inlineStr" s="38">') || updated.includes('<c r="E7" s="38" t="inlineStr">'), 'solo debe cambiar el estilo de Serial Number');
});

test('distingue estilos de serie seleccionados, desmarcados y desconocidos', () => {
  const styleInfo = { greenStyleIds: new Set([38]), greenByBase: { 3: 38 }, baseByGreen: { 38: 3 } };
  equal(InventoryXlsx.seriesStyleState(styleInfo, 38), 'selected', 'verde reconocido');
  equal(InventoryXlsx.seriesStyleState(styleInfo, 3), 'unselected', 'base reconocida');
  equal(InventoryXlsx.seriesStyleState(styleInfo, 99), 'unknown', 'formato desconocido');
});

test('crea un estilo verde compatible cuando el libro no contiene uno', () => {
  const styles = '<?xml version="1.0"?><styleSheet><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border/><border><left style="thin"/></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="2" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs></styleSheet>';
  const created = InventoryXlsx.appendGreenStyle(styles, 1);
  equal(created.fillId, 2, 'nuevo relleno al final');
  equal(created.styleId, 2, 'nuevo estilo al final');
  assert(created.xml.includes('<fills count="3">'), 'actualiza el total de rellenos');
  assert(created.xml.includes('<cellXfs count="3">'), 'actualiza el total de estilos');
  assert(created.xml.includes('<fgColor rgb="FF00B050"/>'), 'usa el verde interno compatible');
  assert(created.xml.includes('<xf numFmtId="49" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf>'), 'clona todos los atributos y la alineación del estilo base');

  const verified = '<c r="E7" s="1" t="inlineStr"><is><t>LM0012</t></is></c>';
  const untouched = '<c r="E8" s="1" t="inlineStr"><is><t>LM0013</t></is></c>';
  const sheet = `<worksheet><sheetData><row r="7">${verified}</row><row r="8">${untouched}</row></sheetData></worksheet>`;
  const updated = InventoryXlsx.setCellStyle(sheet, 7, 'E', created.styleId);
  assert(updated.includes('<c r="E7" t="inlineStr" s="2">') || updated.includes('<c r="E7" s="2" t="inlineStr">'), 'marca solo la serie verificada');
  assert(updated.includes(untouched), 'la serie no verificada conserva exactamente su formato');
});

test('nunca reemplaza una fórmula original por un valor de conteo', () => {
  const xml = '<worksheet><sheetData><row r="2"><c r="H2"><f>SUM(A2:B2)</f><v>3</v></c></row></sheetData></worksheet>';
  let error = null;
  try {
    InventoryXlsx.replaceCell(xml, 2, 'H', '4', true);
  } catch (caught) {
    error = caught;
  }
  assert(error && /contiene una fórmula/.test(error.message), 'debe proteger la fórmula de H2');
  assert(xml.includes('<f>SUM(A2:B2)</f><v>3</v>'), 'la fórmula fuente permanece intacta');
});

test('marca el libro para recálculo completo sin modificar fórmulas', () => {
  const xml = '<workbook><sheets/><calcPr calcId="191029" calcMode="manual"/><definedNames><definedName name="Formula">Inventario!$A$1</definedName></definedNames></workbook>';
  const updated = InventoryXlsx.markWorkbookForFullCalculation(xml);
  assert(updated.includes('calcMode="auto"'), 'cálculo automático');
  assert(updated.includes('fullCalcOnLoad="1"'), 'recálculo completo al abrir');
  assert(updated.includes('forceFullCalc="1"'), 'recálculo forzado');
  assert(updated.includes('calcId="0"'), 'identificador de cálculo invalidado');
  equal((updated.match(/<calcPr\b/g) || []).length, 1, 'una sola configuración de cálculo');
  assert(/<calcPr\b[^>]*\/>/.test(updated), 'calcPr debe conservar un cierre XML válido');
  assert(updated.includes('<definedName name="Formula">Inventario!$A$1</definedName>'), 'contenido original intacto');
});

test('conserva el cierre explícito de calcPr', () => {
  const xml = '<workbook><calcPr calcId="191029"></calcPr><sheets/></workbook>';
  const updated = InventoryXlsx.markWorkbookForFullCalculation(xml);
  assert(/<calcPr\b[^>]*><\/calcPr>/.test(updated), 'calcPr con cierre explícito debe seguir bien formado');
  assert(updated.includes('fullCalcOnLoad="1"'), 'recálculo completo al abrir');
});

test('elimina únicamente las referencias a calcChain', () => {
  const relationships = '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>';
  const contentTypes = '<Types><Override PartName="/xl/workbook.xml" ContentType="book"/><Override PartName="/xl/calcChain.xml" ContentType="chain"/></Types>';
  const nextRelationships = InventoryXlsx.stripCalculationChainReference(relationships, 'relationships');
  const nextContentTypes = InventoryXlsx.stripCalculationChainReference(contentTypes, 'contentTypes');
  assert(nextRelationships.includes('worksheets/sheet1.xml'), 'relación de hoja intacta');
  assert(!nextRelationships.includes('calcChain'), 'relación de cadena eliminada');
  assert(nextContentTypes.includes('/xl/workbook.xml'), 'tipo del libro intacto');
  assert(!nextContentTypes.includes('calcChain'), 'tipo de cadena eliminado');
});

test('agrega la política de recálculo cuando calcPr no existe', () => {
  const updated = InventoryXlsx.markWorkbookForFullCalculation('<workbook><sheets/></workbook>');
  assert(updated.includes('<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcId="0"/>'), 'calcPr nuevo');
});

print(`${passed} pruebas aprobadas`);
