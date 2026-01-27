/**************
 * Write items
 **************/

function writeItemsToSheet(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  clearItemsTable(sheet);

  if (!items || items.length === 0) {
    PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, '0');
    updateHeaderCompletionCell_(sheet);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const numItems = items.length;

  // Copy template row formatting down
  const templateRange = sheet.getRange(START_ROW, 1, 1, lastCol);
  if (numItems > 1) {
    templateRange.copyTo(
      sheet.getRange(START_ROW + 1, 1, numItems - 1, lastCol),
      { contentsOnly: false }
    );
  }

  items.forEach((item, index) => {
    const row = START_ROW + index;
    const fields = item.fields || item;

    // Fill mapped fields
    Object.keys(FIELD_COLUMN_MAP).forEach((fieldName) => {
      if (!(fieldName in fields)) return;

      const colIndex = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);
      let value = fields[fieldName];

      // Booleans must be Si/No
      if (typeof value === 'boolean') {
        value = value ? 'Si' : 'No';
      } else if (typeof value === 'string') {
        const v = value.trim().toLowerCase();
        if (v === 'true') value = 'Si';
        else if (v === 'false') value = 'No';
      }

      sheet.getRange(row, colIndex).setValue(value);
    });

    // Warnings
    const warnings = normalizeWarnings_(getItemWarnings_(item));
    const warningFieldNames = applyWarnings_(sheet, row, warnings);

    // Manual highlights (only if empty)
    const missingManualFields = highlightManualFields_(sheet, row);

    // Store per-row metadata in DocumentProperties (no debug column)
    setRowMeta_(row, {
      warning_fields: warningFieldNames,
      warning_details: warnings
    });

    // Checkbox tick cell (summary/action lives here)
    setupTickCell_(sheet, row, warnings, missingManualFields);
  });

  PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, String(numItems));
  updateHeaderCompletionCell_(sheet);
}

/**
 * Escribe resultados de tarjeta sin borrar la base del estado de cuenta.
 * Solo sobrescribe valores no vacíos y re-aplica warnings/ticks.
 */
function writeTarjetaItemsToSheet_(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  if (!items || items.length === 0) {
    return { ok: false, message: 'No hay comprobantes para escribir.' };
  }

  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  if (!lastCount) {
    writeItemsToSheet(items);
    return { ok: true, message: `Planilla actualizada: ${items.length || 0} filas escritas.`, payload: { count: items.length } };
  }

  const lastCol = sheet.getLastColumn();
  const numItems = items.length;

  if (numItems > lastCount) {
    const templateRange = sheet.getRange(START_ROW, 1, 1, lastCol);
    templateRange.copyTo(
      sheet.getRange(START_ROW + lastCount, 1, numItems - lastCount, lastCol),
      { contentsOnly: false }
    );
  }

  items.forEach((item, index) => {
    const row = START_ROW + index;
    const fields = item.fields || item;

    clearRowFormatting_(sheet, row);
    clearRowNotes_(sheet, row);

    Object.keys(FIELD_COLUMN_MAP).forEach((fieldName) => {
      if (!(fieldName in fields)) return;
      let value = fields[fieldName];
      if (value === null || value === undefined || value === '') return;

      const colIndex = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);

      if (typeof value === 'boolean') {
        value = value ? 'Si' : 'No';
      } else if (typeof value === 'string') {
        const v = value.trim().toLowerCase();
        if (v === 'true') value = 'Si';
        else if (v === 'false') value = 'No';
      }

      sheet.getRange(row, colIndex).setValue(value);
    });

    const warnings = normalizeWarnings_(getItemWarnings_(item));
    const warningFieldNames = applyWarnings_(sheet, row, warnings);
    const missingManualFields = highlightManualFields_(sheet, row);

    setRowMeta_(row, {
      warning_fields: warningFieldNames,
      warning_details: warnings
    });

    setupTickCell_(sheet, row, warnings, missingManualFields);
  });

  const nextCount = Math.max(lastCount, numItems);
  PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, String(nextCount));
  updateHeaderCompletionCell_(sheet);

  return { ok: true, message: `Planilla actualizada: ${numItems || 0} filas escritas.`, payload: { count: numItems } };
}


/**
 * Escribe filas base del estado de cuenta usando escrituras por columna.
 * Recibe el JSON del estado (schema estado/v0) y escribe solo columnas base.
 * No aplica warnings ni metadata (se reescribe luego con los comprobantes).
 */
function writeStatementToSheet_(statementData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  clearItemsTable(sheet);

  const txs = Array.isArray(statementData?.transacciones) ? statementData.transacciones : [];
  if (!txs.length) {
    PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, '0');
    updateHeaderCompletionCell_(sheet);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const numItems = txs.length;

  // Copy template row formatting down
  const templateRange = sheet.getRange(START_ROW, 1, 1, lastCol);
  if (numItems > 1) {
    templateRange.copyTo(
      sheet.getRange(START_ROW + 1, 1, numItems - 1, lastCol),
      { contentsOnly: false }
    );
  }

  const colFecha = colLetterToIndex(FIELD_COLUMN_MAP['Fecha de factura']);
  const colProv = colLetterToIndex(FIELD_COLUMN_MAP['Proveedor']);
  const colImporte = colLetterToIndex(FIELD_COLUMN_MAP['Importe a rendir']);
  const colMoneda = colLetterToIndex(FIELD_COLUMN_MAP['Moneda']);

  const fechas = txs.map((tx) => [tx?.fecha || '']);
  const provs = txs.map((tx) => [tx?.detalle || '']);
  const importes = txs.map((tx) => [
    tx?.importe_uyu != null
      ? tx.importe_uyu
      : (tx?.importe_usd != null ? tx.importe_usd : (tx?.importe_origen || ''))
  ]);
  const monedas = txs.map((tx) => {
    const hasAny = tx?.importe_uyu != null || tx?.importe_usd != null || tx?.importe_origen != null;
    if (tx?.importe_uyu != null) return ['UYU'];
    return [hasAny ? 'USD' : ''];
  });

  sheet.getRange(START_ROW, colFecha, numItems, 1).setValues(fechas);
  sheet.getRange(START_ROW, colProv, numItems, 1).setValues(provs);
  sheet.getRange(START_ROW, colImporte, numItems, 1).setValues(importes);
  sheet.getRange(START_ROW, colMoneda, numItems, 1).setValues(monedas);

  PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, String(numItems));
  updateHeaderCompletionCell_(sheet);
}


function clearItemsTable(sheet) {
  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  if (!lastCount) return;

  // clear row meta too
  deleteRowMetaRange_(START_ROW, lastCount);

  const cols = getWriteColumnIndexes_();
  cols.forEach((colIndex) => {
    const rng = sheet.getRange(START_ROW, colIndex, lastCount, 1);
    rng.clearContent();
    rng.clearNote();
    restoreBackgroundFromTemplateColumn_(sheet, colIndex, START_ROW, lastCount);
  });
}

function getWriteColumnIndexes_() {
  const cols = new Set();
  Object.values(FIELD_COLUMN_MAP).forEach(letter => cols.add(colLetterToIndex(letter)));
  cols.add(colLetterToIndex(WARNINGS_OK_COL_LETTER)); // tick col
  cols.add(colLetterToIndex(STATUS_COL_LETTER)); // status col
  return Array.from(cols).sort((a, b) => a - b);
}

function ensureRowCapacity_(sheet, requiredRows) {
  if (!requiredRows || requiredRows <= 0) return;
  const requiredLastRow = START_ROW + requiredRows - 1;
  const maxRows = sheet.getMaxRows();
  if (requiredLastRow > maxRows) {
    sheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
  }

  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  const baselineRows = Math.max(DEFAULT_MAX_ROWS, lastCount || 0, 1);
  if (requiredRows <= baselineRows) return;

  const lastCol = sheet.getLastColumn();
  const templateRange = sheet.getRange(START_ROW, 1, 1, lastCol);
  templateRange.copyTo(
    sheet.getRange(START_ROW + baselineRows, 1, requiredRows - baselineRows, lastCol),
    { contentsOnly: false }
  );
}

function buildBaseMatrix_(rowsPlan, totalCols) {
  const statusCol = colLetterToIndex(STATUS_COL_LETTER);
  const out = [];
  (rowsPlan || []).forEach((rowPlan) => {
    const row = new Array(totalCols).fill('');
    const fields = rowPlan?.fields || rowPlan || {};
    Object.keys(fields).forEach((fieldName) => {
      if (fieldName === '__status' || fieldName === 'STATUS') return;
      const colLetter = FIELD_COLUMN_MAP[fieldName];
      if (!colLetter) return;
      row[colLetterToIndex(colLetter) - 1] = fields[fieldName];
    });
    const status = rowPlan?.__status || fields?.STATUS || '';
    row[statusCol - 1] = status;
    out.push(row);
  });
  return out;
}

function writeBaseTableLowWrites_(rowsPlan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  const numRows = (rowsPlan || []).length;
  if (!numRows) {
    PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, '0');
    updateHeaderCompletionCell_(sheet);
    return { ok: true, message: 'No hay filas base para escribir.', payload: { count: 0 } };
  }

  ensureRowCapacity_(sheet, numRows);

  const totalCols = colLetterToIndex(STATUS_COL_LETTER);
  const baseMatrix = buildBaseMatrix_(rowsPlan, totalCols);

  sheet.getRange(START_ROW, 1, numRows, totalCols).setValues(baseMatrix);
  PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, String(numRows));
  updateHeaderCompletionCell_(sheet);

  return { ok: true, message: `Planilla base escrita: ${numRows} filas.`, payload: { count: numRows } };
}

function applyPatchesLowWrites_(rowPatches, rowStatus, baseRowCount, options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  const numRows = baseRowCount || 0;
  if (!numRows) {
    return { ok: true, message: 'No hay filas base para aplicar patches.', payload: { count: 0 } };
  }

  const totalCols = colLetterToIndex(STATUS_COL_LETTER);
  const baseRange = sheet.getRange(START_ROW, 1, numRows, totalCols);
  const matrix = baseRange.getValues();

  const statusCol = colLetterToIndex(STATUS_COL_LETTER);
  const baseLocked = new Set(['Importe a rendir', 'Moneda']);
  const fieldConflicts = [];

  const patches = rowPatches || {};
  const statuses = rowStatus || {};
  const hasPatches = Object.keys(patches).length > 0;
  const hasStatuses = Object.keys(statuses).length > 0;
  if (!hasPatches && !hasStatuses) {
    return { ok: true, message: 'No hay patches para aplicar.', payload: { count: 0, fieldConflicts: [] } };
  }
  Object.keys(patches).forEach((rowKey) => {
    const rowIndex = parseInt(rowKey, 10) - START_ROW;
    if (rowIndex < 0 || rowIndex >= matrix.length) return;
    const row = matrix[rowIndex];
    const patch = patches[rowKey] || {};

    Object.keys(patch).forEach((fieldName) => {
      const colLetter = FIELD_COLUMN_MAP[fieldName];
      if (!colLetter) return;
      if (baseLocked.has(fieldName)) return;

      const colIndex = colLetterToIndex(colLetter) - 1;
      const incoming = patch[fieldName];
      const existing = row[colIndex];
      if (incoming === null || incoming === undefined || incoming === '') return;

      if (MANUAL_FIELDS.indexOf(fieldName) !== -1) {
        if (!isEmptyValue_(existing)) return;
        row[colIndex] = incoming;
        return;
      }

      if (isEmptyValue_(existing)) {
        row[colIndex] = incoming;
        return;
      }

      const same = normalizeCellValue_(existing) === normalizeCellValue_(incoming);
      if (!same) {
        fieldConflicts.push({
          row: START_ROW + rowIndex,
          field: fieldName,
          existing,
          incoming
        });
      }
    });
  });

  Object.keys(statuses).forEach((rowKey) => {
    const rowIndex = parseInt(rowKey, 10) - START_ROW;
    if (rowIndex < 0 || rowIndex >= matrix.length) return;
    matrix[rowIndex][statusCol - 1] = statuses[rowKey];
  });

  baseRange.setValues(matrix);

  return {
    ok: true,
    message: `Patches aplicados: ${Object.keys(patches).length || 0} filas.`,
    payload: {
      count: Object.keys(patches).length || 0,
      fieldConflicts
    }
  };
}

function writeOrphansConflictsSection_(orphans, conflicts, baseRowCount, prevStartRow, prevCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  const totalCols = colLetterToIndex(STATUS_COL_LETTER);
  const rows = [];

  if ((orphans && orphans.length) || (conflicts && conflicts.length)) {
    rows.push(['ORPHANS / CONFLICTS', 'Tipo', 'Archivo', 'Match idx', 'Observacion', 'Proveedor', 'Fecha', 'Numero', 'Importe', 'Moneda']);
    (orphans || []).forEach((o) => {
      rows.push([
        '',
        'ORPHAN',
        o.source || '',
        o.matchIndex || '',
        o.observacion || '',
        o.fields?.['Proveedor'] || '',
        o.fields?.['Fecha de factura'] || '',
        o.fields?.['Numero de Factura'] || '',
        o.fields?.['Importe a rendir'] || '',
        o.fields?.['Moneda'] || ''
      ]);
    });
    (conflicts || []).forEach((c) => {
      rows.push([
        '',
        'CONFLICT',
        c.source || '',
        c.matchIndex || '',
        c.observacion || c.reason || '',
        c.fields?.['Proveedor'] || '',
        c.fields?.['Fecha de factura'] || '',
        c.fields?.['Numero de Factura'] || '',
        c.fields?.['Importe a rendir'] || '',
        c.fields?.['Moneda'] || ''
      ]);
    });
  }

  const startRow = START_ROW + (baseRowCount || 0);
  const targetRows = rows.length || 0;
  const prevRows = prevCount || 0;
  if (!targetRows && !prevRows) {
    return { ok: true, message: 'No hay orphans ni conflicts.', payload: { count: 0 } };
  }

  const prevRow = prevStartRow || startRow;
  const rangeStart = Math.min(startRow, prevRow);
  const rangeEnd = Math.max(startRow + targetRows, prevRow + prevRows);
  const rangeRows = Math.max(0, rangeEnd - rangeStart);
  if (!rangeRows) return { ok: true, message: 'No hay orphans ni conflicts.', payload: { count: 0 } };

  const matrix = [];
  const offset = startRow - rangeStart;
  for (let i = 0; i < rangeRows; i++) {
    const row = new Array(totalCols).fill('');
    const srcIndex = i - offset;
    if (srcIndex >= 0 && rows[srcIndex]) {
      for (let c = 0; c < rows[srcIndex].length && c < totalCols; c++) {
        row[c] = rows[srcIndex][c];
      }
      if (rows[srcIndex][1] === 'ORPHAN') row[totalCols - 1] = 'ORPHAN';
      if (rows[srcIndex][1] === 'CONFLICT') row[totalCols - 1] = 'CONFLICT';
    }
    matrix.push(row);
  }

  sheet.getRange(rangeStart, 1, rangeRows, totalCols).setValues(matrix);
  return { ok: true, message: `Orphans/conflicts escritos: ${rows.length || 0}.`, payload: { count: rows.length || 0 } };
}
