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
  return Array.from(cols).sort((a, b) => a - b);
}
