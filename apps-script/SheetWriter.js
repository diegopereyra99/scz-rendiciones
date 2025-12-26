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
