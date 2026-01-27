/**************
 * onEdit:
 * - Manual fields: if user fills -> remove the manual NOTE (but keep green until tick)
 * - Tick: only allow if row ready; if tick true clear highlights; if false reapply
 **************/

function onEdit(e) {
  try {
    const range = e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const startRow = range.getRow();
    const startCol = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    if (startRow + numRows - 1 < START_ROW) return;

    const okCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);
    const manualCols = new Set(
      MANUAL_FIELDS
        .map((f) => colLetterToIndex(FIELD_COLUMN_MAP[f]))
        .filter((c) => !!c)
    );

    const values = range.getValues();
    let tickEdited = false;
    let firstBlocked = null;

    for (let r = 0; r < numRows; r++) {
      const rowIndex = startRow + r;
      if (rowIndex < START_ROW) continue;

      for (let c = 0; c < numCols; c++) {
        const colIndex = startCol + c;
        const cellValue = values[r][c];

        if (manualCols.has(colIndex) && !isEmptyValue_(cellValue)) {
          range.getCell(r + 1, c + 1).setNote('');
        }

        if (colIndex === okCol) {
          tickEdited = true;
          const res = handleTickEdit_(sheet, rowIndex, cellValue === true);
          if (res?.blocked && !firstBlocked) firstBlocked = res;
        }
      }
    }

    if (firstBlocked) {
      SpreadsheetApp.getActive().toast(
        `No se puede marcar: faltan campos manuales en la fila ${firstBlocked.row}.`,
        'Automation',
        5
      );
    }

    if (tickEdited) {
      updateHeaderCompletionCell_(sheet);
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err);
    SpreadsheetApp.getActive().toast('Error en onEdit: ' + err, 'Automation', 5);
  }
}


/**************
 * Warnings + manual
 **************/

function getItemWarnings_(item) {
  return item.Warnings || item.warnings || [];
}

function normalizeWarnings_(warnings) {
  if (!Array.isArray(warnings)) return [];
  return warnings.map((w) => {
    if (w && typeof w === 'object') {
      const campo = w.campo || w.field || w.field_id || w.fieldName || null;
      const tipo = w.tipo || 'Advertencia';
      const mensaje = w.mensaje || w.message || w.msg || '';
      // const fullMessage = (tipo ? (tipo + ': ') : '') + (mensaje || '');
      return { fieldName: campo, message: mensaje };
    }
    if (typeof w === 'string') return { fieldName: null, message: w };
    return { fieldName: null, message: String(w) };
  });
}

/**
 * Paint warning cells yellow and add note on each cell.
 * Returns list of fieldNames that are mapped and were highlighted.
 */
function applyWarnings_(sheet, rowIndex, warnings) {
  const highlighted = [];
  let hasGeneral = false;

  warnings.forEach((w) => {
    const fieldName = w.fieldName;
    const message = w.message || 'Revisar este valor.';

    if (fieldName === 'general' || fieldName === null) {
      applyGeneralWarningRow_(sheet, rowIndex, message);
      highlighted.push('general');
      hasGeneral = true;
      return;
    }
    if (fieldName && FIELD_COLUMN_MAP[fieldName]) {
      const colIndex = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);
      const cell = sheet.getRange(rowIndex, colIndex);

      cell.setBackground(COLOR_WARNING);
      const prevNote = cell.getNote();
      const noteText = buildWarningNote_(message);
      cell.setNote((prevNote ? prevNote + '\n' : '') + noteText);

      highlighted.push(fieldName);
    }
  });

  if (hasGeneral) {
    const okCell = sheet.getRange(rowIndex, colLetterToIndex(WARNINGS_OK_COL_LETTER));
    const prev = okCell.getNote();
    okCell.setNote((prev ? prev + '\n' : '') + buildWarningNote_('Advertencia general en la fila.'));
  }

  return highlighted;
}

function applyGeneralWarningRow_(sheet, rowIndex, message) {
  const cols = getWriteColumnIndexes_();
  cols.forEach((colIndex) => {
    const cell = sheet.getRange(rowIndex, colIndex);
    cell.setFontLine('underline');
  });
  if (message) {
    const cell = sheet.getRange(rowIndex, colLetterToIndex(WARNINGS_OK_COL_LETTER));
    const prevNote = cell.getNote();
    cell.setNote((prevNote ? prevNote + '\n' : '') + buildWarningNote_(message));
  }
}

/**
 * Highlight manual fields green if empty.
 * Returns list of manual field names that are currently missing.
 */
function highlightManualFields_(sheet, rowIndex) {
  const missing = [];

  MANUAL_FIELDS.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;

    const colIndex = colLetterToIndex(colLetter);
    const cell = sheet.getRange(rowIndex, colIndex);
    const value = cell.getValue();

    if (isEmptyValue_(value)) {
      missing.push(fieldName);
      cell.setBackground(COLOR_MANUAL);
      if (!cell.getNote()) cell.setNote('Campo a completar manualmente.');
    }
  });

  return missing;
}


/**************
 * Tick behavior + row meta (DocumentProperties)
 **************/

function handleTickEdit_(sheet, rowIndex, checked) {
  const okColIndex = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const tickCell = sheet.getRange(rowIndex, okColIndex);

  if (checked) {
    const missing = getMissingManualFieldsNow_(sheet, rowIndex);

    if (missing.length) {
      tickCell.setValue(false);
      tickCell.setNote(buildMissingManualNote_(missing));
      return { blocked: true, row: rowIndex, missing };
    }

    clearRowFormatting_(sheet, rowIndex);
    clearRowNotes_(sheet, rowIndex);
    tickCell.setNote('✅ Fila validada. Si destildás, vuelven los resaltados.');
  } else {
    reapplyRowHighlights_(sheet, rowIndex);
  }

  return { blocked: false, row: rowIndex, missing: [] };
}

function setupTickCell_(sheet, rowIndex, warnings, missingManualFields) {
  const okColIndex = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const okCell = sheet.getRange(rowIndex, okColIndex);

  okCell.insertCheckboxes();
  okCell.setValue(false);

  const warnCount = (warnings || []).length;
  const missing = missingManualFields || [];

  const lines = [];
  lines.push('Estado de la fila: PENDIENTE');
  if (missing.length) {
    lines.push('');
    lines.push('Faltan campos manuales:');
    missing.forEach((f) => lines.push('• ' + f));
  }
  if (warnCount) {
    lines.push('');
    lines.push('Warnings detectados: ' + warnCount);
    lines.push('• Revisá las celdas amarillas.');
  }
  lines.push('');
  lines.push('✅ Para cerrar la fila, marcá este tick.');
  lines.push('⚠ No se puede marcar si faltan campos manuales.');

  okCell.setNote(lines.join('\n'));
}




/**************
 * Header completion cell
 **************/

function updateHeaderCompletionCell_(sheet) {
  const headerRow = START_ROW - 1;
  const okCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const headerCell = sheet.getRange(headerRow, okCol);

  const props = PropertiesService.getDocumentProperties();
  const count = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  if (!count) {
    headerCell.clearContent();
    headerCell.clearNote();
    restoreBackgroundFromTemplateCell_(sheet, headerRow, okCol);
    return;
  }

  // Keep header empty; no banner or status
  headerCell.clearContent();
  headerCell.clearNote();
  restoreBackgroundFromTemplateCell_(sheet, headerRow, okCol);
}



/**************
 * Row readiness / highlight toggle
 **************/

function getMissingManualFieldsNow_(sheet, rowIndex) {
  const missing = [];
  MANUAL_FIELDS.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;
    const v = sheet.getRange(rowIndex, colLetterToIndex(colLetter)).getValue();
    if (isEmptyValue_(v)) missing.push(fieldName);
  });
  return missing;
}

function reapplyRowHighlights_(sheet, rowIndex) {
  const meta = getRowMeta_(rowIndex);
  const warningFields = meta.warning_fields || [];
  const warningDetails = meta.warning_details || [];

  // Reapply warnings
  warningDetails.forEach((w) => {
    const fieldName = w.fieldName;
    if (fieldName === 'general' || fieldName === null) {
      applyGeneralWarningRow_(sheet, rowIndex, w.message);
      return;
    }
    if (!fieldName || !FIELD_COLUMN_MAP[fieldName]) return;
    const colIndex = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);
    const cell = sheet.getRange(rowIndex, colIndex);
    cell.setBackground(COLOR_WARNING);
    cell.setNote(buildWarningNote_(w.message));
  });

  // Reapply manual green for still-empty fields
  const missingNow = getMissingManualFieldsNow_(sheet, rowIndex);
  missingNow.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;
    const colIndex = colLetterToIndex(colLetter);
    const cell = sheet.getRange(rowIndex, colIndex);
    cell.setBackground(COLOR_MANUAL);
    if (!cell.getNote()) cell.setNote('Campo a completar manualmente.');
  });

  // Update tick note
  const okCell = sheet.getRange(rowIndex, colLetterToIndex(WARNINGS_OK_COL_LETTER));
  okCell.setNote(buildTickPendingNote_(warningFields.length, missingNow));
}

function buildTickPendingNote_(warnCount, missingFields) {
  const lines = [];
  lines.push('Estado de la fila: PENDIENTE');

  if (missingFields && missingFields.length) {
    lines.push('');
    lines.push('Faltan campos manuales:');
    missingFields.forEach((f) => lines.push('• ' + f));
  }

  if (warnCount) {
    lines.push('');
    lines.push('Warnings detectados: ' + warnCount);
    lines.push('• Revisá las celdas amarillas.');
  }

  lines.push('');
  lines.push('✅ Para cerrar la fila, marcá este tick.');
  lines.push('⚠ No se puede marcar si faltan campos manuales.');

  return lines.join('\n');
}

function buildMissingManualNote_(missing) {
  const lines = [];
  lines.push('❌ No se puede marcar ✅ todavía.');
  lines.push('Faltan campos manuales:');
  (missing || []).forEach((f) => lines.push('• ' + f));
  lines.push('');
  lines.push('Completalos y volvé a marcar el tick.');
  return lines.join('\n');
}

function clearRowNotes_(sheet, rowIndex) {
  const cols = getWriteColumnIndexes_();
  cols.forEach((col) => {
    sheet.getRange(rowIndex, col).setNote('');
  });
}

function clearRowFormatting_(sheet, rowIndex) {
  const cols = getWriteColumnIndexes_();
  cols.forEach((col) => {
    restoreBackgroundFromTemplateOrGray_(sheet, rowIndex, col);
    sheet.getRange(rowIndex, col).setFontLine('none');
  });
}

function buildWarningNote_(message) {
  return (
    '⚠ ' + (message || 'Revisar este valor.') +
    '\n\nImportante:' +
    '\n• Podés corregir el valor editando la celda.' +
    '\n• El amarillo NO se va hasta que marques ✅ en el tick de la fila.'
  );
}

function clearAllRows() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  const totalRows = Math.max(DEFAULT_MAX_ROWS, lastCount || 0);
  if (totalRows <= 0) {
    if (typeof resetJobState_ === 'function') {
      resetJobState_();
    }
    if (typeof resetCache_ === 'function') {
      resetCache_();
    }
    return;
  }

  const cols = getWriteColumnIndexes_();
  const tickCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const st = typeof jobStateGet_ === 'function' ? jobStateGet_() : {};
  const baseRowCount = st?.baseRowCount || lastCount || 0;
  const orphansCount = st?.lastOrphansRowCount || 0;

  // Clear values/notes while preserving any formulas that might exist
  cols.forEach((col) => {
    const rng = sheet.getRange(START_ROW, col, totalRows, 1);
    const formulas = rng.getFormulas();

    rng.clearContent(); // wipes values + formulas
    rng.setFormulas(formulas); // restore any formulas that were there
    rng.clearNote();

    if (col === tickCol) {
      rng.clearDataValidations(); // remove checkboxes
    }

    restoreBackgroundFromTemplateColumn_(sheet, col, START_ROW, totalRows);
  });

  if (orphansCount && baseRowCount) {
    cols.forEach((col) => {
      const rng = sheet.getRange(START_ROW + baseRowCount, col, orphansCount, 1);
      rng.clearContent();
      rng.clearNote();
      restoreBackgroundFromTemplateColumn_(sheet, col, START_ROW + baseRowCount, orphansCount);
    });
  }

  // Reset metadata/state
  deleteRowMetaRange_(START_ROW, totalRows);
  PropertiesService.getDocumentProperties().setProperty(PROP_LAST_ITEM_COUNT, '0');
  updateHeaderCompletionCell_(sheet);

  if (typeof resetJobState_ === 'function') {
    resetJobState_();
  }
  if (typeof resetCache_ === 'function') {
    resetCache_();
  }
}

function isEmptyValue_(v) {
  if (v === '' || v === null) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  return false;
}

function getLastDataRow_(sh) {
  const anchorLetter = FIELD_COLUMN_MAP['Importe a rendir'] || 'G';
  const col = colLetterToIndex(anchorLetter);
  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  const last = Math.max(STOP_ROW, START_ROW + (lastCount || 0) - 1);
  if (last < START_ROW) return START_ROW - 1;

  const vals = sh.getRange(START_ROW, col, last - START_ROW + 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = vals[i][0];
    if (v !== '' && v !== null) return START_ROW + i;
  }
  return START_ROW - 1;
}

function assertAllChecksOk_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja: ${SHEET_NAME}`);

  const props = PropertiesService.getDocumentProperties();
  const lastCount = parseInt(props.getProperty(PROP_LAST_ITEM_COUNT) || '0', 10);
  const st = typeof jobStateGet_ === 'function' ? jobStateGet_() : {};
  const baseCount = Math.max(st?.baseRowCount || 0, lastCount || 0);
  if (!baseCount) throw new Error('No hay filas con datos para finalizar.');

  const checkCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const checks = sh.getRange(START_ROW, checkCol, baseCount, 1).getValues();

  const bad = [];
  for (let i = 0; i < checks.length; i++) {
    const v = checks[i][0];
    if (v !== true) bad.push(START_ROW + i);
  }
  if (bad.length) {
    throw new Error(`No podés generar: hay checks sin marcar en filas: ${bad.slice(0, 25).join(', ')}${bad.length > 25 ? '…' : ''}`);
  }
}

/**************
 * Clear table safely (only columns we write)
 **************/




/**************
 * Background restore (keep your behavior)
 **************/

function restoreBackgroundFromTemplateOrGray_(sheet, rowIndex, colIndex) {
  const colLetter = indexToColLetter_(colIndex);
  const rng = sheet.getRange(rowIndex, colIndex);
  if (GRAY_COLUMNS.includes(colLetter)) {
    rng.setBackground('#dadada');
  } else {
    rng.setBackground(null);
  }
}

function restoreBackgroundFromTemplateColumn_(sheet, colIndex, fromRow, numRows) {
  const colLetter = indexToColLetter_(colIndex);
  const bg = GRAY_COLUMNS.includes(colLetter) ? '#dadada' : null;
  sheet.getRange(fromRow, colIndex, numRows, 1).setBackground(bg);
}

function restoreBackgroundFromTemplateCell_(sheet, rowIndex, colIndex) {
  const colLetter = indexToColLetter_(colIndex);
  const bg = GRAY_COLUMNS.includes(colLetter) ? '#dadada' : null;
  sheet.getRange(rowIndex, colIndex).setBackground(bg);
}
