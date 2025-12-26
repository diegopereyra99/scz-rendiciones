/**************
 * onEdit:
 * - Manual fields: if user fills -> remove the manual NOTE (but keep green until tick)
 * - Tick: only allow if row ready; if tick true clear highlights; if false reapply
 **************/

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  if (row < START_ROW) return;

  const okCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);

  // 0) If user edits a manual field: remove note when filled (without removing green)
  const fieldName = fieldNameFromColumn_(col);
  if (fieldName && MANUAL_FIELDS.includes(fieldName)) {
    const v = range.getValue();
    if (v !== '' && v !== null) {
      if (range.getNote() === 'Campo a completar manualmente.') range.setNote('');
    }
    // Do not return; user might also be editing tick column (handled below)
  }

  // 1) Tick behavior
  if (col === okCol) {
    const checked = range.getValue() === true;

    if (checked) {
      const missing = getMissingManualFieldsNow_(sheet, row);

      if (missing.length) {
        range.setValue(false);

        const lines = []; 
        lines.push('❌ No se puede marcar ✅ todavía.');
        lines.push('Faltan campos manuales:');
        missing.forEach((f) => lines.push('• ' + f));
        lines.push('');
        lines.push('Completalos y volvé a marcar el tick.');
        range.setNote(lines.join('\n'));

        SpreadsheetApp.getActive().toast('No se puede marcar: faltan campos manuales en la fila.', 'Automation', 5);
        return;
      }

      clearRowHighlights_(sheet, row);
      range.setNote('✅ Fila validada. Si destildás, vuelven los resaltados.');
    } else {
      reapplyRowHighlights_(sheet, row);
    }

    updateHeaderCompletionCell_(sheet);
    return;
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

  warnings.forEach((w) => {
    const fieldName = w.fieldName;
    const message = w.message || 'Revisar este valor.';

    const noteText =
      '⚠ ' + message +
      '\n\nImportante:' +
      '\n• Podés corregir el valor editando la celda.' +
      '\n• El amarillo NO se va hasta que marques ✅ en el tick de la fila.';

    if (fieldName && FIELD_COLUMN_MAP[fieldName]) {
      const colIndex = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);
      const cell = sheet.getRange(rowIndex, colIndex);

      cell.setBackground(COLOR_WARNING);
      const prevNote = cell.getNote();
      cell.setNote((prevNote ? prevNote + '\n' : '') + noteText);

      highlighted.push(fieldName);
    }
  });

  return highlighted;
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

    if (value === '' || value === null) {
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

  let allChecked = true;
  for (let i = 0; i < count; i++) {
    const row = START_ROW + i;
    const v = sheet.getRange(row, okCol).getValue();
    if (v !== true) { allChecked = false; break; }
  }

  if (allChecked) {
    headerCell.clearContent();
    headerCell.clearNote();
    restoreBackgroundFromTemplateCell_(sheet, headerRow, okCol);
  } else {
    headerCell.setValue('COMPLETAR');
    headerCell.setNote('Debe quedar todo tildado ✅ para considerar la rendición pronta.');
  }
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
    if (v === '' || v === null) missing.push(fieldName);
  });
  return missing;
}

function clearRowHighlights_(sheet, rowIndex) {
  const meta = getRowMeta_(rowIndex);
  const warningFields = meta.warning_fields || [];

  // Restore warning cells
  warningFields.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;
    const colIndex = colLetterToIndex(colLetter);
    restoreBackgroundFromTemplateOrGray_(sheet, rowIndex, colIndex);
  });

  // Restore manual cells (remove green)
  MANUAL_FIELDS.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;
    const colIndex = colLetterToIndex(colLetter);
    restoreBackgroundFromTemplateOrGray_(sheet, rowIndex, colIndex);
  });
}

function reapplyRowHighlights_(sheet, rowIndex) {
  const meta = getRowMeta_(rowIndex);
  const warningFields = meta.warning_fields || [];

  // Reapply warnings
  warningFields.forEach((fieldName) => {
    const colLetter = FIELD_COLUMN_MAP[fieldName];
    if (!colLetter) return;
    sheet.getRange(rowIndex, colLetterToIndex(colLetter)).setBackground(COLOR_WARNING);
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

function getLastDataRow_(sh) {
  const anchorLetter = FIELD_COLUMN_MAP['Importe total'] || 'G';
  const col = colLetterToIndex(anchorLetter);
  const last = STOP_ROW
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

  const lastRow = getLastDataRow_(sh);
  if (lastRow < START_ROW) throw new Error('No hay filas con datos para finalizar.');

  const checkCol = colLetterToIndex(WARNINGS_OK_COL_LETTER);
  const checks = sh.getRange(START_ROW, checkCol, lastRow - START_ROW + 1, 1).getValues();

  const bad = [];
  for (let i = 0; i < checks.length; i++) {
    const v = checks[i][0];
    if (v !== true) bad.push(START_ROW + i);
  }
  // TODO; Correct this check
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
  if (GRAY_COLUMNS.includes(colLetter)) {
    sheet.getRange(rowIndex, colIndex).setBackground('#dadada');
    return;
  }
  restoreBackgroundFromTemplateCell_(sheet, rowIndex, colIndex);
}

function getTemplateBackground_(sheet, colIndex) {
  return sheet.getRange(START_ROW, colIndex).getBackground();
}

function restoreBackgroundFromTemplateColumn_(sheet, colIndex, fromRow, numRows) {
  const bg = getTemplateBackground_(sheet, colIndex);
  sheet.getRange(fromRow, colIndex, numRows, 1).setBackground(bg);
}

// You changed this to null; keep it
function restoreBackgroundFromTemplateCell_(sheet, rowIndex, colIndex) {
  sheet.getRange(rowIndex, colIndex).setBackground(null);
}
