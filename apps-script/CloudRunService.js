function callCloudRunJson_(path, payload) {
  // const token = getServiceAccountAccessToken();
  const url = SERVICE_URL.replace(/\/+$/, '') + path;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    // headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText() || '';
  if (code < 200 || code >= 300) throw new Error(`Cloud Run ${path} failed (${code}): ${text}`);
  return text ? JSON.parse(text) : null;
}


function jobStateGet_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('REN_JOB_STATE');
  return raw ? JSON.parse(raw) : {};
}

function jobStateSet_(obj) {
  PropertiesService.getDocumentProperties().setProperty('REN_JOB_STATE', JSON.stringify(obj || {}));
}

function getCellValueByA1_(sh, a1) {
  return sh.getRange(a1).getValue();
}

function buildXlsmValuesFromSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const lastRow = getLastDataRow_(sh);
  if (lastRow < START_ROW) return [];

  const fields = Object.keys(FIELD_COLUMN_MAP);
  const colIdxs = fields.map(f => colLetterToIndex(FIELD_COLUMN_MAP[f]));

  // Leer en bloque: rango desde minCol a maxCol para eficiencia
  const minCol = Math.min(...colIdxs);
  const maxCol = Math.max(...colIdxs);
  const width = maxCol - minCol + 1;

  const values = sh.getRange(START_ROW, minCol, lastRow - START_ROW + 1, width).getValues();

  const out = [];
  for (let r = 0; r < values.length; r++) {
    const sheetRow = START_ROW + r;

    for (let i = 0; i < fields.length; i++) {
      const fieldName = fields[i];
      const col = colIdxs[i];
      const v = values[r][col - minCol];

      // opcional: no mandar vacíos
      if (v === '' || v === null) continue;

      const colNum = colLetterToIndex(FIELD_COLUMN_MAP[fieldName]);

      out.push({
        fieldId: fieldName,
        sheet: SHEET_NAME,
        row: sheetRow,
        col: colNum,
        value: normalizeCellValue_(v)
      });
    }
  }

  const headerFields = [
    { fieldId: 'Fecha', a1: 'G6' },
    { fieldId: 'Titular', a1: 'G7' },
    { fieldId: 'Dias trabajados (Comercial)', a1: 'G8' }
  ];

  headerFields.forEach(f => {
    const v = getCellValueByA1_(sh, f.a1);
    if (v === '' || v === null) return;

    const colLetter = f.a1.replace(/[0-9]/g, '');
    const row = parseInt(f.a1.replace(/[A-Z]/gi, ''), 10);

    out.push({
      fieldId: f.fieldId,
      sheet: SHEET_NAME,
      row,
      col: colLetterToIndex(colLetter),
      value: normalizeCellValue_(v)
    });
  });

  return out;
}


function normalizeCellValue_(v) {
  // Apps Script devuelve Date/Number/Boolean/String; normalizamos a JSON friendly
  if (v instanceof Date) return v.toISOString();
  return v;
}

function buildCoverPdf_(spreadsheetId, sheetId, pdfFileName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);

  ss.getSheets().forEach(sh => {
    sh.getDataRange().clearNote(); // borra TODAS las notas de la hoja
  });

  // Export endpoint de Google Sheets (requiere OAuth del usuario / script)
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
    `?format=pdf` +
    `&gid=${sheetId}` +
    `&size=A4` +
    `&portrait=false` +          // <-- landscape
    `&fitw=true` +               // fit to width
    // `&scale=2` +                 // <-- 1..4 (probá 2 o 3)
    `&sheetnames=false&printtitle=false&pagenumbers=false` +
    `&gridlines=false&fzr=false` +
    `&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2`;

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) throw new Error(`Cover export failed (${code}): ${resp.getContentText()}`);

  const pdfBlob = resp.getBlob().setName(pdfFileName);
  return pdfBlob; // <- Blob PDF listo (para Drive o GCS)
}


