
/**************
 * Col + mapping utils
 **************/

function colLetterToIndex(colLetter) {
  colLetter = colLetter.toUpperCase();
  let index = 0;
  for (let i = 0; i < colLetter.length; i++) {
    index = index * 26 + (colLetter.charCodeAt(i) - 64);
  }
  return index;
}

function indexToColLetter_(index) {
  let s = '';
  while (index > 0) {
    const m = (index - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    index = Math.floor((index - 1) / 26);
  }
  return s;
}

function fieldNameFromColumn_(colIndex) {
  return Object.keys(FIELD_COLUMN_MAP).find((name) => {
    return colLetterToIndex(FIELD_COLUMN_MAP[name]) === colIndex;
  }) || null;
}

function listDriveFiles_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    // salteá basura típica
    const name = f.getName();
    if (name.startsWith('~$') || name === '.DS_Store') continue;
    files.push({ id: f.getId(), name, created: f.getDateCreated().getTime() });
  }
  // orden determinístico: por nombre, y fallback por fecha
  files.sort((a, b) => (a.name.localeCompare(b.name) || a.created - b.created));
  return files;
}

function newRendicionId_() {
  // suficientemente único para tus flujos
  return Utilities.getUuid().slice(0, 8);
}

function buildRendicionId_() {
  const mm = String(RENDICION_MONTH).padStart(2, '0');
  // simple, estable, legible
  return `${RENDICION_YEAR}_${mm}_${RENDICION_USER}`;
}

// Nombre del zip en GCS
function buildInputsZipObjectName_() {
  // rendiciones/<rendicionId>/inputs/inputs_<timestamp>.zip
  const ts = Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss');
  return `rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${String(RENDICION_MONTH).padStart(2,'0')}/inputs/inputs_${ts}.zip`;
}

function buildZipFromDriveFolder_(folderId) {
  const files = listDriveFiles_(folderId);
  if (!files.length) throw new Error('No hay archivos para zipear.');

  const blobs = [];
  for (const f of files) {
    const file = DriveApp.getFileById(f.id);
    // Ojo: el nombre dentro del zip
    const b = file.getBlob().setName(file.getName());
    blobs.push(b);
  }

  // utilities.zip crea un blob ZIP
  return Utilities.zip(blobs, 'inputs.zip');
}

function getSheetGidByName_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) throw new Error(`No existe la hoja: ${sheetName}`);
  return sheet.getSheetId(); // ← ESTE es el gid
}


function clearAllNotes() {
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    sh.getDataRange().clearNote(); // borra TODAS las notas de la hoja
  });
}


