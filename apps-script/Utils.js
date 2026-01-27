
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
    files.push({
      id: f.getId(),
      name,
      created: f.getDateCreated().getTime(),
      size: f.getSize(),
      updated: f.getLastUpdated().getTime()
    });
  }
  // orden determinístico: por nombre, y fallback por fecha
  files.sort((a, b) => (a.name.localeCompare(b.name) || a.created - b.created));
  return files;
}

function getBaseFolder_() {
  const ssFile = DriveApp.getFileById(SS_ID);
  const parents = ssFile.getParents();
  if (!parents.hasNext()) throw new Error('El formulario no tiene carpeta padre.');
  return parents.next();
}

function getSubfolderIdByName_(parentFolder, folderName, required) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    if (required) throw new Error(`No existe la carpeta: ${folderName}`);
    return null;
  }
  return folders.next().getId();
}

function getModeFolderId_(mode) {
  const base = getBaseFolder_();
  if (mode === 'efectivo') return getSubfolderIdByName_(base, FOLDER_EFECTIVO_NAME, true);
  if (mode === 'tarjeta') return getSubfolderIdByName_(base, FOLDER_TARJETA_NAME, true);
  throw new Error(`Modo inválido: ${mode}`);
}

function getStatementFolderId_() {
  const base = getBaseFolder_();
  return getSubfolderIdByName_(base, FOLDER_ESTADO_NAME, false);
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
function buildInputsZipObjectName_(mode) {
  // rendiciones/<rendicionId>/inputs/inputs_<timestamp>.zip
  const ts = Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss');
  const suffix = mode ? `inputs/${mode}/inputs_${ts}.zip` : `inputs/inputs_${ts}.zip`;
  return `rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${String(RENDICION_MONTH).padStart(2,'0')}/${suffix}`;
}

function buildStatementObjectName_(fingerprint, filename) {
  const safeName = (filename || 'estado.pdf').replace(/[^A-Za-z0-9._-]/g, '_');
  const safeFp = fingerprint ? String(fingerprint).replace(/[^A-Za-z0-9._-]/g, '_') : Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss');
  return `rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${String(RENDICION_MONTH).padStart(2,'0')}/inputs/estado/${safeFp}_${safeName}`;
}

function buildZipFromDriveFolder_(folderId, filesOrIds) {
  let files = [];
  if (Array.isArray(filesOrIds) && filesOrIds.length) {
    files = filesOrIds.map((f) => (typeof f === 'string' ? { id: f } : f))
      .filter((f) => f && f.id);
  } else {
    files = listDriveFiles_(folderId);
  }
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
