/**************
 * CONFIG
 **************/

const SHEET_NAME = 'Formulario';
const START_ROW = 14;
const STOP_ROW = 49;
const GCS_BUCKET = 'scz-uy-rendiciones';

const GEMINI_MODEL = 'gemini-2.5-flash';
// https://us-central1-aiplatform.googleapis.com/v1/projects/diesel-thunder-465918-r4/locations/us-central1/publishers/google/models/gemini-2.5-flash:generateContent

const GEMINI_ENDPOINT =
  'https://us-central1-aiplatform.googleapis.com/v1/projects/diesel-thunder-465918-r4/locations/us-central1/publishers/google/models/' +
  GEMINI_MODEL +
  ':generateContent';

const PROMPT_DOC_ID = '1kP0ZIsnkHwSFppuCc-Gl9-Hk39zaNWeaT--tf67AQ8w';
const SYSTEM_DOC_ID = '1_FzB_GNV-Fm_aEfGE0TORtTtrI9AkB0wA8ZXPMFM6b4';
const SCHEMA_FILE_ID = '1ai53TISXDSZBG-TiGCbAKdSczd523C8j';

// ✅ checkbox column (you changed this)
const WARNINGS_OK_COL_LETTER = 'C';

const FIELD_COLUMN_MAP = {
  'Fecha de factura': 'E',
  'Numero de Factura': 'G',
  'OC': 'H',
  'Contiene RUT comprador': 'I',
  'Factura exterior': 'J',
  'Proveedor': 'K',
  'Tipo de gasto': 'L',
  'Imputación gasto': 'N',
  'Importe a rendir': 'P',
  'Moneda': 'Q',
  'Base 22': 'S',
  // 'Iva 22': 'T',
  'Base 10': 'U',
  // 'Iva 10': 'V',
  'Exento': 'W',
  'Descuentos': 'X',
  'Propina': 'Y',
  'Descripcion': 'AA'
};

const MANUAL_FIELDS = [
  'OC',
  'Imputación gasto'
];

const COLOR_WARNING = '#fff2cc';
const COLOR_MANUAL  = '#d9ead3';
const COLOR_ORPHAN  = '#f4cccc';

const MAX_RESOLUTION = 2000;
const PROCESS_BATCH_SIZE = 8;

const PROP_LAST_ITEM_COUNT = 'LAST_ITEM_COUNT';

// Row meta keys will be: ROW_META_<rowNumber>
const PROP_ROW_META_PREFIX = 'ROW_META_';

// You changed this; keep as-is
const GRAY_COLUMNS = ['L', 'N', 'Q'];

const SERVICE_URL = 'https://rendiciones-service-213593806678.us-central1.run.app';

// Identidad de la rendición (por usuario/mes/año)
const RENDICION_USER = 'diego';     // <- cada usuario lo cambia
const RENDICION_YEAR = 2025;        // <- o new Date().getFullYear()
const RENDICION_MONTH = 12;         // 1-12 (cargado en config del usuario)

const SS_ID = SpreadsheetApp.getActive().getId();

// Drive: carpeta donde van los outputs (PDF + XLSM)
let OUTPUT_FOLDER_ID = null; // Se resuelve lazy para no romper en triggers simples (onEdit)

// Cloud Run: dónde escribir en GCS
const GCS_PREFIX = `gs://${GCS_BUCKET}/rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${String(RENDICION_MONTH).padStart(2,'0')}`;

const DRIVE_API_ENABLED = true;   // <-- toggle

function getOutputFolderId_() {
  if (OUTPUT_FOLDER_ID) return OUTPUT_FOLDER_ID;
  OUTPUT_FOLDER_ID = getBaseFolder_().getId();
  return OUTPUT_FOLDER_ID;
}

const FOLDER_EFECTIVO_NAME = 'Comprobantes efectivo';
const FOLDER_TARJETA_NAME = 'Comprobantes tarjeta corporativa';
const FOLDER_ESTADO_NAME = 'Estado de cuenta';
