function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automatización')
    .addItem('1) Procesar archivos (IA → planilla)', 'ui_process_all')
    .addItem('2) Generar Excel + PDF', 'ui_stage3_finalize')
    .addItem('Reiniciar todo', 'clearAllRows')
    .addToUi();
}

/** ---------- UI wrappers (abren HTML) ---------- */

function ui_process_all() {
  return showModal_('stage_wait', {
    title: 'Procesando archivos…',
    action: 'stage_upload_and_process',
    description: 'Preprocesando, analizando con IA y escribiendo la planilla.'
  });
}

function ui_stage3_finalize() {
  return showModal_('stage_wait', {
    title: 'Generando entregables…',
    action: 'stage3_finalize',
    description: 'Generando PDF final y Excel.'
  });
}

/** ---------- Modal helper ---------- */

function showModal_(htmlFile, data) {
  const tpl = HtmlService.createTemplateFromFile(htmlFile);
  tpl.DATA = data || {};
  const html = tpl.evaluate().setWidth(520).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, data?.title || 'Automatización');
}

/** ---------- Backend handlers (llamados desde HTML) ---------- */

function stage1_upload() {
  return runStage1Upload_();
}

function stage_upload_and_process() {
  const uploadRes = runStage1Upload_();
  if (!uploadRes?.ok) return uploadRes;

  const analyzeRes = stage2_process_analyze();
  if (!analyzeRes?.ok) return analyzeRes;

  const writeRes = stage2_write_results(analyzeRes.items || analyzeRes.payload?.items || []);
  if (!writeRes?.ok) return writeRes;

  return {
    ok: true,
    message: writeRes.message || 'Archivos procesados y planilla actualizada.',
    payload: {
      upload: uploadRes.payload || null,
      process: analyzeRes.payload || null,
      write: writeRes.payload || null
    }
  };
}

function runStage1Upload_() {
  const rendicionId = buildRendicionId_();
  const driveFiles = listDriveFiles_(FOLDER_ID);
  if (!driveFiles.length) return { ok: false, message: 'No hay archivos en la carpeta.' };

  let input = null;

  if (DRIVE_API_ENABLED) {
    input = { driveFileIds: driveFiles.map(f => f.id) };
  } else {
    const zipBlob = buildZipFromDriveFolder_(FOLDER_ID);
    const objectName = buildInputsZipObjectName_();
    const gcsUri = uploadBlobToGCS_(GCS_BUCKET, objectName, zipBlob);

    input = { zipGcsUri: gcsUri }; // <- si tu servicio lo llama distinto, lo ajustamos acá
  }

  const req = {
    rendicionId,
    input,
    output: { gcsPrefix: GCS_PREFIX },
    options: {
      jpgQuality: 95,
      maxSidePx: MAX_RESOLUTION,
      uploadOriginals: false,
      pdfMode: "keep"
    }
  };

  const res = callCloudRunJson_('/v1/normalize', req);
  if (!res?.ok) throw new Error(`Normalize ok=false: ${JSON.stringify(res)}`);

  jobStateSet_({
    rendicionId,
    gcsPrefix: GCS_PREFIX,
    manifestGcsUri: res.manifestGcsUri || null,
    normalizedItems: (res.items || []).map(it => ({
      gcsUri: it.normalized?.gcsUri,
      mimeType: it.normalized?.mime
      // gcsUri: it.normalized?.originalGcsUri,
      // mimeType: it.normalized?.originalMime
    })).filter(x => !!x.gcsUri),
  });

  Logger.log(res)
  return {
    ok: true,
    message: `Archivos subidos: ${res.items?.length || 0}. Pasá a etapa 2.`,
    payload: {
      rendicionId,
      mode: DRIVE_API_ENABLED ? 'driveFileIds' : 'gcsZip',
      // manifestGcsUri: res.manifestGcsUri || null
    }
  };
}


// function stage1_upload() {

//   const uploadedFiles = uploadDriveFolderToGCS_(FOLDER_ID, GCS_BUCKET, `${GCS_PREFIX}/raw/`)

//   const rendicionId = buildRendicionId_()

//   jobStateSet_({
//     rendicionId,
//     gcsPrefix: GCS_PREFIX,
//     normalizedItems: (uploadedFiles || []).map(it => ({
//       gcsUri: it.gcsUri,
//       mimeType: it.mimeType
//     })).filter(x => !!x.gcsUri),
//   });

//   return {
//     ok: true,
//     message: `Archivos subidos: ${uploadedFiles.length || 0}. Pasá a etapa 2.`
//   };
// }




function stage2_process() {
  const analyzeRes = stage2_process_analyze();
  if (!analyzeRes?.ok) return analyzeRes;
  return stage2_write_results(analyzeRes.items || analyzeRes.payload?.items || []);
}

function stage2_process_analyze() {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems guardados).' };
  }

  const uploadedFiles = st.normalizedItems;
  const items = callGemini(uploadedFiles);   // GeminiClient.gs (tu implementación)

  return {
    ok: true,
    message: `Comprobantes detectados: ${items?.length || 0}.`,
    items,
    payload: { items }
  };
}

function stage2_write_results(items) {
  if (!items || !items.length) {
    return { ok: false, message: 'No hay comprobantes para escribir.' };
  }
  writeItemsToSheet(items);                  // SheetWriter.gs (tu implementación)
  return { ok: true, message: `Planilla actualizada: ${items?.length || 0} comprobantes escritos.`, payload: { count: items.length } };
}


function stage3_finalize() {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems).' };
  }

  const coverRes = stage3_create_cover();
  if (!coverRes?.ok) return coverRes;

  const genRes = stage3_generate_outputs(coverRes.payload?.coverGcsUri || null);
  if (!genRes?.ok) return genRes;

  const dlRes = stage3_download_outputs(genRes.payload?.pdfUri || null, genRes.payload?.xlsmUri || null);
  if (!dlRes?.ok) return dlRes;

  return { ok: true, message: 'Completado: PDF + Excel generados (Ver Drive).', payload: dlRes.payload || {} };
}

function stage3_create_cover() {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems).' };
  }

  // Bloqueante: no seguimos si faltan ticks
  assertAllChecksOk_();

  const pdfBlob = buildCoverPdf_(SS_ID, getSheetGidByName_('Caratula'), 'cover.pdf');
  const coverGcsUri = uploadBlobToGCS_(GCS_BUCKET, `rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${RENDICION_MONTH}/cover.pdf`, pdfBlob);

  return { ok: true, message: 'Carátula creada.', payload: { coverGcsUri } };
}

function stage3_generate_outputs(coverGcsUri) {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay items).' };
  }

  // 1) checks
  assertAllChecksOk_();

  // 2) celdas a mandar (solo FIELD_COLUMN_MAP)
  const xlsmValues = buildXlsmValuesFromSheet_();
  if (!xlsmValues.length) throw new Error('No hay valores para mandar al XLSM.');

  const req = {
    rendicionId: st.rendicionId,
    inputs: {
      cover: coverGcsUri ? { gcsUri: coverGcsUri } : undefined,
      normalizedItems: st.normalizedItems,
      xlsmTemplate: {gcsUri: 'gs://scz-uy-rendiciones/templates/rendiciones_macro_template.xlsm'}, 
      xlsmValues
    },
    output: {
      // driveFolderId: OUTPUT_FOLDER_ID,
      gcsPrefix: st.gcsPrefix
    },
    options: {
      pdfName: `RENDREQ-XXXX_${st.rendicionId}.pdf`,
      xlsmName: `RENDREQ-XXXX_${st.rendicionId}.xlsm`,
      mergeOrder: "cover_first"
    }
  };

  const cleanReq = JSON.parse(JSON.stringify(req));
  const res = callCloudRunJson_('/v1/finalize', cleanReq);
  if (!res?.ok) throw new Error(`Finalize ok=false: ${JSON.stringify(res)}`);

  const pdfUri = res.pdf?.gcsUri || null;
  const xlsmUri = res.xlsm?.gcsUri || null;

  return {
    ok: true,
    message: 'PDF/Excel generados en GCS.',
    payload: { pdfUri, xlsmUri }
  };
}

function stage3_download_outputs(pdfUri, xlsmUri) {
  const st = jobStateGet_();
  if (!st?.rendicionId) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay rendicionId).' };
  }

  const saved = {};

  if (pdfUri) {
    saveGcsFileToDrive_(pdfUri, `RENDREQ-XXXX_${st.rendicionId}.pdf`, getOutputFolderId_());
    saved.pdf = true;
  }
  if (xlsmUri) {
    saveGcsFileToDrive_(xlsmUri, `RENDREQ-XXXX_${st.rendicionId}.xlsm`, getOutputFolderId_());
    saved.xlsm = true;
  }

  return { ok: true, message: 'Archivos descargados a Drive.', payload: { saved, pdfUri, xlsmUri } };
}



// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('Automation')
//     .addItem('Process test folder', 'showProcessingDialog')
//     .addToUi();
// }

// function showProcessingDialog() {
//   const html = HtmlService.createHtmlOutputFromFile('Processing')
//     .setWidth(320)
//     .setHeight(180);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Procesando…');
// }

// function processFolder() {
//   const ss = SpreadsheetApp.getActive();
//   ss.toast('Procesando archivos de la carpeta...', 'Automation', 60);

//   const uploadedFiles = uploadDriveFolderToGCS_(FOLDER_ID, GCS_BUCKET, 'facturas')

//   const items = callGemini(uploadedFiles);     // GeminiClient.gs
//   writeItemsToSheet(items);                // SheetWriter.gs

//   ss.toast('Procesamiento completo: ' + items.length + ' comprobantes.', 'Automation', 10);
//   return { ok: true, count: items.length, startRow: START_ROW };
// }

// function debugWhoAmI() {
//   const token = getServiceAccountAccessToken();

//   const resp = UrlFetchApp.fetch(
//     'https://oauth2.googleapis.com/tokeninfo?access_token=' + token
//   );

//   Logger.log(resp.getContentText());
// }
