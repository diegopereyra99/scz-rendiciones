function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automatización')
    .addItem('1) Subir archivos (para la IA)', 'ui_stage1_upload')
    .addItem('2) Procesar (IA → completar planilla)', 'ui_stage2_process')
    .addItem('3) Generar Excel + PDF', 'ui_stage3_finalize')
    .addSeparator()
    // .addItem('Process test folder', 'showProcessingDialog') // tu item viejo si querés
    .addToUi();
}

/** ---------- UI wrappers (abren HTML) ---------- */

function ui_stage1_upload() {
  return showModal_('stage_wait', {
    title: 'Subiendo archivos…',
    action: 'stage1_upload',
    description: 'Subiendo y normalizando adjuntos para que la IA los procese.'
  });
}

function ui_stage2_process() {
  return showModal_('stage_wait', {
    title: 'Procesando con IA…',
    action: 'stage2_process',
    description: 'Extracción + warnings + volcado a la planilla.'
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
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems guardados).' };
  }

  // Si tu callGemini esperaba "uploadedFiles" con signedUrl / driveId,
  // acá adaptás. Por ahora le pasamos los GCS URIs + mime.
  const uploadedFiles = st.normalizedItems;

  // REUSO de tu pipeline existente:
  const items = callGemini(uploadedFiles);   // GeminiClient.gs (tu implementación)
  writeItemsToSheet(items);                  // SheetWriter.gs (tu implementación)

  return { ok: true, message: `Planilla actualizada: ${items?.length || 0} comprobantes encontrados.` };
}


function stage3_finalize() {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems).' };
  }

  // 1) checks
  assertAllChecksOk_();

  // 2) celdas a mandar (solo FIELD_COLUMN_MAP)
  const xlsmValues = buildXlsmValuesFromSheet_();
  if (!xlsmValues.length) throw new Error('No hay valores para mandar al XLSM.');

  // 3) cover: por ahora lo dejo opcional (si ya lo generás en tu Cloud Run, genial)
  const pdfBlob = buildCoverPdf_(SS_ID, getSheetGidByName_('Formulario'), 'cover.pdf');
  const coverGcsUri = uploadBlobToGCS_(GCS_BUCKET, `rendiciones/${RENDICION_YEAR}/${RENDICION_USER}/${RENDICION_MONTH}/cover.pdf`, pdfBlob);

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

  // limpia undefined para no romper payloads estrictos
  const cleanReq = JSON.parse(JSON.stringify(req));

  const res = callCloudRunJson_('/v1/finalize', cleanReq);
  if (!res?.ok) throw new Error(`Finalize ok=false: ${JSON.stringify(res)}`);

  // Descargar el pdf y xlsm a OUTPUT_FOLDER_ID
  if (res.pdf?.gcsUri) {
    saveGcsFileToDrive_(
      res.pdf.gcsUri,
      `RENDREQ-XXXX_${st.rendicionId}.pdf`,
      OUTPUT_FOLDER_ID
    );
  }

  if (res.xlsm?.gcsUri) {
    saveGcsFileToDrive_(
      res.xlsm.gcsUri,
      `RENDREQ-XXXX_${st.rendicionId}.xlsm`,
      OUTPUT_FOLDER_ID
    );
  }

  return { ok: true, message: 'Completado: PDF + Excel generados (Ver Drive).'};
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

