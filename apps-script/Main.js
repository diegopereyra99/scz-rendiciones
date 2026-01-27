function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automatización')
    .addItem('1) Procesar Comprobantes efectivo', 'ui_process_cash')
    .addItem('1) Procesar Comprobantes tarjeta corporativa', 'ui_process_card')
    .addItem('2) Generar Excel + PDF', 'ui_stage3_finalize')
    .addItem('Reiniciar todo', 'clearAllRows')
    .addToUi();
}

/** ---------- UI wrappers (abren HTML) ---------- */

function ui_process_cash() {
  return showModal_('stage_wait', {
    title: 'Procesando archivos…',
    action: 'stage_upload_and_process_cash',
    description: 'Preprocesando, analizando con IA y escribiendo la planilla (efectivo).'
  });
}

function ui_process_card() {
  return showModal_('stage_wait', {
    title: 'Procesando archivos…',
    action: 'stage_upload_and_process_card',
    description: 'Preprocesando, analizando con IA y escribiendo la planilla (tarjeta).'
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

function stage1_upload(mode) {
  return runStage1Upload_(mode || 'efectivo');
}

function stage_upload_and_process_cash() {
  return stage_upload_and_process_('efectivo');
}

function stage_upload_and_process_card() {
  return stage_card_low_writes_();
}

function stage_upload_and_process_(mode) {
  const uploadRes = runStage1Upload_(mode);
  if (!uploadRes?.ok) return uploadRes;

  const analyzeRes = stage2_process_analyze_by_mode_(mode);
  if (!analyzeRes?.ok) return analyzeRes;

  const items = analyzeRes.items || analyzeRes.payload?.items || [];
  let writeRes = { ok: true, message: analyzeRes.message || null, payload: null };
  if (!analyzeRes.skipWrite) {
    writeRes = stage2_write_results(items);
    if (!writeRes?.ok) return writeRes;
  }

  return {
    ok: true,
    message: writeRes.message || analyzeRes.message || 'Archivos procesados y planilla actualizada.',
    payload: {
      upload: uploadRes.payload || null,
      process: analyzeRes.payload || null,
      write: writeRes.payload || null
    }
  };
}

/** ---------- Tarjeta low-writes pipeline ---------- */

function stage_card_low_writes_() {
  try {
    const rendicionId = buildRendicionId_();
    const statementFile = getSingleStatementFile_();
    const statementFingerprint = buildStatementFingerprint_(statementFile);

    let st = jobStateGet_();
    let resetNotice = null;
    if (st?.statementFingerprint && st.statementFingerprint !== statementFingerprint) {
      resetNotice = 'Cambió el estado de cuenta, se reinicia la rendición.';
      resetFullJob_();
      st = jobStateGet_();
    }

    const prevBaseRowCount = st?.baseRowCount || 0;
    const prevOrphansCount = st?.lastOrphansRowCount || 0;

    const folderId = getModeFolderId_('tarjeta');
    const receiptFiles = listDriveFiles_(folderId);
    const receiptsFingerprint = buildReceiptsFingerprint_(receiptFiles);

    const baseState = Object.assign({}, st, {
      rendicionId,
      mode: 'tarjeta',
      sourceFolderId: folderId,
      statementFingerprint,
      receiptsFingerprint,
      lastRunAt: new Date().toISOString()
    });
    jobStateSet_(baseState);

    const statementRes = processStatementLowWrites_(statementFile, statementFingerprint, baseState);
    if (!statementRes?.ok) return statementRes;
    const statementData = statementRes.statementData || {};

    const planRes = buildRowsPlanFromStatement_(statementData);
    if (!planRes?.ok) return planRes;
    const rowsPlan = planRes.rowsPlan || [];
    const lineToSheetRow = planRes.lineToSheetRow || {};

    jobStateSet_(Object.assign({}, baseState, {
      lastEstadoCuenta: statementData,
      lastStatementFile: {
        id: statementFile.id,
        size: statementFile.size,
        updated: statementFile.updated
      },
      lineToSheetRow,
      reductionsMeta: planRes.reductionsMeta || null,
      baseRowCount: rowsPlan.length
    }));

    const baseWriteRes = writeBaseTableLowWrites_(rowsPlan);
    if (!baseWriteRes?.ok) return baseWriteRes;

    const uploadRes = runStage1Upload_('tarjeta');
    if (!uploadRes?.ok) return uploadRes;

    const stAfterUpload = jobStateGet_();
    const normalizedItems = getNormalizedItemsForMode_(stAfterUpload, 'tarjeta');

    const receiptsCacheKey = buildReceiptsCacheKey_(statementFingerprint, receiptsFingerprint);
    const patchesRes = processReceiptsToPatches_(statementData, normalizedItems, lineToSheetRow, stAfterUpload, receiptsCacheKey);
    if (!patchesRes?.ok) return patchesRes;

    const applyRes = applyPatchesLowWrites_(patchesRes.rowPatches, patchesRes.rowStatus, rowsPlan.length);
    if (!applyRes?.ok) return applyRes;

    const prevOrphansStart = START_ROW + prevBaseRowCount;
    const orphansRes = writeOrphansConflictsSection_(
      patchesRes.orphans,
      patchesRes.conflicts,
      rowsPlan.length,
      prevOrphansStart,
      prevOrphansCount
    );
    if (!orphansRes?.ok) return orphansRes;

    jobStateSet_(Object.assign({}, jobStateGet_(), {
      baseRowCount: rowsPlan.length,
      lineToSheetRow,
      assignedReceiptBySheetRow: patchesRes.assignedReceiptBySheetRow || null,
      orphansSummary: patchesRes.orphans || [],
      conflictsSummary: patchesRes.conflicts || [],
      fieldConflictsSummary: applyRes?.payload?.fieldConflicts || [],
      lastOrdenArchivos: patchesRes.orderList && patchesRes.orderList.length ? patchesRes.orderList : null,
      lastOrphansRowCount: orphansRes?.payload?.count || 0,
      statementFingerprint,
      receiptsFingerprint,
      lastRunAt: new Date().toISOString()
    }));

    const msgParts = [];
    if (resetNotice) msgParts.push(resetNotice);
    msgParts.push(`Base: ${rowsPlan.length} filas.`);
    msgParts.push(`Comprobantes: ${(normalizedItems || []).length}.`);
    msgParts.push(`Orphans: ${(patchesRes.orphans || []).length}, Conflicts: ${(patchesRes.conflicts || []).length}.`);

    return {
      ok: true,
      message: msgParts.join(' '),
      payload: {
        base: baseWriteRes.payload || null,
        upload: uploadRes.payload || null,
        orphans: patchesRes.orphans?.length || 0,
        conflicts: patchesRes.conflicts?.length || 0
      }
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function resetFullJob_() {
  clearAllRows();
}

function resetJobState_() {
  PropertiesService.getDocumentProperties().deleteProperty('REN_JOB_STATE');
}

const CACHE_KEYS_PROP = 'REN_CACHE_KEYS';

function cacheGetJson_(key) {
  if (!key) return null;
  const cache = CacheService.getDocumentCache();
  const raw = cache.get(key);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function cachePutJson_(key, value, ttlSeconds) {
  if (!key) return;
  const cache = CacheService.getDocumentCache();
  const raw = JSON.stringify(value || {});
  if (raw.length > 90000) return; // avoid cache limit issues
  cache.put(key, raw, ttlSeconds || 21600);
  trackCacheKey_(key);
}

function trackCacheKey_(key) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(CACHE_KEYS_PROP);
  const list = raw ? JSON.parse(raw) : [];
  if (list.indexOf(key) === -1) list.push(key);
  props.setProperty(CACHE_KEYS_PROP, JSON.stringify(list));
}

function resetCache_() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(CACHE_KEYS_PROP);
  if (!raw) return;
  const list = JSON.parse(raw);
  const cache = CacheService.getDocumentCache();
  list.forEach((key) => cache.remove(key));
  props.deleteProperty(CACHE_KEYS_PROP);
}

function buildStatementFingerprint_(statementFile) {
  if (!statementFile) return '';
  return `${statementFile.id}:${statementFile.updated}:${statementFile.size}`;
}

function buildReceiptsFingerprint_(files) {
  const hash = buildFilesDigest_(files || []);
  return hash ? `receipts:${hash}` : 'receipts:empty';
}

function buildFilesDigest_(files) {
  const parts = (files || []).map((f) => `${f.id}:${f.updated || ''}:${f.size || ''}`);
  parts.sort();
  const raw = parts.join('|');
  if (!raw) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return digest.map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function buildReceiptsCacheKey_(statementFingerprint, receiptsFingerprint) {
  if (!statementFingerprint || !receiptsFingerprint) return '';
  return `REN_RECEIPTS_${statementFingerprint}_${receiptsFingerprint}`;
}

function processStatementLowWrites_(statementFile, fingerprint, st) {
  const cacheKey = `REN_STATEMENT_${fingerprint}`;
  const cached = cacheGetJson_(cacheKey);
  if (cached) {
    return { ok: true, statementData: cached, reused: true, cache: true };
  }
  if (st?.statementFingerprint === fingerprint && st?.lastEstadoCuenta) {
    return { ok: true, statementData: st.lastEstadoCuenta, reused: true, cache: false };
  }

  const statementReq = {
    rendicionId: st?.rendicionId || buildRendicionId_(),
    statement: {
      driveFileId: statementFile.id,
      mime: statementFile.mimeType || null
    }
  };

  const statementRes = callCloudRunJson_('/v1/process_statement', statementReq);
  if (!statementRes?.ok) throw new Error(`process_statement ok=false: ${JSON.stringify(statementRes)}`);
  const statementData = statementRes.data || {};

  cachePutJson_(cacheKey, statementData, 21600);
  return { ok: true, statementData, reused: false, cache: false };
}

function buildRowsPlanFromStatement_(statementData) {
  const txs = Array.isArray(statementData?.transacciones) ? statementData.transacciones : [];
  if (!txs.length) {
    return { ok: false, message: 'No se pudieron extraer líneas del estado de cuenta.' };
  }

  const rowsPlan = [];
  const lineToSheetRow = {};
  const reductionsMeta = [];

  txs.forEach((tx, idx) => {
    const lineIndex = idx + 1;
    const detail = tx?.detalle || tx?.descripcion || tx?.concepto || '';
    const amountRes = mapStatementAmount_(tx);
    const isReduc = isReducIvaLine_(detail);

    if (isReduc && rowsPlan.length) {
      const last = rowsPlan[rowsPlan.length - 1];
      const prev = getNumeric_(last['Descuentos']) || 0;
      const add = Math.abs(Number(amountRes.amount || 0));
      last['Descuentos'] = prev + add;
      lineToSheetRow[lineIndex] = START_ROW + rowsPlan.length - 1;
      reductionsMeta.push({ lineIndex, appliedTo: lineToSheetRow[lineIndex], amount: add, detail });
      return;
    }

    if (isReduc && !rowsPlan.length) {
      reductionsMeta.push({ lineIndex, appliedTo: null, amount: Math.abs(Number(amountRes.amount || 0)), detail, orphan: true });
      return;
    }

    const row = {
      'Fecha de factura': tx?.fecha || null,
      'Proveedor': detail || null,
      'Importe a rendir': amountRes.amount,
      'Moneda': amountRes.currency,
      'Descuentos': 0
    };
    row.__status = 'MISSING';
    rowsPlan.push(row);
    lineToSheetRow[lineIndex] = START_ROW + rowsPlan.length - 1;
  });

  return { ok: true, rowsPlan, lineToSheetRow, reductionsMeta };
}

function isReducIvaLine_(text) {
  if (!text) return false;
  return String(text).toLowerCase().indexOf('reduc. iva ley') !== -1;
}

function processReceiptsToPatches_(statementData, normalizedItems, lineToSheetRow, st, cacheKey) {
  if (cacheKey) {
    const cached = cacheGetJson_(cacheKey);
    if (cached) return Object.assign({ ok: true, cached: true }, cached);
  }
  const items = normalizedItems || [];
  if (!items.length) {
    return {
      ok: true,
      message: 'No hay comprobantes para procesar.',
      rowPatches: {},
      rowStatus: {},
      orphans: [],
      conflicts: [],
      assignedReceiptBySheetRow: {}
    };
  }

  const receipts = items.map((it) => ({
    gcsUri: it.gcsUri,
    mime: it.mime || it.mimeType || null
  })).filter((it) => !!it.gcsUri);

  const rows = [];
  const batches = chunkArray_(receipts, PROCESS_BATCH_SIZE);
  batches.forEach((batch) => {
    const receiptsReq = {
      rendicionId: st?.rendicionId || buildRendicionId_(),
      mode: 'tarjeta',
      statement: { parsed: statementData },
      receipts: batch
    };
    const receiptsRes = callCloudRunJson_('/v1/process_receipts_batch', receiptsReq);
    if (!receiptsRes?.ok) throw new Error(`process_receipts_batch ok=false: ${JSON.stringify(receiptsRes)}`);
    if (Array.isArray(receiptsRes.rows)) rows.push.apply(rows, receiptsRes.rows);
  });

  const receiptItems = flattenDocflowRows_(rows);
  const orderList = buildOrderListFromItems_(receiptItems);
  const rowPatches = {};
  const rowStatus = {};
  const orphans = [];
  const conflicts = [];
  const assignedReceiptBySheetRow = {};

  receiptItems.forEach((item) => {
    const idx = extractStatementIdx_(item);
    const source = extractSourceKey_(item) || '';
    const targetRow = idx && lineToSheetRow ? lineToSheetRow[idx] : null;

    if (targetRow) {
      if (assignedReceiptBySheetRow[targetRow] && assignedReceiptBySheetRow[targetRow] !== source) {
        rowStatus[targetRow] = 'CONFLICT';
        conflicts.push({
          type: 'duplicate_match',
          source,
          matchIndex: idx,
          targetRow,
          fields: scrubInternalKeys_(Object.assign({}, item))
        });
        return;
      }
      assignedReceiptBySheetRow[targetRow] = source;
      const patch = buildReceiptPatch_(item);
      if (!rowPatches[targetRow]) rowPatches[targetRow] = {};
      Object.keys(patch).forEach((k) => {
        if (rowPatches[targetRow][k] === undefined || rowPatches[targetRow][k] === '' || rowPatches[targetRow][k] === null) {
          rowPatches[targetRow][k] = patch[k];
        }
      });
      if (!rowStatus[targetRow]) rowStatus[targetRow] = 'MATCHED';
      return;
    }

    orphans.push({
      source,
      matchIndex: idx || null,
      observacion: extractReceiptObservation_(item),
      fields: scrubInternalKeys_(Object.assign({}, item))
    });
  });

  const result = {
    ok: true,
    message: `Comprobantes procesados: ${receiptItems.length || 0}.`,
    rowPatches,
    rowStatus,
    orphans,
    conflicts,
    assignedReceiptBySheetRow,
    orderList
  };
  if (cacheKey) {
    cachePutJson_(cacheKey, result, 21600);
  }
  return result;
}

function buildReceiptPatch_(item) {
  const fields = item?.fields || item || {};
  const out = {};
  Object.keys(fields).forEach((key) => {
    if (key === 'Warnings' || key === 'warnings') return;
    if (key === 'Estado de cuenta' || key === 'Estado_de_cuenta') return;
    if (key === '__statement_idx' || key === '__doc_names') return;
    out[key] = fields[key];
  });
  return out;
}

function extractReceiptObservation_(item) {
  const st = item?.['Estado de cuenta'] || item?.estadoCuenta || item?.estado_cuenta || null;
  const obs = st?.observacion || item?.observacion || item?.observation || null;
  return obs || '';
}

function runStage1Upload_(mode) {
  const rendicionId = buildRendicionId_();
  const folderId = getModeFolderId_(mode);
  const driveFiles = listDriveFiles_(folderId);
  if (!driveFiles.length) return { ok: false, message: 'No hay archivos en la carpeta.' };

  const st = jobStateGet_();
  const prevByMode = st?.normalizedItemsByMode || {};
  const prevItems = prevByMode[mode] || [];
  const prevIds = prevItems.map((it) => it.driveFileId).filter((id) => !!id);
  const prevSet = new Set(prevIds);

  const currentIds = driveFiles.map((f) => f.id);
  const currentSet = new Set(currentIds);
  const addedIds = currentIds.filter((id) => !prevSet.has(id));
  const removedIds = prevIds.filter((id) => !currentSet.has(id));

  if (!addedIds.length && !removedIds.length && prevItems.length) {
    jobStateSet_(Object.assign({}, st, { rendicionId, mode, sourceFolderId: folderId }));
    return {
      ok: true,
      message: 'No hay archivos nuevos; se reutilizan los normalizados existentes.',
      payload: { rendicionId, mode, reused: true }
    };
  }

  let input = null;
  let filesForNormalize = driveFiles;

  if (DRIVE_API_ENABLED) {
    filesForNormalize = driveFiles.filter((f) => addedIds.indexOf(f.id) !== -1);
    if (!filesForNormalize.length && removedIds.length) {
      const remaining = prevItems.filter((it) => !removedIds.includes(it.driveFileId));
      const nextByMode = Object.assign({}, prevByMode, { [mode]: remaining });
      jobStateSet_(Object.assign({}, st, { rendicionId, mode, sourceFolderId: folderId, normalizedItemsByMode: nextByMode }));
      return {
        ok: true,
        message: 'Se actualizaron eliminaciones; no hay archivos nuevos.',
        payload: { rendicionId, mode, removed: removedIds }
      };
    }
    input = { driveFileIds: filesForNormalize.map(f => f.id) };
  } else {
    filesForNormalize = driveFiles.filter((f) => addedIds.indexOf(f.id) !== -1);
    if (!filesForNormalize.length && removedIds.length) {
      const remaining = prevItems.filter((it) => !removedIds.includes(it.driveFileId));
      const nextByMode = Object.assign({}, prevByMode, { [mode]: remaining });
      jobStateSet_(Object.assign({}, st, { rendicionId, mode, sourceFolderId: folderId, normalizedItemsByMode: nextByMode }));
      return {
        ok: true,
        message: 'Se actualizaron eliminaciones; no hay archivos nuevos.',
        payload: { rendicionId, mode, removed: removedIds }
      };
    }
    const zipBlob = buildZipFromDriveFolder_(folderId, filesForNormalize);
    const objectName = buildInputsZipObjectName_(mode);
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

  const driveIdByName = new Map(driveFiles.map((f) => [f.name, f.id]));
  const newItems = (res.items || []).map((it, idx) => ({
    gcsUri: it.normalized?.gcsUri,
    mime: it.normalized?.mime,
    mimeType: it.normalized?.mime, // compat
    normalizedIndex: String(idx).padStart(4, '0'),
    originalName: it.source?.originalName || null,
    driveFileId: it.source?.driveFileId || (it.source?.originalName ? driveIdByName.get(it.source.originalName) : null)
  })).filter(x => !!x.gcsUri);

  const remaining = prevItems.filter((it) => !removedIds.includes(it.driveFileId));
  const merged = remaining.concat(newItems);
  const nextByMode = Object.assign({}, prevByMode, { [mode]: merged });

  jobStateSet_(Object.assign({}, st, {
    rendicionId,
    gcsPrefix: GCS_PREFIX,
    mode,
    sourceFolderId: folderId,
    manifestGcsUri: res.manifestGcsUri || null,
    normalizedItemsByMode: nextByMode
  }));

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
  const mode = st?.mode || 'efectivo';
  const analyzeRes = stage2_process_analyze_by_mode_(mode);
  if (!analyzeRes?.ok) return analyzeRes;
  if (analyzeRes.skipWrite) return analyzeRes;
  return stage2_write_results(analyzeRes.items || analyzeRes.payload?.items || []);
}

function stage2_process_analyze_by_mode_(mode) {
  if (mode === 'efectivo') return stage2_process_analyze_efectivo_();
  if (mode === 'tarjeta') {
    return stage2_process_analyze_tarjeta_();
  }
  return { ok: false, message: `Modo inválido: ${mode}` };
}

function stage2_process_analyze_by_mode(mode) {
  return stage2_process_analyze_by_mode_(mode);
}

function stage2_process_statement_() {
  const st = jobStateGet_();
  if (!st?.rendicionId) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay rendicionId).' };
  }

  const statementFile = getSingleStatementFile_();
  if (!statementFile) return { ok: false, message: 'No hay archivo de estado de cuenta.' };
  const prevFile = st?.lastStatementFile || null;
  const sameFile = !!(prevFile &&
    prevFile.id === statementFile.id &&
    prevFile.size === statementFile.size &&
    prevFile.updated === statementFile.updated);

  if (sameFile && st?.lastEstadoCuenta) {
    return {
      ok: true,
      message: 'Estado de cuenta sin cambios; se reutiliza el último.',
      items: buildStatementItems_(st.lastEstadoCuenta),
      payload: { statement: st.lastEstadoCuenta, reused: true }
    };
  }

  const statementReq = {
    rendicionId: st.rendicionId,
    statement: {
      driveFileId: statementFile.id,
      mime: statementFile.mimeType || null
    }
  };

  const statementRes = callCloudRunJson_('/v1/process_statement', statementReq);
  if (!statementRes?.ok) throw new Error(`process_statement ok=false: ${JSON.stringify(statementRes)}`);

  const statementData = statementRes.data || {};
  const statementItems = buildStatementItems_(statementData);
  if (!statementItems.length) {
    return { ok: false, message: 'No se pudieron extraer líneas del estado de cuenta.' };
  }

  writeStatementToSheet_(statementData);
  jobStateSet_(Object.assign({}, st, {
    lastEstadoCuenta: statementData || null,
    lastStatementFile: {
      id: statementFile.id,
      size: statementFile.size,
      updated: statementFile.updated
    }
  }));

  return {
    ok: true,
    message: 'Estado de cuenta procesado.',
    items: statementItems,
    payload: { statement: statementData }
  };
}

function stage2_process_statement() {
  return stage2_process_statement_();
}

function stage2_process_receipts_tarjeta_(options) {
  const st = jobStateGet_();
  if (!st?.rendicionId) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay rendicionId).' };
  }

  const statementData = st?.lastEstadoCuenta || null;
  if (!statementData) {
    return { ok: false, message: 'Primero procesá el estado de cuenta.' };
  }

  const statementItems = buildStatementItems_(statementData);
  const normalizedItems = getNormalizedItemsForMode_(st, 'tarjeta');
  const receipts = (normalizedItems || []).map((it) => ({
    gcsUri: it.gcsUri,
    mime: it.mime || it.mimeType || null
  })).filter((it) => !!it.gcsUri);

  if (!receipts.length) {
    return {
      ok: true,
      message: 'Estado de cuenta escrito. No hay comprobantes para procesar.',
      items: statementItems,
      orphanRowIndexes: [],
      skipWrite: true,
      noReceipts: true,
      payload: { statement: statementData }
    };
  }

  const rows = [];
  const batches = chunkArray_(receipts, PROCESS_BATCH_SIZE);
  batches.forEach((batch) => {
    const receiptsReq = {
      rendicionId: st.rendicionId,
      mode: 'tarjeta',
      statement: { parsed: statementData },
      receipts: batch
    };
    const receiptsRes = callCloudRunJson_('/v1/process_receipts_batch', receiptsReq);
    if (!receiptsRes?.ok) throw new Error(`process_receipts_batch ok=false: ${JSON.stringify(receiptsRes)}`);
    if (Array.isArray(receiptsRes.rows)) rows.push.apply(rows, receiptsRes.rows);
  });

  const receiptItems = flattenDocflowRows_(rows);
  const orderList = buildOrderListFromItems_(receiptItems);
  const merged = mergeStatementWithReceipts_(statementItems, receiptItems);
  const doWrite = !(options && options.write === false);

  if (doWrite) {
    writeItemsToSheet(merged.items);
    markRowsColor_(merged.orphanRowIndexes, COLOR_ORPHAN);
  }

  jobStateSet_(Object.assign({}, st, {
    lastEstadoCuenta: statementData || null,
    lastOrdenArchivos: orderList && orderList.length ? orderList : st.lastOrdenArchivos || null
  }));

  return {
    ok: true,
    message: `Tarjeta procesada: ${merged.items.length || 0} filas.`,
    items: merged.items,
    orphanRowIndexes: merged.orphanRowIndexes,
    skipWrite: !doWrite,
    payload: {
      statement: statementData,
      rows: rows
    }
  };
}

function stage2_process_receipts_tarjeta(options) {
  return stage2_process_receipts_tarjeta_(options);
}

function stage2_process_analyze_efectivo_() {
  const st = jobStateGet_();
  const normalizedItems = getNormalizedItemsForMode_(st, 'efectivo');
  if (!st?.rendicionId || !normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems guardados).' };
  }

  const receipts = (normalizedItems || []).map((it) => ({
    gcsUri: it.gcsUri,
    mime: it.mime || it.mimeType || null
  })).filter((it) => !!it.gcsUri);

  const rows = [];
  const batches = chunkArray_(receipts, PROCESS_BATCH_SIZE);
  batches.forEach((batch) => {
    const req = {
      rendicionId: st.rendicionId,
      mode: 'efectivo',
      receipts: batch
    };
    const res = callCloudRunJson_('/v1/process_receipts_batch', req);
    if (!res?.ok) throw new Error(`process_receipts_batch ok=false: ${JSON.stringify(res)}`);
    if (Array.isArray(res.rows)) rows.push.apply(rows, res.rows);
  });

  const items = sortItemsByDate_(flattenDocflowRows_(rows));
  const orderList = buildOrderListFromItems_(items);
  jobStateSet_(Object.assign({}, st, { lastOrdenArchivos: orderList || null }));

  return {
    ok: true,
    message: `Comprobantes detectados: ${items.length || 0}.`,
    items,
    payload: {
      rows,
      items
    }
  };
}

function stage2_process_analyze_tarjeta_() {
  const statementRes = stage2_process_statement_();
  if (!statementRes?.ok) return statementRes;

  const receiptsRes = stage2_process_receipts_tarjeta_();
  if (!receiptsRes?.ok) return receiptsRes;

  return Object.assign({}, receiptsRes, { skipWrite: true });
}

function stage2_process_analyze() {
  const st = jobStateGet_();
  if (!st?.rendicionId || !st?.normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems guardados).' };
  }

  const uploadedFiles = st.normalizedItems;
  const gemRes = callGemini(uploadedFiles);   // GeminiClient.gs (tu implementación)
  const comprobantes = extractComprobantes_(gemRes);
  const ordenArchivos = extractOrdenArchivos_(gemRes);
  const estadoCuenta = extractEstadoCuenta_(gemRes);

  // Guardamos orden y estado para usarlos en finalize
  jobStateSet_(Object.assign({}, st, {
    lastOrdenArchivos: ordenArchivos || null,
    lastEstadoCuenta: estadoCuenta || null
  }));

  return {
    ok: true,
    message: `Comprobantes detectados: ${comprobantes?.length || 0}.`,
    items: comprobantes,
    payload: {
      items: comprobantes,
      ordenArchivos,
      estadoCuenta
    }
  };
}

function flattenDocflowRows_(rows) {
  if (!Array.isArray(rows)) return [];
  const out = [];
  rows.forEach((row) => {
    if (!row) return;
    const data = row.data !== undefined ? row.data : row;
    const docNames = row?.meta?.docs || null;
    if (Array.isArray(data)) {
      data.forEach((item) => {
        if (docNames) item.__doc_names = docNames;
        out.push(item);
      });
    } else if (data) {
      if (docNames) data.__doc_names = docNames;
      out.push(data);
    }
  });
  return out;
}

function getNormalizedItemsForMode_(st, mode) {
  if (!st) return [];
  const byMode = st.normalizedItemsByMode || {};
  if (Array.isArray(byMode[mode]) && byMode[mode].length) return byMode[mode];
  if (Array.isArray(st.normalizedItems) && st.normalizedItems.length) return st.normalizedItems;
  return [];
}

function chunkArray_(arr, size) {
  const out = [];
  const n = Math.max(1, size || 1);
  for (let i = 0; i < (arr || []).length; i += n) {
    out.push(arr.slice(i, i + n));
  }
  return out;
}

function parseIsoDate_(value) {
  if (!value) return null;
  const s = String(value).trim();
  if (!s) return null;
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return d;
}

function sortItemsByDate_(items) {
  const list = (items || []).map((item, idx) => {
    const d = parseIsoDate_(item?.['Fecha de factura'] || item?.fecha);
    return { item, idx, date: d };
  });
  list.sort((a, b) => {
    if (!a.date && !b.date) return a.idx - b.idx;
    if (!a.date) return 1;
    if (!b.date) return -1;
    if (a.date.getTime() === b.date.getTime()) return a.idx - b.idx;
    return a.date.getTime() - b.date.getTime();
  });
  return list.map((e) => e.item);
}

function buildOrderListFromItems_(items) {
  const firstDateByKey = new Map();
  (items || []).forEach((item) => {
    const key = extractSourceKey_(item);
    if (!key) return;
    const d = parseIsoDate_(item?.['Fecha de factura'] || item?.fecha);
    if (!d) return;
    if (!firstDateByKey.has(key) || d < firstDateByKey.get(key)) {
      firstDateByKey.set(key, d);
    }
  });
  const entries = Array.from(firstDateByKey.entries());
  entries.sort((a, b) => a[1] - b[1]);
  return entries.map((e) => e[0]);
}

function extractSourceKey_(item) {
  if (!item) return null;
  const docs = item.__doc_names;
  if (Array.isArray(docs) && docs.length) {
    const base = extractDocBasename_(docs[0]);
    const driveId = extractDriveIdFromDocName_(base);
    if (driveId) return driveId;
    return base || extractIndexFromDocName_(docs[0]) || docs[0];
  }
  if (typeof docs === 'string') {
    const base = extractDocBasename_(docs);
    const driveId = extractDriveIdFromDocName_(base);
    return driveId || base || extractIndexFromDocName_(docs) || docs;
  }
  return null;
}

function extractIndexFromDocName_(name) {
  if (!name) return null;
  const s = String(name).trim();
  if (!s) return null;
  const m = s.match(/(^|\/)(\d{4})_/);
  if (m) return `f${m[2]}`;
  const m2 = s.match(/f(\d{4})/i);
  if (m2) return `f${m2[1]}`;
  return null;
}

function extractDocBasename_(name) {
  if (!name) return null;
  const s = String(name).trim();
  if (!s) return null;
  const parts = s.split(/[\\/]/);
  return parts[parts.length - 1] || s;
}

function extractDriveIdFromDocName_(name) {
  if (!name) return null;
  const s = String(name).trim();
  if (!s) return null;
  const m = s.match(/drive_([A-Za-z0-9_-]+)/i);
  if (m) return m[1];
  return null;
}

function getSingleStatementFile_() {
  const folderId = getStatementFolderId_();
  if (!folderId) throw new Error(`No existe la carpeta: ${FOLDER_ESTADO_NAME}`);
  const files = listDriveFiles_(folderId);
  if (!files.length) throw new Error('No hay archivos en Estado de cuenta.');
  if (files.length > 1) throw new Error('Debe existir un solo archivo en Estado de cuenta.');
  const file = DriveApp.getFileById(files[0].id);
  return {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    size: file.getSize(),
    updated: file.getLastUpdated().getTime()
  };
}

function buildStatementItems_(data) {
  const txs = Array.isArray(data?.transacciones) ? data.transacciones : [];
  const out = [];
  txs.forEach((tx, idx) => {
    const res = mapStatementAmount_(tx);
    const item = {
      'Fecha de factura': tx?.fecha || null,
      'Proveedor': tx?.detalle || null,
      'Importe a rendir': res.amount,
      'Moneda': res.currency,
      'Warnings': res.warnings || []
    };
    item.__statement_idx = idx + 1;
    out.push(item);
  });
  return out;
}

function mapStatementAmount_(tx) {
  const warnings = [];
  if (tx?.importe_uyu !== null && tx?.importe_uyu !== undefined) {
    return { amount: tx.importe_uyu, currency: 'UYU', warnings };
  }
  if (tx?.importe_usd !== null && tx?.importe_usd !== undefined) {
    return { amount: tx.importe_usd, currency: 'USD', warnings };
  }
  if (tx?.importe_origen !== null && tx?.importe_origen !== undefined) {
    return { amount: tx.importe_origen, currency: 'USD', warnings };
  }
  warnings.push({ campo: 'Importe a rendir', mensaje: 'Importe no identificado en estado de cuenta.' });
  return { amount: null, currency: null, warnings };
}

function mergeStatementWithReceipts_(statementItems, receiptItems) {
  const out = statementItems.map((it) => Object.assign({}, it));
  const matched = new Set();
  const baseReceiptByIdx = new Map();
  const docNamesByIdx = new Map();
  const orphanRows = [];

  receiptItems.forEach((item) => {
    const idx = extractStatementIdx_(item);
    if (idx && idx >= 1 && idx <= out.length && !matched.has(idx)) {
      const base = out[idx - 1];
      checkStatementAmountMatch_(base, item);
      out[idx - 1] = mergeItemsPreferNonNull_(base, item);
      matched.add(idx);
      baseReceiptByIdx.set(idx, item);
      docNamesByIdx.set(idx, collectDocNames_(item));
      return;
    }
    if (idx && idx >= 1 && idx <= out.length && matched.has(idx)) {
      const base = baseReceiptByIdx.get(idx);
      const baseStatement = out[idx - 1];
      const baseDocs = docNamesByIdx.get(idx) || [];
      const nextDocs = collectDocNames_(item);
      const allDocs = baseDocs.concat(nextDocs).filter((v, i, a) => a.indexOf(v) === i);
      docNamesByIdx.set(idx, allDocs);
      if (!base || areItemsConcordant_(base, item)) {
        checkStatementAmountMatch_(baseStatement, item);
        out[idx - 1] = mergeItemsPreferNonNull_(out[idx - 1], item);
      } else {
        addWarning_(out[idx - 1], 'general',
          `Comprobantes con el mismo idx no concuerdan: ${allDocs.join(', ') || 'sin nombre'}.`);
      }
      return;
    }
    addWarning_(item, 'general', 'Comprobante sin movimiento en estado de cuenta.');
    addStatementObservation_(item, 'Comprobante sin movimiento en estado de cuenta.');
    out.push(item);
    orphanRows.push(out.length);
  });

  out.forEach((item, i) => {
    if (i >= statementItems.length) return;
    if (!matched.has(i + 1)) {
      addWarning_(item, 'general', 'Falta comprobante asociado a esta línea del estado de cuenta.');
      addStatementObservation_(item, 'Falta comprobante asociado a esta línea del estado de cuenta.');
    }
  });

  return { items: out.map((it) => scrubInternalKeys_(it)), orphanRowIndexes: orphanRows.map((n) => START_ROW + n - 1) };
}

function extractStatementIdx_(item) {
  const st = item?.['Estado de cuenta'] || item?.estadoCuenta || item?.estado_cuenta || null;
  const idx = st?.idx;
  if (idx === null || idx === undefined) return null;
  const num = parseInt(idx, 10);
  return Number.isFinite(num) ? num : null;
}

function collectDocNames_(item) {
  if (!item) return [];
  const docs = item.__doc_names;
  if (Array.isArray(docs)) return docs.map(String);
  if (docs) return [String(docs)];
  return [];
}

function areItemsConcordant_(a, b) {
  if (!a || !b) return true;
  const pairs = [
    ['Proveedor', 'Proveedor'],
    ['Moneda', 'Moneda'],
    ['Fecha de factura', 'Fecha de factura'],
    ['Numero de Factura', 'Numero de Factura'],
    ['Importe facturado', 'Importe facturado'],
    ['Importe a rendir', 'Importe a rendir']
  ];
  for (let i = 0; i < pairs.length; i++) {
    const keyA = pairs[i][0];
    const keyB = pairs[i][1];
    const va = a[keyA];
    const vb = b[keyB];
    if (va === null || va === undefined || va === '') continue;
    if (vb === null || vb === undefined || vb === '') continue;
    if (typeof va === 'number' && typeof vb === 'number') {
      if (Math.abs(va - vb) > 0.01) return false;
    } else {
      const sa = String(va).trim().toLowerCase();
      const sb = String(vb).trim().toLowerCase();
      if (sa && sb && sa !== sb) return false;
    }
  }
  return true;
}

function mergeItemsPreferNonNull_(base, extra) {
  const out = Object.assign({}, base);
  Object.keys(extra || {}).forEach((k) => {
    if (k === 'Warnings') return;
    const v = extra[k];
    if (k === 'Estado de cuenta' && v && typeof v === 'object') {
      const merged = mergeEstadoCuenta_(out[k], v);
      if (merged) out[k] = merged;
      return;
    }
    if (k === 'Importe a rendir' && out[k] !== null && out[k] !== undefined && out[k] !== '') {
      return;
    }
    if (v !== null && v !== undefined && v !== '') {
      out[k] = v;
    }
  });
  const mergedWarnings = [];
  const baseWarnings = base?.Warnings || [];
  const extraWarnings = extra?.Warnings || [];
  if (Array.isArray(baseWarnings)) mergedWarnings.push.apply(mergedWarnings, baseWarnings);
  if (Array.isArray(extraWarnings)) mergedWarnings.push.apply(mergedWarnings, extraWarnings);
  if (mergedWarnings.length) out.Warnings = mergedWarnings;
  return out;
}

function mergeEstadoCuenta_(base, extra) {
  if (!base && !extra) return null;
  const out = Object.assign({}, base || {});
  Object.keys(extra || {}).forEach((k) => {
    const v = extra[k];
    if (v !== null && v !== undefined && v !== '') {
      out[k] = v;
    }
  });
  if (base?.observacion && extra?.observacion) {
    const b = String(base.observacion).trim();
    const e = String(extra.observacion).trim();
    if (b && e && b !== e) out.observacion = `${b} ${e}`;
  }
  return out;
}

function addWarning_(item, campo, mensaje) {
  if (!item) return;
  if (!Array.isArray(item.Warnings)) item.Warnings = [];
  item.Warnings.push({ campo: campo || 'general', mensaje: mensaje || '' });
}

function addStatementObservation_(item, message) {
  if (!item || !message) return;
  if (!item['Estado de cuenta'] || typeof item['Estado de cuenta'] !== 'object') {
    item['Estado de cuenta'] = {};
  }
  const st = item['Estado de cuenta'];
  const prev = st.observacion ? String(st.observacion).trim() : '';
  st.observacion = prev ? `${prev} ${message}` : message;
}

function checkStatementAmountMatch_(statementItem, receiptItem) {
  if (!statementItem || !receiptItem) return;
  const stmt = getNumeric_(statementItem['Importe a rendir']);
  const rec = getNumeric_(receiptItem['Importe a rendir']);
  if (stmt === null || rec === null) return;
  if (Math.abs(stmt - rec) > 0.01) {
    const msg = `Importe a rendir no coincide con estado de cuenta (${stmt} vs ${rec}).`;
    addWarning_(statementItem, 'Importe a rendir', msg);
    addStatementObservation_(statementItem, msg);
  }
}

function getNumeric_(value) {
  if (value === null || value === undefined || value === '') return null;
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function scrubInternalKeys_(item) {
  if (!item) return item;
  if (item.__statement_idx !== undefined) delete item.__statement_idx;
  if (item.__doc_names !== undefined) delete item.__doc_names;
  return item;
}

function markRowsColor_(rowIndexes, color) {
  if (!Array.isArray(rowIndexes) || !rowIndexes.length) return;
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);
  const cols = getWriteColumnIndexes_();
  rowIndexes.forEach((row) => {
    cols.forEach((col) => {
      sheet.getRange(row, col).setBackground(color);
    });
  });
}

function stage2_write_results(items) {
  if (!items || !items.length) {
    return { ok: false, message: 'No hay comprobantes para escribir.' };
  }
  writeItemsToSheet(items);                  // SheetWriter.gs (tu implementación)
  return { ok: true, message: `Planilla actualizada: ${items?.length || 0} comprobantes escritos.`, payload: { count: items.length } };
}

function stage2_write_tarjeta_results_(items, orphanRowIndexes) {
  const res = writeTarjetaItemsToSheet_(items);
  if (!res?.ok) return res;
  if (Array.isArray(orphanRowIndexes) && orphanRowIndexes.length) {
    markRowsColor_(orphanRowIndexes, COLOR_ORPHAN);
  }
  return res;
}

function stage2_write_tarjeta_results(items, orphanRowIndexes) {
  return stage2_write_tarjeta_results_(items, orphanRowIndexes);
}


function stage3_finalize() {
  const st = jobStateGet_();
  const mode = st?.mode || 'efectivo';
  const normalizedItems = getNormalizedItemsForMode_(st, mode);
  if (!st?.rendicionId || !normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay normalizedItems).' };
  }

  const coverRes = stage3_create_cover();
  if (!coverRes?.ok) return coverRes;

  const genRes = stage3_generate_outputs(coverRes.payload?.coverGcsUri || null);
  if (!genRes?.ok) return genRes;

  const dlRes = stage3_download_outputs(
    genRes.payload?.pdfUri || null,
    genRes.payload?.xlsmUri || null,
    genRes.payload?.pdfDriveId || null,
    genRes.payload?.xlsmDriveId || null
  );
  if (!dlRes?.ok) return dlRes;

  return { ok: true, message: 'Completado: PDF + Excel generados (Ver Drive).', payload: dlRes.payload || {} };
}

function stage3_create_cover() {
  const st = jobStateGet_();
  const mode = st?.mode || 'efectivo';
  const normalizedItems = getNormalizedItemsForMode_(st, mode);
  if (!st?.rendicionId || !normalizedItems?.length) {
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
  const mode = st?.mode || 'efectivo';
  const normalizedItems = getNormalizedItemsForMode_(st, mode);
  if (!st?.rendicionId || !normalizedItems?.length) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay items).' };
  }

  // 1) checks
  assertAllChecksOk_();

  // 2) celdas a mandar (solo FIELD_COLUMN_MAP)
  const orderedItems = reorderNormalizedItemsByOrder_(
    normalizedItems || [],
    st.lastOrdenArchivos || st.ordenArchivos || []
  );

  const xlsmValues = buildXlsmValuesFromSheet_();
  if (!xlsmValues.length) throw new Error('No hay valores para mandar al XLSM.');

  const statementInput = mode === 'tarjeta' ? buildStatementInputForFinalize_(st) : null;
  const finalItems = statementInput ? [statementInput].concat(orderedItems) : orderedItems;

  const output = DRIVE_API_ENABLED
    ? { driveFolderId: getOutputFolderId_() }
    : { gcsPrefix: st.gcsPrefix || GCS_PREFIX };

  const req = {
    rendicionId: st.rendicionId,
    inputs: {
      cover: coverGcsUri ? { gcsUri: coverGcsUri } : undefined,
      normalizedItems: finalItems,
      xlsmTemplate: {gcsUri: 'gs://scz-uy-rendiciones/templates/rendiciones_macro_template.xlsm'},
      xlsmValues
    },
    output,
    options: {
      pdfName: `RENDREQ-XXXX_${st.rendicionId}.pdf`,
      xlsmName: `RENDREQ-XXXX_${st.rendicionId}.xlsm`,
      mergeOrder: "cover_first"
    }
  };

  const res = callCloudRunJson_('/v1/finalize', req);
  if (!res?.ok) throw new Error(`Finalize ok=false: ${JSON.stringify(res)}`);

  const pdfUri = res.pdf?.gcsUri || null;
  const xlsmUri = res.xlsm?.gcsUri || null;
  const pdfDriveId = res.pdf?.driveFileId || null;
  const xlsmDriveId = res.xlsm?.driveFileId || null;

  return {
    ok: true,
    message: 'PDF/Excel generados.',
    payload: { pdfUri, xlsmUri, pdfDriveId, xlsmDriveId }
  };
}

function stage3_download_outputs(pdfUri, xlsmUri, pdfDriveId, xlsmDriveId) {
  const st = jobStateGet_();
  if (!st?.rendicionId) {
    return { ok: false, message: 'Primero corré etapa 1 (no hay rendicionId).' };
  }

  const saved = {};

  if (pdfDriveId) {
    saved.pdf = true;
    saved.pdfDriveId = pdfDriveId;
  } else if (pdfUri) {
    saveGcsFileToDrive_(pdfUri, `RENDREQ-XXXX_${st.rendicionId}.pdf`, getOutputFolderId_());
    saved.pdf = true;
  }
  if (xlsmDriveId) {
    saved.xlsm = true;
    saved.xlsmDriveId = xlsmDriveId;
  } else if (xlsmUri) {
    saveGcsFileToDrive_(xlsmUri, `RENDREQ-XXXX_${st.rendicionId}.xlsm`, getOutputFolderId_());
    saved.xlsm = true;
  }

  return { ok: true, message: 'Archivos generados.', payload: { saved, pdfUri, xlsmUri, pdfDriveId, xlsmDriveId } };
}

function buildStatementInputForFinalize_(st) {
  const cachedFingerprint = st?.statementFingerprint || null;
  const cachedGcs = st?.statementGcsUri || null;
  const cachedFile = st?.lastStatementFile || null;

  if (cachedFingerprint && cachedGcs) {
    return { gcsUri: cachedGcs, mime: cachedFile?.mimeType || null };
  }

  let file = cachedFile;
  if (!file || !file.id) {
    try {
      file = getSingleStatementFile_();
    } catch (err) {
      return null;
    }
  }

  const fingerprint = buildStatementFingerprint_(file);
  if (st?.statementGcsUri && st?.statementFingerprint === fingerprint) {
    return { gcsUri: st.statementGcsUri, mime: file?.mimeType || null };
  }

  const gcsUri = uploadStatementFileToGCS_(file, fingerprint);
  jobStateSet_(Object.assign({}, jobStateGet_(), {
    statementFingerprint: fingerprint,
    statementGcsUri: gcsUri,
    lastStatementFile: {
      id: file.id,
      name: file.name,
      mimeType: file.mimeType || null,
      size: file.size,
      updated: file.updated
    }
  }));
  return { gcsUri, mime: file?.mimeType || null };
}

function uploadStatementFileToGCS_(file, fingerprint) {
  if (!file || !file.id) throw new Error('Estado de cuenta: archivo inválido para subir a GCS.');
  const driveFile = DriveApp.getFileById(file.id);
  const blob = driveFile.getBlob().setName(file.name || 'estado.pdf');
  const objectName = buildStatementObjectName_(fingerprint, file.name || 'estado.pdf');
  return uploadBlobToGCS_(GCS_BUCKET, objectName, blob);
}


/** ---------- Helpers ---------- */

function extractComprobantes_(gemRes) {
  if (!gemRes) return [];
  if (Array.isArray(gemRes)) return gemRes;
  if (Array.isArray(gemRes?.data)) return gemRes.data;
  if (Array.isArray(gemRes?.comprobantes)) return gemRes.comprobantes;
  if (Array.isArray(gemRes?.['comprobantes'])) return gemRes['comprobantes'];
  return [];
}

function extractOrdenArchivos_(gemRes) {
  if (!gemRes) return null;
  const order = gemRes['orden archivos'] || gemRes.ordenArchivos || gemRes.orden_archivos || gemRes.order || null;
  if (!Array.isArray(order)) return null;
  const cleaned = order.map((o) => String(o || '').trim()).filter((o) => !!o);
  return cleaned.length ? cleaned : null;
}

function extractEstadoCuenta_(gemRes) {
  if (!gemRes) return null;
  return gemRes['estado de cuenta'] || gemRes.estadoCuenta || gemRes.estado_cuenta || null;
}

function reorderNormalizedItemsByOrder_(items, orderList) {
  if (!Array.isArray(items) || !items.length) return items || [];
  if (!Array.isArray(orderList) || !orderList.length) return items;

  const order = orderList
    .map((o) => normalizeOrderToken_(o))
    .filter((o) => !!o);
  if (!order.length) return items;

  const itemsWithIndex = items.map((item, idx) => {
    return {
      item,
      idx,
      key: getItemNormalizedIndex_(item, idx),
      driveFileId: item?.driveFileId || null,
      docName: getItemDocName_(item)
    };
  });

  const byKey = new Map();
  itemsWithIndex.forEach((entry) => {
    if (entry.key && !byKey.has(entry.key)) byKey.set(entry.key, entry);
    if (entry.driveFileId && !byKey.has(entry.driveFileId)) byKey.set(entry.driveFileId, entry);
    if (entry.docName && !byKey.has(entry.docName)) byKey.set(entry.docName, entry);
  });

  const orderedEntries = [];
  const consumed = new Set();

  order.forEach((k) => {
    const entry = byKey.get(k);
    if (entry && !consumed.has(entry)) {
      orderedEntries.push(entry);
      consumed.add(entry);
    }
  });

  itemsWithIndex.forEach((entry) => {
    if (!consumed.has(entry)) {
      orderedEntries.push(entry);
      consumed.add(entry);
    }
  });

  return orderedEntries.map((e) => e.item);
}

function normalizeOrderToken_(token) {
  if (token === null || token === undefined) return null;
  const raw = String(token).trim();
  if (!raw) return null;
  const driveMatch = raw.match(/drive_([A-Za-z0-9_-]+)/i);
  if (driveMatch) return driveMatch[1];
  const str = raw.toLowerCase();
  if (!str) return null;
  const cleaned = str.startsWith('f') ? str.slice(1) : str;
  const m = cleaned.match(/^(\d{1,4})$/);
  if (m) return m[1].padStart(4, '0');
  return raw;
}

function getItemDocName_(item) {
  if (!item) return null;
  const uri = item.gcsUri;
  if (!uri) return null;
  const s = String(uri).trim();
  if (!s) return null;
  const parts = s.split('/');
  return parts[parts.length - 1] || null;
}

function getItemNormalizedIndex_(item, fallbackIdx) {
  if (!item) return null;
  if (item.normalizedIndex !== undefined && item.normalizedIndex !== null) {
    const s = String(item.normalizedIndex).trim();
    if (s) return s.padStart(4, '0');
  }
  const uri = String(item.gcsUri || '').toLowerCase();
  const m = uri.match(/\/(\d{4})_/);
  if (m) return m[1];
  if (fallbackIdx !== undefined && fallbackIdx !== null) {
    return String(fallbackIdx).padStart(4, '0');
  }
  return null;
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
