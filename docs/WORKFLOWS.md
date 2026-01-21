# WORKFLOWS

## Botones disponibles (Apps Script)
- Procesar comprobantes de efectivo.
- Procesar comprobantes de tarjeta corporativa.
- Generar PDF + XLSM.
- Reiniciar todo (borrado de planilla; hoy es lento y se puede optimizar).

## Carpetas esperadas (Drive)
Carpeta base = carpeta donde está el Spreadsheet.
Subcarpetas por nombre:
- `Comprobantes efectivo`
- `Comprobantes tarjeta corporativa` (puede no existir)
- `Estado de cuenta` (puede no existir)

## Autenticación Cloud Run (Apps Script)
- `callCloudRunJson_` usa ID token generado con la Service Account definida en `SERVICE_ACCOUNT_KEY`.
- Requisitos IAM:
  - `roles/iam.serviceAccountTokenCreator` sobre la misma SA.
  - `roles/run.invoker` sobre `rendiciones-service`.
- API requerida: `iamcredentials.googleapis.com`.

## Flujo: Procesar comprobantes de efectivo
1) UI: botón **Procesar comprobantes de efectivo** → `stage_upload_and_process_cash`.
2) Resolución de carpeta:
   - `getBaseFolder_()` toma el padre del Spreadsheet.
   - `getModeFolderId_('efectivo')` busca `Comprobantes efectivo`.
3) Etapa 1 (normalize):
   - `listDriveFiles_` arma lista ordenada.
   - Si `DRIVE_API_ENABLED=true`, se llama `/v1/normalize` con `driveFileIds`.
   - Si no, se crea ZIP y se sube a GCS, luego `/v1/normalize` con `zipGcsUri`.
   - Se guarda `normalizedItems` en `jobState`.
4) Etapa 2 (extract):
   - `/v1/process_receipts_batch` con `mode=efectivo` y `receipts` desde `normalizedItems` (gcsUri + mime).
   - `flattenDocflowRows_` aplana `rows[].data` a items.
5) Escritura:
   - `writeItemsToSheet(items)` escribe campos y warnings por fila.
   - `Importe a rendir` es el único importe visible en planilla.

## Flujo: Procesar comprobantes de tarjeta corporativa
1) UI: botón **Procesar comprobantes de tarjeta corporativa** → `stage_upload_and_process_card`.
2) Estado de cuenta (obligatorio):
   - `getStatementFolderId_()` busca `Estado de cuenta`.
   - `getSingleStatementFile_()` exige 1 archivo (si 0 o >1, error).
3) Etapa 2 (estado):
   - `/v1/process_statement` con `driveFileId` del estado.
   - El servicio rasteriza PDFs multipágina y usa DocFlow en modo aggregate.
   - `buildStatementItems_` escribe filas base: fecha, proveedor (detalle), importe y moneda.
4) Comprobantes:
   - `getModeFolderId_('tarjeta')` busca `Comprobantes tarjeta corporativa`.
   - `/v1/normalize` genera `normalizedItems`.
5) Etapa 2 (comprobantes):
   - `/v1/process_receipts_batch` con `mode=tarjeta` y `statement.parsed`.
   - `mergeStatementWithReceipts_` mergea por `Estado de cuenta.idx`.
6) Cobertura y conflictos:
   - Línea de estado sin comprobante → warning `general` (fila subrayada).
   - Comprobante sin movimiento → warning `general` y fila marcada en rojo.
   - Varios comprobantes mismo `idx`:
     - Si concuerdan, se mergea.
     - Si no concuerdan, warning `general` con filenames en conflicto.
7) Escritura final:
   - `writeItemsToSheet(merged.items)` y `markRowsColor_` para sobrantes.

## Flujo: Generar PDF + XLSM
1) Verifica que todas las filas estén validadas (checkbox).
2) Genera carátula PDF desde la hoja.
3) Llama `/v1/finalize` para crear PDF + XLSM con la plantilla.
4) Guarda artefactos en GCS y opcionalmente en Drive.

## Flujo: Reiniciar todo
- `clearAllRows` limpia valores, notas y metadata por fila.
- Nota: recorre columnas por fila; es funcional pero lento para lotes grandes.

## Observaciones
- `Importe facturado` queda en el JSON; en la planilla se usa `Importe a rendir`.
- Warnings de extracción se guardan por fila; warnings de API van por fuera del payload.
