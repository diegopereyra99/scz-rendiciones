# CONTRACTS

## Contratos existentes (vigentes en repo)

### Cloud Run /v1/normalize
- Request/Response: ver `service/README.md`.
- Output: normalizados en `gs://.../normalized/` + manifest.
- Estado: implementado.

### Cloud Run /v1/finalize
- Request/Response: ver `service/README.md`.
- Output: PDF + XLSM en `gs://.../outputs/` (o Drive).
- Estado: implementado.

### DocFlow SDK (local)
- Perfil `lineas_gastos/v0`: retorna envelope `data` + `meta` por archivo (ver `sample_output.json`).
- Modo esperado: `per_file`.
- Estado: probado via CLI; pendiente comparar con SDK (script `docflow_sdk_compare.py`).

## Contratos propuestos (Etapa 2)

### Cloud Run /v1/process_statement (propuesto)
**Request**
```json
{
  "rendicionId": "2025_10_user",
  "statement": {
    "driveFileId": "fileId123",
    "mime": "application/pdf"
  },
  "options": {
    "profile": "estado/v0",
    "model": "gemini-2.5-flash"
  }
}
```
**Response**
```json
{
  "ok": true,
  "rendicionId": "2025_10_user",
  "data": { "fecha_emision": "...", "transacciones": [ ... ] },
  "meta": { "model": "...", "usage": {"input_tokens": 0, "output_tokens": 0}, "docs": ["f0000"], "profile": "estado/v0" },
  "warnings": []
}
```
Notas:
- `data` es el output de DocFlow (perfil `estado/v0`).
- `warnings` es lista de warnings de infraestructura (no warnings del modelo).
- `statement` acepta `gcsUri`, `driveFileId` o `signedUrl` (exactamente uno).

### Cloud Run /v1/process_receipts_batch (propuesto)
**Request**
```json
{
  "rendicionId": "2025_10_user",
  "mode": "efectivo",
  "statement": {
    "parsed": { "transacciones": [ ... ] }
  },
  "receipts": [
    { "gcsUri": "gs://.../normalized/0002_x.jpg", "mime": "image/jpeg" }
  ],
  "options": {
    "profile": "lineas_gastos/v0",
    "model": "gemini-2.5-flash"
  }
}
```
**Response**
```json
{
  "ok": true,
  "rendicionId": "2025_10_user",
  "rows": [
    {
      "data": [ { "Fecha de factura": "...", "Warnings": [], "Estado de cuenta": {"idx": 12, "match_status": "matched"} } ],
      "meta": { "model": "...", "docs": ["f0002"], "profile": "lineas_gastos/v0", "mode": "per_file" }
    }
  ],
  "warnings": []
}
```
Notas:
- `rows` usa el mismo envelope `data`/`meta` que DocFlow SDK (per-file).
- Si hay estado de cuenta, se pasa en el prompt y se completa el campo `Estado de cuenta`.
- `receipts[]` acepta `gcsUri`, `driveFileId` o `signedUrl` (exactamente uno).

## Contratos que faltan definir
- Campo exacto para asociacion con estado de cuenta: nombre final y formato del `idx` y `match_status`.
- Envelope final para apps script: se usa `rows[]` como arriba o se aplana a `items[]`?
- Error schema y codigos de warning/errores a nivel API (estandarizar).
- Layout GCS definitivo para `normalized/` y archivos de estado de cuenta.
- Politica incremental (se reporta `removed[]`? se marca en planilla?).
- Tratamiento de lineas "Reduc. IVA" y como se reflejan en el output.

## Puntos a aclarar
- Versionado de perfiles DocFlow en Cloud Run (tag fijo o hash).
- Donde se guarda el estado de cuenta parseado (memoria, GCS, cache).
- Criterio de matching cuando hay multiples movimientos candidatos.
- Consumo de warnings en Apps Script (formato exacto por campo).
