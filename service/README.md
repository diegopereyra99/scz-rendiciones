# Rendiciones Cloud Run Service

Implements the contract in `Servicio_CloudRun_Rendiciones_SPEC.md` via FastAPI. Normalization and finalization pipelines are functional (GCS-first; Drive supported when enabled) with remaining TODOs noted below.

## Layout
- `Dockerfile` – container image for Cloud Run.
- `requirements.txt` – runtime deps (FastAPI, GCS/Drive SDKs, PDF/XLSM helpers).
- `src/` – FastAPI app, Pydantic models, service layers.
- `templates/rendiciones_macro_template.xlsm` – baked-in XLSM template (renamed from the provided TEST 2.xlsm).
- `tests/` – placeholder for contract and integration tests.

## Local development
```bash
cd service
python -m venv .venv
./.venv/Scripts/Activate.ps1  # or source .venv/bin/activate on macOS/Linux
pip install -r requirements.txt
uvicorn src.main:app --reload --host 0.0.0.0 --port 8080
```
Endpoints:
- `GET /healthz`
- `POST /v1/normalize`
- `POST /v1/finalize`

## Build and deploy (Cloud Run)
```bash
PROJECT_ID="your-gcp-project"
REGION="us-central1"
SERVICE_NAME="rendiciones-service"
IMAGE="gcr.io/${PROJECT_ID}/${SERVICE_NAME}:v0"

cd service
gcloud builds submit --tag "${IMAGE}"
gcloud run deploy "${SERVICE_NAME}" \
  --image "${IMAGE}" \
  --platform managed \
  --region "${REGION}" \
  --allow-unauthenticated=false \
  --set-env-vars REN_ENVIRONMENT=prod,REN_GCS_BUCKET=your-bucket
```
If Drive is required, enable Drive API and share folders/files with the Cloud Run service account.

## Configuration
Environment variables (prefix `REN_`):
- `REN_ENVIRONMENT` – env label (`local|dev|prod`).
- `REN_GCS_BUCKET` – default bucket when only prefixes are given.
- `REN_DRIVE_ENABLED` – set `true` when Drive API + permissions are available.
- `REN_DEFAULT_JPG_QUALITY`, `REN_DEFAULT_MAX_SIDE_PX`, `REN_DEFAULT_PDF_MODE` – defaults for normalization.
- `REN_DEFAULT_SIGNED_URL_TTL` – TTL in seconds for signed URLs when requested.
- `REN_XLSM_TEMPLATE_PATH` – path to the XLSM template inside the image (defaults to `templates/rendiciones_macro_template.xlsm`).

## Normalize endpoint (inputs/outputs)
- One-of input sources: `driveFileIds[]` (preferred ordered list), inline `files[]` (filename + base64), `zipBase64`, `zipGcsUri`, or `driveFolderId` (fallback). Drive paths require `REN_DRIVE_ENABLED=true` and SA access.
- Options: `jpgQuality` (default 90), `maxSidePx` (default 2000), `pdfMode` (`keep` only; rasterize not implemented), `uploadOriginals` (default false).
- Output: uploads to `<gcsPrefix>/normalized/` (and optionally `<gcsPrefix>/originals/`), returns items with `gcsUri`, `mime`, `sha256`, `bytes`, `pageCount`, optional `originalGcsUri`; manifest at `<gcsPrefix>/manifests/normalize_manifest.json` when written.

Example request:
```json
{
  "rendicionId": "abc123",
  "input": { "driveFileIds": ["fileId1", "fileId2"] },
  "output": { "gcsPrefix": "gs://bucket/rendiciones/2025/usr_x/01/" },
  "options": { "jpgQuality": 90, "maxSidePx": 2000, "uploadOriginals": false }
}
```
Example response:
```json
{
  "ok": true,
  "rendicionId": "abc123",
  "items": [
    {
      "source": { "driveFileId": "fileId1", "originalName": "1.png" },
      "normalized": {
        "gcsUri": "gs://bucket/.../normalized/0000_hash.jpg",
        "mime": "image/jpeg",
        "sha256": "hash",
        "bytes": 12345,
        "pageCount": null,
        "originalGcsUri": null
      }
    }
  ],
  "manifestGcsUri": "gs://bucket/.../manifests/normalize_manifest.json",
  "warnings": null,
  "error": null
}
```

## Finalize endpoint (inputs/outputs)
- Inputs: `cover` (one-of `gcsUri` | `driveFileId` | `signedUrl`), `normalizedItems[]` (GCS/Drive/signed URL; ordered), `xlsmTemplate` (one-of `gcsUri` | `driveFileId`; optional—if omitted, the embedded template at `REN_XLSM_TEMPLATE_PATH` is used), `xlsmValues[]` (cell writes: `sheet`, `row`, `col`, `value`).
- Output target: one-of `driveFolderId` or `gcsPrefix` (recommended).
- Options: `pdfName`, `xlsmName`, `mergeOrder` (`cover_first`), `signedUrlTtlSeconds` (>=60; omit/0 to skip signed URLs).
- Output: PDF and XLSM artifacts (GCS URIs and optional signed URLs, or Drive IDs); warnings included when signing fails or items are skipped.

Example request:
```json
{
  "rendicionId": "abc123",
  "inputs": {
    "cover": { "gcsUri": "gs://bucket/rendiciones/2025/usr_x/01/cover.pdf" },
    "normalizedItems": [
      { "gcsUri": "gs://bucket/.../normalized/0000_hash.jpg", "mime": "image/jpeg" }
    ],
    "xlsmTemplate": { "gcsUri": "gs://bucket/templates/rendiciones_macro_template.xlsm" },
    "xlsmValues": [
      { "sheet": "Formulario E14", "row": 14, "col": 2, "value": "2025-12-18" },
      { "sheet": "Formulario E14", "row": 14, "col": 7, "value": "A-1" }
    ]
  },
  "output": { "gcsPrefix": "gs://bucket/rendiciones/2025/usr_x/01/" },
  "options": { "pdfName": "rendicion.pdf", "xlsmName": "rendicion.xlsm", "signedUrlTtlSeconds": 3600 }
}
```
Example response:
```json
{
  "ok": true,
  "rendicionId": "abc123",
  "pdf": { "gcsUri": "gs://bucket/.../outputs/rendicion.pdf", "signedUrl": null },
  "xlsm": { "gcsUri": "gs://bucket/.../outputs/rendicion.xlsm", "signedUrl": null },
  "warnings": [
    { "code": "SIGNED_URL_FAILED", "message": "Failed to generate signed URL for PDF" }
  ],
  "error": null
}
```

## How to swap the XLSM template
Drop your macros-enabled template under `service/templates/` and update `REN_XLSM_TEMPLATE_PATH` (or overwrite the existing filename). Rebuild and redeploy the image so Cloud Run containers ship with the new file.

## TODOs / open work
- `/v1/normalize`: add proper PDF rasterize mode, stronger MIME/type detection, retries/streaming for large inputs, and richer logging/metrics; optionally allow sanitizing names for originals upload.
- `/v1/finalize`: harden Drive upload path (scopes/fields), add chunked uploads for large PDFs/XLSM, and better MIME inference for items lacking extensions.
- Add GCS + Drive client helpers with retries, idempotency, and IAM/error mapping (now using direct SDK calls).
- Wire observability (structured logs with rendicionId/endpoint/step/duration, counters for files/bytes) and add tracing hooks.
- Add tests listed in the spec (normalize format mix, EXIF rotation, corrupt file handling, finalize happy path to GCS/Drive).
- Publish shared schemas package (or JSON Schemas) so Apps Script and the service share one contract source of truth.
- Add infra as code (Terraform) to provision Cloud Run, buckets, service account, IAM, secrets, and deploy pipelines.
