from __future__ import annotations

import base64
import io
import json
import zipfile
from datetime import datetime, timezone
from typing import Any, List, Tuple

from google.auth import default as google_auth_default  # type: ignore
from google.auth.transport.requests import Request as GoogleAuthRequest  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.http import MediaIoBaseDownload  # type: ignore
from pypdf import PdfReader
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed

from .. import gcs
from ..config import Settings
from ..models import (
    ErrorPayload,
    NormalizeItem,
    NormalizeRequest,
    NormalizeResponse,
    NormalizedArtifact,
    SourceInfo,
    Warning,
)
from ..utils import (
    SUPPORTED_IMAGE_EXTS,
    MIME_TYPE_MAP,
    decode_zip_base64,
    image_to_jpeg_bytes,
    apply_exif_orientation,
    resize_image_max_side,
    ensure_rgb,
    sha256_bytes,
)


def _exception_details(exc: Exception) -> dict[str, Any]:
    return {"error": str(exc), "exceptionType": exc.__class__.__name__}


def _drive_service():
    creds, _ = google_auth_default(scopes=["https://www.googleapis.com/auth/drive.readonly"])
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(GoogleAuthRequest())
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _list_drive_folder(folder_id: str) -> List[Tuple[str, str]]:
    """
    Returns list of tuples (name, fileId).
    """
    service = _drive_service()
    page_token = None
    files: List[Tuple[str, str]] = []
    while True:
        response = (
            service.files()
            .list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
            )
            .execute()
        )
        for f in response.get("files", []):
            files.append((f["name"], f["id"]))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return files


def _download_drive_file(file_id: str) -> bytes:
    service = _drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()


def _get_drive_file_name(file_id: str) -> str:
    service = _drive_service()
    meta = service.files().get(fileId=file_id, fields="name").execute()
    return meta.get("name", file_id)


def _download_drive_entry(file_id: str, name: str | None) -> Tuple[str, bytes, SourceInfo]:
    if not name:
        name = _get_drive_file_name(file_id)
    data = _download_drive_file(file_id)
    return name, data, SourceInfo(driveFileId=file_id, originalName=name)


def _normalize_entry(
    idx: int,
    name: str,
    data: bytes,
    source: SourceInfo,
    *,
    jpg_quality: int,
    max_side: int,
    pdf_mode: str,
    upload_originals: bool,
    gcs_prefix: str,
) -> Tuple[int, NormalizeItem | None, List[Warning]]:
    warnings: List[Warning] = []
    ext = name.split(".")[-1].lower() if "." in name else ""
    normalized_bytes: bytes | None = None
    mime = ""
    page_count = None
    target_ext = ""
    original_uri = None

    try:
        if ext in SUPPORTED_IMAGE_EXTS:
            img = Image.open(io.BytesIO(data))
            img = apply_exif_orientation(img)
            img = resize_image_max_side(img, max_side)
            img = ensure_rgb(img)
            normalized_bytes = image_to_jpeg_bytes(img, jpg_quality)
            mime = "image/jpeg"
            target_ext = "jpg"
        elif ext in {"pdf"}:
            if pdf_mode == "rasterize":
                warnings.append(
                    Warning(
                        code="PDF_RASTERIZE_NOT_IMPLEMENTED",
                        message="pdfMode=rasterize not implemented; keeping PDF as-is.",
                        details={"file": name},
                    )
                )
            normalized_bytes = data
            mime = "application/pdf"
            target_ext = "pdf"
            try:
                reader = PdfReader(io.BytesIO(data))
                page_count = len(reader.pages)
            except Exception:
                page_count = None
        else:
            warnings.append(
                Warning(
                    code="UNSUPPORTED_FILE_TYPE",
                    message=f"Skipping unsupported file: {name}",
                    details={"extension": ext, "filename": name},
                )
            )
            return idx, None, warnings

        sha = sha256_bytes(normalized_bytes)
        object_path = f"{gcs_prefix}normalized/{idx:04d}_{sha}.{target_ext}"
        gcs_uri = gcs.upload_bytes(normalized_bytes, object_path)

        if upload_originals:
            try:
                original_path = f"{gcs_prefix}originals/{idx:04d}_{name}"
                original_uri = gcs.upload_bytes(data, original_path)
            except Exception as exc:  # noqa: BLE001
                warnings.append(
                    Warning(
                        code="ORIGINAL_UPLOAD_FAILED",
                        message=f"Failed to upload original for {name}",
                        details={
                            "filename": name,
                            "targetPath": original_path,
                            **_exception_details(exc),
                        },
                    )
                )

        item = NormalizeItem(
            source=source,
            normalized=NormalizedArtifact(
                gcsUri=gcs_uri,
                mime=mime,
                sha256=sha,
                bytes=len(normalized_bytes),
                pageCount=page_count,
                originalGcsUri=original_uri,
                originalMime=MIME_TYPE_MAP.get(ext, ""),
            ),
        )
        return idx, item, warnings
    except Exception as exc:  # noqa: BLE001
        warnings.append(
            Warning(
                code="NORMALIZATION_FAILED",
                message=f"Failed to process {name}",
                details={
                    "filename": name,
                    "source": source.model_dump(exclude_none=True),
                    "extension": ext,
                    **_exception_details(exc),
                },
            )
        )
        return idx, None, warnings


async def run_normalize(request: NormalizeRequest, settings: Settings) -> NormalizeResponse:
    jpg_quality = request.options.jpgQuality if request.options else settings.default_jpg_quality
    max_side = request.options.maxSidePx if request.options else settings.default_max_side_px
    pdf_mode = request.options.pdfMode if request.options else settings.default_pdf_mode
    upload_originals = request.options.uploadOriginals if request.options else False

    warnings: List[Warning] = []
    items: List[NormalizeItem] = []

    try:
        # Resolve inputs
        raw_entries: List[Tuple[str, bytes, SourceInfo]] = []
        sort_entries = False
        if request.input.driveFileIds:
            if not settings.drive_enabled:
                return NormalizeResponse(
                    ok=False,
                    rendicionId=request.rendicionId,
                    items=[],
                    warnings=[],
                    error=ErrorPayload(
                        code="DRIVE_API_DISABLED",
                        message="Drive API is disabled; enable REN_DRIVE_ENABLED to true and configure ADC.",
                        details={},
                    ),
                )
            file_ids = request.input.driveFileIds
            if file_ids:
                workers = min(len(file_ids), settings.normalize_workers)
                entries: List[Tuple[str, bytes, SourceInfo] | None] = [None] * len(file_ids)
                with ThreadPoolExecutor(max_workers=workers) as executor:
                    futures = {
                        executor.submit(_download_drive_entry, fid, None): idx
                        for idx, fid in enumerate(file_ids)
                    }
                    for fut in as_completed(futures):
                        idx = futures[fut]
                        fid = file_ids[idx]
                        try:
                            entries[idx] = fut.result()
                        except Exception as exc:  # noqa: BLE001
                            warnings.append(
                                Warning(
                                    code="DRIVE_DOWNLOAD_FAILED",
                                    message=f"Failed to download {fid}",
                                    details={"fileId": fid, **_exception_details(exc)},
                                )
                            )
                raw_entries.extend([entry for entry in entries if entry])
        elif request.input.files:
            for file in request.input.files:
                try:
                    data = base64.b64decode(file.contentBase64)
                    raw_entries.append((file.filename, data, SourceInfo(originalName=file.filename)))
                except Exception as exc:  # noqa: BLE001
                    warnings.append(
                        Warning(
                            code="INVALID_ARGUMENT",
                            message=f"Failed to decode file {file.filename}",
                            details={"filename": file.filename, **_exception_details(exc)},
                        )
                    )
        elif request.input.zipBase64:
            sort_entries = True
            try:
                zf = decode_zip_base64(request.input.zipBase64)
            except Exception as exc:  # noqa: BLE001
                return NormalizeResponse(
                    ok=False,
                    rendicionId=request.rendicionId,
                    items=[],
                    warnings=[],
                    error=ErrorPayload(
                        code="INVALID_ARGUMENT",
                        message="Invalid zipBase64 input",
                        details={"zipBase64Length": len(request.input.zipBase64), **_exception_details(exc)},
                    ),
                )
            for name in zf.namelist():
                if name.endswith("/"):
                    continue
                with zf.open(name) as f:
                    raw_entries.append((name, f.read(), SourceInfo(originalName=name)))
        elif request.input.zipGcsUri:
            sort_entries = True
            try:
                zip_bytes = gcs.download_bytes(request.input.zipGcsUri)
                zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
            except Exception as exc:  # noqa: BLE001
                return NormalizeResponse(
                    ok=False,
                    rendicionId=request.rendicionId,
                    items=[],
                    warnings=[],
                    error=ErrorPayload(
                        code="INVALID_ARGUMENT",
                        message="Invalid zipGcsUri input",
                        details={"zipGcsUri": request.input.zipGcsUri, **_exception_details(exc)},
                    ),
                )
            for name in zf.namelist():
                if name.endswith("/"):
                    continue
                with zf.open(name) as f:
                    raw_entries.append((name, f.read(), SourceInfo(originalName=name)))
        elif request.input.driveFolderId:
            sort_entries = True
            if not settings.drive_enabled:
                return NormalizeResponse(
                    ok=False,
                    rendicionId=request.rendicionId,
                    items=[],
                    warnings=[],
                    error=ErrorPayload(
                        code="DRIVE_API_DISABLED",
                        message="Drive API is disabled; enable REN_DRIVE_ENABLED to true and configure ADC.",
                        details={},
                    ),
                )
            try:
                drive_files = _list_drive_folder(request.input.driveFolderId)
            except Exception as exc:  # noqa: BLE001
                return NormalizeResponse(
                    ok=False,
                    rendicionId=request.rendicionId,
                    items=[],
                    warnings=[],
                    error=ErrorPayload(
                        code="DRIVE_ACCESS_DENIED",
                        message=str(exc),
                        details={"driveFolderId": request.input.driveFolderId, **_exception_details(exc)},
                    ),
                )
            if drive_files:
                workers = min(len(drive_files), settings.normalize_workers)
                entries: List[Tuple[str, bytes, SourceInfo] | None] = [None] * len(drive_files)
                with ThreadPoolExecutor(max_workers=workers) as executor:
                    futures = {
                        executor.submit(_download_drive_entry, fid, fname): idx
                        for idx, (fname, fid) in enumerate(drive_files)
                    }
                    for fut in as_completed(futures):
                        idx = futures[fut]
                        fname, fid = drive_files[idx]
                        try:
                            entries[idx] = fut.result()
                        except Exception as exc:  # noqa: BLE001
                            warnings.append(
                                Warning(
                                    code="DRIVE_DOWNLOAD_FAILED",
                                    message=f"Failed to download {fname}",
                                    details={"fileId": fid, "fileName": fname, **_exception_details(exc)},
                                )
                            )
                raw_entries.extend([entry for entry in entries if entry])

        if sort_entries:
            raw_entries = sorted(raw_entries, key=lambda entry: entry[0].lower())
    except Exception as exc:  # noqa: BLE001
        return NormalizeResponse(
            ok=False,
            rendicionId=request.rendicionId,
            items=[],
            warnings=[],
            error=ErrorPayload(
                code="NORMALIZATION_FAILED",
                message="Failed while preparing inputs",
                details={"stage": "prepare_inputs", **_exception_details(exc)},
            ),
        )

    gcs_prefix = gcs.normalize_prefix(request.output.gcsPrefix)

    if raw_entries:
        workers = min(len(raw_entries), settings.normalize_workers)
        items_by_idx: List[NormalizeItem | None] = [None] * len(raw_entries)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    _normalize_entry,
                    idx,
                    name,
                    data,
                    source,
                    jpg_quality=jpg_quality,
                    max_side=max_side,
                    pdf_mode=pdf_mode,
                    upload_originals=upload_originals,
                    gcs_prefix=gcs_prefix,
                ): idx
                for idx, (name, data, source) in enumerate(raw_entries)
            }
            for fut in as_completed(futures):
                idx = futures[fut]
                try:
                    _, item, entry_warnings = fut.result()
                except Exception as exc:  # noqa: BLE001
                    warnings.append(
                        Warning(
                            code="NORMALIZATION_FAILED",
                            message="Failed to process item",
                            details={"index": idx, **_exception_details(exc)},
                        )
                    )
                    continue
                if item:
                    items_by_idx[idx] = item
                if entry_warnings:
                    warnings.extend(entry_warnings)
        items = [item for item in items_by_idx if item]

    # Manifest (optional)
    manifest_uri = None
    try:
        manifest = {
            "rendicionId": request.rendicionId,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "items": [
                {
                    "source": item.source.model_dump(),
                    "normalized": item.normalized.model_dump(),
                }
                for item in items
            ],
        }
        manifest_bytes = json.dumps(manifest, ensure_ascii=True, indent=2).encode("utf-8")
        manifest_path = f"{gcs_prefix}manifests/normalize_manifest.json"
        manifest_uri = gcs.upload_bytes(manifest_bytes, manifest_path, content_type="application/json")
    except Exception as exc:  # noqa: BLE001
        warnings.append(
            Warning(
                code="MANIFEST_WRITE_FAILED",
                message="Failed to write manifest",
                details={"targetPath": f"{gcs_prefix}manifests/normalize_manifest.json", **_exception_details(exc)},
            )
        )

    failed_warning_codes = {"NORMALIZATION_FAILED", "DRIVE_DOWNLOAD_FAILED"}
    failed_warnings = [w for w in warnings if w.code in failed_warning_codes]
    ok = len(failed_warnings) == 0

    error = (
        None
        if ok
        else ErrorPayload(
            code="PARTIAL_FAILURE",
            message=f"{len(failed_warnings)} item(s) failed during normalization",
            details={
                "failedWarnings": [w.model_dump() for w in failed_warnings],
                "failedCount": len(failed_warnings),
                "totalWarnings": len(warnings),
                "attemptedItems": len(raw_entries),
                "successfulItems": len(items),
            },
        )
    )

    return NormalizeResponse(
        ok=ok,
        rendicionId=request.rendicionId,
        items=items,
        manifestGcsUri=manifest_uri,
        warnings=warnings or None,
        error=error,
    )
