from __future__ import annotations

import io
import urllib.request

from google.auth import default as google_auth_default  # type: ignore
from google.auth.transport.requests import Request as GoogleAuthRequest  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.http import MediaIoBaseDownload  # type: ignore

from . import gcs
from .config import Settings


def _drive_service() -> any:
    # Relies on Application Default Credentials; for Workload Identity, ADC will be used.
    creds, _ = google_auth_default(scopes=["https://www.googleapis.com/auth/drive"])
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(GoogleAuthRequest())
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def fetch_bytes_from_drive(file_id: str) -> bytes:
    service = _drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()


def fetch_bytes(ref: dict, settings: Settings) -> bytes:
    """
    ref accepts keys: gcsUri, signedUrl, driveFileId
    """
    if ref.get("gcsUri"):
        return gcs.download_bytes(ref["gcsUri"])
    if ref.get("signedUrl"):
        with urllib.request.urlopen(ref["signedUrl"]) as resp:  # nosec B310
            return resp.read()
    if ref.get("driveFileId"):
        if not settings.drive_enabled:
            raise RuntimeError("Drive API disabled")
        return fetch_bytes_from_drive(ref["driveFileId"])
    raise ValueError("No valid reference provided")
