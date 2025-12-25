from __future__ import annotations

import re
from datetime import timedelta
from typing import Tuple

from google.cloud import storage
from google.cloud.storage import Blob


_GCS_URI_RE = re.compile(r"^gs://(?P<bucket>[^/]+)/(?P<path>.+)$")


def parse_gcs_uri(uri: str) -> Tuple[str, str]:
    match = _GCS_URI_RE.match(uri)
    if not match:
        raise ValueError(f"Invalid GCS URI: {uri}")
    return match.group("bucket"), match.group("path")


def normalize_prefix(prefix: str) -> str:
    if not prefix.endswith("/"):
        return prefix + "/"
    return prefix


def upload_bytes(
    data: bytes,
    gcs_uri: str,
    content_type: str | None = None,
) -> str:
    bucket_name, blob_path = parse_gcs_uri(gcs_uri)
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob: Blob = bucket.blob(blob_path)
    blob.upload_from_string(data, content_type=content_type)
    return f"gs://{bucket_name}/{blob_path}"


def download_bytes(gcs_uri: str) -> bytes:
    bucket_name, blob_path = parse_gcs_uri(gcs_uri)
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob: Blob = bucket.blob(blob_path)
    return blob.download_as_bytes()


def maybe_signed_url(gcs_uri: str, ttl_seconds: int | None) -> str | None:
    if not ttl_seconds:
        return None
    bucket_name, blob_path = parse_gcs_uri(gcs_uri)
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob: Blob = bucket.blob(blob_path)
    return blob.generate_signed_url(expiration=timedelta(seconds=ttl_seconds))
