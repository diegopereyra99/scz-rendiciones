from __future__ import annotations

import base64
import hashlib
import io
import zipfile
from typing import Iterable, List, Tuple

from PIL import Image, ImageOps


SUPPORTED_IMAGE_EXTS = {"jpg", "jpeg", "png", "webp", "bmp", "tiff"}
MIME_TYPE_MAP = {
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "png": "image/png",
    "webp": "image/webp",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "pdf": "application/pdf",
}
PDF_EXTS = {"pdf"}


def decode_zip_base64(encoded: str) -> zipfile.ZipFile:
    data = base64.b64decode(encoded)
    return zipfile.ZipFile(io.BytesIO(data))


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def ensure_rgb(image: Image.Image) -> Image.Image:
    if image.mode in ("RGBA", "LA", "P"):
        return image.convert("RGB")
    if image.mode == "CMYK":
        return image.convert("RGB")
    if image.mode == "L":
        return image.convert("RGB")
    return image


def apply_exif_orientation(image: Image.Image) -> Image.Image:
    try:
        return ImageOps.exif_transpose(image)
    except Exception:
        return image


def resize_image_max_side(image: Image.Image, max_side: int) -> Image.Image:
    width, height = image.size
    if max(width, height) <= max_side:
        return image
    scale = max_side / float(max(width, height))
    new_size = (int(width * scale), int(height * scale))
    return image.resize(new_size, Image.LANCZOS)


def image_to_jpeg_bytes(image: Image.Image, quality: int) -> bytes:
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def sort_entries(entries: Iterable[Tuple[str, bytes]]) -> List[Tuple[str, bytes]]:
    return sorted(entries, key=lambda kv: kv[0].lower())
