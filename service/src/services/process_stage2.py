from __future__ import annotations

import json
import mimetypes
from dataclasses import replace
from pathlib import Path
from typing import Any, List, Tuple
from urllib.parse import urlparse

from docflow.core.extraction.engine import ExtractionResult, extract
from docflow.core.providers.base import ProviderOptions
from docflow.core.providers.gemini import GeminiProvider
from docflow.sdk import profiles
from docflow.sdk.config import SdkConfig

from ..config import Settings
from ..fetch import fetch_bytes
from ..models import (
    DocflowRow,
    DocumentRef,
    ErrorPayload,
    ProcessReceiptsBatchRequest,
    ProcessReceiptsBatchResponse,
    ProcessStatementRequest,
    ProcessStatementResponse,
    Warning,
)
from concurrent.futures import ThreadPoolExecutor, as_completed
import random
import time


class BytesSource:
    def __init__(self, name: str, data: bytes) -> None:
        self._name = name
        self._data = data

    def load(self) -> bytes:
        return self._data

    def display_name(self) -> str:
        return self._name


def _profile_dir(settings: Settings) -> Path:
    base = Path(__file__).resolve().parents[2]
    raw = settings.docflow_profile_dir
    path = Path(raw)
    if not path.is_absolute():
        path = (base / path).resolve()
    return path


def _load_profile(profile_name: str, settings: Settings):
    cfg = SdkConfig(profile_dir=_profile_dir(settings))
    return profiles.load_profile(profile_name, cfg)


def _provider(settings: Settings) -> GeminiProvider:
    return GeminiProvider(project=settings.docflow_project, location=settings.docflow_location)


def _name_from_ref(ref: DocumentRef) -> str:
    name = "document"
    if ref.gcsUri:
        name = ref.gcsUri.rsplit("/", 1)[-1] or name
    elif ref.signedUrl:
        path = urlparse(ref.signedUrl).path
        name = path.rsplit("/", 1)[-1] or name
    elif ref.driveFileId:
        name = f"drive_{ref.driveFileId}"

    if ref.mime:
        ext = mimetypes.guess_extension(ref.mime) or ""
        if ext and not name.lower().endswith(ext):
            name = f"{name}{ext}"
    return name


def _ref_to_source(ref: DocumentRef, settings: Settings) -> BytesSource:
    data = fetch_bytes(ref.model_dump(), settings)
    name = _name_from_ref(ref)
    return BytesSource(name, data)


def _is_pdf_bytes(data: bytes, name: str | None, mime: str | None) -> bool:
    if mime and mime.lower() == "application/pdf":
        return True
    if name and name.lower().endswith(".pdf"):
        return True
    return data[:4] == b"%PDF"


def _rasterize_pdf_bytes(pdf_bytes: bytes, base_name: str, max_side: int) -> List[BytesSource]:
    try:
        import fitz  # PyMuPDF
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError("PyMuPDF no está instalado") from exc

    sources: List[BytesSource] = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        if doc.page_count < 1:
            raise ValueError("PDF sin páginas")
        for i in range(doc.page_count):
            page = doc[i]
            width, height = page.rect.width, page.rect.height
            scale = max_side / max(width, height)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            img_bytes = pix.tobytes("png")
            name = f"{base_name}_page_{i+1}.png"
            sources.append(BytesSource(name, img_bytes))
    return sources


def _build_statement_prompt(base_prompt: str | None, statement_parsed: dict[str, Any]) -> str:
    payload = json.dumps(statement_parsed, ensure_ascii=True, indent=2)
    # TODO: convert parsed statement to csv table with 1-based int index before writing the prompt
    parts = []
    if base_prompt:
        parts.append(base_prompt.strip())
    parts.append("ESTADO DE CUENTA PARSEADO (JSON):\n" + payload)
    parts.append(
        "Usa esta info para completar el campo 'Estado de cuenta' en cada item."
    )
    return "\n\n".join(p for p in parts if p)


def _run_extract(
    docs: List[BytesSource],
    profile,
    settings: Settings,
    model: str | None = None,
    multi_mode: str = "per_file",
) -> List[ExtractionResult]:
    options = ProviderOptions(model_name=model) if model else None
    result = _extract_with_retry(
        docs=docs,
        profile=profile,
        settings=settings,
        options=options,
        multi_mode=multi_mode,
    )
    if multi_mode == "aggregate":
        return [result] if not isinstance(result, list) else result
    if isinstance(result, list):
        return result
    if hasattr(result, "per_file"):
        return list(result.per_file)
    return [result]


def _run_extract_single(
    doc: BytesSource,
    profile,
    settings: Settings,
    model: str | None = None,
) -> ExtractionResult:
    options = ProviderOptions(model_name=model) if model else None
    result = _extract_with_retry(
        docs=[doc],
        profile=profile,
        settings=settings,
        options=options,
        multi_mode="per_file",
    )
    if isinstance(result, list):
        return result[0]
    if hasattr(result, "per_file"):
        return list(result.per_file)[0]
    return result


def _run_extract_parallel(
    docs: List[BytesSource],
    profile,
    settings: Settings,
    model: str | None = None,
) -> List[ExtractionResult]:
    if len(docs) <= 1:
        return [_run_extract_single(docs[0], profile, settings, model=model)]
    workers = min(len(docs), settings.docflow_workers)
    results: List[Tuple[int, ExtractionResult]] = []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(_run_extract_single, doc, profile, settings, model): idx
            for idx, doc in enumerate(docs)
        }
        for fut in as_completed(futures):
            idx = futures[fut]
            results.append((idx, fut.result()))
    results.sort(key=lambda x: x[0])
    return [r[1] for r in results]


def _chunk_list(items: List[BytesSource], size: int) -> List[List[BytesSource]]:
    if not items:
        return []
    n = max(1, size or 1)
    return [items[i : i + n] for i in range(0, len(items), n)]


def _extract_with_retry(
    docs: List[BytesSource],
    profile,
    settings: Settings,
    options: ProviderOptions | None,
    multi_mode: str,
):
    attempts = max(1, settings.docflow_retry_max_attempts)
    last_exc: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            return extract(
                docs=docs,
                profile=profile,
                provider=_provider(settings),
                options=options,
                multi_mode=multi_mode,
            )
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            if not _is_retryable_error(exc) or attempt >= attempts:
                raise
            delay = min(
                settings.docflow_retry_max_delay,
                settings.docflow_retry_base_delay
                * (settings.docflow_retry_backoff ** (attempt - 1)),
            )
            jitter = random.random() * 0.3
            time.sleep(delay + jitter)
    if last_exc:
        raise last_exc
    raise RuntimeError("DocFlow extract failed without exception")


def _is_retryable_error(exc: Exception) -> bool:
    msg = str(exc).lower()
    if "resource exhausted" in msg:
        return True
    if "rate limit" in msg:
        return True
    if "429" in msg:
        return True
    if "quota" in msg:
        return True
    return False


def run_process_statement(request: ProcessStatementRequest, settings: Settings) -> ProcessStatementResponse:
    warnings: List[Warning] = []
    profile_name = request.options.profile if request.options and request.options.profile else "estado/v0"
    model = request.options.model if request.options else None

    try:
        profile = _load_profile(profile_name, settings)
        ref = request.statement
        data = fetch_bytes(ref.model_dump(), settings)
        name = _name_from_ref(ref)
        if _is_pdf_bytes(data, name, ref.mime):
            base = Path(name).stem or "statement"
            docs = _rasterize_pdf_bytes(data, base, settings.default_max_side_px)
        else:
            docs = [BytesSource(name, data)]
        multi_mode = "aggregate" if len(docs) > 1 else "per_file"
        results = _run_extract(docs, profile, settings, model=model, multi_mode=multi_mode)
        if not results:
            raise ValueError("No extraction results returned")
        result = results[0]
        return ProcessStatementResponse(
            ok=True,
            rendicionId=request.rendicionId,
            data=result.data,
            meta=result.meta,
            warnings=warnings or None,
            error=None,
        )
    except Exception as exc:  # noqa: BLE001
        return ProcessStatementResponse(
            ok=False,
            rendicionId=request.rendicionId,
            data=None,
            meta=None,
            warnings=warnings or None,
            error=ErrorPayload(code="PROCESS_STATEMENT_FAILED", message=str(exc), details={}),
        )


def run_process_receipts_batch(
    request: ProcessReceiptsBatchRequest, settings: Settings
) -> ProcessReceiptsBatchResponse:
    warnings: List[Warning] = []
    profile_name = request.options.profile if request.options and request.options.profile else "lineas_gastos/v0"
    model = request.options.model if request.options else None

    try:
        profile = _load_profile(profile_name, settings)

        if request.statement and request.statement.parsed:
            prompt = _build_statement_prompt(profile.prompt, request.statement.parsed)
            profile = replace(profile, prompt=prompt)

        docs = [_ref_to_source(it, settings) for it in request.receipts]
        batches = _chunk_list(docs, settings.docflow_batch_size)
        results: List[ExtractionResult] = []
        for batch in batches:
            results.extend(_run_extract_parallel(batch, profile, settings, model=model))

        rows = [DocflowRow(data=res.data, meta=res.meta) for res in results]
        return ProcessReceiptsBatchResponse(
            ok=True,
            rendicionId=request.rendicionId,
            rows=rows,
            warnings=warnings or None,
            error=None,
        )
    except Exception as exc:  # noqa: BLE001
        return ProcessReceiptsBatchResponse(
            ok=False,
            rendicionId=request.rendicionId,
            rows=None,
            warnings=warnings or None,
            error=ErrorPayload(code="PROCESS_RECEIPTS_FAILED", message=str(exc), details={}),
        )
