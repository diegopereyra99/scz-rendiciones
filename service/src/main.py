from fastapi import Depends, FastAPI, HTTPException

from .config import Settings, get_settings
from .models import (
    FinalizeRequest,
    FinalizeResponse,
    NormalizeRequest,
    NormalizeResponse,
    ProcessReceiptsBatchRequest,
    ProcessReceiptsBatchResponse,
    ProcessStatementRequest,
    ProcessStatementResponse,
)
from .services.finalize import run_finalize
from .services.normalize import run_normalize
from .services.process_stage2 import run_process_receipts_batch, run_process_statement

app = FastAPI(
    title="Rendiciones Cloud Run Service",
    version="0.1.0",
    description="Skeleton implementation for normalize/finalize endpoints.",
)


@app.get("/healthz")
async def healthcheck():
    return {"ok": True}


@app.post("/v1/normalize", response_model=NormalizeResponse)
async def normalize(
    request: NormalizeRequest, settings: Settings = Depends(get_settings)
) -> NormalizeResponse:
    response = await run_normalize(request, settings)
    if not response.ok and response.error:
        status = 400 if response.error.code in {"INVALID_ARGUMENT"} else 500
        raise HTTPException(status_code=status, detail=response.error.model_dump())
    return response


@app.post("/v1/finalize", response_model=FinalizeResponse)
async def finalize(
    request: FinalizeRequest, settings: Settings = Depends(get_settings)
) -> FinalizeResponse:
    response = await run_finalize(request, settings)
    if not response.ok and response.error:
        status = 400 if response.error.code in {"INVALID_ARGUMENT"} else 500
        raise HTTPException(status_code=status, detail=response.error.model_dump())
    return response


@app.post("/v1/process_statement", response_model=ProcessStatementResponse)
async def process_statement(
    request: ProcessStatementRequest, settings: Settings = Depends(get_settings)
) -> ProcessStatementResponse:
    response = run_process_statement(request, settings)
    if not response.ok and response.error:
        raise HTTPException(status_code=500, detail=response.error.model_dump())
    return response


@app.post("/v1/process_receipts_batch", response_model=ProcessReceiptsBatchResponse)
async def process_receipts_batch(
    request: ProcessReceiptsBatchRequest, settings: Settings = Depends(get_settings)
) -> ProcessReceiptsBatchResponse:
    response = run_process_receipts_batch(request, settings)
    if not response.ok and response.error:
        raise HTTPException(status_code=500, detail=response.error.model_dump())
    return response


# NOTE: uvicorn entrypoint is declared in Dockerfile; keep for local dev.
def get_app() -> FastAPI:
    return app


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("src.main:app", host="0.0.0.0", port=8080, reload=True)
