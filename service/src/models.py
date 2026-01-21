from typing import Any, List, Literal, Optional

from pydantic import BaseModel, Field, model_validator


class Warning(BaseModel):
    code: str
    message: str
    details: dict[str, Any] | None = None


class ErrorPayload(BaseModel):
    code: str
    message: str
    details: dict[str, Any] | None = None


# -------- Normalize --------
class InlineFile(BaseModel):
    filename: str
    contentBase64: str


class NormalizeInput(BaseModel):
    driveFolderId: str | None = Field(default=None)
    zipBase64: str | None = Field(default=None)
    zipFilename: str | None = Field(default=None)
    zipGcsUri: str | None = Field(default=None)
    driveFileIds: List[str] | None = Field(default=None)
    files: List[InlineFile] | None = Field(default=None)

    @model_validator(mode="after")
    def validate_one_of(cls, values: "NormalizeInput") -> "NormalizeInput":
        provided = [
            bool(values.driveFolderId),
            bool(values.zipBase64),
            bool(values.zipGcsUri),
            bool(values.driveFileIds),
            bool(values.files),
        ]
        if sum(provided) != 1:
            raise ValueError(
                "Exactly one of driveFolderId, zipBase64, zipGcsUri, driveFileIds, or files must be provided"
            )
        return values


class NormalizeOutput(BaseModel):
    gcsPrefix: str = Field(
        ...,
        description="gs://bucket/path/prefix/",
    )


class NormalizeOptions(BaseModel):
    jpgQuality: int = Field(default=90, ge=1, le=100)
    maxSidePx: int = Field(default=2000, ge=1)
    pdfMode: Literal["keep", "rasterize"] = Field(default="keep")
    uploadOriginals: bool = Field(
        default=False,
        description="If true, upload originals under originals/; default false.",
    )


class NormalizeRequest(BaseModel):
    rendicionId: str
    input: NormalizeInput
    output: NormalizeOutput
    options: NormalizeOptions | None = None


class NormalizedArtifact(BaseModel):
    gcsUri: str
    mime: str
    sha256: str
    bytes: int | None = None
    pageCount: int | None = None
    originalGcsUri: str | None = None
    originalMime: str | None = None


class SourceInfo(BaseModel):
    driveFileId: str | None = None
    originalName: str | None = None


class NormalizeItem(BaseModel):
    source: SourceInfo
    normalized: NormalizedArtifact


class NormalizeResponse(BaseModel):
    ok: bool
    rendicionId: str
    items: List[NormalizeItem]
    manifestGcsUri: str | None = None
    warnings: List[Warning] | None = None
    error: ErrorPayload | None = None


# -------- Finalize --------
class CoverRef(BaseModel):
    gcsUri: str | None = None
    driveFileId: str | None = None
    signedUrl: str | None = None

    @model_validator(mode="after")
    def validate_one_of(cls, values: "CoverRef") -> "CoverRef":
        provided = [bool(values.gcsUri), bool(values.driveFileId), bool(values.signedUrl)]
        if sum(provided) != 1:
            raise ValueError("Exactly one of cover.gcsUri, cover.driveFileId, or cover.signedUrl is required")
        return values


class XlsmTemplateRef(BaseModel):
    gcsUri: str | None = None
    driveFileId: str | None = None

    @model_validator(mode="after")
    def validate_one_of(cls, values: "XlsmTemplateRef") -> "XlsmTemplateRef":
        provided = [bool(values.gcsUri), bool(values.driveFileId)]
        if sum(provided) != 1:
            raise ValueError("Exactly one of xlsmTemplate.gcsUri or xlsmTemplate.driveFileId is required")
        return values


class FinalizeNormalizedItem(BaseModel):
    gcsUri: str
    mime: str | None = None
    originalName: str | None = None


class XlsmValue(BaseModel):
    sheet: str
    row: int
    col: int
    value: Any


class FinalizeInputs(BaseModel):
    normalizedItems: List[FinalizeNormalizedItem]
    cover: CoverRef
    xlsmTemplate: XlsmTemplateRef | None = None
    xlsmValues: List[XlsmValue] = Field(default_factory=list)


class FinalizeOutput(BaseModel):
    driveFolderId: str | None = None
    gcsPrefix: str | None = None

    @model_validator(mode="after")
    def validate_one_of(cls, values: "FinalizeOutput") -> "FinalizeOutput":
        provided = [bool(values.driveFolderId), bool(values.gcsPrefix)]
        if sum(provided) != 1:
            raise ValueError("Exactly one of driveFolderId or gcsPrefix is required")
        return values


class FinalizeOptions(BaseModel):
    pdfName: str = Field(default="rendicion.pdf")
    xlsmName: str = Field(default="rendicion.xlsm")
    mergeOrder: Literal["cover_first"] = Field(default="cover_first")
    signedUrlTtlSeconds: int | None = Field(default=3600, ge=60)


class FinalizeRequest(BaseModel):
    rendicionId: str
    inputs: FinalizeInputs
    output: FinalizeOutput
    options: FinalizeOptions | None = None


class FinalizeArtifact(BaseModel):
    driveFileId: str | None = None
    gcsUri: str | None = None
    signedUrl: str | None = None


class FinalizeResponse(BaseModel):
    ok: bool
    rendicionId: str
    pdf: FinalizeArtifact
    xlsm: FinalizeArtifact
    warnings: List[Warning] | None = None
    error: ErrorPayload | None = None


# -------- Stage 2 --------
class DocumentRef(BaseModel):
    gcsUri: str | None = None
    signedUrl: str | None = None
    driveFileId: str | None = None
    mime: str | None = None

    @model_validator(mode="after")
    def validate_one_of(cls, values: "DocumentRef") -> "DocumentRef":
        provided = [bool(values.gcsUri), bool(values.signedUrl), bool(values.driveFileId)]
        if sum(provided) != 1:
            raise ValueError("Exactly one of gcsUri, signedUrl, or driveFileId is required")
        return values


class ProcessOptions(BaseModel):
    profile: str | None = None
    model: str | None = None


class ProcessStatementRequest(BaseModel):
    rendicionId: str
    statement: DocumentRef
    options: ProcessOptions | None = None


class ProcessStatementResponse(BaseModel):
    ok: bool
    rendicionId: str
    data: Any | None = None
    meta: dict[str, Any] | None = None
    warnings: List[Warning] | None = None
    error: ErrorPayload | None = None


class StatementContext(BaseModel):
    parsed: dict[str, Any] | None = None


class DocflowRow(BaseModel):
    data: Any
    meta: dict[str, Any] | None = None


class ProcessReceiptsBatchRequest(BaseModel):
    rendicionId: str
    mode: Literal["efectivo", "tarjeta"]
    receipts: List[DocumentRef]
    statement: StatementContext | None = None
    options: ProcessOptions | None = None

    @model_validator(mode="after")
    def validate_statement_for_mode(cls, values: "ProcessReceiptsBatchRequest") -> "ProcessReceiptsBatchRequest":
        if values.mode == "tarjeta" and not values.statement:
            raise ValueError("statement.parsed is required when mode=tarjeta")
        return values


class ProcessReceiptsBatchResponse(BaseModel):
    ok: bool
    rendicionId: str
    rows: List[DocflowRow] | None = None
    warnings: List[Warning] | None = None
    error: ErrorPayload | None = None
