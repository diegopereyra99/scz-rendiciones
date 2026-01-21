from functools import lru_cache
from pydantic import Field
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    app_name: str = Field("run-rendiciones-service", description="Service identifier for logs/metrics.")
    environment: str = Field("local", description="Deployment environment name.")

    gcs_bucket: str | None = Field(
        default=None,
        description="Default bucket to use when clients only provide prefixes.",
    )
    drive_enabled: bool = Field(
        default=False,
        description="Whether Drive API is enabled and credentials can read/write Drive.",
    )

    default_jpg_quality: int = Field(90, description="JPEG quality used when none is provided.")
    default_max_side_px: int = Field(2000, description="Max side in pixels when resizing images.")
    default_pdf_mode: str = Field(
        "keep",
        description='Default PDF handling strategy ("keep" | "rasterize").',
    )
    default_signed_url_ttl: int = Field(
        3600,
        description="TTL in seconds for signed URLs when requested.",
    )

    xlsm_template_path: str = Field(
        "templates/rendiciones_macro_template.xlsm",
        description="Path to the baked-in XLSM template (copied into the image).",
    )
    docflow_profile_dir: str = Field(
        ".",
        description="DocFlow catalog root (should contain profiles/).",
    )
    docflow_project: str | None = Field(
        default=None,
        description="GCP project for Vertex AI (optional).",
    )
    docflow_location: str | None = Field(
        default=None,
        description="GCP location for Vertex AI (optional).",
    )
    docflow_workers: int = Field(
        4,
        description="Parallel workers for DocFlow per-file extraction.",
        ge=1,
    )
    docflow_batch_size: int = Field(
        6,
        description="Max docs per extraction batch inside process_receipts_batch.",
        ge=1,
    )
    docflow_retry_max_attempts: int = Field(
        4,
        description="Max retry attempts for DocFlow extract calls.",
        ge=1,
    )
    docflow_retry_base_delay: float = Field(
        1.5,
        description="Base delay (seconds) for DocFlow retry backoff.",
        ge=0,
    )
    docflow_retry_max_delay: float = Field(
        15.0,
        description="Max delay (seconds) for DocFlow retry backoff.",
        ge=0,
    )
    docflow_retry_backoff: float = Field(
        2.0,
        description="Backoff multiplier for DocFlow retries.",
        ge=1.0,
    )
    normalize_workers: int = Field(
        4,
        description="Parallel workers for /v1/normalize download + processing.",
        ge=1,
    )

    class Config:
        env_prefix = "REN_"
        case_sensitive = False


@lru_cache
def get_settings() -> Settings:
    return Settings()
