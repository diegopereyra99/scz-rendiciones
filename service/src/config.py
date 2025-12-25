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

    class Config:
        env_prefix = "REN_"
        case_sensitive = False


@lru_cache
def get_settings() -> Settings:
    return Settings()
