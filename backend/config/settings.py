# backend/config/settings.py
"""
Application settings and configuration
"""
import os
import tempfile
from pathlib import Path
from typing import Annotated, Any, List
from pydantic import field_validator
from pydantic_settings import BaseSettings, NoDecode


def default_database_url() -> str:
    rota_db = str(os.getenv("ROTA_DB") or "").strip()
    if rota_db:
        if "://" in rota_db:
            return rota_db
        path = writable_sqlite_path(rota_db)
        return f"sqlite+aiosqlite:///{path.as_posix()}"
    return "sqlite+aiosqlite:///./rotadb.db"


def writable_sqlite_path(value: str) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = path.resolve()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        probe = path.parent / ".rotahub_write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return path
    except Exception:
        fallback = Path(tempfile.gettempdir()) / "rotahub" / (path.name or "rotadb.db")
        fallback.parent.mkdir(parents=True, exist_ok=True)
        return fallback


class Settings(BaseSettings):
    """Application settings"""

    # Environment
    ENVIRONMENT: str = os.getenv("ENVIRONMENT", "development")
    DEBUG: bool = os.getenv("DEBUG", "true").lower() == "true"

    # Server
    HOST: str = os.getenv("HOST", "0.0.0.0")
    PORT: int = int(os.getenv("PORT", "8000"))
    ALLOWED_HOSTS: Annotated[List[str], NoDecode] = os.getenv(
        "ALLOWED_HOSTS",
        "localhost,127.0.0.1,10.0.2.2",
    ).split(",")

    # CORS
    CORS_ORIGINS: Annotated[List[str], NoDecode] = os.getenv(
        "CORS_ORIGINS",
        "http://localhost:3000,http://127.0.0.1:3000,http://localhost:8000,http://10.0.2.2:8000"
    ).split(",")

    # Database
    DATABASE_URL: str = os.getenv("DATABASE_URL") or default_database_url()
    ROTA_DB: str = os.getenv("ROTA_DB", "")

    # Security
    SECRET_KEY: str = os.getenv("SECRET_KEY", "your-secret-key-change-in-production")
    JWT_SECRET_KEY: str = os.getenv("JWT_SECRET_KEY", SECRET_KEY)
    JWT_ALGORITHM: str = "HS256"
    JWT_ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    JWT_REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    @field_validator("ALLOWED_HOSTS", "CORS_ORIGINS", mode="before")
    @classmethod
    def parse_csv_list(cls, value: Any) -> List[str]:
        if isinstance(value, str):
            return [item.strip() for item in value.split(",") if item.strip()]
        return value

    @field_validator("DEBUG", mode="before")
    @classmethod
    def parse_debug(cls, value: Any) -> bool:
        if isinstance(value, bool):
            return value
        if value is None or value == "":
            return True

        text = str(value).strip().lower()
        if text in {"1", "true", "t", "yes", "y", "on", "debug", "development", "dev"}:
            return True
        if text in {"0", "false", "f", "no", "n", "off", "release", "production", "prod"}:
            return False

        raise ValueError("DEBUG must be a boolean-like value")

    # Redis (for cache and rate limiting)
    REDIS_URL: str = os.getenv("REDIS_URL", "redis://localhost:6379")

    # Rate Limiting
    RATE_LIMIT_REQUESTS: int = int(os.getenv("RATE_LIMIT_REQUESTS", "100"))
    RATE_LIMIT_WINDOW: int = int(os.getenv("RATE_LIMIT_WINDOW", "60"))  # seconds

    # File Upload
    MAX_UPLOAD_SIZE: int = int(os.getenv("MAX_UPLOAD_SIZE", "10485760"))  # 10MB

    # SaaS Settings
    DEFAULT_PLAN_VEHICLES: int = 5
    MAX_FREE_VEHICLES: int = 2
    OWNER_ADMIN_USERS: Annotated[List[str], NoDecode] = os.getenv("OWNER_ADMIN_USERS", "ADMIN").split(",")
    ROTA_ENABLE_LEGACY_MOBILE_API: bool = (
        os.getenv("ROTA_ENABLE_LEGACY_MOBILE_API", "0").lower()
        in {"1", "true", "yes", "y", "sim", "on"}
    )

    @field_validator("OWNER_ADMIN_USERS", mode="before")
    @classmethod
    def parse_owner_users(cls, value: Any) -> List[str]:
        if isinstance(value, str):
            return [item.strip().upper() for item in value.split(",") if item.strip()]
        return [str(item).strip().upper() for item in (value or []) if str(item).strip()]

    class Config:
        env_file = ".env"
        case_sensitive = True


# Global settings instance
settings = Settings()
