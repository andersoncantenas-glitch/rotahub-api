from dataclasses import dataclass


@dataclass(frozen=True)
class BaseConfig:
    APP_ENV: str = "development"
    ENABLE_API_SYNC: bool = False
    ENABLE_SQL_MIRROR: bool = False
    LOG_LEVEL: str = "INFO"
    UPDATE_CHANNEL: str = "disabled"
    TENANT_MODE: str = "database-per-tenant"
    ALLOW_SEED_DB: bool = False
    ALLOW_REMOTE_WRITE: bool = False
    ALLOW_VERSION_UPDATE: bool = False


@dataclass(frozen=True)
class DevelopmentConfig(BaseConfig):
    APP_ENV: str = "development"
    LOG_LEVEL: str = "DEBUG"


@dataclass(frozen=True)
class StagingConfig(BaseConfig):
    APP_ENV: str = "staging"
    ENABLE_API_SYNC: bool = True
    ENABLE_SQL_MIRROR: bool = True
    UPDATE_CHANNEL: str = "staging"
    ALLOW_REMOTE_WRITE: bool = True
    ALLOW_VERSION_UPDATE: bool = True


@dataclass(frozen=True)
class ProductionConfig(BaseConfig):
    APP_ENV: str = "production"
    ENABLE_API_SYNC: bool = True
    ENABLE_SQL_MIRROR: bool = True
    UPDATE_CHANNEL: str = "stable"
    ALLOW_REMOTE_WRITE: bool = True
    ALLOW_VERSION_UPDATE: bool = True
