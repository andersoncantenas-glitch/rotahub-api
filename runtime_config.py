import json
import logging
import os
import re
import sys
from dataclasses import asdict, dataclass
from typing import Any, Dict, Optional, Type
from version import APP_VERSION


class BaseConfig:
    APP_ENV = "production"
    API_BASE_URL = "https://rotahub-api.onrender.com"
    ENABLE_API_SYNC = True
    ENABLE_SQL_MIRROR = True
    REQUIRE_SERVER_BINDING = True
    ALLOW_REMOTE_WRITE = True
    ALLOW_REMOTE_READ = True
    ALLOW_DEV_DATA_UPLOAD = True
    ALLOW_SEED_DB = False
    ALLOW_VERSION_UPDATE = True
    TENANT_MODE = "database-per-tenant"
    UPDATE_CHANNEL = "stable"
    SOURCE_OF_TRUTH = "api-central"
    API_TIMEOUT = 60.0
    DESKTOP_DB_NAME = "rotahub_desktop.db"
    SERVER_DB_NAME = "rotahub_server.db"


class DevelopmentConfig(BaseConfig):
    APP_ENV = "development"
    API_BASE_URL = "http://127.0.0.1:8000"
    ENABLE_API_SYNC = False
    ENABLE_SQL_MIRROR = False
    REQUIRE_SERVER_BINDING = False
    ALLOW_REMOTE_WRITE = False
    ALLOW_REMOTE_READ = False
    ALLOW_DEV_DATA_UPLOAD = False
    ALLOW_VERSION_UPDATE = False
    UPDATE_CHANNEL = "disabled"
    SOURCE_OF_TRUTH = "sqlite-local"


class StagingConfig(BaseConfig):
    APP_ENV = "staging"
    API_BASE_URL = "https://staging.rotahub-api.example.com"
    UPDATE_CHANNEL = "staging"


class ProductionConfig(BaseConfig):
    APP_ENV = "production"


ENVIRONMENT_CONFIGS: Dict[str, Type[BaseConfig]] = {
    "development": DevelopmentConfig,
    "staging": StagingConfig,
    "production": ProductionConfig,
}


def get_environment_profile(app_env: str) -> Type[BaseConfig]:
    return ENVIRONMENT_CONFIGS.get(str(app_env or "").strip().lower(), ProductionConfig)


def _slugify(value: str, fallback: str) -> str:
    raw = str(value or "").strip().lower()
    raw = re.sub(r"[^a-z0-9._-]+", "-", raw)
    raw = raw.strip("-._")
    return raw or fallback


def _read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception:
        logging.debug("Falha ao ler JSON de configuracao: %s", path, exc_info=True)
        return {}


def _merge_dicts(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    merged = dict(base or {})
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_dicts(merged[key], value)
        else:
            merged[key] = value
    return merged


def _normalize_env(value: Optional[str], *, is_frozen: bool) -> str:
    raw = str(value or "").strip().lower()
    aliases = {
        "dev": "development",
        "development": "development",
        "local": "development",
        "homolog": "staging",
        "staging": "staging",
        "stage": "staging",
        "prod": "production",
        "production": "production",
    }
    normalized = aliases.get(raw, raw)
    if normalized in {"development", "staging", "production"}:
        return normalized
    return "production" if is_frozen else "development"


@dataclass
class AppConfig:
    app_kind: str
    app_env: str
    app_version: str
    is_frozen: bool
    app_dir: str
    resource_dir: str
    config_file: str
    config_source: str
    env_file: str
    data_root: str
    runtime_dir: str
    db_path: str
    api_base_url: str
    update_manifest_url: str
    setup_download_url: str
    changelog_url: str
    support_whatsapp: str
    support_email: str
    log_level: str
    api_sync_timeout: float
    sync_enabled: bool
    sql_mirror_api: bool
    require_server_binding: bool
    allow_remote_write: bool
    allow_dev_data_upload: bool
    allow_seed_db: bool
    allow_version_update: bool
    allow_remote_read: bool
    update_channel: str
    tenant_mode: str
    source_of_truth: str
    desktop_secret: str
    app_title: str
    tenant_id: str
    company_id: str
    sync_mode: str
    cache_dir: str
    updates_dir: str
    temp_dir: str
    schema_version: int

    def diagnostics(self) -> Dict[str, Any]:
        data = asdict(self)
        data["desktop_secret"] = "***" if self.desktop_secret else ""
        return data


def _build_default_db_path(*, app_kind: str, data_root: str, app_env: str, tenant_id: str, is_frozen: bool) -> str:
    tenant_key = _slugify(tenant_id, "default")
    profile = get_environment_profile(app_env)
    if app_kind == "server":
        return os.path.join(data_root, "server", app_env, tenant_key, profile.SERVER_DB_NAME)
    db_name = profile.DESKTOP_DB_NAME if is_frozen else "rota_granja.db"
    return os.path.join(data_root, "desktop", app_env, tenant_key, db_name)


def _default_data_root(*, app_kind: str, app_dir: str, is_frozen: bool) -> str:
    if app_kind == "server":
        return os.path.join(app_dir, ".rotahub_runtime")
    if is_frozen:
        base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
        return os.path.join(base, "RotaHubDesktop")
    return os.path.join(app_dir, ".rotahub_runtime")


def _default_config_file(*, app_kind: str, app_dir: str, data_root: str, is_frozen: bool) -> str:
    env_path = os.environ.get("ROTA_CONFIG_FILE", "").strip()
    if env_path:
        return env_path
    if app_kind == "server":
        return os.path.join(app_dir, "config", "server.runtime.json")
    if is_frozen:
        return os.path.join(data_root, "config.json")
    return os.path.join(app_dir, "config", "desktop.runtime.json")


def _default_desktop_secret(*, app_kind: str, is_frozen: bool, app_env: str) -> str:
    if app_kind == "desktop" and is_frozen and app_env in {"staging", "production"}:
        return "rota-secreta"
    return ""


def resolve_database_path(
    *,
    explicit_db_path: str,
    runtime_cfg: Dict[str, Any],
    app_kind: str,
    data_root: str,
    app_env: str,
    tenant_id: str,
    is_frozen: bool,
) -> str:
    candidate = str(explicit_db_path or runtime_cfg.get("db_path") or "").strip()
    if not candidate:
        candidate = _build_default_db_path(
            app_kind=app_kind,
            data_root=data_root,
            app_env=app_env,
            tenant_id=tenant_id,
            is_frozen=is_frozen,
        )
    candidate = os.path.abspath(candidate)
    lower = candidate.lower()
    if app_env == "development" and ("production" in lower or "\\dist\\" in lower):
        raise RuntimeError(f"DB_PATH inseguro para development: {candidate}")
    if app_env == "staging" and "\\development\\" in lower:
        raise RuntimeError(f"DB_PATH inseguro para staging: {candidate}")
    if app_kind == "server" and app_env in {"staging", "production"} and os.path.basename(lower) == "rota_granja.db":
        raise RuntimeError(f"Servidor publicado nao pode usar banco desktop legado: {candidate}")
    if app_kind == "desktop" and is_frozen and "\\desktop\\development\\" in lower:
        raise RuntimeError(f"Desktop publicado nao pode usar banco development: {candidate}")
    return candidate


def validate_runtime_guardrails(config: AppConfig) -> None:
    db_lower = os.path.abspath(config.db_path).lower()
    app_dir_lower = os.path.abspath(config.app_dir).lower()

    if config.app_env == "development":
        if config.allow_dev_data_upload:
            raise RuntimeError("APP_ENV=development nao pode habilitar ALLOW_DEV_DATA_UPLOAD.")
        if config.sync_enabled or config.sql_mirror_api or config.allow_remote_write:
            raise RuntimeError("APP_ENV=development nao pode sincronizar ou gravar remotamente.")

    if config.app_kind == "server" and config.app_env in {"staging", "production"}:
        if "\\desktop\\development\\" in db_lower or db_lower.startswith(app_dir_lower + os.sep) and db_lower.endswith("\\rota_granja.db"):
            raise RuntimeError(f"Servidor publicado apontando para base development/local: {config.db_path}")

    if config.app_kind == "desktop" and config.is_frozen:
        if "\\desktop\\development\\" in db_lower or db_lower.startswith(app_dir_lower + os.sep):
            raise RuntimeError(f"Desktop publicado apontando para workspace local/development: {config.db_path}")

    if config.app_kind == "desktop" and config.is_frozen and config.app_env in {"staging", "production"}:
        if config.source_of_truth != "api-central":
            raise RuntimeError("Desktop publicado deve usar source_of_truth='api-central'.")
        if not config.sync_enabled or not config.allow_remote_read:
            raise RuntimeError("Desktop publicado deve manter leitura remota ativa para usar somente a base publicada.")


def load_app_config(app_kind: str = "desktop") -> AppConfig:
    is_frozen = getattr(sys, "frozen", False)
    custom_config_requested = bool(os.environ.get("ROTA_CONFIG_FILE", "").strip())
    local_desktop_mode = app_kind == "desktop" and not is_frozen and not custom_config_requested
    local_desktop_custom_mode = app_kind == "desktop" and not is_frozen and custom_config_requested
    desktop_locked_runtime = app_kind == "desktop" and is_frozen
    app_dir = os.path.dirname(os.path.abspath(__file__))
    resource_dir = getattr(sys, "_MEIPASS", app_dir)
    config_root = os.path.join(app_dir, "config")

    data_root = _default_data_root(app_kind=app_kind, app_dir=app_dir, is_frozen=is_frozen)
    config_file = _default_config_file(
        app_kind=app_kind,
        app_dir=app_dir,
        data_root=data_root,
        is_frozen=is_frozen,
    )
    file_config = _read_json(config_file)

    def env_value(name: str) -> Optional[str]:
        if local_desktop_mode or local_desktop_custom_mode or desktop_locked_runtime:
            return None
        return os.environ.get(name)

    app_env = _normalize_env(
        ("development" if local_desktop_mode else (env_value("ROTA_APP_ENV") or file_config.get("app_env"))),
        is_frozen=is_frozen,
    )
    env_file = os.path.join(config_root, "environments", f"{app_env}.json")
    env_defaults = _read_json(env_file)
    merged = _merge_dicts(env_defaults, file_config)
    profile = get_environment_profile(app_env)

    runtime_cfg = merged.get("runtime") if isinstance(merged.get("runtime"), dict) else {}
    api_cfg = merged.get("api") if isinstance(merged.get("api"), dict) else {}
    update_cfg = merged.get("update") if isinstance(merged.get("update"), dict) else {}
    support_cfg = merged.get("support") if isinstance(merged.get("support"), dict) else {}
    tenant_cfg = merged.get("tenant") if isinstance(merged.get("tenant"), dict) else {}
    logging_cfg = merged.get("logging") if isinstance(merged.get("logging"), dict) else {}

    company_default = "dev-local" if local_desktop_mode or app_env == "development" else "default-company"
    company_id = str(
        env_value("ROTA_COMPANY_ID")
        or env_value("ROTA_TENANT_ID")
        or tenant_cfg.get("company_id")
        or tenant_cfg.get("tenant_id")
        or company_default
    ).strip()
    tenant_id = str(
        env_value("ROTA_TENANT_ID")
        or tenant_cfg.get("tenant_id")
        or company_id
    ).strip()
    tenant_id = _slugify(tenant_id, "default-company")
    company_id = _slugify(company_id, tenant_id)

    configured_data_root = str((env_value("ROTA_DATA_ROOT")) or runtime_cfg.get("data_root") or data_root).strip()
    data_root = configured_data_root or data_root

    db_path = resolve_database_path(
        explicit_db_path=env_value("ROTA_DB") or "",
        runtime_cfg={} if local_desktop_mode else runtime_cfg,
        app_kind=app_kind,
        data_root=data_root,
        app_env=app_env,
        tenant_id=tenant_id,
        is_frozen=is_frozen,
    )

    default_api_url = profile.API_BASE_URL
    default_sync = profile.ENABLE_API_SYNC
    default_sql_mirror = profile.ENABLE_SQL_MIRROR
    default_binding = profile.REQUIRE_SERVER_BINDING
    default_update_channel = profile.UPDATE_CHANNEL
    default_allow_remote_write = profile.ALLOW_REMOTE_WRITE
    default_allow_version_update = profile.ALLOW_VERSION_UPDATE
    default_allow_seed_db = profile.ALLOW_SEED_DB
    default_allow_remote_read = profile.ALLOW_REMOTE_READ
    default_source_of_truth = profile.SOURCE_OF_TRUTH
    default_allow_dev_data_upload = profile.ALLOW_DEV_DATA_UPLOAD

    api_base_url = str((env_value("ROTA_SERVER_URL")) or ("" if local_desktop_mode else api_cfg.get("base_url")) or default_api_url).strip().rstrip("/")
    try:
        api_sync_timeout = float((env_value("ROTA_SYNC_TIMEOUT")) or api_cfg.get("timeout") or 60)
    except Exception:
        api_sync_timeout = 60.0

    sync_enabled_raw = env_value("ROTA_DESKTOP_SYNC_API")
    if local_desktop_mode:
        sync_enabled = False
    elif sync_enabled_raw is None:
        sync_enabled = bool(runtime_cfg.get("sync_enabled", default_sync))
    else:
        sync_enabled = str(sync_enabled_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    sql_mirror_raw = env_value("ROTA_SQL_MIRROR_API")
    if local_desktop_mode:
        sql_mirror_api = False
    elif sql_mirror_raw is None:
        sql_mirror_api = bool(runtime_cfg.get("sql_mirror_api", default_sql_mirror))
    else:
        sql_mirror_api = str(sql_mirror_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    binding_raw = env_value("ROTA_REQUIRE_SERVER_BINDING")
    if local_desktop_mode:
        require_server_binding = False
    elif binding_raw is None:
        require_server_binding = bool(runtime_cfg.get("require_server_binding", default_binding))
    else:
        require_server_binding = str(binding_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    desktop_secret = str(
        (env_value("ROTA_SECRET"))
        or runtime_cfg.get("desktop_secret")
        or _default_desktop_secret(app_kind=app_kind, is_frozen=is_frozen, app_env=app_env)
        or ""
    ).strip()
    if app_kind == "desktop" and is_frozen and not desktop_secret:
        sql_mirror_api = False
    allow_remote_write_raw = env_value("ROTA_ALLOW_REMOTE_WRITE")
    if local_desktop_mode:
        allow_remote_write = False
    elif allow_remote_write_raw is None:
        allow_remote_write = bool(runtime_cfg.get("allow_remote_write", default_allow_remote_write))
    else:
        allow_remote_write = str(allow_remote_write_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_dev_upload_raw = env_value("ROTA_ALLOW_DEV_DATA_UPLOAD")
    if local_desktop_mode:
        allow_dev_data_upload = False
    elif allow_dev_upload_raw is None:
        allow_dev_data_upload = bool(runtime_cfg.get("allow_dev_data_upload", default_allow_dev_data_upload))
    else:
        allow_dev_data_upload = str(allow_dev_upload_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_remote_read_raw = env_value("ROTA_ALLOW_REMOTE_READ")
    if local_desktop_mode:
        allow_remote_read = False
    elif allow_remote_read_raw is None:
        allow_remote_read = bool(runtime_cfg.get("allow_remote_read", default_allow_remote_read))
    else:
        allow_remote_read = str(allow_remote_read_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_seed_db_raw = env_value("ROTA_ALLOW_SEED_DB")
    if local_desktop_mode:
        allow_seed_db = False
    elif allow_seed_db_raw is None:
        allow_seed_db = bool(runtime_cfg.get("allow_seed_db", default_allow_seed_db))
    else:
        allow_seed_db = str(allow_seed_db_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_version_update_raw = env_value("ROTA_ALLOW_VERSION_UPDATE")
    if local_desktop_mode:
        allow_version_update = False
    elif allow_version_update_raw is None:
        allow_version_update = bool(runtime_cfg.get("allow_version_update", default_allow_version_update))
    else:
        allow_version_update = str(allow_version_update_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    update_channel = str(
        env_value("ROTA_UPDATE_CHANNEL")
        or update_cfg.get("channel")
        or default_update_channel
    ).strip() or default_update_channel
    tenant_mode = str(runtime_cfg.get("tenant_mode") or "database-per-tenant").strip() or "database-per-tenant"
    log_level = str(logging_cfg.get("level") or env_value("ROTA_LOG_LEVEL") or ("DEBUG" if app_env == "development" else "INFO")).strip().upper()
    source_of_truth = str(runtime_cfg.get("source_of_truth") or default_source_of_truth).strip() or "sqlite-local"
    update_manifest_url = str(
        env_value("ROTA_UPDATE_MANIFEST_URL")
        or update_cfg.get("manifest_url")
        or "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/version.json"
    ).strip()
    setup_download_url = str(env_value("ROTA_SETUP_URL") or update_cfg.get("setup_url") or "").strip()
    changelog_url = str(env_value("ROTA_CHANGELOG_URL") or update_cfg.get("changelog_url") or "").strip()
    support_whatsapp = str(env_value("ROTA_SUPPORT_WHATSAPP") or support_cfg.get("whatsapp") or "").strip()
    support_email = str(env_value("ROTA_SUPPORT_EMAIL") or support_cfg.get("email") or "").strip()
    app_title = str(merged.get("app_title") or "ROTAHUB DESKTOP").strip()
    runtime_dir = os.path.dirname(db_path) or data_root
    sync_mode = "local-only" if not sync_enabled else ("tenant-server-sync" if tenant_id else "server-sync")
    cache_dir = os.path.join(runtime_dir, "cache")
    updates_dir = os.path.join(data_root, "updates", app_env, tenant_id)
    temp_dir = os.path.join(runtime_dir, "tmp")
    schema_version = int(runtime_cfg.get("schema_version") or 1)

    os.makedirs(runtime_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    os.makedirs(updates_dir, exist_ok=True)
    os.makedirs(temp_dir, exist_ok=True)

    config = AppConfig(
        app_kind=app_kind,
        app_env=app_env,
        app_version=APP_VERSION,
        is_frozen=is_frozen,
        app_dir=app_dir,
        resource_dir=resource_dir,
        config_file=config_file,
        config_source=config_file if os.path.exists(config_file) else env_file,
        env_file=env_file,
        data_root=data_root,
        runtime_dir=runtime_dir,
        db_path=db_path,
        api_base_url=api_base_url,
        update_manifest_url=update_manifest_url,
        setup_download_url=setup_download_url,
        changelog_url=changelog_url,
        support_whatsapp=support_whatsapp,
        support_email=support_email,
        log_level=log_level,
        api_sync_timeout=api_sync_timeout,
        sync_enabled=sync_enabled,
        sql_mirror_api=sql_mirror_api and sync_enabled,
        require_server_binding=require_server_binding and sync_enabled,
        allow_remote_write=allow_remote_write and sync_enabled,
        allow_dev_data_upload=allow_dev_data_upload and app_env != "development",
        allow_seed_db=allow_seed_db,
        allow_version_update=allow_version_update,
        allow_remote_read=allow_remote_read,
        update_channel=update_channel,
        tenant_mode=tenant_mode,
        source_of_truth=source_of_truth,
        desktop_secret=desktop_secret,
        app_title=app_title,
        tenant_id=tenant_id,
        company_id=company_id,
        sync_mode=sync_mode,
        cache_dir=cache_dir,
        updates_dir=updates_dir,
        temp_dir=temp_dir,
        schema_version=schema_version,
    )
    validate_runtime_guardrails(config)
    return config


def apply_process_environment(config: AppConfig) -> None:
    env_values = {
        "ROTA_APP_ENV": config.app_env,
        "ROTA_DB": config.db_path,
        "ROTA_SERVER_URL": config.api_base_url,
        "ROTA_SYNC_TIMEOUT": str(config.api_sync_timeout),
        "ROTA_DESKTOP_SYNC_API": "1" if config.sync_enabled else "0",
        "ROTA_SQL_MIRROR_API": "1" if config.sql_mirror_api else "0",
        "ROTA_REQUIRE_SERVER_BINDING": "1" if config.require_server_binding else "0",
        "ROTA_ALLOW_REMOTE_WRITE": "1" if config.allow_remote_write else "0",
        "ROTA_ALLOW_DEV_DATA_UPLOAD": "1" if config.allow_dev_data_upload else "0",
        "ROTA_ALLOW_REMOTE_READ": "1" if config.allow_remote_read else "0",
        "ROTA_ALLOW_SEED_DB": "1" if config.allow_seed_db else "0",
        "ROTA_ALLOW_VERSION_UPDATE": "1" if config.allow_version_update else "0",
        "ROTA_UPDATE_CHANNEL": config.update_channel,
        "ROTA_TENANT_MODE": config.tenant_mode,
        "ROTA_SOURCE_OF_TRUTH": config.source_of_truth,
        "ROTA_LOG_LEVEL": config.log_level,
        "ROTA_TENANT_ID": config.tenant_id,
        "ROTA_COMPANY_ID": config.company_id,
        "ROTA_UPDATE_MANIFEST_URL": config.update_manifest_url,
        "ROTA_SETUP_URL": config.setup_download_url,
        "ROTA_CHANGELOG_URL": config.changelog_url,
        "ROTA_SUPPORT_WHATSAPP": config.support_whatsapp,
        "ROTA_SUPPORT_EMAIL": config.support_email,
    }
    if config.desktop_secret:
        env_values["ROTA_SECRET"] = config.desktop_secret
    for key, value in env_values.items():
        os.environ[key] = str(value)


def ensure_runtime_files(config: AppConfig) -> None:
    os.makedirs(os.path.dirname(config.config_file) or ".", exist_ok=True)
    if not os.path.exists(config.config_file):
        template = {
            "app_env": config.app_env,
            "tenant": {
                "tenant_id": config.tenant_id,
                "company_id": config.company_id,
            },
            "runtime": {
                "data_root": config.data_root,
                "db_path": config.db_path,
                "sync_enabled": config.sync_enabled,
                "sql_mirror_api": config.sql_mirror_api,
                "require_server_binding": config.require_server_binding,
                "allow_remote_write": config.allow_remote_write,
                "allow_dev_data_upload": config.allow_dev_data_upload,
                "allow_remote_read": config.allow_remote_read,
                "allow_seed_db": config.allow_seed_db,
                "allow_version_update": config.allow_version_update,
                "tenant_mode": config.tenant_mode,
                "source_of_truth": config.source_of_truth,
                "schema_version": config.schema_version,
                "desktop_secret": config.desktop_secret,
            },
            "api": {
                "base_url": config.api_base_url,
                "timeout": config.api_sync_timeout,
            },
            "update": {
                "channel": config.update_channel,
                "manifest_url": config.update_manifest_url,
                "setup_url": config.setup_download_url,
                "changelog_url": config.changelog_url,
            },
            "logging": {
                "level": config.log_level,
            },
            "support": {
                "whatsapp": config.support_whatsapp,
                "email": config.support_email,
            },
        }
        try:
            with open(config.config_file, "w", encoding="utf-8") as handle:
                json.dump(template, handle, ensure_ascii=False, indent=2)
                handle.write("\n")
        except Exception:
            logging.debug("Falha ao gravar config template: %s", config.config_file, exc_info=True)
