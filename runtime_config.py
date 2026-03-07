import json
import logging
import os
import re
import sys
from dataclasses import asdict, dataclass
from typing import Any, Dict, Optional
from version import APP_VERSION


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


def _build_default_db_path(*, app_kind: str, data_root: str, app_env: str, tenant_id: str) -> str:
    tenant_key = _slugify(tenant_id, "default")
    if app_kind == "server":
        return os.path.join(data_root, "server", app_env, tenant_key, "rotahub_server.db")
    return os.path.join(data_root, "desktop", app_env, tenant_key, "rota_granja.db")


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


def resolve_database_path(
    *,
    explicit_db_path: str,
    runtime_cfg: Dict[str, Any],
    app_kind: str,
    data_root: str,
    app_env: str,
    tenant_id: str,
) -> str:
    candidate = str(explicit_db_path or runtime_cfg.get("db_path") or "").strip()
    if not candidate:
        candidate = _build_default_db_path(
            app_kind=app_kind,
            data_root=data_root,
            app_env=app_env,
            tenant_id=tenant_id,
        )
    candidate = os.path.abspath(candidate)
    lower = candidate.lower()
    if app_env == "development" and ("production" in lower or "\\dist\\" in lower):
        raise RuntimeError(f"DB_PATH inseguro para development: {candidate}")
    if app_env == "staging" and "\\development\\" in lower:
        raise RuntimeError(f"DB_PATH inseguro para staging: {candidate}")
    return candidate


def load_app_config(app_kind: str = "desktop") -> AppConfig:
    is_frozen = getattr(sys, "frozen", False)
    local_desktop_mode = app_kind == "desktop" and not is_frozen
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

    app_env = _normalize_env(
        ("development" if local_desktop_mode else (os.environ.get("ROTA_APP_ENV") or file_config.get("app_env"))),
        is_frozen=is_frozen,
    )
    env_file = os.path.join(config_root, "environments", f"{app_env}.json")
    env_defaults = _read_json(env_file)
    merged = _merge_dicts(env_defaults, file_config)

    runtime_cfg = merged.get("runtime") if isinstance(merged.get("runtime"), dict) else {}
    api_cfg = merged.get("api") if isinstance(merged.get("api"), dict) else {}
    update_cfg = merged.get("update") if isinstance(merged.get("update"), dict) else {}
    support_cfg = merged.get("support") if isinstance(merged.get("support"), dict) else {}
    tenant_cfg = merged.get("tenant") if isinstance(merged.get("tenant"), dict) else {}
    logging_cfg = merged.get("logging") if isinstance(merged.get("logging"), dict) else {}

    company_default = "dev-local" if local_desktop_mode or app_env == "development" else "default-company"
    company_id = str(
        (None if local_desktop_mode else os.environ.get("ROTA_COMPANY_ID"))
        or (None if local_desktop_mode else os.environ.get("ROTA_TENANT_ID"))
        or tenant_cfg.get("company_id")
        or tenant_cfg.get("tenant_id")
        or company_default
    ).strip()
    tenant_id = str(
        (None if local_desktop_mode else os.environ.get("ROTA_TENANT_ID"))
        or tenant_cfg.get("tenant_id")
        or company_id
    ).strip()
    tenant_id = _slugify(tenant_id, "default-company")
    company_id = _slugify(company_id, tenant_id)

    configured_data_root = str((None if local_desktop_mode else os.environ.get("ROTA_DATA_ROOT")) or runtime_cfg.get("data_root") or data_root).strip()
    data_root = configured_data_root or data_root

    db_path = resolve_database_path(
        explicit_db_path=(None if local_desktop_mode else os.environ.get("ROTA_DB")) or "",
        runtime_cfg={} if local_desktop_mode else runtime_cfg,
        app_kind=app_kind,
        data_root=data_root,
        app_env=app_env,
        tenant_id=tenant_id,
    )

    if app_env == "development":
        default_api_url = "http://127.0.0.1:8000"
        default_sync = False
        default_sql_mirror = False
        default_binding = False
        default_update_channel = "disabled"
        default_allow_remote_write = False
        default_allow_version_update = False
        default_allow_seed_db = False
        default_allow_remote_read = False
        default_source_of_truth = "sqlite-local"
    else:
        default_api_url = "https://rotahub-api.onrender.com"
        default_sync = True
        default_sql_mirror = True
        default_binding = True
        default_update_channel = "staging" if app_env == "staging" else "stable"
        default_allow_remote_write = True
        default_allow_version_update = True
        default_allow_seed_db = False
        default_allow_remote_read = False
        default_source_of_truth = "sqlite-local"

    api_base_url = str((None if local_desktop_mode else os.environ.get("ROTA_SERVER_URL")) or ("" if local_desktop_mode else api_cfg.get("base_url")) or default_api_url).strip().rstrip("/")
    try:
        api_sync_timeout = float((None if local_desktop_mode else os.environ.get("ROTA_SYNC_TIMEOUT")) or api_cfg.get("timeout") or 60)
    except Exception:
        api_sync_timeout = 60.0

    sync_enabled_raw = None if local_desktop_mode else os.environ.get("ROTA_DESKTOP_SYNC_API")
    if local_desktop_mode:
        sync_enabled = False
    elif sync_enabled_raw is None:
        sync_enabled = bool(runtime_cfg.get("sync_enabled", default_sync))
    else:
        sync_enabled = str(sync_enabled_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    sql_mirror_raw = None if local_desktop_mode else os.environ.get("ROTA_SQL_MIRROR_API")
    if local_desktop_mode:
        sql_mirror_api = False
    elif sql_mirror_raw is None:
        sql_mirror_api = bool(runtime_cfg.get("sql_mirror_api", default_sql_mirror))
    else:
        sql_mirror_api = str(sql_mirror_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    binding_raw = None if local_desktop_mode else os.environ.get("ROTA_REQUIRE_SERVER_BINDING")
    if local_desktop_mode:
        require_server_binding = False
    elif binding_raw is None:
        require_server_binding = bool(runtime_cfg.get("require_server_binding", default_binding))
    else:
        require_server_binding = str(binding_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    desktop_secret = str((None if local_desktop_mode else os.environ.get("ROTA_SECRET")) or runtime_cfg.get("desktop_secret") or "").strip()
    allow_remote_write_raw = None if local_desktop_mode else os.environ.get("ROTA_ALLOW_REMOTE_WRITE")
    if local_desktop_mode:
        allow_remote_write = False
    elif allow_remote_write_raw is None:
        allow_remote_write = bool(runtime_cfg.get("allow_remote_write", default_allow_remote_write))
    else:
        allow_remote_write = str(allow_remote_write_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_remote_read_raw = None if local_desktop_mode else os.environ.get("ROTA_ALLOW_REMOTE_READ")
    if local_desktop_mode:
        allow_remote_read = False
    elif allow_remote_read_raw is None:
        allow_remote_read = bool(runtime_cfg.get("allow_remote_read", default_allow_remote_read))
    else:
        allow_remote_read = str(allow_remote_read_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_seed_db_raw = None if local_desktop_mode else os.environ.get("ROTA_ALLOW_SEED_DB")
    if local_desktop_mode:
        allow_seed_db = False
    elif allow_seed_db_raw is None:
        allow_seed_db = bool(runtime_cfg.get("allow_seed_db", default_allow_seed_db))
    else:
        allow_seed_db = str(allow_seed_db_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    allow_version_update_raw = None if local_desktop_mode else os.environ.get("ROTA_ALLOW_VERSION_UPDATE")
    if local_desktop_mode:
        allow_version_update = False
    elif allow_version_update_raw is None:
        allow_version_update = bool(runtime_cfg.get("allow_version_update", default_allow_version_update))
    else:
        allow_version_update = str(allow_version_update_raw).strip().lower() in {"1", "true", "yes", "y", "sim", "on"}

    update_channel = str(
        (None if local_desktop_mode else os.environ.get("ROTA_UPDATE_CHANNEL"))
        or update_cfg.get("channel")
        or default_update_channel
    ).strip() or default_update_channel
    tenant_mode = str(runtime_cfg.get("tenant_mode") or "database-per-tenant").strip() or "database-per-tenant"
    log_level = str(logging_cfg.get("level") or os.environ.get("ROTA_LOG_LEVEL") or ("DEBUG" if app_env == "development" else "INFO")).strip().upper()
    source_of_truth = str(runtime_cfg.get("source_of_truth") or default_source_of_truth).strip() or "sqlite-local"
    update_manifest_url = str(
        os.environ.get("ROTA_UPDATE_MANIFEST_URL")
        or update_cfg.get("manifest_url")
        or "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/version.json"
    ).strip()
    setup_download_url = str(os.environ.get("ROTA_SETUP_URL") or update_cfg.get("setup_url") or "").strip()
    changelog_url = str(os.environ.get("ROTA_CHANGELOG_URL") or update_cfg.get("changelog_url") or "").strip()
    support_whatsapp = str(os.environ.get("ROTA_SUPPORT_WHATSAPP") or support_cfg.get("whatsapp") or "").strip()
    support_email = str(os.environ.get("ROTA_SUPPORT_EMAIL") or support_cfg.get("email") or "").strip()
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

    return AppConfig(
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
                "allow_remote_read": config.allow_remote_read,
                "allow_seed_db": config.allow_seed_db,
                "allow_version_update": config.allow_version_update,
                "tenant_mode": config.tenant_mode,
                "source_of_truth": config.source_of_truth,
                "schema_version": config.schema_version,
                "desktop_secret": "",
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
