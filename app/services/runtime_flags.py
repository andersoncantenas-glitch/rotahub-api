import os

from runtime_config import load_app_config


_APP_CONFIG = load_app_config("desktop")
_ENABLE_API_SYNC = bool(_APP_CONFIG.sync_enabled)
_ALLOW_REMOTE_READ = bool(_APP_CONFIG.allow_remote_read)
_SOURCE_OF_TRUTH = _APP_CONFIG.source_of_truth


def is_desktop_api_sync_enabled() -> bool:
    """Controla sincronizacao automatica Desktop <-> API central.
    O runtime define o valor padrao por ambiente; a variavel de ambiente
    serve apenas como override explicito do processo atual.
    """
    raw = str(os.environ.get("ROTA_DESKTOP_SYNC_API", "1" if _ENABLE_API_SYNC else "0") or "").strip().lower()
    return raw in {"1", "true", "yes", "y", "sim", "on"}


def can_read_from_api() -> bool:
    """Permite leitura remota apenas quando o runtime libera a origem central."""
    if not is_desktop_api_sync_enabled():
        return False
    raw = str(os.environ.get("ROTA_ALLOW_REMOTE_READ", "1" if _ALLOW_REMOTE_READ else "0") or "").strip().lower()
    if raw not in {"1", "true", "yes", "y", "sim", "on"}:
        return False
    source = str(os.environ.get("ROTA_SOURCE_OF_TRUTH", _SOURCE_OF_TRUTH or "") or "").strip().lower()
    return source in {"api-central", "server", "remote", "hybrid", "api"}