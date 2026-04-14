import json
import logging
import time
import urllib.error
import urllib.request


class SyncError(Exception):
    """Erro específico ao tentar sincronizar com a API mobile."""


API_BASE_URL = ""
API_SYNC_TIMEOUT = 0
APP_ENV = ""
TENANT_ID = ""
COMPANY_ID = ""

API_GET_CACHE = {}
API_CACHE_TTLS = {
    "desktop/cadastros/": 30.0,
    "desktop/clientes/base": 20.0,
    "desktop/programacoes": 8.0,
    "desktop/monitoramento/rotas": 5.0,
    "desktop/overview": 5.0,
}


def configure_api_client(
    *,
    api_base_url,
    api_sync_timeout,
    app_env,
    tenant_id,
    company_id,
):
    global API_BASE_URL
    global API_SYNC_TIMEOUT
    global APP_ENV
    global TENANT_ID
    global COMPANY_ID

    previous_config = (
        API_BASE_URL,
        API_SYNC_TIMEOUT,
        APP_ENV,
        TENANT_ID,
        COMPANY_ID,
    )
    API_BASE_URL = api_base_url
    API_SYNC_TIMEOUT = api_sync_timeout
    APP_ENV = app_env
    TENANT_ID = tenant_id
    COMPANY_ID = company_id
    if previous_config != (
        API_BASE_URL,
        API_SYNC_TIMEOUT,
        APP_ENV,
        TENANT_ID,
        COMPANY_ID,
    ):
        _invalidate_api_cache()


def _api_cache_ttl(path: str) -> float:
    normalized = str(path or "").strip().lstrip("/")
    for prefix, ttl in API_CACHE_TTLS.items():
        if normalized.startswith(prefix):
            return ttl
    return 0.0


def _api_cache_key(method: str, path: str, token: str = None, extra_headers: dict = None):
    headers = tuple(sorted((str(k), str(v)) for k, v in (extra_headers or {}).items() if k and v is not None))
    return (str(method or "").upper(), str(path or "").strip(), str(token or ""), headers)


def _invalidate_api_cache(path: str = ""):
    normalized = str(path or "").strip().lstrip("/")
    keys = list(API_GET_CACHE.keys())
    for key in keys:
        try:
            _, cached_path, _, _ = key
        except Exception:
            cached_path = ""
        if (not normalized) or str(cached_path).startswith(normalized):
            API_GET_CACHE.pop(key, None)


def _build_api_url(path: str) -> str:
    path = (path or "").strip().lstrip("/")
    if path:
        return f"{API_BASE_URL}/{path}"
    return API_BASE_URL


def _call_api(method: str, path: str, payload=None, token: str = None, extra_headers: dict = None):
    method = str(method or "GET").upper()
    url = _build_api_url(path)
    headers = {
        "Accept": "application/json",
        "X-App-Env": APP_ENV,
        "X-Tenant-ID": TENANT_ID,
        "X-Company-ID": COMPANY_ID,
    }
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if extra_headers:
        for k, v in (extra_headers or {}).items():
            if k and v is not None:
                headers[str(k)] = str(v)

    cache_ttl = _api_cache_ttl(path) if method == "GET" and payload is None else 0.0
    cache_key = _api_cache_key(method, path, token=token, extra_headers=extra_headers) if cache_ttl > 0 else None
    if cache_key:
        cached = API_GET_CACHE.get(cache_key)
        if cached:
            expires_at, cached_payload = cached
            if expires_at > time.time():
                logging.info("API cache hit | path=%s", path)
                return cached_payload
            API_GET_CACHE.pop(cache_key, None)

    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    started = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=API_SYNC_TIMEOUT) as resp:
            body = resp.read()
            elapsed_ms = (time.perf_counter() - started) * 1000.0
            logging.info("API %s %s | %.0f ms", method, path, elapsed_ms)
            if not body:
                return {}
            text = body.decode("utf-8")
            parsed = json.loads(text)
            if cache_key:
                API_GET_CACHE[cache_key] = (time.time() + cache_ttl, parsed)
            elif method != "GET":
                _invalidate_api_cache()
            return parsed
    except urllib.error.HTTPError as exc:
        elapsed_ms = (time.perf_counter() - started) * 1000.0
        logging.warning("API %s %s | HTTPError %s | %.0f ms", method, path, exc.code, elapsed_ms)
        body = exc.read()
        detail = ""
        if body:
            try:
                payload = json.loads(body.decode("utf-8"))
                if isinstance(payload, dict):
                    detail = payload.get("detail") or payload.get("message") or str(payload)
                else:
                    detail = str(payload)
            except Exception:
                detail = body.decode("utf-8", errors="ignore")
        raise SyncError(f"{exc.code} {exc.reason}: {detail or 'Sem detalhes'}")
    except urllib.error.URLError as exc:
        elapsed_ms = (time.perf_counter() - started) * 1000.0
        logging.warning("API %s %s | URLError | %.0f ms", method, path, elapsed_ms)
        raise SyncError(f"Falha ao conectar-se a {url}: {exc.reason}")
    except Exception as exc:
        elapsed_ms = (time.perf_counter() - started) * 1000.0
        logging.warning("API %s %s | erro inesperado | %.0f ms", method, path, elapsed_ms)
        raise SyncError(f"Erro inesperado ao chamar a API de sincronização: {exc}")
    return None


def _friendly_sync_error(exc: Exception, default_message: str = "Falha na operação com a API central.") -> str:
    msg = str(exc or "").strip()
    if not msg:
        return default_message
    prefixes = (
        "Erro inesperado ao chamar a API de sincronização:",
        "Falha ao conectar-se a",
    )
    for prefix in prefixes:
        if msg.startswith(prefix):
            msg = msg[len(prefix):].strip(" :")
    return msg or default_message


__all__ = [
    "API_CACHE_TTLS",
    "API_GET_CACHE",
    "SyncError",
    "_api_cache_key",
    "_api_cache_ttl",
    "_build_api_url",
    "_call_api",
    "_friendly_sync_error",
    "_invalidate_api_cache",
    "configure_api_client",
]
