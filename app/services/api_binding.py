import logging
import os
import time
from tkinter import messagebox

from app.services.api_client import _call_api
from app.services.runtime_flags import is_desktop_api_sync_enabled
from runtime_config import load_app_config


_APP_CONFIG = load_app_config("desktop")
_APP_ENV = _APP_CONFIG.app_env
_API_BASE_URL = _APP_CONFIG.api_base_url
_API_BINDING_CACHE = {"ok": False, "checked_at": 0.0, "error": ""}


def ensure_system_api_binding(context: str = "Operacao", parent=None, force_probe: bool = False) -> bool:
    """Garante vinculo obrigatorio Desktop <-> API para operacoes criticas."""
    if not is_desktop_api_sync_enabled():
        if os.environ.get("ROTA_REQUIRE_SERVER_BINDING", "0").strip().lower() not in {"1", "true", "yes", "y", "sim", "on"}:
            logging.info("Operacao liberada sem integracao obrigatoria | env=%s | contexto=%s", _APP_ENV, context)
            return True
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "A integracao Desktop<->Servidor esta desativada para este ambiente.\n\n"
            f"Operacao bloqueada: {context}\n"
            "Ative a sincronizacao externa apenas em staging/producao.",
            parent=parent,
        )
        return False

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "ROTA_SECRET nao configurada.\n\n"
            f"Operacao bloqueada: {context}\n"
            "Defina a chave da estacao para manter o fluxo unico Desktop/Servidor/Dispositivo.",
            parent=parent,
        )
        return False

    now = time.time()
    if (not force_probe) and _API_BINDING_CACHE.get("ok") and (now - float(_API_BINDING_CACHE.get("checked_at") or 0.0) <= 20.0):
        return True

    try:
        _call_api(
            "GET",
            "admin/motoristas/acesso",
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )
        _API_BINDING_CACHE["ok"] = True
        _API_BINDING_CACHE["checked_at"] = now
        _API_BINDING_CACHE["error"] = ""
        return True
    except Exception as exc:
        _API_BINDING_CACHE["ok"] = False
        _API_BINDING_CACHE["checked_at"] = now
        _API_BINDING_CACHE["error"] = str(exc or "")
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "Nao foi possivel validar conexao com a API central.\n\n"
            f"Operacao bloqueada: {context}\n"
            f"URL base: {_API_BASE_URL}\n"
            f"Detalhe: {str(exc or 'Falha de conectividade')}",
            parent=parent,
        )
        return False