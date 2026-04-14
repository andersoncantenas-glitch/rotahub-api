# -*- coding: utf-8 -*-
import logging
import os

from app.repositories.vendedor_repository import update_vendedor_senha_hash_local
from app.services.api_client import _call_api


def _service_result(*, ok: bool, data=None, error: str = None, source: str = "local"):
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def _error_message(exc: Exception, default_message: str) -> str:
    msg = str(exc or "").strip()
    return msg or str(default_message or "Falha inesperada.")


def fetch_vendedores_rows(*, fields, can_read_from_api):
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and can_read_from_api():
        try:
            api_rows = _call_api(
                "GET",
                "desktop/cadastros/vendedores",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            if isinstance(api_rows, list):
                rows = []
                for r in api_rows:
                    if not isinstance(r, dict):
                        continue
                    row_map = {
                        "codigo": str(r.get("codigo") or ""),
                        "nome": str(r.get("nome") or ""),
                        "senha": "",
                        "telefone": str(r.get("telefone") or ""),
                        "cidade_base": str(r.get("cidade_base") or ""),
                        "status": str(r.get("status") or "ATIVO"),
                    }
                    row = [int(r.get("id") or 0)]
                    for c, _ in fields:
                        row.append(row_map.get(c, ""))
                    rows.append(tuple(row))
                return _service_result(ok=True, data=rows, source="api")
            return _service_result(
                ok=False,
                data=None,
                error="Resposta de vendedores fora do formato esperado (lista).",
                source="api",
            )
        except Exception as exc:
            logging.debug("Falha ao carregar vendedores via API; usando fallback local.", exc_info=True)
            return _service_result(
                ok=False,
                data=None,
                error=_error_message(exc, "Falha ao carregar vendedores na API."),
                source="api",
            )
    return _service_result(ok=True, data=None, source="local")


def sync_vendedor_upsert_api(data: dict, *, is_desktop_api_sync_enabled, norm, normalize_phone):
    if not is_desktop_api_sync_enabled():
        return _service_result(ok=True, data=None, source="local")

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        return _service_result(ok=True, data=None, source="local")

    codigo = norm(data.get("codigo"))
    nome = norm(data.get("nome"))
    if not codigo or not nome:
        return _service_result(ok=True, data=None, source="local")

    payload = {
        "codigo": codigo,
        "nome": nome,
        "telefone": normalize_phone(data.get("telefone")),
        "cidade_base": norm(data.get("cidade_base")),
        "status": norm(data.get("status") or "ATIVO"),
    }
    senha = data.get("senha")
    if senha:
        payload["senha"] = senha

    try:
        resp = _call_api(
            "POST",
            "desktop/cadastros/vendedores/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )
        return _service_result(ok=True, data=resp, source="api")
    except Exception as exc:
        return _service_result(
            ok=False,
            data=None,
            error=_error_message(exc, "Falha ao sincronizar vendedor na API."),
            source="api",
        )


def update_vendedor_password_hash_local(vendedor_id: int, senha_hash: str):
    try:
        update_vendedor_senha_hash_local(vendedor_id, senha_hash)
        return _service_result(ok=True, data=None, source="local")
    except Exception as exc:
        return _service_result(
            ok=False,
            data=None,
            error=_error_message(exc, "Falha ao atualizar senha local do vendedor."),
            source="local",
        )
