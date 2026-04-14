# -*- coding: utf-8 -*-
import logging
import os
import re

from app.repositories.motorista_repository import (
    fetch_motorista_access_snapshot_by_codigo,
    fetch_motorista_codigos_local,
    fetch_motoristas_cache_local_by_codigo,
)
from app.services.api_client import _call_api


_MOTORISTA_SEQ_RE = re.compile(r"^MOT-(\d+)$")


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


def extract_motorista_seq(codigo: str) -> int:
    s = str(codigo or "").strip().upper()
    m = _MOTORISTA_SEQ_RE.match(s)
    if not m:
        return 0
    try:
        return int(m.group(1))
    except Exception:
        return 0


def next_motorista_codigo(*, can_read_from_api, cur=None):
    max_seq = 0
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    source = "local"

    if desktop_secret and can_read_from_api():
        try:
            rows = _call_api(
                "GET",
                "desktop/cadastros/motoristas",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            if isinstance(rows, list):
                source = "both"
                for r in rows:
                    if not isinstance(r, dict):
                        continue
                    max_seq = max(max_seq, extract_motorista_seq(r.get("codigo")))
        except Exception:
            logging.debug("Falha ao calcular proximo codigo de motorista via API; usando local.", exc_info=True)

    try:
        for codigo in fetch_motorista_codigos_local(cur=cur):
            max_seq = max(max_seq, extract_motorista_seq(codigo))
    except Exception:
        logging.debug("Falha ao calcular proximo codigo de motorista local.", exc_info=True)

    return _service_result(ok=True, data=f"MOT-{max(1, max_seq + 1):02d}", source=source)


def fetch_motoristas_rows(*, fields, can_read_from_api):
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and can_read_from_api():
        try:
            api_rows = _call_api(
                "GET",
                "desktop/cadastros/motoristas",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            if isinstance(api_rows, list):
                motorista_cache = fetch_motoristas_cache_local_by_codigo()
                rows = []
                for r in api_rows:
                    if not isinstance(r, dict):
                        continue
                    cod = str(r.get("codigo") or "").strip()
                    local_row = motorista_cache.get(cod, {})
                    senha_val = str(r.get("senha") or r.get("senha_hash") or local_row.get("senha") or "")
                    cpf_val = str(r.get("cpf") or local_row.get("cpf") or "")
                    tel_val = str(r.get("telefone") or local_row.get("telefone") or "")
                    row_map = {
                        "nome": str(r.get("nome") or ""),
                        "codigo": cod,
                        "senha": senha_val,
                        "cpf": cpf_val,
                        "telefone": tel_val,
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
                error="Resposta de motoristas fora do formato esperado (lista).",
                source="api",
            )
        except Exception as exc:
            logging.debug("Falha ao carregar motoristas via API; usando fallback local.", exc_info=True)
            return _service_result(
                ok=False,
                data=None,
                error=_error_message(exc, "Falha ao carregar motoristas na API."),
                source="api",
            )
    return _service_result(ok=True, data=None, source="local")


def sync_motorista_upsert_api(
    data: dict,
    *,
    is_desktop_api_sync_enabled,
    norm,
):
    if not is_desktop_api_sync_enabled():
        return _service_result(ok=True, data=None, source="local")

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        return _service_result(ok=True, data=None, source="local")

    codigo = norm(data.get("codigo"))
    nome = norm(data.get("nome"))
    if not codigo or not nome:
        return _service_result(ok=True, data=None, source="local")

    acesso_liberado_raw = data.get("acesso_liberado", None)
    acesso_liberado_por = data.get("acesso_liberado_por", None)
    acesso_obs = data.get("acesso_obs", None)

    if acesso_liberado_raw in (None, ""):
        try:
            snapshot = fetch_motorista_access_snapshot_by_codigo(codigo)
            if snapshot:
                acesso_liberado_raw = snapshot.get("acesso_liberado")
                if not acesso_liberado_por:
                    acesso_liberado_por = snapshot.get("acesso_liberado_por")
                if not acesso_obs:
                    acesso_obs = snapshot.get("acesso_obs")
        except Exception:
            logging.debug("Falha ao preservar acesso_liberado do motorista para sync", exc_info=True)

    if acesso_liberado_raw in (None, ""):
        acesso_liberado_payload = None
    else:
        try:
            acesso_liberado_payload = bool(int(acesso_liberado_raw or 0))
        except Exception:
            acesso_liberado_payload = bool(acesso_liberado_raw)

    payload = {
        "codigo": codigo,
        "nome": nome,
        "telefone": norm(data.get("telefone")),
        "cpf": norm(data.get("cpf")),
        "status": norm(data.get("status") or "ATIVO"),
        "senha": data.get("senha") or None,
        "acesso_liberado": acesso_liberado_payload,
        "acesso_liberado_por": norm(acesso_liberado_por or "DESKTOP_SYNC"),
        "acesso_obs": norm(acesso_obs or "Sincronizado via Desktop"),
    }
    try:
        resp = _call_api(
            "POST",
            "desktop/cadastros/motoristas/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )
        return _service_result(ok=True, data=resp, source="api")
    except Exception as exc:
        return _service_result(
            ok=False,
            data=None,
            error=_error_message(exc, "Falha ao sincronizar motorista na API."),
            source="api",
        )
