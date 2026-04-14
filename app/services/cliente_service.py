# -*- coding: utf-8 -*-
import logging
import os

try:
    import pandas as pd
except Exception:
    pd = None

from app.repositories.cliente_repository import fetch_clientes_rows_local, upsert_clientes_local
from app.services.api_client import _call_api
from app.utils.excel_helpers import excel_engine_for, guess_col, upper


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


def _api_mode(is_desktop_api_sync_enabled):
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    return bool(desktop_secret and is_desktop_api_sync_enabled()), desktop_secret


def sync_cliente_upsert_api(cod: str, nome: str, endereco: str, telefone: str, vendedor: str, *, is_desktop_api_sync_enabled):
    if not is_desktop_api_sync_enabled():
        return _service_result(ok=True, data=None, source="local")
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        return _service_result(ok=True, data=None, source="local")

    cod_n = upper(str(cod or "").strip())
    nome_n = upper(str(nome or "").strip())
    if not cod_n or not nome_n:
        return _service_result(ok=True, data=None, source="local")

    payload = {
        "cod_cliente": cod_n,
        "nome_cliente": nome_n,
        "endereco": upper(str(endereco or "").strip()),
        "telefone": upper(str(telefone or "").strip()),
        "vendedor": upper(str(vendedor or "").strip()),
    }
    try:
        resp = _call_api(
            "POST",
            "desktop/cadastros/clientes/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )
        return _service_result(ok=True, data=resp, source="api")
    except Exception as exc:
        return _service_result(
            ok=False,
            data=None,
            error=_error_message(exc, "Falha ao sincronizar cliente na API."),
            source="api",
        )


def fetch_clientes_rows(*, is_desktop_api_sync_enabled):
    rows = []
    api_mode, desktop_secret = _api_mode(is_desktop_api_sync_enabled)
    source = "local"
    if api_mode:
        try:
            resp = _call_api(
                "GET",
                "desktop/clientes/base?ordem=nome&limit=1000",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            if isinstance(resp, list):
                rows = [
                    (
                        str(r.get("cod_cliente") or ""),
                        str(r.get("nome_cliente") or ""),
                        str(r.get("endereco") or ""),
                        str(r.get("telefone") or ""),
                        str(r.get("vendedor") or ""),
                    )
                    for r in resp
                    if isinstance(r, dict)
                ]
                if rows:
                    source = "api"
        except Exception:
            logging.debug("Falha ao carregar clientes via API; usando fallback local.", exc_info=True)

    if not rows:
        rows = fetch_clientes_rows_local(limit=5000)
        source = "local"
    return _service_result(ok=True, data=rows, source=source)


def _build_fail_detail(sync_falhas: int, falhas_refs):
    detalhe = ", ".join(falhas_refs)
    if sync_falhas > len(falhas_refs):
        detalhe += f" e mais {sync_falhas - len(falhas_refs)}"
    return detalhe


def salvar_clientes_linhas(linhas, *, is_desktop_api_sync_enabled):
    total = 0
    sync_falhas = 0
    falhas_refs = []
    api_mode, _desktop_secret = _api_mode(is_desktop_api_sync_enabled)

    if api_mode:
        for cod, nome, endereco, telefone, vendedor in (linhas or []):
            result = sync_cliente_upsert_api(
                cod, nome, endereco, telefone, vendedor,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            )
            if isinstance(result, dict) and bool(result.get("ok", False)):
                total += 1
            else:
                sync_falhas += 1
                if len(falhas_refs) < 10:
                    falhas_refs.append(str(cod or nome or "?"))
        if sync_falhas:
            detalhe = _build_fail_detail(sync_falhas, falhas_refs)
            return _service_result(
                ok=False,
                data=None,
                error=(
                    "Falha ao salvar clientes na API central. "
                    "Nenhuma confirmação local foi aplicada.\n\n"
                    f"Clientes com falha: {detalhe}"
                ),
                source="api",
            )
    else:
        upsert_clientes_local(linhas or [])
        total = len(linhas or [])
        for cod, nome, endereco, telefone, vendedor in (linhas or []):
            result = sync_cliente_upsert_api(
                cod, nome, endereco, telefone, vendedor,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            )
            if isinstance(result, dict) and not bool(result.get("ok", False)):
                sync_falhas += 1

    msg = f"Clientes salvos/atualizados: {total}"
    if sync_falhas:
        msg += f"\nFalhas de sincronização API: {sync_falhas}"
    return _service_result(ok=True, data=msg, source=("api" if api_mode else "both"))


def importar_clientes_excel(path: str, *, is_desktop_api_sync_enabled):
    if pd is None:
        return _service_result(ok=False, data=None, error="Pandas indisponível para importação de clientes.", source="local")

    df = pd.read_excel(path, engine=excel_engine_for(path))

    col_cod = guess_col(df.columns, ["cod", "cód", "codigo", "cliente", "cod cliente"])
    col_nome = guess_col(df.columns, ["nome", "cliente"])
    col_end = guess_col(df.columns, ["endereco", "endereço", "rua", "logradouro"])
    col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
    col_vendedor = guess_col(df.columns, ["vendedor", "vend", "representante"])

    if not col_cod or not col_nome:
        cols = list(df.columns or [])
        if len(cols) >= 2:
            col_cod = col_cod or cols[0]
            col_nome = col_nome or cols[1]
        else:
            return _service_result(
                ok=False,
                data=None,
                error="Nao identifiquei as colunas de codigo e nome do cliente no Excel.",
                source="local",
            )

    total = 0
    sync_falhas = 0
    falhas_refs = []
    api_mode, _desktop_secret = _api_mode(is_desktop_api_sync_enabled)

    linhas_local = []
    for _, r in df.iterrows():
        cod = str(r.get(col_cod, "")).strip()
        nome = str(r.get(col_nome, "")).strip()
        if not cod or not nome:
            continue

        endereco = str(r.get(col_end, "")).strip() if col_end else ""
        telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
        vendedor = str(r.get(col_vendedor, "")).strip() if col_vendedor else ""

        if api_mode:
            result = sync_cliente_upsert_api(
                cod, nome, endereco, telefone, vendedor,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            )
            if isinstance(result, dict) and bool(result.get("ok", False)):
                total += 1
            else:
                sync_falhas += 1
                if len(falhas_refs) < 10:
                    falhas_refs.append(str(cod or nome or "?"))
        else:
            linhas_local.append((upper(cod), upper(nome), upper(endereco), upper(telefone), upper(vendedor)))

    if api_mode and sync_falhas:
        detalhe = _build_fail_detail(sync_falhas, falhas_refs)
        return _service_result(
            ok=False,
            data=None,
            error=(
                "Falha ao importar clientes na API central. "
                "Nenhuma confirmação local foi aplicada.\n\n"
                f"Clientes com falha: {detalhe}"
            ),
            source="api",
        )

    if not api_mode:
        upsert_clientes_local(linhas_local)
        total = len(linhas_local)
        for cod, nome, endereco, telefone, vendedor in linhas_local:
            result = sync_cliente_upsert_api(
                cod, nome, endereco, telefone, vendedor,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            )
            if isinstance(result, dict) and not bool(result.get("ok", False)):
                sync_falhas += 1

    msg = f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}"
    if sync_falhas:
        msg += f"\nFalhas de sincronização API: {sync_falhas}"
    return _service_result(ok=True, data=msg, source=("api" if api_mode else "both"))
