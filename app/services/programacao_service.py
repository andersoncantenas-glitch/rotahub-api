# -*- coding: utf-8 -*-
import logging

from app.repositories.programacao_repository import fetch_programacao_itens_local


def _service_result(*, ok: bool, data=None, error: str = None, source: str = "local"):
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def fetch_programacao_itens(
    *,
    codigo_programacao: str,
    limit: int = 8000,
    upper,
    safe_int,
    safe_float,
    get_db,
    call_api,
    quote,
    is_desktop_api_sync_enabled,
    desktop_secret: str,
):
    """Regra de negócio para obter itens da programação com fallback API->DB."""
    codigo_programacao = upper(str(codigo_programacao or "").strip())
    if not codigo_programacao:
        return _service_result(ok=True, data=[], source="local")

    if desktop_secret and is_desktop_api_sync_enabled():
        try:
            resp = call_api(
                "GET",
                f"desktop/rotas/{quote(codigo_programacao)}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            clientes = resp.get("clientes") if isinstance(resp, dict) else []
            out_api = []
            for d in (clientes or [])[: max(int(limit or 0), 1)]:
                if not isinstance(d, dict):
                    continue
                out_api.append(
                    {
                        "cod_cliente": str(d.get("cod_cliente") or "").strip().upper(),
                        "nome_cliente": str(d.get("nome_cliente") or "").strip().upper(),
                        "endereco": str(d.get("endereco") or "").strip().upper(),
                        "produto": str(d.get("produto") or "").strip().upper(),
                        "qnt_caixas": safe_int(d.get("qnt_caixas"), 0),
                        "kg": safe_float(d.get("kg"), 0.0),
                        "preco": safe_float(d.get("preco"), 0.0),
                        "vendedor": str(d.get("vendedor") or "").strip().upper(),
                        "pedido": str(d.get("pedido") or "").strip().upper(),
                        "obs": str(d.get("obs") or d.get("observacao") or "").strip(),
                        "status_pedido": str(d.get("status_pedido") or "PENDENTE").strip().upper(),
                        "caixas_atual": safe_int(d.get("caixas_atual"), 0),
                        "preco_atual": safe_float(d.get("preco_atual"), 0.0),
                        "alterado_em": str(d.get("alterado_em") or "").strip(),
                        "alterado_por": str(d.get("alterado_por") or "").strip().upper(),
                        "mortalidade_aves": safe_int(d.get("mortalidade_aves"), 0),
                        "peso_previsto": safe_float(d.get("peso_previsto"), 0.0),
                        "valor_recebido": safe_float(d.get("valor_recebido"), 0.0),
                        "forma_recebimento": str(d.get("forma_recebimento") or "").strip().upper(),
                        "obs_recebimento": str(d.get("obs_recebimento") or "").strip(),
                        "alteracao_tipo": str(d.get("alteracao_tipo") or "").strip().upper(),
                        "alteracao_detalhe": str(d.get("alteracao_detalhe") or "").strip(),
                    }
                )
            if out_api:
                return _service_result(ok=True, data=out_api, source="api")
        except Exception:
            logging.debug("Falha ao buscar itens da programacao na API", exc_info=True)

    try:
        out_local = fetch_programacao_itens_local(
            codigo_programacao=codigo_programacao,
            limit=limit,
            get_db=get_db,
            safe_int=safe_int,
            safe_float=safe_float,
        )
        source = "both" if (desktop_secret and is_desktop_api_sync_enabled()) else "local"
        return _service_result(ok=True, data=(out_local or []), source=source)
    except Exception:
        return _service_result(ok=False, data=[], error="Falha ao obter itens da programação.", source="local")
