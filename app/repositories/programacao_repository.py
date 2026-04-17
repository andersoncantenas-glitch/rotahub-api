# -*- coding: utf-8 -*-
import re


_SAFE_SQL_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _safe_ident(name: str) -> str:
    """Valida identificador SQL (tabela/coluna) para uso em PRAGMA."""
    name = str(name or "").strip()
    if not _SAFE_SQL_IDENT.match(name):
        raise ValueError(f"Identificador SQL invalido: {name!r}")
    return name


def db_has_column(cur, table_name: str, column_name: str) -> bool:
    """Verifica se uma coluna existe numa tabela SQLite."""
    try:
        table_name = _safe_ident(table_name)
        column_name = str(column_name or "").strip().lower()
        cur.execute(f"PRAGMA table_info({table_name})")
        cols = [str(r[1]).lower() for r in cur.fetchall()]
        return column_name in cols
    except Exception:
        return False


def fetch_programacao_itens_local(
    *,
    codigo_programacao: str,
    limit: int,
    get_db,
    safe_int,
    safe_float,
    db_has_column_fn=None,
):
    """Busca itens de programacao no banco local (schema-flexível)."""
    db_has_column_impl = db_has_column_fn or db_has_column

    with get_db() as conn:
        cur = conn.cursor()

        has_obs = db_has_column_impl(cur, "programacao_itens", "obs") or db_has_column_impl(cur, "programacao_itens", "observacao")
        has_vendedor = db_has_column_impl(cur, "programacao_itens", "vendedor")
        has_pedido = db_has_column_impl(cur, "programacao_itens", "pedido")
        has_end = db_has_column_impl(cur, "programacao_itens", "endereco") or db_has_column_impl(cur, "programacao_itens", "endereco")
        has_produto = db_has_column_impl(cur, "programacao_itens", "produto")

        has_status = db_has_column_impl(cur, "programacao_itens", "status_pedido")
        has_caixas_atual = db_has_column_impl(cur, "programacao_itens", "caixas_atual")
        has_preco_atual = db_has_column_impl(cur, "programacao_itens", "preco_atual")
        has_alt_em = db_has_column_impl(cur, "programacao_itens", "alterado_em")
        has_alt_por = db_has_column_impl(cur, "programacao_itens", "alterado_por")
        has_ordem_sugerida = db_has_column_impl(cur, "programacao_itens", "ordem_sugerida")
        has_eta = db_has_column_impl(cur, "programacao_itens", "eta")
        has_distancia = db_has_column_impl(cur, "programacao_itens", "distancia")
        has_confianca_localizacao = db_has_column_impl(cur, "programacao_itens", "confianca_localizacao")

        has_ctrl = db_has_column_impl(cur, "programacao_itens_controle", "codigo_programacao")
        has_ctrl_pedido = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "pedido")
        has_ctrl_alt_tipo = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "alteracao_tipo")
        has_ctrl_alt_detalhe = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "alteracao_detalhe")
        has_ctrl_ordem_sugerida = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "ordem_sugerida")
        has_ctrl_eta = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "eta")
        has_ctrl_distancia = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "distancia")
        has_ctrl_confianca_localizacao = has_ctrl and db_has_column_impl(cur, "programacao_itens_controle", "confianca_localizacao")

        status_expr = ("pi.status_pedido" if has_status else "''")
        caixas_atual_expr = ("pi.caixas_atual" if has_caixas_atual else "NULL")
        preco_atual_expr = ("pi.preco_atual" if has_preco_atual else "NULL")
        alterado_em_expr = ("pi.alterado_em" if has_alt_em else "NULL")
        alterado_por_expr = ("pi.alterado_por" if has_alt_por else "NULL")
        ordem_sugerida_expr = ("pi.ordem_sugerida" if has_ordem_sugerida else "NULL")
        eta_expr = ("pi.eta" if has_eta else "''")
        distancia_expr = ("pi.distancia" if has_distancia else "NULL")
        confianca_localizacao_expr = ("pi.confianca_localizacao" if has_confianca_localizacao else "NULL")

        if has_ctrl:
            status_expr = "COALESCE(NULLIF(TRIM(pc.status_pedido), ''), NULLIF(TRIM(" + status_expr + "), ''), 'PENDENTE')"
            caixas_atual_expr = (
                "CASE "
                "WHEN pc.caixas_atual IS NOT NULL THEN pc.caixas_atual "
                "WHEN " + caixas_atual_expr + " IS NOT NULL THEN " + caixas_atual_expr + " "
                "ELSE COALESCE(pi.qnt_caixas, 0) "
                "END"
            )
            preco_atual_expr = "COALESCE(pc.preco_atual, " + preco_atual_expr + ", 0)"
            alterado_em_expr = "COALESCE(NULLIF(TRIM(pc.alterado_em), ''), NULLIF(TRIM(" + alterado_em_expr + "), ''), '')"
            alterado_por_expr = "COALESCE(NULLIF(TRIM(pc.alterado_por), ''), NULLIF(TRIM(" + alterado_por_expr + "), ''), '')"
            ordem_sugerida_expr = (
                "COALESCE(pc.ordem_sugerida, " + ordem_sugerida_expr + ")"
                if has_ctrl_ordem_sugerida
                else "COALESCE(" + ordem_sugerida_expr + ", NULL)"
            )
            eta_expr = (
                "COALESCE(NULLIF(TRIM(pc.eta), ''), NULLIF(TRIM(" + eta_expr + "), ''), '')"
                if has_ctrl_eta
                else "COALESCE(NULLIF(TRIM(" + eta_expr + "), ''), '')"
            )
            distancia_expr = (
                "COALESCE(pc.distancia, " + distancia_expr + ")"
                if has_ctrl_distancia
                else "COALESCE(" + distancia_expr + ", NULL)"
            )
            confianca_localizacao_expr = (
                "COALESCE(pc.confianca_localizacao, " + confianca_localizacao_expr + ")"
                if has_ctrl_confianca_localizacao
                else "COALESCE(" + confianca_localizacao_expr + ", NULL)"
            )
        else:
            status_expr = "COALESCE(NULLIF(TRIM(" + status_expr + "), ''), 'PENDENTE')"
            caixas_atual_expr = "COALESCE(" + caixas_atual_expr + ", COALESCE(pi.qnt_caixas, 0))"
            preco_atual_expr = "COALESCE(" + preco_atual_expr + ", 0)"
            alterado_em_expr = "COALESCE(NULLIF(TRIM(" + alterado_em_expr + "), ''), '')"
            alterado_por_expr = "COALESCE(NULLIF(TRIM(" + alterado_por_expr + "), ''), '')"
            ordem_sugerida_expr = "COALESCE(" + ordem_sugerida_expr + ", NULL)"
            eta_expr = "COALESCE(NULLIF(TRIM(" + eta_expr + "), ''), '')"
            distancia_expr = "COALESCE(" + distancia_expr + ", NULL)"
            confianca_localizacao_expr = "COALESCE(" + confianca_localizacao_expr + ", NULL)"

        select_cols = [
            "pi.cod_cliente",
            "pi.nome_cliente",
            ("pi.endereco" if has_end else "'' as endereco"),
            ("pi.produto" if has_produto else "'' as produto"),
            "pi.qnt_caixas",
            "pi.kg",
            "pi.preco",
            ("pi.vendedor" if has_vendedor else "'' as vendedor"),
            ("pi.pedido" if has_pedido else "'' as pedido"),
            (
                "pi.obs"
                if db_has_column_impl(cur, "programacao_itens", "obs")
                else ("pi.observacao" if db_has_column_impl(cur, "programacao_itens", "observacao") else "'' as obs")
            ),
            status_expr + " as status_pedido",
            caixas_atual_expr + " as caixas_atual",
            preco_atual_expr + " as preco_atual",
            alterado_em_expr + " as alterado_em",
            alterado_por_expr + " as alterado_por",
            ordem_sugerida_expr + " as ordem_sugerida",
            eta_expr + " as eta",
            distancia_expr + " as distancia",
            confianca_localizacao_expr + " as confianca_localizacao",
        ]

        join_ctrl = ""
        if has_ctrl:
            select_cols.extend([
                "COALESCE(pc.mortalidade_aves, 0) as mortalidade_aves",
                "COALESCE(pc.peso_previsto, 0) as peso_previsto",
                "COALESCE(pc.valor_recebido, 0) as valor_recebido",
                "COALESCE(pc.forma_recebimento, '') as forma_recebimento",
                "COALESCE(pc.obs_recebimento, '') as obs_recebimento",
                ("COALESCE(pc.alteracao_tipo, '') as alteracao_tipo" if has_ctrl_alt_tipo else "'' as alteracao_tipo"),
                ("COALESCE(pc.alteracao_detalhe, '') as alteracao_detalhe" if has_ctrl_alt_detalhe else "'' as alteracao_detalhe"),
            ])
            join_ctrl = (
                "LEFT JOIN programacao_itens_controle pc "
                "ON pc.codigo_programacao = pi.codigo_programacao "
                "AND UPPER(pc.cod_cliente)=UPPER(pi.cod_cliente)"
            )
            if has_pedido and has_ctrl_pedido:
                join_ctrl += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"
        else:
            select_cols.extend([
                "0 as mortalidade_aves",
                "0 as peso_previsto",
                "0 as valor_recebido",
                "'' as forma_recebimento",
                "'' as obs_recebimento",
                "'' as alteracao_tipo",
                "'' as alteracao_detalhe",
            ])

        cur.execute(
            f"""
            SELECT {", ".join(select_cols)}
            FROM programacao_itens pi
            {join_ctrl}
            WHERE pi.codigo_programacao=?
            ORDER BY pi.id ASC
            LIMIT ?
            """,
            (codigo_programacao, limit),
        )

        rows = cur.fetchall()

    out = []
    for r in rows:
        out.append(
            {
                "cod_cliente": r[0] or "",
                "nome_cliente": r[1] or "",
                "endereco": r[2] or "",
                "produto": r[3] or "",
                "qnt_caixas": safe_int(r[4], 0),
                "kg": safe_float(r[5], 0.0),
                "preco": safe_float(r[6], 0.0),
                "vendedor": r[7] or "",
                "pedido": r[8] or "",
                "obs": (r[9] or "") if has_obs else (r[9] or ""),
                "status_pedido": r[10] or "",
                "caixas_atual": safe_int(r[11], 0),
                "preco_atual": safe_float(r[12], 0.0),
                "alterado_em": r[13] or "",
                "alterado_por": r[14] or "",
                "mortalidade_aves": safe_int(r[15], 0),
                "peso_previsto": safe_float(r[16], 0.0),
                "valor_recebido": safe_float(r[17], 0.0),
                "forma_recebimento": r[18] or "",
                "obs_recebimento": r[19] or "",
                "alteracao_tipo": r[20] or "",
                "alteracao_detalhe": r[21] or "",
                "ordem_sugerida": safe_int(r[22], 0) if r[22] not in (None, "") else None,
                "eta": r[23] or "",
                "distancia": safe_float(r[24], 0.0) if r[24] not in (None, "") else None,
                "confianca_localizacao": safe_float(r[25], 0.0) if r[25] not in (None, "") else None,
            }
        )
    return out
