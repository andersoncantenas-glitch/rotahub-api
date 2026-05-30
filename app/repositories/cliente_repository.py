# -*- coding: utf-8 -*-
from app.db.connection import get_db
from app.utils.formatters import safe_float, safe_int


def _table_exists(cur, table: str) -> bool:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
    return cur.fetchone() is not None


def _columns(cur, table: str):
    try:
        cur.execute(f"PRAGMA table_info({table})")
        return {str(r[1]) for r in cur.fetchall() or []}
    except Exception:
        return set()


def _col_expr(cols, alias: str, names, fallback="''"):
    for name in names:
        if name in cols:
            return f"{alias}.{name}"
    return fallback


def fetch_clientes_rows_local(limit: int = 5000):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT cod_cliente, nome_cliente, endereco, telefone, vendedor
            FROM clientes
            ORDER BY nome_cliente ASC
            LIMIT ?
            """,
            (int(limit or 5000),),
        )
        return cur.fetchall()


def upsert_clientes_local(linhas):
    with get_db() as conn:
        cur = conn.cursor()
        for cod, nome, endereco, telefone, vendedor in (linhas or []):
            cur.execute(
                """
                INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(cod_cliente) DO UPDATE SET
                    nome_cliente=excluded.nome_cliente,
                    endereco=COALESCE(NULLIF(excluded.endereco,''), clientes.endereco),
                    telefone=COALESCE(NULLIF(excluded.telefone,''), clientes.telefone),
                    vendedor=COALESCE(NULLIF(excluded.vendedor,''), clientes.vendedor)
                """,
                (cod, nome, endereco, telefone, vendedor),
            )


def fetch_clientes_dashboard_local():
    with get_db() as conn:
        cur = conn.cursor()
        total_clientes = 0
        clientes_com_historico = 0
        amostras_localizacao = 0
        clientes_com_localizacao = 0

        if _table_exists(cur, "clientes"):
            cur.execute("SELECT COUNT(*) FROM clientes")
            total_clientes = safe_int(cur.fetchone()[0], 0)

        if _table_exists(cur, "programacao_itens"):
            cur.execute(
                """
                SELECT COUNT(DISTINCT NULLIF(TRIM(COALESCE(cod_cliente,'')), ''))
                FROM programacao_itens
                WHERE NULLIF(TRIM(COALESCE(cod_cliente,'')), '') IS NOT NULL
                """
            )
            clientes_com_historico = safe_int(cur.fetchone()[0], 0)

        if _table_exists(cur, "cliente_localizacao_amostras"):
            cur.execute("SELECT COUNT(*) FROM cliente_localizacao_amostras")
            amostras_localizacao = safe_int(cur.fetchone()[0], 0)
            cur.execute(
                """
                SELECT COUNT(DISTINCT NULLIF(TRIM(COALESCE(cod_cliente,'')), ''))
                FROM cliente_localizacao_amostras
                WHERE (latitude IS NOT NULL OR longitude IS NOT NULL)
                  AND NULLIF(TRIM(COALESCE(cod_cliente,'')), '') IS NOT NULL
                """
            )
            clientes_com_localizacao = safe_int(cur.fetchone()[0], 0)

        return {
            "total_clientes": total_clientes,
            "clientes_com_historico": clientes_com_historico,
            "amostras_localizacao": amostras_localizacao,
            "clientes_com_localizacao": clientes_com_localizacao,
        }


def fetch_clientes_lookup_local(limit: int = 5000):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT cod_cliente, COALESCE(NULLIF(nome_cliente,''), nome, '') AS nome_cliente
            FROM clientes
            WHERE NULLIF(TRIM(COALESCE(cod_cliente,'')), '') IS NOT NULL
            ORDER BY nome_cliente ASC
            LIMIT ?
            """,
            (int(limit or 5000),),
        )
        return [(str(r[0] or ""), str(r[1] or "")) for r in cur.fetchall() or []]


def fetch_cliente_historico_local(cod_cliente: str, limit: int = 300):
    cod = str(cod_cliente or "").strip()
    if not cod:
        return {"resumo": {}, "rows": []}

    with get_db() as conn:
        cur = conn.cursor()
        if not _table_exists(cur, "programacao_itens"):
            return {"resumo": {}, "rows": []}

        pi_cols = _columns(cur, "programacao_itens")
        pc_cols = _columns(cur, "programacao_itens_controle") if _table_exists(cur, "programacao_itens_controle") else set()
        p_cols = _columns(cur, "programacoes") if _table_exists(cur, "programacoes") else set()

        prog_expr = _col_expr(pi_cols, "pi", ["codigo_programacao"], "''")
        pedido_expr = _col_expr(pi_cols, "pi", ["pedido"], "''")
        nome_expr = _col_expr(pi_cols, "pi", ["nome_cliente"], "''")
        vendedor_expr = _col_expr(pi_cols, "pi", ["vendedor"], "''")
        caixas_prog_expr = _col_expr(pi_cols, "pi", ["qnt_caixas", "caixas"], "0")
        kg_prog_expr = _col_expr(pi_cols, "pi", ["kg", "kg_cliente"], "0")
        preco_prog_expr = _col_expr(pi_cols, "pi", ["preco"], "0")
        status_pi = _col_expr(pi_cols, "pi", ["status_pedido"], "NULL")
        caixas_pi = _col_expr(pi_cols, "pi", ["caixas_atual"], "NULL")
        preco_pi = _col_expr(pi_cols, "pi", ["preco_atual"], "NULL")
        alt_pi = _col_expr(pi_cols, "pi", ["alteracao_tipo", "alteracao_detalhe"], "NULL")
        data_expr = _col_expr(p_cols, "p", ["data_criacao", "data", "saida_data"], "''")
        motorista_expr = _col_expr(p_cols, "p", ["motorista", "motorista_codigo", "codigo_motorista"], "''")

        has_pc = bool(pc_cols)
        join_pc = ""
        status_pc = "NULL"
        caixas_pc = "NULL"
        preco_pc = "NULL"
        alt_pc = "NULL"
        mortalidade_expr = "0"
        peso_previsto_expr = "NULL"
        valor_recebido_expr = "0"
        lat_expr = "NULL"
        lon_expr = "NULL"
        alterado_em_expr = "''"
        if has_pc:
            join_pc = (
                "LEFT JOIN programacao_itens_controle pc "
                "ON UPPER(TRIM(COALESCE(pc.codigo_programacao,''))) = UPPER(TRIM(COALESCE(pi.codigo_programacao,''))) "
                "AND UPPER(TRIM(COALESCE(pc.cod_cliente,''))) = UPPER(TRIM(COALESCE(pi.cod_cliente,''))) "
                "AND UPPER(TRIM(COALESCE(pc.pedido,''))) = UPPER(TRIM(COALESCE(pi.pedido,'')))"
            )
            status_pc = _col_expr(pc_cols, "pc", ["status_pedido"], "NULL")
            caixas_pc = _col_expr(pc_cols, "pc", ["caixas_atual"], "NULL")
            preco_pc = _col_expr(pc_cols, "pc", ["preco_atual"], "NULL")
            alt_pc = _col_expr(pc_cols, "pc", ["alteracao_tipo", "alteracao_detalhe"], "NULL")
            mortalidade_expr = _col_expr(pc_cols, "pc", ["mortalidade_aves"], "0")
            peso_previsto_expr = _col_expr(pc_cols, "pc", ["peso_previsto"], "NULL")
            valor_recebido_expr = _col_expr(pc_cols, "pc", ["valor_recebido"], "0")
            lat_expr = _col_expr(pc_cols, "pc", ["lat_entrega", "lat_evento"], "NULL")
            lon_expr = _col_expr(pc_cols, "pc", ["lon_entrega", "lon_evento"], "NULL")
            alterado_em_expr = _col_expr(pc_cols, "pc", ["alterado_em", "updated_at"], "''")

        join_p = ""
        if p_cols:
            p_code = _col_expr(p_cols, "p", ["codigo_programacao", "codigo"], "''")
            join_p = f"LEFT JOIN programacoes p ON UPPER(TRIM(COALESCE({p_code},''))) = UPPER(TRIM(COALESCE({prog_expr},'')))"

        sql = f"""
            SELECT
                {data_expr} AS data_ref,
                {prog_expr} AS codigo_programacao,
                {pedido_expr} AS pedido,
                {nome_expr} AS nome_cliente,
                {vendedor_expr} AS vendedor,
                COALESCE(NULLIF(TRIM({status_pc}), ''), NULLIF(TRIM({status_pi}), ''), 'PENDENTE') AS status_pedido,
                COALESCE({caixas_prog_expr}, 0) AS caixas_programadas,
                COALESCE({caixas_pc}, {caixas_pi}, {caixas_prog_expr}, 0) AS caixas_atuais,
                COALESCE({kg_prog_expr}, 0) AS kg_programado,
                COALESCE({peso_previsto_expr}, {kg_prog_expr}, 0) AS kg_atual,
                COALESCE({mortalidade_expr}, 0) AS mortalidade_aves,
                COALESCE({preco_prog_expr}, 0) AS preco_programado,
                COALESCE({preco_pc}, {preco_pi}, {preco_prog_expr}, 0) AS preco_atual,
                COALESCE({valor_recebido_expr}, 0) AS valor_recebido,
                {lat_expr} AS latitude,
                {lon_expr} AS longitude,
                COALESCE({alt_pc}, {alt_pi}, '') AS alteracao,
                {alterado_em_expr} AS alterado_em,
                {motorista_expr} AS motorista
            FROM programacao_itens pi
            {join_pc}
            {join_p}
            WHERE UPPER(TRIM(COALESCE(pi.cod_cliente,''))) = UPPER(TRIM(?))
            ORDER BY COALESCE(NULLIF(data_ref,''), {prog_expr}) DESC, pi.id DESC
            LIMIT ?
        """
        cur.execute(sql, (cod, int(limit or 300)))
        rows = []
        for r in cur.fetchall() or []:
            status = str(r["status_pedido"] or "PENDENTE").strip().upper()
            kg_prog = safe_float(r["kg_programado"], 0.0)
            kg_atual = safe_float(r["kg_atual"], kg_prog)
            entregue = status in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
            cancelado = status in {"CANCELADO", "CANCELADA"}
            alterado = bool(str(r["alteracao"] or "").strip()) or safe_int(r["caixas_atuais"], 0) != safe_int(r["caixas_programadas"], 0)
            kg_recebido = kg_atual if entregue and not cancelado else 0.0
            kg_descontado = kg_prog if cancelado else max(kg_prog - kg_atual, 0.0)
            rows.append({
                "data_ref": str(r["data_ref"] or ""),
                "codigo_programacao": str(r["codigo_programacao"] or ""),
                "pedido": str(r["pedido"] or ""),
                "nome_cliente": str(r["nome_cliente"] or ""),
                "vendedor": str(r["vendedor"] or ""),
                "status_pedido": status,
                "caixas_programadas": safe_int(r["caixas_programadas"], 0),
                "caixas_atuais": safe_int(r["caixas_atuais"], 0),
                "kg_programado": kg_prog,
                "kg_recebido": kg_recebido,
                "kg_descontado": kg_descontado,
                "mortalidade_aves": safe_int(r["mortalidade_aves"], 0),
                "preco_atual": safe_float(r["preco_atual"], 0.0),
                "valor_recebido": safe_float(r["valor_recebido"], 0.0),
                "latitude": "" if r["latitude"] is None else str(r["latitude"]),
                "longitude": "" if r["longitude"] is None else str(r["longitude"]),
                "alterado": alterado,
                "alterado_em": str(r["alterado_em"] or ""),
                "motorista": str(r["motorista"] or ""),
            })

        resumo = {
            "total_programacoes": len(rows),
            "entregues": sum(1 for r in rows if r["status_pedido"] in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}),
            "canceladas": sum(1 for r in rows if r["status_pedido"] in {"CANCELADO", "CANCELADA"}),
            "alteradas": sum(1 for r in rows if r["alterado"]),
            "mortalidade_aves": sum(safe_int(r["mortalidade_aves"], 0) for r in rows),
            "kg_recebidos": sum(safe_float(r["kg_recebido"], 0.0) for r in rows),
            "kg_descontados": sum(safe_float(r["kg_descontado"], 0.0) for r in rows),
        }
        return {"resumo": resumo, "rows": rows}


def fetch_cliente_localizacoes_local(cod_cliente: str, limit: int = 200):
    cod = str(cod_cliente or "").strip()
    if not cod:
        return {"resumo": {}, "rows": []}

    with get_db() as conn:
        cur = conn.cursor()
        if not _table_exists(cur, "cliente_localizacao_amostras"):
            return {"resumo": {}, "rows": []}
        cur.execute(
            """
            SELECT codigo_programacao, pedido, latitude, longitude, endereco, cidade, bairro,
                   status_pedido, motorista_codigo, motorista_nome, origem, registrado_em
            FROM cliente_localizacao_amostras
            WHERE UPPER(TRIM(COALESCE(cod_cliente,''))) = UPPER(TRIM(?))
            ORDER BY registrado_em DESC, id DESC
            LIMIT ?
            """,
            (cod, int(limit or 200)),
        )
        rows = [dict(r) for r in cur.fetchall() or []]
        resumo = {
            "amostras": len(rows),
            "ultima": rows[0] if rows else {},
            "com_coordenada": sum(1 for r in rows if r.get("latitude") is not None or r.get("longitude") is not None),
        }
        return {"resumo": resumo, "rows": rows}
