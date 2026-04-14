# -*- coding: utf-8 -*-
from app.db.connection import get_db


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
