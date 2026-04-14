# -*- coding: utf-8 -*-
from app.db.connection import get_db


def fetch_motorista_codigos_local(cur=None):
    if cur is not None:
        cur.execute("SELECT COALESCE(codigo,'') FROM motoristas")
        return [str((row[0] if row else "") or "") for row in (cur.fetchall() or [])]

    with get_db() as conn:
        c2 = conn.cursor()
        c2.execute("SELECT COALESCE(codigo,'') FROM motoristas")
        return [str((row[0] if row else "") or "") for row in (c2.fetchall() or [])]


def fetch_motoristas_cache_local_by_codigo():
    out = {}
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT codigo, senha, cpf, telefone FROM motoristas")
        for codigo, senha, cpf, telefone in (cur.fetchall() or []):
            out[str(codigo or "").strip()] = {
                "senha": str(senha or ""),
                "cpf": str(cpf or ""),
                "telefone": str(telefone or ""),
            }
    return out


def fetch_motorista_access_snapshot_by_codigo(codigo: str):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if "acesso_liberado" not in cols_m:
            return None

        cur.execute(
            """
            SELECT
                COALESCE(acesso_liberado,1) AS acesso_liberado,
                COALESCE(acesso_liberado_por,'') AS acesso_liberado_por,
                COALESCE(acesso_obs,'') AS acesso_obs
            FROM motoristas
            WHERE UPPER(COALESCE(codigo,''))=UPPER(?)
            ORDER BY id DESC
            LIMIT 1
            """,
            (str(codigo or "").strip(),),
        )
        row = cur.fetchone()
        if not row:
            return None
        return {
            "acesso_liberado": int(row[0] or 0),
            "acesso_liberado_por": str(row[1] or "").strip(),
            "acesso_obs": str(row[2] or "").strip(),
        }


def fetch_motorista_nome_by_id(motorista_id: int):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(nome,'') FROM motoristas WHERE id=? LIMIT 1", (int(motorista_id or 0),))
        rr = cur.fetchone()
        return str((rr[0] if rr else "") or "")


def update_motorista_status_local(motorista_id: int, status: str):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE motoristas SET status=? WHERE id=?",
            (str(status or "").strip(), int(motorista_id or 0)),
        )
