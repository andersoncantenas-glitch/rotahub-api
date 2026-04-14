# -*- coding: utf-8 -*-
from app.db.connection import get_db


def update_vendedor_senha_hash_local(vendedor_id: int, senha_hash: str):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE vendedores SET senha=? WHERE id=?",
            (str(senha_hash or ""), int(vendedor_id or 0)),
        )
