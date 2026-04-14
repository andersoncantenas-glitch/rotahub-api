import logging
import os
import tkinter as tk
from tkinter import messagebox

from app.db.connection import get_db
from app.security.passwords import hash_password_pbkdf2, verify_password_pbkdf2


def _service_result(*, ok: bool, data=None, error: str = None, source: str = "local"):
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def autenticar_usuario(login: str, senha: str):
    """
    Login por NOME + SENHA
    Migração automática:
       - Se senha no DB já for HASH: valida HASH
       - Se senha no DB for PURA: valida pura e, se OK, converte e salva HASH
    Retorna contrato de serviço padronizado.
    """
    login = (login or "").strip()
    senha = (senha or "").strip()
    if not login or not senha:
        return _service_result(ok=False, data=None, error="Credenciais inválidas.", source="local")

    with get_db() as conn:
        cur = conn.cursor()

        # Descobre colunas disponíveis
        cur.execute("PRAGMA table_info(usuarios)")
        cols = [str(r[1] or "").lower() for r in cur.fetchall()]

        has_permissoes = "permissoes" in cols
        has_cpf = "cpf" in cols
        has_telefone = "telefone" in cols
        has_senha = "senha" in cols

        if not has_senha:
            return _service_result(ok=False, data=None, error="Base sem coluna de senha.", source="local")

        # Puxa o usuário pelo nome (case-insensitive)
        select_parts = ["id", "nome", "senha"]
        select_parts.append("permissoes" if has_permissoes else "'' as permissoes")
        select_parts.append("cpf" if has_cpf else "'' as cpf")
        select_parts.append("telefone" if has_telefone else "'' as telefone")

        cur.execute(
            f"""
            SELECT {", ".join(select_parts)}
            FROM usuarios
            WHERE UPPER(nome)=UPPER(?)
            LIMIT 1
        """,
            (login,),
        )
        row = cur.fetchone()

        if not row:
            return _service_result(ok=False, data=None, error="Credenciais inválidas.", source="local")

        user_id = row[0]
        nome = row[1] or ""
        senha_db = row[2] or ""
        permissoes = row[3] if len(row) > 3 else ""
        cpf = row[4] if len(row) > 4 else ""
        telefone = row[5] if len(row) > 5 else ""

        # 1) Se já é hash, valida por hash
        if str(senha_db).startswith("pbkdf2_sha256$"):
            if not verify_password_pbkdf2(senha, senha_db):
                return _service_result(ok=False, data=None, error="Credenciais inválidas.", source="local")

        # 2) Senha pura: valida pura e MIGRA para hash automaticamente
        else:
            if str(senha_db) != senha:
                return _service_result(ok=False, data=None, error="Credenciais inválidas.", source="local")

            try:
                novo_hash = hash_password_pbkdf2(senha)
                cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (novo_hash, user_id))
            except Exception as e:
                # Se falhar a migração, pelo menos deixa logar (já validou a senha pura)
                logging.exception("Falha ao migrar senha do usuario id=%s: %s", user_id, e)

    is_admin = ("ADMIN" in (permissoes or "").upper()) or (nome.strip().upper() == "ADMIN")

    return _service_result(
        ok=True,
        data={
            "id": user_id,
            "nome": nome,
            "permissoes": permissoes,
            "cpf": cpf,
            "telefone": telefone,
            "is_admin": is_admin,
        },
        source="local",
    )


def _notify_admin_seed(password: str):
    """Notifica senha temporária do ADMIN sem gravar em arquivo."""
    msg = (
        "ADMIN criado automaticamente.\n"
        f"Senha temporaria: {password}\n\n"
        "Troque a senha no primeiro acesso."
    )
    try:
        if tk._default_root is not None:
            messagebox.showwarning("ADMIN CRIADO", msg)
            return
    except Exception:
        logging.debug("Falha ignorada")
    print(msg)


def ensure_admin_user():
    """Garante que exista ADMIN (sem sobrescrever senha existente)."""
    with get_db() as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(usuarios)")
        cols = [str(r[1] or "").lower() for r in cur.fetchall()]
        if "senha" not in cols:
            return _service_result(ok=False, data=None, error="Base sem coluna de senha.", source="local")

        has_permissoes = "permissoes" in cols

        cur.execute("SELECT id, senha FROM usuarios WHERE UPPER(nome)=?", ("ADMIN",))
        row = cur.fetchone()

        if not row:
            senha_plana = (
                os.environ.get("ROTA_ADMIN_PASS")
                or os.environ.get("ROTA_ADMIN_PASSWORD")
                or "123456"
            ).strip()
            if not senha_plana:
                senha_plana = "123456"
            senha_hash = hash_password_pbkdf2(senha_plana)

            if has_permissoes:
                cur.execute(
                    "INSERT INTO usuarios (nome, senha, permissoes) VALUES (?, ?, ?)",
                    ("ADMIN", senha_hash, "ADMIN"),
                )
            else:
                cur.execute(
                    "INSERT INTO usuarios (nome, senha) VALUES (?, ?)",
                    ("ADMIN", senha_hash),
                )
            _notify_admin_seed(senha_plana)
        else:
            admin_id = row[0]
            senha_db = row[1] or ""
            # se admin ainda estiver com senha pura, converte para hash mantendo o mesmo valor
            if senha_db and not str(senha_db).startswith("pbkdf2_sha256$"):
                try:
                    novo_hash = hash_password_pbkdf2(str(senha_db))
                    if has_permissoes:
                        cur.execute(
                            "UPDATE usuarios SET senha=?, permissoes=? WHERE id=?",
                            (novo_hash, "ADMIN", admin_id),
                        )
                    else:
                        cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (novo_hash, admin_id))
                except Exception as e:
                    logging.exception("Falha ao migrar senha do ADMIN id=%s: %s", admin_id, e)

    return _service_result(ok=True, data=None, source="local")


__all__ = [
    "_notify_admin_seed",
    "autenticar_usuario",
    "ensure_admin_user",
]
