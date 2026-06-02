# api_server.py
import os
import json
import sqlite3
import base64
import hmac
import hashlib
import time
import math
import logging
import re
from uuid import uuid4
from datetime import datetime
from typing import Optional, List, Dict, Any, Iterator, Tuple
from contextlib import contextmanager

from fastapi import FastAPI, HTTPException, Depends, Query, Header, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, Field
from app.middleware.billing_middleware import BillingProtectionMiddleware
from app.middleware.feature_middleware import FeatureGateMiddleware
from app.middleware.tenant_middleware import TenantContextMiddleware
from app.services.billing_automation_service import suspend_overdue_subscriptions_conn
from app.services.vehicle_limit_service import check_vehicle_limit, vehicle_usage_snapshot
from app.utils.validators import (
    is_valid_cpf,
    is_valid_motorista_codigo,
    is_valid_motorista_senha,
    is_valid_phone,
    normalize_cpf,
    normalize_phone,
    validate_placa,
)
from version import APP_VERSION
from db_bootstrap import ensure_admin_user as ensure_admin_user_bootstrap, ensure_core_schema
from database_runtime import log_startup_diagnostics
from runtime_config import apply_process_environment, ensure_runtime_files, load_app_config


# =========================================================
# CONFIG
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_CONFIG = load_app_config("server")
apply_process_environment(APP_CONFIG)
ensure_runtime_files(APP_CONFIG)

# âœ… Prioridade:
# 1) variável de ambiente ROTA_DB
# 2) rota_granja.db na pasta do projeto
DB_PATH = APP_CONFIG.db_path

SECRET_KEY = os.environ.get("ROTA_SECRET")
if not SECRET_KEY:
    raise RuntimeError("ROTA_SECRET nao definido. Configure a variavel de ambiente para iniciar a API.")
TOKEN_TTL_SECONDS = 60 * 60 * 24 * 7  # 7 dias

app = FastAPI(title="Rota Granja API", version=APP_VERSION)

def _cors_origins_from_env() -> List[str]:
    raw = os.environ.get("ROTA_CORS_ORIGINS", "").strip()
    if raw:
        return [o.strip() for o in raw.split(",") if o.strip()]
    # Defaults para desenvolvimento local (Flutter web / front local)
    return [
        "http://127.0.0.1:3000",
        "http://localhost:3000",
        "http://127.0.0.1:5173",
        "http://localhost:5173",
        "http://127.0.0.1:8000",
        "http://localhost:8000",
    ]


CORS_ORIGINS = _cors_origins_from_env()
CORS_ALLOW_CREDENTIALS = os.environ.get("ROTA_CORS_ALLOW_CREDENTIALS", "0").strip() in ("1", "true", "TRUE")
ENABLE_ROTAS_ATIVAS_TODAS = os.environ.get("ROTA_ENABLE_ROTAS_ATIVAS_TODAS", "1").strip() in ("1", "true", "TRUE")
ENABLE_START_GPS_GATE = os.environ.get("ROTA_ENABLE_START_GPS_GATE", "0").strip() in ("1", "true", "TRUE")

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=CORS_ALLOW_CREDENTIALS,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer()

# =========================================================
# DB HELPERS
# =========================================================
@contextmanager
def get_conn() -> Iterator[sqlite3.Connection]:
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA busy_timeout = 5000")
    except Exception:
        pass
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def col_exists(conn: sqlite3.Connection, table: str, col: str) -> bool:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return col in cols


def _default_company_id(cur: sqlite3.Cursor) -> int:
    try:
        cur.execute("SELECT id FROM companies ORDER BY id ASC LIMIT 1")
        row = cur.fetchone()
        if row:
            return int(row[0])
    except Exception:
        pass
    return 1


def _row_company_id(row: Any, default_company_id: int = 1) -> int:
    if row is None:
        return int(default_company_id or 1)
    try:
        value = row["company_id"]
    except Exception:
        try:
            value = row.get("company_id")
        except Exception:
            value = None
    try:
        return int(value or default_company_id or 1)
    except Exception:
        return int(default_company_id or 1)


def table_exists(cur: sqlite3.Cursor, table: str) -> bool:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
    return bool(cur.fetchone())


def _ensure_fornecedor_perfis_schema(cur: sqlite3.Cursor) -> None:
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS fornecedor_perfis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            nome TEXT NOT NULL,
            categoria TEXT DEFAULT 'OUTROS',
            status TEXT DEFAULT 'ATIVO',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    for codigo, nome, categoria in (
        ("FRANGO_VIVO", "Frango vivo", "FRANGO"),
        ("PRESTADOR_SERVICO", "Prestador de servico", "SERVICO"),
        ("PECAS", "Pecas", "MANUTENCAO"),
        ("MECANICO", "Mecanico", "MANUTENCAO"),
        ("BORRACHEIRO", "Borracheiro", "MANUTENCAO"),
        ("LAVADOR_CAIXAS", "Lavador de caixas", "OPERACIONAL"),
        ("PNEUS", "Pneus", "MANUTENCAO"),
        ("OLEO_LUBRIFICANTES", "Oleo e lubrificantes", "MANUTENCAO"),
        ("COMBUSTIVEL", "Combustivel", "OPERACIONAL"),
        ("SERVICO_SEM_NF", "Servico sem NF", "SERVICO"),
        ("OUTROS", "Outros", "OUTROS"),
    ):
        cur.execute(
            """
            INSERT OR IGNORE INTO fornecedor_perfis (codigo, nome, categoria, status)
            VALUES (?, ?, ?, 'ATIVO')
            """,
            (codigo, nome, categoria),
        )


def _ensure_compras_app_schema(cur: sqlite3.Cursor) -> None:
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS compras_nfe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chave_acesso TEXT UNIQUE,
            serie TEXT,
            numero TEXT,
            fornecedor_documento TEXT,
            fornecedor_razao TEXT,
            emissao TEXT,
            valor_total REAL DEFAULT 0,
            situacao_nfe TEXT DEFAULT 'Autorizado',
            nsu TEXT,
            natureza_operacao TEXT,
            fornecedor_id INTEGER,
            origem TEXT DEFAULT 'XML',
            xml_path TEXT,
            pdf_path TEXT,
            estoque_fiscal_status TEXT DEFAULT 'PENDENTE',
            estoque_fisico_status TEXT DEFAULT 'PENDENTE',
            estoque_kg_entrada REAL DEFAULT 0,
            estoque_kg_saldo REAL DEFAULT 0,
            estoque_caixas_entrada INTEGER DEFAULT 0,
            produto_id INTEGER,
            produto TEXT,
            codigo_programacao TEXT,
            vinculada_em TEXT,
            vinculada_por TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT
        )
        """
    )
    for col, decl in {
        "chave_acesso": "TEXT",
        "serie": "TEXT",
        "numero": "TEXT",
        "fornecedor_documento": "TEXT",
        "fornecedor_razao": "TEXT",
        "emissao": "TEXT",
        "valor_total": "REAL DEFAULT 0",
        "situacao_nfe": "TEXT DEFAULT 'Autorizado'",
        "nsu": "TEXT",
        "natureza_operacao": "TEXT",
        "fornecedor_id": "INTEGER",
        "origem": "TEXT DEFAULT 'XML'",
        "xml_path": "TEXT",
        "pdf_path": "TEXT",
        "estoque_fiscal_status": "TEXT DEFAULT 'PENDENTE'",
        "estoque_fisico_status": "TEXT DEFAULT 'PENDENTE'",
        "estoque_kg_entrada": "REAL DEFAULT 0",
        "estoque_kg_saldo": "REAL DEFAULT 0",
        "estoque_caixas_entrada": "INTEGER DEFAULT 0",
        "produto_id": "INTEGER",
        "produto": "TEXT",
        "codigo_programacao": "TEXT",
        "vinculada_em": "TEXT",
        "vinculada_por": "TEXT",
        "created_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
        "updated_at": "TEXT",
    }.items():
        if not col_exists(cur.connection, "compras_nfe", col):
            cur.execute(f"ALTER TABLE compras_nfe ADD COLUMN {col} {decl}")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_numero ON compras_nfe(numero)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_vinculo ON compras_nfe(codigo_programacao)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_fornecedor ON compras_nfe(fornecedor_documento)")


def _normalize_nf_numero(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or "").strip())


def _nf_frango_clause() -> str:
    return """
        (
            UPPER(TRIM(COALESCE(f.perfil_fornecedor, '')))='FRANGO_VIVO'
            OR UPPER(TRIM(COALESCE(c.produto, ''))) LIKE '%FRANGO%'
            OR UPPER(TRIM(COALESCE(c.produto, ''))) LIKE '%AVE%'
            OR UPPER(TRIM(COALESCE(c.natureza_operacao, ''))) LIKE '%FRANGO%'
            OR UPPER(TRIM(COALESCE(c.natureza_operacao, ''))) LIKE '%AVE%'
        )
    """


def _fetch_nf_compra_by_numero(cur: sqlite3.Cursor, nf_numero: str) -> Optional[sqlite3.Row]:
    _ensure_compras_app_schema(cur)
    numero = _normalize_nf_numero(nf_numero)
    if not numero:
        return None
    cur.execute(
        f"""
        SELECT c.*,
               COALESCE(f.perfil_fornecedor, '') AS perfil_fornecedor,
               COALESCE(f.razao_social, c.fornecedor_razao, '') AS fornecedor_nome
          FROM compras_nfe c
          LEFT JOIN fornecedores f
            ON (
                (c.fornecedor_id IS NOT NULL AND f.id=c.fornecedor_id)
                OR (TRIM(COALESCE(c.fornecedor_documento,''))<>'' AND TRIM(COALESCE(f.documento,''))=TRIM(COALESCE(c.fornecedor_documento,'')))
            )
         WHERE REPLACE(REPLACE(REPLACE(TRIM(COALESCE(c.numero,'')), '.', ''), '-', ''), '/', '')=?
         ORDER BY c.id DESC
         LIMIT 1
        """,
        (numero,),
    )
    return cur.fetchone()


def _nf_compra_is_frango(row: sqlite3.Row) -> bool:
    text = " ".join(
        str(row[key] or "").upper()
        for key in ("perfil_fornecedor", "produto", "natureza_operacao")
        if key in row.keys()
    )
    return any(token in text for token in ("FRANGO_VIVO", "FRANGO", "AVE"))


def _nf_compra_payload(row: sqlite3.Row) -> Dict[str, Any]:
    kg = float(row["estoque_kg_entrada"] or 0.0)
    caixas = int(row["estoque_caixas_entrada"] or 0)
    valor_total = float(row["valor_total"] or 0.0)
    preco_kg = round(valor_total / kg, 6) if kg > 0 and valor_total > 0 else 0.0
    return {
        "id": int(row["id"] or 0),
        "numero": str(row["numero"] or "").strip(),
        "serie": str(row["serie"] or "").strip(),
        "chave_acesso": str(row["chave_acesso"] or "").strip(),
        "fornecedor": str(row["fornecedor_nome"] or row["fornecedor_razao"] or "").strip(),
        "fornecedor_documento": str(row["fornecedor_documento"] or "").strip(),
        "perfil_fornecedor": str(row["perfil_fornecedor"] or "").strip().upper(),
        "emissao": str(row["emissao"] or "").strip(),
        "produto": str(row["produto"] or "").strip(),
        "nf_kg": kg,
        "nf_caixas": caixas,
        "nf_preco": preco_kg,
        "valor_total": valor_total,
        "estoque_fiscal_status": str(row["estoque_fiscal_status"] or "").strip(),
        "estoque_fisico_status": str(row["estoque_fisico_status"] or "").strip(),
    }


def _ensure_clientes_columns(cur: sqlite3.Cursor) -> set[str]:
    cur.execute("PRAGMA table_info(clientes)")
    cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    expected = {
        "cod_cliente": "TEXT",
        "nome_cliente": "TEXT",
        "endereco": "TEXT",
        "bairro": "TEXT",
        "cidade": "TEXT",
        "uf": "TEXT",
        "telefone": "TEXT",
        "rota": "TEXT",
        "vendedor": "TEXT",
        "latitude": "REAL",
        "longitude": "REAL",
        "company_id": "INTEGER",
    }
    for col, decl in expected.items():
        if col not in cols:
            cur.execute(f"ALTER TABLE clientes ADD COLUMN {col} {decl}")
            cols.add(col)
    if "company_id" in cols:
        cur.execute(
            "UPDATE clientes SET company_id=? WHERE company_id IS NULL OR company_id=0",
            (_default_company_id(cur),),
        )
    try:
        if "company_id" in cols:
            _drop_single_column_unique_index(cur, "clientes", "cod_cliente")
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_company_cod ON clientes(company_id, cod_cliente)")
        else:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_cod ON clientes(cod_cliente)")
    except Exception:
        logging.debug("Indice unico de clientes(cod_cliente) nao criado; mantendo schema legado.", exc_info=True)
    return cols


def _drop_single_column_unique_index(cur: sqlite3.Cursor, table: str, column: str) -> None:
    target = str(column or "").lower()
    for idx in cur.execute(f"PRAGMA index_list({table})").fetchall() or []:
        idx_name = str(idx[1] if len(idx) > 1 else "")
        is_unique = bool(idx[2] if len(idx) > 2 else 0)
        if not idx_name or not is_unique or idx_name.startswith("sqlite_autoindex"):
            continue
        idx_cols = [
            str(row[2]).lower()
            for row in (cur.execute(f'PRAGMA index_info("{idx_name}")').fetchall() or [])
            if len(row) > 2
        ]
        if idx_cols == [target]:
            cur.execute(f'DROP INDEX IF EXISTS "{idx_name}"')


def _ensure_clientes_table(cur: sqlite3.Cursor) -> set[str]:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
    if not cur.fetchone():
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cod_cliente TEXT,
                nome_cliente TEXT,
                endereco TEXT,
                bairro TEXT,
                cidade TEXT,
                uf TEXT,
                telefone TEXT,
                rota TEXT,
                vendedor TEXT,
                company_id INTEGER
            )
            """
        )
    return _ensure_clientes_columns(cur)


PBKDF2_ITERATIONS = 200_000

def hash_password_pbkdf2(password: str, *, iterations: int = PBKDF2_ITERATIONS) -> str:
    password = str(password or "")
    if password == "":
        raise ValueError("Senha vazia.")
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)
    return "pbkdf2_sha256${}${}${}".format(
        iterations,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(dk).decode("ascii"),
    )

def verify_password_pbkdf2(password: str, stored: str) -> bool:
    try:
        password = str(password or "")
        stored = str(stored or "")

        if not stored.startswith("pbkdf2_sha256$"):
            return False

        _, iters_s, salt_b64, hash_b64 = stored.split("$", 3)
        iterations = int(iters_s)

        salt = base64.b64decode(salt_b64.encode("ascii"))
        expected = base64.b64decode(hash_b64.encode("ascii"))

        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=len(expected))
        return hmac.compare_digest(dk, expected)
    except Exception:
        return False


def _motorista_login_candidates(codigo: str) -> List[str]:
    raw = str(codigo or "").strip().upper()
    if not raw:
        return []

    out: List[str] = []

    def _push(value: str) -> None:
        v = str(value or "").strip().upper()
        if v and v not in out:
            out.append(v)

    _push(raw)
    _push(raw.replace(" ", ""))

    digits = re.sub(r"\D+", "", raw)
    if digits:
        try:
            num = int(digits)
        except Exception:
            num = 0
        if num > 0:
            _push(f"MT{num:03d}")
            _push(f"MT{num:02d}")
            _push(f"MOT-{num:02d}")
            _push(f"MOT-{num:03d}")
            _push(f"MOT{num:02d}")
            _push(f"MOT{num:03d}")

    return out


def authenticate_motorista(cur: sqlite3.Cursor, codigo: str, senha: str) -> tuple[Optional[sqlite3.Row], Optional[str]]:
    has_acesso = col_exists(cur.connection, "motoristas", "acesso_liberado")
    has_perfil_app = col_exists(cur.connection, "motoristas", "perfil_app")
    has_company_id = col_exists(cur.connection, "motoristas", "company_id")
    default_company_id = _default_company_id(cur)
    select_parts = [
        "id",
        "nome",
        "codigo",
        "senha",
        "COALESCE(acesso_liberado, 0) AS acesso_liberado" if has_acesso else "1 AS acesso_liberado",
        "UPPER(TRIM(COALESCE(perfil_app, 'MOTORISTA'))) AS perfil_app" if has_perfil_app else "'MOTORISTA' AS perfil_app",
        f"COALESCE(company_id, {default_company_id}) AS company_id" if has_company_id else f"{default_company_id} AS company_id",
    ]
    select_cols = ", ".join(select_parts)
    row = None
    for candidate in _motorista_login_candidates(codigo):
        cur.execute(
            f"""
            SELECT {select_cols}
            FROM motoristas
            WHERE UPPER(TRIM(codigo))=?
            LIMIT 1
            """,
            (candidate,),
        )
        row = cur.fetchone()
        if row:
            break

    if not row:
        nome_candidate = str(codigo or "").strip().upper()
        if nome_candidate:
            cur.execute(
                f"""
                SELECT {select_cols}
                FROM motoristas
                WHERE UPPER(TRIM(nome))=?
                LIMIT 1
                """,
                (nome_candidate,),
            )
            row = cur.fetchone()
    if not row:
        return None, "not_found"

    if has_acesso and int(row["acesso_liberado"] or 0) != 1:
        return None, "blocked"

    senha_db = row["senha"] or ""
    if str(senha_db).startswith("pbkdf2_sha256$"):
        if not verify_password_pbkdf2(senha, senha_db):
            return None, "invalid_password"
    else:
        if str(senha_db).strip() != senha:
            return None, "invalid_password"
        try:
            novo_hash = hash_password_pbkdf2(senha)
            cur.execute("UPDATE motoristas SET senha=? WHERE id=?", (novo_hash, row["id"]))
        except Exception:
            pass

    return row, None


def authenticate_vendedor(cur: sqlite3.Cursor, codigo: str, senha: str) -> tuple[Optional[sqlite3.Row], Optional[str]]:
    identificador = (codigo or "").strip()
    identificador_norm = identificador.casefold()
    has_company_id = col_exists(cur.connection, "vendedores", "company_id")
    default_company_id = _default_company_id(cur)
    company_expr = f"COALESCE(company_id, {default_company_id}) AS company_id" if has_company_id else f"{default_company_id} AS company_id"
    cur.execute(
        f"""
        SELECT id, nome, codigo, senha, COALESCE(status, 'ATIVO') AS status, {company_expr}
        FROM vendedores
        """
    )
    rows = cur.fetchall() or []
    row = None
    for candidate in rows:
        nome_norm = _clean_text(candidate["nome"]).casefold()
        codigo_norm = _clean_text(candidate["codigo"]).casefold()
        if nome_norm == identificador_norm:
            row = candidate
            break
        if row is None and codigo_norm == identificador_norm:
            row = candidate
    if not row:
        return None, "not_found"

    status = str(row["status"] or "ATIVO").strip().upper()
    if status not in {"ATIVO", ""}:
        return None, "blocked"

    senha_db = row["senha"] or ""
    if str(senha_db).startswith("pbkdf2_sha256$"):
        if not verify_password_pbkdf2(senha, senha_db):
            return None, "invalid_password"
    else:
        if str(senha_db).strip() != senha:
            return None, "invalid_password"
        try:
            novo_hash = hash_password_pbkdf2(senha)
            cur.execute("UPDATE vendedores SET senha=? WHERE id=?", (novo_hash, row["id"]))
        except Exception:
            pass

    return row, None


def row_to_dict(row: sqlite3.Row) -> Dict[str, Any]:
    if row is None:
        return {}
    return {k: row[k] for k in row.keys()}


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()

def _first_name(value: Any) -> str:
    txt = _clean_text(value)
    if not txt:
        return ""
    return txt.split()[0].upper()


def _split_people_tokens(raw: str) -> List[str]:
    txt = _clean_text(raw)
    if not txt:
        return []
    for sep in ("|", "/", ",", ";"):
        txt = txt.replace(sep, "|")
    return [p.strip() for p in txt.split("|") if p.strip()]


def _resolve_ajudante_primeiro_nome(cur: Optional[sqlite3.Cursor], raw: Any) -> str:
    value = _clean_text(raw)
    if not value:
        return ""

    parts = _split_people_tokens(value)
    if len(parts) > 1:
        resolved_parts = []
        for p in parts:
            rp = _resolve_ajudante_primeiro_nome(cur, p)
            if rp and rp not in resolved_parts:
                resolved_parts.append(rp)
        return " / ".join(resolved_parts)

    if cur is not None:
        # Novo modelo: equipe guarda ids de ajudantes.
        try:
            cur.execute(
                """
                SELECT nome, sobrenome
                FROM ajudantes
                WHERE UPPER(TRIM(CAST(id AS TEXT))) = UPPER(TRIM(?))
                LIMIT 1
                """,
                (value,),
            )
            row = cur.fetchone()
            if row:
                nome = _clean_text(row["nome"]) if hasattr(row, "keys") else _clean_text(row[0])
                sobrenome = _clean_text(row["sobrenome"]) if hasattr(row, "keys") else _clean_text(row[1])
                resolved = _first_name(nome) or _first_name(sobrenome)
                if resolved:
                    return resolved
        except Exception:
            pass

        # Fallback: valor pode ser id/codigo de equipe.
        try:
            cur.execute(
                """
                SELECT ajudante1, ajudante2, ajudante_1, ajudante_2
                FROM equipes
                WHERE UPPER(TRIM(CAST(id AS TEXT))) = UPPER(TRIM(?))
                   OR UPPER(TRIM(codigo)) = UPPER(TRIM(?))
                LIMIT 1
                """,
                (value, value),
            )
            eq = cur.fetchone()
            if eq:
                nomes_eq: List[str] = []
                for k in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2"):
                    if k in eq.keys():
                        n = _resolve_ajudante_primeiro_nome(cur, eq[k])
                        if n and n not in nomes_eq:
                            nomes_eq.append(n)
                if nomes_eq:
                    return " / ".join(nomes_eq)
        except Exception:
            pass

    return _first_name(value)


def _load_equipes_map(cur: sqlite3.Cursor) -> Dict[str, str]:
    mapa: Dict[str, str] = {}
    try:
        cur.execute("SELECT codigo, ajudante1, ajudante2 FROM equipes")
        for r in cur.fetchall() or []:
            codigo = _clean_text(r["codigo"]).upper()
            a1 = _resolve_ajudante_primeiro_nome(cur, r["ajudante1"])
            a2 = _resolve_ajudante_primeiro_nome(cur, r["ajudante2"])
            nomes = " / ".join([n for n in [a1, a2] if n])
            if codigo and nomes:
                mapa[codigo] = nomes
    except Exception:
        pass
    return mapa

def _apply_equipe_nome(row: Dict[str, Any], equipes_map: Dict[str, str], cur: Optional[sqlite3.Cursor] = None) -> Dict[str, Any]:
    equipe_raw = row.get("equipe")
    key = _clean_text(equipe_raw).upper()
    if key and key in equipes_map:
        row["equipe"] = equipes_map[key]
    elif key:
        row["equipe"] = _resolve_ajudante_primeiro_nome(cur, equipe_raw)
    return row


def _format_equipe_ajudantes(row: Dict[str, Any], cur: Optional[sqlite3.Cursor] = None) -> str:
    if row is None:
        return ""
    names = []
    for key in ("ajudante1", "ajudante_1", "ajudante2", "ajudante_2"):
        candidate = _resolve_ajudante_primeiro_nome(cur, row.get(key))
        if candidate:
            names.append(candidate)
    if names:
        return " / ".join(names)

    # Fallback robusto: quando "equipe" vier como codigo (ex.: EQP-01),
    # tenta resolver na tabela equipes para retornar nomes dos ajudantes.
    equipe_raw = _clean_text(row.get("equipe"))
    if equipe_raw and cur is not None:
        try:
            cur.execute("PRAGMA table_info(equipes)")
            cols_eq = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            cand_cols = [c for c in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2") if c in cols_eq]
            if not cand_cols:
                return _resolve_ajudante_primeiro_nome(cur, equipe_raw)

            select_cols = ", ".join(cand_cols)
            cur.execute(
                f"""
                SELECT {select_cols}
                FROM equipes
                WHERE UPPER(TRIM(codigo)) = UPPER(TRIM(?))
                LIMIT 1
                """,
                (equipe_raw,),
            )
            eq = cur.fetchone()
            if eq:
                nomes_eq: List[str] = []
                for k in cand_cols:
                    n = _resolve_ajudante_primeiro_nome(cur, eq[k] if k in eq.keys() else None)
                    if n and n not in nomes_eq:
                        nomes_eq.append(n)
                if nomes_eq:
                    return " / ".join(nomes_eq)
        except Exception:
            pass

    return _resolve_ajudante_primeiro_nome(cur, equipe_raw)


def _decorate_rota_row(row: Dict[str, Any], cur: Optional[sqlite3.Cursor] = None) -> Dict[str, Any]:
    row["equipe_ajudantes"] = _format_equipe_ajudantes(row, cur)
    if row.get("equipe_ajudantes"):
        row["equipe"] = row["equipe_ajudantes"]

    def _first_non_empty(*keys: str):
        for k in keys:
            if k in row:
                v = row.get(k)
                if v is not None and str(v).strip() != "":
                    return v
        return None

    loc_rota = _first_non_empty(
        "local_rota",
        "tipo_rota",
        "local",
    )
    if loc_rota is not None:
        loc_rota_txt = str(loc_rota).strip()
        row["local_rota"] = loc_rota_txt
        row["tipo_rota"] = loc_rota_txt

    # Alias de local de carregamento (prioridade mobile).
    loc_car = _first_non_empty(
        "local_carregamento",
        "granja_carregada",
        "local_carregado",
        "local_carreg",
        "carregou_em",
    )
    if loc_car is None:
        loc_car = loc_rota
    if loc_car is not None:
        loc_car_txt = str(loc_car).strip()
        row["local_carregamento"] = loc_car_txt
        row["granja_carregada"] = loc_car_txt
        row["local_carregado"] = loc_car_txt
        row["local_carreg"] = loc_car_txt

    # Aliases esperados pelo desktop (compatibilidade retroativa).
    sd = _first_non_empty("saida_data", "data_saida")
    if sd is not None:
        row["saida_data"] = sd
    sh = _first_non_empty("saida_hora", "hora_saida")
    if sh is not None:
        row["saida_hora"] = sh
    fd = _first_non_empty("fim_data", "data_chegada", "data_fim")
    if fd is not None:
        row["fim_data"] = fd
    fh = _first_non_empty("fim_hora", "hora_chegada", "hora_fim")
    if fh is not None:
        row["fim_hora"] = fh

    cx = _first_non_empty("nf_caixas", "caixas_carregadas", "qnt_cx_carregada", "total_caixas")
    if cx is not None:
        row["nf_caixas"] = cx
    kg = _first_non_empty("nf_kg_carregado", "kg_carregado", "nf_kg")
    if kg is not None:
        row["nf_kg_carregado"] = kg
    md = _first_non_empty("media", "media_carregada")
    if md is not None:
        row["media"] = md
    cf = _first_non_empty("caixa_final", "aves_caixa_final", "qnt_aves_caixa_final")
    if cf is not None:
        row["caixa_final"] = cf

    tipo_estimativa = str(row.get("tipo_estimativa") or "KG").strip().upper()
    if tipo_estimativa not in ("KG", "CX"):
        tipo_estimativa = "KG"
    row["tipo_estimativa"] = tipo_estimativa
    row["unidade_estimativa"] = "CAIXAS" if tipo_estimativa == "CX" else "KG"
    operacao_tipo = str(row.get("operacao_tipo") or "").strip().upper().replace("-", "_").replace(" ", "_")
    if operacao_tipo not in ("TRANSBORDO", "VENDA"):
        operacao_tipo = "TRANSBORDO" if tipo_estimativa == "CX" else "VENDA"
    row["operacao_tipo"] = operacao_tipo
    row["tipo_operacao"] = "TRANSBORDO" if operacao_tipo == "TRANSBORDO" else ("EMPRESA_BUSCA" if tipo_estimativa == "CX" else "CIF")
    row["transbordo"] = operacao_tipo == "TRANSBORDO"
    transbordo_modalidade = str(row.get("transbordo_modalidade") or ("EMPRESA_BUSCA" if operacao_tipo == "TRANSBORDO" else "CIF")).strip().upper()
    row["transbordo_modalidade"] = "EMPRESA_BUSCA" if transbordo_modalidade == "FOB" else transbordo_modalidade
    row["transbordo_observacao"] = str(row.get("transbordo_observacao") or "").strip()
    row["transbordo_grupo"] = str(row.get("transbordo_grupo") or row.get("codigo_programacao") or "").strip().upper() if operacao_tipo == "TRANSBORDO" else ""
    if tipo_estimativa == "CX":
        try:
            row["estimativa_valor"] = int(float(row.get("caixas_estimado") or 0))
        except Exception:
            row["estimativa_valor"] = 0
    else:
        try:
            row["estimativa_valor"] = float(row.get("kg_estimado") or 0.0)
        except Exception:
            row["estimativa_valor"] = 0.0

    for key in (
        "codigo_programacao",
        "status",
        "motorista",
        "veiculo",
        "equipe",
        "local_rota",
        "local_carregamento",
        "data_criacao",
    ):
        row[key] = str(row.get(key) or "").strip()

    return row


def _status_operacional_normalizado(v: Any) -> str:
    return str(v or "").strip().upper().replace(" ", "_")


def _status_operacional_especial(row: Dict[str, Any], pend_substituicao: bool = False) -> Optional[str]:
    if pend_substituicao:
        return "EM_TRANSFERENCIA"
    st = _status_operacional_normalizado(row.get("status_operacional"))
    # Blindagem contra legado inconsistente:
    # status_operacional nao deve carregar estados finais de rota.
    if st in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
        return None
    if st in ("", "NORMAL", "OK", "SEM_PROBLEMA"):
        return None
    return st


def _norm_pedido_key(v: Any) -> str:
    s = (str(v or "")).strip()
    if not s:
        return ""
    s_num = s.replace(" ", "").replace(",", ".")
    try:
        f = float(s_num)
        if f.is_integer():
            return str(int(f))
        return ("%f" % f).rstrip("0").rstrip(".")
    except Exception:
        if s.endswith(".0"):
            base = s[:-2].strip()
            if base:
                return base
        return s


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None:
            return float(default)
        text = str(value).strip().replace(",", ".")
        if not text:
            return float(default)
        return float(text)
    except Exception:
        return float(default)


def _local_rota_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "local_rota"):
        candidates.append("NULLIF(TRIM(p.local_rota), '')")
    if col_exists(conn, "programacoes", "tipo_rota"):
        candidates.append("NULLIF(TRIM(p.tipo_rota), '')")
    if col_exists(conn, "programacoes", "local"):
        candidates.append("NULLIF(TRIM(p.local), '')")

    if not candidates:
        return "'-' AS local_rota"
    return f"COALESCE({', '.join(candidates)}, '-') AS local_rota"


def _local_carregamento_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "local_carregamento"):
        candidates.append("NULLIF(TRIM(p.local_carregamento), '')")
    if col_exists(conn, "programacoes", "granja_carregada"):
        candidates.append("NULLIF(TRIM(p.granja_carregada), '')")
    if col_exists(conn, "programacoes", "local_carregado"):
        candidates.append("NULLIF(TRIM(p.local_carregado), '')")
    if col_exists(conn, "programacoes", "local_carreg"):
        candidates.append("NULLIF(TRIM(p.local_carreg), '')")
    if col_exists(conn, "programacoes", "carregou_em"):
        candidates.append("NULLIF(TRIM(p.carregou_em), '')")
    if col_exists(conn, "programacoes", "local_rota"):
        candidates.append("NULLIF(TRIM(p.local_rota), '')")
    if col_exists(conn, "programacoes", "tipo_rota"):
        candidates.append("NULLIF(TRIM(p.tipo_rota), '')")
    if col_exists(conn, "programacoes", "local"):
        candidates.append("NULLIF(TRIM(p.local), '')")

    if not candidates:
        return "'-' AS local_carregamento"
    return f"COALESCE({', '.join(candidates)}, '-') AS local_carregamento"


def _media_carregada_expr(conn: sqlite3.Connection) -> str:
    if col_exists(conn, "programacoes", "media"):
        return "COALESCE(p.media, 0) AS media_carregada"
    return "0 AS media_carregada"


def _kg_carregado_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "kg_carregado"):
        candidates.append("p.kg_carregado")
    if col_exists(conn, "programacoes", "nf_kg_carregado"):
        candidates.append("p.nf_kg_carregado")
    if not candidates:
        return "0 AS kg_carregado"
    return f"COALESCE({', '.join(candidates)}, 0) AS kg_carregado"


def _caixas_carregadas_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "caixas_carregadas"):
        candidates.append("p.caixas_carregadas")
    if col_exists(conn, "programacoes", "qnt_cx_carregada"):
        candidates.append("p.qnt_cx_carregada")
    if col_exists(conn, "programacoes", "nf_caixas"):
        candidates.append("p.nf_caixas")
    if not candidates:
        return "0 AS caixas_carregadas"
    return f"COALESCE({', '.join(candidates)}, 0) AS caixas_carregadas"


def _caixa_final_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "aves_caixa_final"):
        candidates.append("p.aves_caixa_final")
    if col_exists(conn, "programacoes", "qnt_aves_caixa_final"):
        candidates.append("p.qnt_aves_caixa_final")
    if not candidates:
        return "0 AS caixa_final"
    return f"COALESCE({', '.join(candidates)}, 0) AS caixa_final"


def _caixas_saldo_subquery(conn: sqlite3.Connection, prog_alias: str = "p") -> str:
    cols_pi = set()
    cols_pc = set()
    cols_prog = set()
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {r[1] for r in cur.fetchall() or []}
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_pi = {r[1] for r in cur.fetchall() or []}
        cur.execute("PRAGMA table_info(programacao_itens_controle)")
        cols_pc = {r[1] for r in cur.fetchall() or []}
    except Exception:
        return "0 AS caixas_saldo"

    has_pi_pedido = "pedido" in cols_pi
    has_pc_pedido = "pedido" in cols_pc
    has_pi_status = "status_pedido" in cols_pi
    has_pc_status = "status_pedido" in cols_pc
    has_pi_caixas_atual = "caixas_atual" in cols_pi
    has_pc_caixas_atual = "caixas_atual" in cols_pc
    has_pi_qnt_caixas = "qnt_caixas" in cols_pi
    has_pi_transferencia = "transferencia_origem_id" in cols_pi
    has_pi_alt_tipo = "alteracao_tipo" in cols_pi
    has_pc_alt_tipo = "alteracao_tipo" in cols_pc

    join_on = "pc.codigo_programacao = pi.codigo_programacao AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))"
    if has_pi_pedido and has_pc_pedido:
        join_on += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"

    st_pi_expr = "COALESCE(NULLIF(TRIM(pi.status_pedido),''), 'PENDENTE')" if has_pi_status else "'PENDENTE'"
    st_pc_expr = "NULLIF(TRIM(pc.status_pedido),'')" if has_pc_status else "NULL"
    base_expr = "COALESCE(pi.qnt_caixas, 0)" if has_pi_qnt_caixas else "0"

    if has_pc_caixas_atual and has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = (
            "CASE "
            "WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual "
            "WHEN pi.caixas_atual IS NOT NULL AND pi.caixas_atual > 0 THEN pi.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pc_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = (
            "CASE "
            "WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = (
            "CASE "
            "WHEN pi.caixas_atual IS NOT NULL AND pi.caixas_atual > 0 THEN pi.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pc_caixas_atual:
        caixas_raw = "CASE WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual ELSE 0 END"
    else:
        caixas_raw = base_expr

    status_eff = f"COALESCE({st_pc_expr}, {st_pi_expr}, 'PENDENTE')"
    saldo_item = f"CASE WHEN UPPER({status_eff}) IN ('ENTREGUE','CANCELADO') THEN 0 ELSE COALESCE({caixas_raw},0) END"
    converted_filters: List[str] = []
    if has_pi_transferencia:
        converted_filters.append("TRIM(COALESCE(pi.transferencia_origem_id,''))<>''")
    if has_pi_alt_tipo:
        converted_filters.append("UPPER(TRIM(COALESCE(pi.alteracao_tipo,'')))='TRANSBORDO'")
    if has_pc_alt_tipo:
        converted_filters.append("UPPER(TRIM(COALESCE(pc.alteracao_tipo,'')))='TRANSBORDO'")
    converted_where = " OR ".join(converted_filters) if converted_filters else "0=1"
    converted_item_saldo_sql = f"""(
                    SELECT SUM({saldo_item})
                    FROM programacao_itens pi
                    LEFT JOIN programacao_itens_controle pc
                      ON {join_on}
                    WHERE pi.codigo_programacao = {prog_alias}.codigo_programacao
                      AND ({converted_where})
                )"""

    loaded_candidates: List[str] = []
    for col in ("caixas_carregadas", "qnt_cx_carregada"):
        if col in cols_prog:
            loaded_candidates.append(f"{prog_alias}.{col}")
    loaded_expr = f"COALESCE({', '.join(loaded_candidates)}, 0)" if loaded_candidates else "0"

    if table_exists(conn.cursor(), "transferencias"):
        accepted_in_sql = f"""(
                    SELECT COALESCE(SUM(MAX(COALESCE(t.qtd_caixas,0) - COALESCE(t.qtd_convertida,0), 0)), 0)
                    FROM transferencias t
                    WHERE UPPER(TRIM(COALESCE(t.codigo_destino,''))) = UPPER(TRIM({prog_alias}.codigo_programacao))
                      AND UPPER(TRIM(COALESCE(t.status,''))) = 'ACEITA'
                )"""
        active_out_sql = f"""(
                    SELECT COALESCE(SUM(MAX(COALESCE(t.qtd_caixas,0) - COALESCE(t.qtd_convertida,0), 0)), 0)
                    FROM transferencias t
                    WHERE UPPER(TRIM(COALESCE(t.codigo_origem,''))) = UPPER(TRIM({prog_alias}.codigo_programacao))
                      AND UPPER(TRIM(COALESCE(t.status,''))) IN ('PENDENTE','ACEITA')
                )"""
    else:
        accepted_in_sql = "0"
        active_out_sql = "0"

    saldo_real_sql = (
        f"MAX(COALESCE({loaded_expr},0) + COALESCE({accepted_in_sql},0) "
        f"+ COALESCE({converted_item_saldo_sql},0) - COALESCE({active_out_sql},0), 0)"
    )

    return f"""{saldo_real_sql} AS caixas_saldo"""


def _ultimo_km_final_veiculo(cur: sqlite3.Cursor, veiculo: str, *, exclude_programacao: str = "") -> Dict[str, Any]:
    v = str(veiculo or "").strip().upper()
    ex = str(exclude_programacao or "").strip().upper()
    if not v:
        return {"veiculo": "", "km_final": 0.0, "codigo_programacao": ""}
    try:
        cur.execute(
            """
            SELECT
                COALESCE(km_final, 0) AS km_final,
                COALESCE(codigo_programacao, '') AS codigo_programacao
            FROM programacoes
            WHERE UPPER(TRIM(COALESCE(veiculo,''))) = UPPER(TRIM(?))
              AND COALESCE(km_final, 0) > 0
              AND UPPER(TRIM(COALESCE(codigo_programacao,''))) <> UPPER(TRIM(?))
            ORDER BY COALESCE(data_chegada, data_saida, data_criacao, data, '') DESC, id DESC
            LIMIT 1
            """,
            (v, ex),
        )
        row = cur.fetchone()
    except Exception:
        logging.debug("Falha ao consultar ultimo KM final do veiculo %s", v, exc_info=True)
        row = None
    return {
        "veiculo": v,
        "km_final": float((row["km_final"] if row else 0.0) or 0.0),
        "codigo_programacao": str((row["codigo_programacao"] if row else "") or "").strip().upper(),
    }


def _attach_ultimo_km_veiculo(cur: sqlite3.Cursor, rota: Dict[str, Any]) -> Dict[str, Any]:
    veiculo = str((rota or {}).get("veiculo") or "").strip().upper()
    codigo = str((rota or {}).get("codigo_programacao") or "").strip().upper()
    ultimo = _ultimo_km_final_veiculo(cur, veiculo, exclude_programacao=codigo)
    km_atual = safe_float((rota or {}).get("km_inicial"), 0.0)
    km_sugerido = km_atual if km_atual > 0 else safe_float(ultimo.get("km_final"), 0.0)
    rota["ultimo_km_veiculo"] = safe_float(ultimo.get("km_final"), 0.0)
    rota["km_inicial_sugerido"] = safe_float(km_sugerido, 0.0)
    rota["ultimo_km_programacao"] = str(ultimo.get("codigo_programacao") or "")
    return rota


def _is_transbordo_row(row: Dict[str, Any]) -> bool:
    operacao = str(row.get("operacao_tipo") or "").strip().upper().replace("-", "_").replace(" ", "_")
    return operacao == "TRANSBORDO" or str(row.get("tipo_estimativa") or "").strip().upper() == "CX"


def _transferencias_resumo(cur: sqlite3.Cursor, codigo_programacao: str) -> Dict[str, int]:
    codigo = str(codigo_programacao or "").strip().upper()
    out = {"transferencias_saida": 0, "transferencias_entrada": 0, "transferencias_pendentes": 0}
    if not codigo:
        return out
    try:
        cur.execute(
            """
            SELECT codigo_origem, codigo_destino, qtd_caixas, qtd_convertida, status
            FROM transferencias
            WHERE UPPER(COALESCE(codigo_origem,''))=?
               OR UPPER(COALESCE(codigo_destino,''))=?
            """,
            (codigo, codigo),
        )
        for row in cur.fetchall() or []:
            status_value = str(row["status"] or "").strip().upper()
            if status_value in ("CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"):
                continue
            qtd = int(row["qtd_convertida"] or row["qtd_caixas"] or 0)
            if qtd <= 0:
                continue
            if status_value in ("", "PENDENTE", "ABERTA", "AGUARDANDO", "SOLICITADA"):
                out["transferencias_pendentes"] += qtd
            if str(row["codigo_origem"] or "").strip().upper() == codigo:
                out["transferencias_saida"] += qtd
            if str(row["codigo_destino"] or "").strip().upper() == codigo:
                out["transferencias_entrada"] += qtd
    except Exception:
        return out
    return out


def _rotas_not_finalizadas_clause(conn: sqlite3.Connection, alias: str = "p") -> str:
    """
    Filtro defensivo para não exibir rotas já encerradas.
    Cobre status legado inconsistente (ex.: status aberto com sinais de finalização).
    """
    parts = [
        f"UPPER(TRIM(COALESCE({alias}.status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')",
        f"UPPER(TRIM(COALESCE({alias}.status_operacional,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')",
    ]
    if col_exists(conn, "programacoes", "finalizada_no_app"):
        parts.append(f"COALESCE({alias}.finalizada_no_app,0)=0")
    if col_exists(conn, "programacoes", "data_chegada"):
        parts.append(f"TRIM(COALESCE({alias}.data_chegada,''))=''")
    if col_exists(conn, "programacoes", "hora_chegada"):
        parts.append(f"TRIM(COALESCE({alias}.hora_chegada,''))=''")
    if col_exists(conn, "programacoes", "km_final"):
        parts.append(f"COALESCE({alias}.km_final,0)=0")
    return " AND ".join(parts)


def _total_caixas_ativas_programacao(cur: sqlite3.Cursor, codigo_programacao: str) -> int:
    def _to_int_db(v: Any) -> int:
        try:
            if v is None:
                return 0
            if isinstance(v, (int, float)):
                return int(float(v))
            s = str(v).strip()
            if not s:
                return 0
            s = s.replace(" ", "")
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            return int(float(s))
        except Exception:
            return 0

    cur.execute("PRAGMA table_info(programacao_itens)")
    cols_pi = {row[1] for row in cur.fetchall() or []}
    cur.execute("PRAGMA table_info(programacao_itens_controle)")
    cols_pc = {row[1] for row in cur.fetchall() or []}

    has_pi_pedido = "pedido" in cols_pi
    has_pc_pedido = "pedido" in cols_pc
    has_pi_status = "status_pedido" in cols_pi
    has_pc_status = "status_pedido" in cols_pc
    has_pi_caixas_atual = "caixas_atual" in cols_pi
    has_pc_caixas_atual = "caixas_atual" in cols_pc
    has_pi_qnt_caixas = "qnt_caixas" in cols_pi

    join_on = "pc.codigo_programacao = pi.codigo_programacao AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))"
    if has_pi_pedido and has_pc_pedido:
        join_on += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"

    st_pi_expr = "COALESCE(NULLIF(TRIM(pi.status_pedido),''), 'PENDENTE')" if has_pi_status else "'PENDENTE'"
    st_pc_expr = "NULLIF(TRIM(pc.status_pedido),'')" if has_pc_status else "NULL"
    pedido_expr = "COALESCE(pi.pedido, '')" if has_pi_pedido else "''"
    nome_expr = "COALESCE(pi.nome_cliente, '')" if "nome_cliente" in cols_pi else "''"
    base_expr = "COALESCE(pi.qnt_caixas, 0)" if has_pi_qnt_caixas else "0"

    if has_pc_caixas_atual and has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_expr = (
            "CASE "
            "WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual "
            "WHEN pi.caixas_atual IS NOT NULL AND pi.caixas_atual > 0 THEN pi.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pc_caixas_atual and has_pi_qnt_caixas:
        caixas_expr = (
            "CASE "
            "WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_expr = (
            "CASE "
            "WHEN pi.caixas_atual IS NOT NULL AND pi.caixas_atual > 0 THEN pi.caixas_atual "
            "ELSE COALESCE(pi.qnt_caixas, 0) END"
        )
    elif has_pc_caixas_atual:
        caixas_expr = "CASE WHEN pc.caixas_atual IS NOT NULL AND pc.caixas_atual > 0 THEN pc.caixas_atual ELSE 0 END"
    else:
        caixas_expr = base_expr

    cur.execute(
        f"""
        SELECT
            COALESCE(pi.cod_cliente, '') AS cod_cliente,
            {nome_expr} AS nome_cliente,
            {pedido_expr} AS pedido,
            {base_expr} AS base_cx,
            {caixas_expr} AS atual_cx,
            COALESCE({st_pc_expr}, {st_pi_expr}, 'PENDENTE') AS status_eff
        FROM programacao_itens pi
        LEFT JOIN programacao_itens_controle pc
          ON {join_on}
        WHERE pi.codigo_programacao=?
        """,
        (codigo_programacao,),
    )

    total_em_aberto = 0
    for it in (cur.fetchall() or []):
        status_eff = str(it["status_eff"] or "PENDENTE").strip().upper()
        atual = _to_int_db(it["atual_cx"])
        if atual < 0:
            atual = 0
        if status_eff in ("ENTREGUE", "CANCELADO"):
            continue
        total_em_aberto += atual
    return total_em_aberto


def _equipe_cols_expr(conn: sqlite3.Connection, alias: str = "e") -> str:
    def col_or_null(col: str) -> str:
        if col_exists(conn, "equipes", col):
            return f"{alias}.{col} AS {col}"
        return f"NULL AS {col}"

    return ", ".join(
        [
            col_or_null("ajudante1"),
            col_or_null("ajudante2"),
            col_or_null("ajudante_1"),
            col_or_null("ajudante_2"),
        ]
    )


def _programacao_itens_select_expr(conn: sqlite3.Connection, alias: str = "pi") -> str:
    def col_or_null(col: str) -> str:
        if col_exists(conn, "programacao_itens", col):
            return f"{alias}.{col} AS {col}"
        return f"NULL AS {col}"

    return ", ".join(
        [
            col_or_null("cod_cliente"),
            col_or_null("nome_cliente"),
            col_or_null("qnt_caixas"),
            col_or_null("kg"),
            col_or_null("preco"),
            col_or_null("endereco"),
            col_or_null("vendedor"),
            col_or_null("pedido"),
            col_or_null("produto"),
            col_or_null("observacao"),
            col_or_null("status_pedido"),
            col_or_null("caixas_atual"),
            col_or_null("preco_atual"),
            col_or_null("ordem_sugerida"),
            col_or_null("eta"),
            col_or_null("distancia"),
            col_or_null("confianca_localizacao"),
            col_or_null("carga_raiz_programacao"),
            col_or_null("carga_origem_imediata"),
            col_or_null("transferencia_origem_id"),
        ]
    )


def now_date_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def now_time_str() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _haversine_m(lat1, lon1, lat2, lon2):
    # distance in meters
    r = 6371000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return r * c


def _gps_distance_last_minutes(cur, codigo_programacao: str, minutes: int = 15) -> float:
    cur.execute(
        """
        SELECT lat, lon, recorded_at
        FROM rota_gps_pings
        WHERE codigo_programacao=?
          AND recorded_at >= datetime('now', ?)
        ORDER BY recorded_at ASC
        """,
        (codigo_programacao, f"-{minutes} minutes"),
    )
    rows = cur.fetchall() or []
    if len(rows) < 2:
        return 0.0
    total = 0.0
    for i in range(1, len(rows)):
        a = rows[i - 1]
        b = rows[i]
        try:
            total += _haversine_m(float(a["lat"]), float(a["lon"]), float(b["lat"]), float(b["lon"]))
        except Exception:
            continue
    return total


def ensure_tables():
    """
    Cria tabelas auxiliares sem quebrar o banco existente.
    """
    with get_conn() as conn:
        cur = conn.cursor()
        ensure_core_schema(conn)

        # Controle por cliente (mortalidade/recebimento futuro etc.)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens_controle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT NOT NULL,
                cod_cliente TEXT NOT NULL,
                pedido TEXT DEFAULT NULL,

                mortalidade_aves INTEGER DEFAULT 0,
                media_aplicada REAL DEFAULT NULL,
                aves_por_caixa INTEGER DEFAULT NULL,
                peso_previsto REAL DEFAULT NULL,

                -- reservado para recebimentos (futuro)
                valor_recebido REAL DEFAULT NULL,
                forma_recebimento TEXT DEFAULT NULL,
                obs_recebimento TEXT DEFAULT NULL,

                -- status/alteracoes
                status_pedido TEXT DEFAULT NULL,
                alteracao_tipo TEXT DEFAULT NULL,
                alteracao_detalhe TEXT DEFAULT NULL,
                caixas_atual INTEGER DEFAULT NULL,
                preco_atual REAL DEFAULT NULL,
                alterado_em TEXT DEFAULT NULL,
                alterado_por TEXT DEFAULT NULL,
                lat_evento REAL DEFAULT NULL,
                lon_evento REAL DEFAULT NULL,
                endereco_evento TEXT DEFAULT NULL,
                cidade_evento TEXT DEFAULT NULL,
                bairro_evento TEXT DEFAULT NULL,

                updated_at TEXT DEFAULT (datetime('now')),

                UNIQUE(codigo_programacao, cod_cliente, pedido)
            )
        """)

        # Migra schema legado da tabela de controle:
        # antes a chave única era (codigo_programacao, cod_cliente), o que colide pedidos do mesmo cliente.
        try:
            cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='programacao_itens_controle'")
            row_tbl = cur.fetchone()
            sql_tbl = (row_tbl[0] if row_tbl and row_tbl[0] else "") or ""
            legacy_unique = (
                "UNIQUE(codigo_programacao, cod_cliente)" in sql_tbl
                and "UNIQUE(codigo_programacao, cod_cliente, pedido)" not in sql_tbl
            )
            if legacy_unique:
                cur.execute("ALTER TABLE programacao_itens_controle RENAME TO programacao_itens_controle_old")
                cur.execute("""
                    CREATE TABLE programacao_itens_controle (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        codigo_programacao TEXT NOT NULL,
                        cod_cliente TEXT NOT NULL,
                        pedido TEXT DEFAULT NULL,
                        mortalidade_aves INTEGER DEFAULT 0,
                        media_aplicada REAL DEFAULT NULL,
                        aves_por_caixa INTEGER DEFAULT NULL,
                        peso_previsto REAL DEFAULT NULL,
                        valor_recebido REAL DEFAULT NULL,
                        forma_recebimento TEXT DEFAULT NULL,
                        obs_recebimento TEXT DEFAULT NULL,
                        status_pedido TEXT DEFAULT NULL,
                        alteracao_tipo TEXT DEFAULT NULL,
                        alteracao_detalhe TEXT DEFAULT NULL,
                        caixas_atual INTEGER DEFAULT NULL,
                        preco_atual REAL DEFAULT NULL,
                        alterado_em TEXT DEFAULT NULL,
                        alterado_por TEXT DEFAULT NULL,
                        lat_evento REAL DEFAULT NULL,
                        lon_evento REAL DEFAULT NULL,
                        endereco_evento TEXT DEFAULT NULL,
                        cidade_evento TEXT DEFAULT NULL,
                        bairro_evento TEXT DEFAULT NULL,
                        updated_at TEXT DEFAULT (datetime('now')),
                        UNIQUE(codigo_programacao, cod_cliente, pedido)
                    )
                """)
                cur.execute("""
                    INSERT INTO programacao_itens_controle
                        (codigo_programacao, cod_cliente, pedido,
                         mortalidade_aves, media_aplicada, aves_por_caixa, peso_previsto,
                         valor_recebido, forma_recebimento, obs_recebimento,
                         status_pedido, alteracao_tipo, alteracao_detalhe,
                         caixas_atual, preco_atual, alterado_em, alterado_por,
                         lat_evento, lon_evento, endereco_evento, cidade_evento, bairro_evento, updated_at)
                    SELECT
                        codigo_programacao,
                        cod_cliente,
                        COALESCE(pedido, ''),
                        COALESCE(mortalidade_aves, 0),
                        media_aplicada,
                        aves_por_caixa,
                        peso_previsto,
                        valor_recebido,
                        forma_recebimento,
                        obs_recebimento,
                        status_pedido,
                        alteracao_tipo,
                        alteracao_detalhe,
                        caixas_atual,
                        preco_atual,
                        alterado_em,
                        alterado_por,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        COALESCE(updated_at, datetime('now'))
                    FROM programacao_itens_controle_old
                """)
                cur.execute("DROP TABLE programacao_itens_controle_old")
        except Exception:
            pass

        # log de sincronizacao
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT NOT NULL,
                cod_cliente TEXT NOT NULL,
                evento TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        for col_name, col_ddl in {
            "pedido": "pedido TEXT",
            "evento": "evento TEXT DEFAULT 'cliente_controle'",
            "payload_json": "payload_json TEXT DEFAULT '{}'",
            "registrado_em": "registrado_em TEXT",
            "created_at": "created_at TEXT",
            "company_id": "company_id INTEGER",
        }.items():
            try:
                cur.execute("PRAGMA table_info(programacao_itens_log)")
                log_cols = {str(row[1]).lower() for row in cur.fetchall() or []}
                if col_name not in log_cols:
                    cur.execute(f"ALTER TABLE programacao_itens_log ADD COLUMN {col_ddl}")
            except Exception:
                pass
        try:
            cur.execute(
                """
                UPDATE programacao_itens_log
                   SET evento=COALESCE(NULLIF(TRIM(evento), ''), 'cliente_controle'),
                       created_at=COALESCE(NULLIF(TRIM(created_at), ''), NULLIF(TRIM(registrado_em), ''), datetime('now'))
                 WHERE evento IS NULL OR TRIM(evento)=''
                    OR created_at IS NULL OR TRIM(created_at)=''
                """
            )
        except Exception:
            pass

        # gps pings da rota (rastreamento)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS rota_gps_pings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT NOT NULL,
                motorista TEXT NOT NULL,
                lat REAL NOT NULL,
                lon REAL NOT NULL,
                speed REAL DEFAULT NULL,
                accuracy REAL DEFAULT NULL,
                recorded_at TEXT DEFAULT (datetime('now')),
                company_id INTEGER
            )
        """)
        cur.execute("PRAGMA table_info(rota_gps_pings)")
        cols_gps = {row[1] for row in cur.fetchall() or []}
        if "company_id" not in cols_gps:
            cur.execute("ALTER TABLE rota_gps_pings ADD COLUMN company_id INTEGER")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS rota_fotos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_foto TEXT UNIQUE,
                codigo_programacao TEXT NOT NULL,
                categoria TEXT NOT NULL,
                tipo_registro TEXT NOT NULL,
                cod_cliente TEXT,
                cliente_nome TEXT,
                pedido TEXT,
                id_vinculo TEXT,
                path_local TEXT,
                storage_path TEXT,
                arquivo_nome TEXT,
                mime_type TEXT,
                tamanho_bytes INTEGER DEFAULT 0,
                motorista_codigo TEXT,
                motorista_nome TEXT,
                registrado_em TEXT DEFAULT (datetime('now')),
                payload_json TEXT DEFAULT '{}',
                company_id INTEGER
            )
        """)
        cur.execute("PRAGMA table_info(rota_fotos)")
        cols_fotos = {row[1] for row in cur.fetchall() or []}
        if "company_id" not in cols_fotos:
            cur.execute("ALTER TABLE rota_fotos ADD COLUMN company_id INTEGER")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rota_fotos_prog_categoria ON rota_fotos(codigo_programacao, categoria)")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS roteiro_operacional (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo_evento TEXT NOT NULL,
                codigo_programacao TEXT NOT NULL,
                origem TEXT,
                destino TEXT,
                motorista_codigo TEXT,
                motorista_nome TEXT,
                pedido TEXT,
                cod_cliente TEXT,
                cliente_nome TEXT,
                caixas INTEGER DEFAULT 0,
                kg REAL DEFAULT 0,
                media REAL DEFAULT 0,
                aves_por_caixa INTEGER DEFAULT 0,
                nf_numero TEXT,
                nf_preco REAL DEFAULT 0,
                lotes TEXT,
                data_hora TEXT,
                observacao TEXT,
                payload_json TEXT DEFAULT '{}',
                created_at TEXT DEFAULT (datetime('now')),
                company_id INTEGER
            )
        """)
        for col_name, col_ddl in {
            "tipo_evento": "tipo_evento TEXT",
            "codigo_programacao": "codigo_programacao TEXT",
            "origem": "origem TEXT",
            "destino": "destino TEXT",
            "motorista_codigo": "motorista_codigo TEXT",
            "motorista_nome": "motorista_nome TEXT",
            "pedido": "pedido TEXT",
            "cod_cliente": "cod_cliente TEXT",
            "cliente_nome": "cliente_nome TEXT",
            "caixas": "caixas INTEGER DEFAULT 0",
            "kg": "kg REAL DEFAULT 0",
            "media": "media REAL DEFAULT 0",
            "aves_por_caixa": "aves_por_caixa INTEGER DEFAULT 0",
            "nf_numero": "nf_numero TEXT",
            "nf_preco": "nf_preco REAL DEFAULT 0",
            "lotes": "lotes TEXT",
            "data_hora": "data_hora TEXT",
            "observacao": "observacao TEXT",
            "payload_json": "payload_json TEXT DEFAULT '{}'",
            "created_at": "created_at TEXT",
            "company_id": "company_id INTEGER",
        }.items():
            try:
                cur.execute("PRAGMA table_info(roteiro_operacional)")
                roteiro_cols = {str(row[1]).lower() for row in cur.fetchall() or []}
                if col_name not in roteiro_cols:
                    cur.execute(f"ALTER TABLE roteiro_operacional ADD COLUMN {col_ddl}")
            except Exception:
                pass
        cur.execute("CREATE INDEX IF NOT EXISTS idx_roteiro_operacional_prog_data ON roteiro_operacional(codigo_programacao, data_hora, id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_roteiro_operacional_tipo ON roteiro_operacional(tipo_evento)")

        # log de override manual de inicio
        cur.execute("""
            CREATE TABLE IF NOT EXISTS rota_gps_override_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT NOT NULL,
                motorista TEXT NOT NULL,
                motivo TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)

        # garante colunas extras em programacao_itens (para status/alteracoes)
        try:
            cur.execute("PRAGMA table_info(programacao_itens)")
            cols = {row[1] for row in cur.fetchall()}

            def add_col(name: str, ddl: str):
                if name not in cols:
                    cur.execute(f"ALTER TABLE programacao_itens ADD COLUMN {ddl}")

            add_col("observacao", "observacao TEXT")
            add_col("status_pedido", "status_pedido TEXT")
            add_col("alteracao_tipo", "alteracao_tipo TEXT")
            add_col("alteracao_detalhe", "alteracao_detalhe TEXT")
            add_col("caixas_atual", "caixas_atual INTEGER")
            add_col("preco_atual", "preco_atual REAL")
            add_col("alterado_em", "alterado_em TEXT")
            add_col("alterado_por", "alterado_por TEXT")
            add_col("ordem_sugerida", "ordem_sugerida INTEGER")
            add_col("eta", "eta TEXT")
            add_col("distancia", "distancia REAL")
            add_col("confianca_localizacao", "confianca_localizacao REAL")
            add_col("carga_raiz_programacao", "carga_raiz_programacao TEXT")
            add_col("carga_origem_imediata", "carga_origem_imediata TEXT")
            add_col("transferencia_origem_id", "transferencia_origem_id TEXT")
            if table_exists(cur, "transferencias") and table_exists(cur, "transferencias_conversoes"):
                cur.execute(
                    """
                    SELECT
                        t.id,
                        t.codigo_origem,
                        t.codigo_destino,
                        t.snapshot,
                        c.pedido_destino,
                        c.cod_cliente_destino
                    FROM transferencias_conversoes c
                    JOIN transferencias t ON t.id=c.transferencia_id
                    WHERE TRIM(COALESCE(c.pedido_destino, ''))<>''
                      AND TRIM(COALESCE(c.cod_cliente_destino, ''))<>''
                    LIMIT 5000
                    """
                )
                for row_conv in cur.fetchall() or []:
                    snapshot = _parse_snapshot(row_conv["snapshot"])
                    raiz = str(
                        snapshot.get("carga_raiz_programacao")
                        or snapshot.get("carga_origem_programacao")
                        or row_conv["codigo_origem"]
                        or ""
                    ).strip().upper()
                    if not raiz:
                        continue
                    cur.execute(
                        """
                        UPDATE programacao_itens
                           SET carga_raiz_programacao=COALESCE(NULLIF(carga_raiz_programacao, ''), ?),
                               carga_origem_imediata=COALESCE(NULLIF(carga_origem_imediata, ''), ?),
                               transferencia_origem_id=COALESCE(NULLIF(transferencia_origem_id, ''), ?)
                         WHERE codigo_programacao=?
                           AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))
                           AND COALESCE(TRIM(COALESCE(pedido,'')),'')=COALESCE(TRIM(?),'')
                        """,
                        (
                            raiz,
                            str(row_conv["codigo_origem"] or "").strip().upper(),
                            str(row_conv["id"] or "").strip(),
                            str(row_conv["codigo_destino"] or "").strip(),
                            str(row_conv["cod_cliente_destino"] or "").strip(),
                            str(row_conv["pedido_destino"] or "").strip(),
                        ),
                    )
        except Exception:
            pass

        # garante colunas extras em programacao_itens_controle (migracao)
        try:
            cur.execute("PRAGMA table_info(programacao_itens_controle)")
            cols = {row[1] for row in cur.fetchall()}

            def add_ctrl_col(name: str, ddl: str):
                if name not in cols:
                    cur.execute(f"ALTER TABLE programacao_itens_controle ADD COLUMN {ddl}")

            add_ctrl_col("mortalidade_aves", "mortalidade_aves INTEGER DEFAULT 0")
            add_ctrl_col("media_aplicada", "media_aplicada REAL")
            add_ctrl_col("aves_por_caixa", "aves_por_caixa INTEGER")
            add_ctrl_col("peso_previsto", "peso_previsto REAL")
            add_ctrl_col("valor_recebido", "valor_recebido REAL")
            add_ctrl_col("forma_recebimento", "forma_recebimento TEXT")
            add_ctrl_col("obs_recebimento", "obs_recebimento TEXT")
            add_ctrl_col("status_pedido", "status_pedido TEXT")
            add_ctrl_col("alteracao_tipo", "alteracao_tipo TEXT")
            add_ctrl_col("alteracao_detalhe", "alteracao_detalhe TEXT")
            add_ctrl_col("pedido", "pedido TEXT")
            add_ctrl_col("caixas_atual", "caixas_atual INTEGER")
            add_ctrl_col("preco_atual", "preco_atual REAL")
            add_ctrl_col("alterado_em", "alterado_em TEXT")
            add_ctrl_col("alterado_por", "alterado_por TEXT")
            add_ctrl_col("lat_evento", "lat_evento REAL")
            add_ctrl_col("lon_evento", "lon_evento REAL")
            add_ctrl_col("lat_entrega", "lat_entrega REAL")
            add_ctrl_col("lon_entrega", "lon_entrega REAL")
            add_ctrl_col("accuracy_entrega", "accuracy_entrega REAL")
            add_ctrl_col("timestamp_entrega", "timestamp_entrega TEXT")
            add_ctrl_col("endereco_evento", "endereco_evento TEXT")
            add_ctrl_col("cidade_evento", "cidade_evento TEXT")
            add_ctrl_col("bairro_evento", "bairro_evento TEXT")
            add_ctrl_col("foto_mortalidade_path", "foto_mortalidade_path TEXT")
            add_ctrl_col("mortalidade_foto_path", "mortalidade_foto_path TEXT")
            add_ctrl_col("foto_mortalidade_ref_json", "foto_mortalidade_ref_json TEXT")
            add_ctrl_col("ordem_sugerida", "ordem_sugerida INTEGER")
            add_ctrl_col("eta", "eta TEXT")
            add_ctrl_col("distancia", "distancia REAL")
            add_ctrl_col("confianca_localizacao", "confianca_localizacao REAL")
            add_ctrl_col("updated_at", "updated_at TEXT")
            add_ctrl_col("company_id", "company_id INTEGER")
        except Exception:
            pass

        try:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cliente_localizacao_amostras (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cod_cliente TEXT NOT NULL,
                    codigo_programacao TEXT,
                    pedido TEXT,
                    latitude REAL,
                    longitude REAL,
                    endereco TEXT,
                    cidade TEXT,
                    bairro TEXT,
                    status_pedido TEXT,
                    motorista_codigo TEXT,
                    motorista_nome TEXT,
                    origem TEXT DEFAULT 'APP',
                    registrado_em TEXT DEFAULT (datetime('now')),
                    company_id INTEGER
                )
            """)
            cur.execute("PRAGMA table_info(cliente_localizacao_amostras)")
            cols_cli_loc = {row[1] for row in cur.fetchall() or []}
            if "company_id" not in cols_cli_loc:
                cur.execute("ALTER TABLE cliente_localizacao_amostras ADD COLUMN company_id INTEGER")
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_cli_loc_amostras_cliente ON cliente_localizacao_amostras(cod_cliente, registrado_em DESC)"
            )
        except Exception:
            pass

        # garante coluna pedido em recebimentos (auditoria por pedido)
        try:
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
            if cur.fetchone() is not None:
                cur.execute("PRAGMA table_info(recebimentos)")
                cols_rec = {row[1] for row in cur.fetchall() or []}
                if "pedido" not in cols_rec:
                    cur.execute("ALTER TABLE recebimentos ADD COLUMN pedido TEXT")
                if "company_id" not in cols_rec:
                    cur.execute("ALTER TABLE recebimentos ADD COLUMN company_id INTEGER")
        except Exception:
            pass

        # garante colunas de carregamento em programacoes (migracao legado)
        try:
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {row[1] for row in cur.fetchall()}

            def add_prog_col(name: str, ddl: str):
                if name not in cols:
                    cur.execute(f"ALTER TABLE programacoes ADD COLUMN {ddl}")

            add_prog_col("aves_caixa_final", "aves_caixa_final INTEGER")
            add_prog_col("qnt_aves_caixa_final", "qnt_aves_caixa_final INTEGER")
            add_prog_col("media_1", "media_1 REAL")
            add_prog_col("media_2", "media_2 REAL")
            add_prog_col("media_3", "media_3 REAL")
            add_prog_col("carregamento_fechado", "carregamento_fechado INTEGER DEFAULT 0")
            add_prog_col("carregamento_salvo_em", "carregamento_salvo_em TEXT")
            add_prog_col("motorista_id", "motorista_id INTEGER")
            add_prog_col("motorista_codigo", "motorista_codigo TEXT")
            add_prog_col("codigo_motorista", "codigo_motorista TEXT")
            # CIF/EMPRESA BUSCA + auditoria (compatível com desktop)
            add_prog_col("tipo_estimativa", "tipo_estimativa TEXT DEFAULT 'KG'")
            add_prog_col("caixas_estimado", "caixas_estimado INTEGER DEFAULT 0")
            add_prog_col("operacao_tipo", "operacao_tipo TEXT DEFAULT 'VENDA'")
            add_prog_col("transbordo_modalidade", "transbordo_modalidade TEXT")
            add_prog_col("transbordo_observacao", "transbordo_observacao TEXT")
            add_prog_col("transbordo_grupo", "transbordo_grupo TEXT")
            add_prog_col("local_rota", "local_rota TEXT")
            add_prog_col("tipo_rota", "tipo_rota TEXT")
            add_prog_col("local_carregamento", "local_carregamento TEXT")
            add_prog_col("granja_carregada", "granja_carregada TEXT")
            add_prog_col("local_carregado", "local_carregado TEXT")
            add_prog_col("local_carreg", "local_carreg TEXT")
            add_prog_col("adiantamento_origem", "adiantamento_origem TEXT")
            add_prog_col("pix_motorista", "pix_motorista REAL DEFAULT 0")
            add_prog_col("usuario_criacao", "usuario_criacao TEXT")
            add_prog_col("usuario_ultima_edicao", "usuario_ultima_edicao TEXT")
            add_prog_col("status_operacional", "status_operacional TEXT")
            add_prog_col("status_operacional_obs", "status_operacional_obs TEXT")
            add_prog_col("status_operacional_em", "status_operacional_em TEXT")
            add_prog_col("status_operacional_por", "status_operacional_por TEXT")
            add_prog_col("foto_doa_path", "foto_doa_path TEXT")
            add_prog_col("doa_foto_path", "doa_foto_path TEXT")
            add_prog_col("mortalidade_transbordo_foto_path", "mortalidade_transbordo_foto_path TEXT")
            add_prog_col("foto_doa_ref_json", "foto_doa_ref_json TEXT")
            add_prog_col("ajudantes_alteracao_motivo", "ajudantes_alteracao_motivo TEXT")
            add_prog_col("ajudantes_alterado_em", "ajudantes_alterado_em TEXT")
            add_prog_col("historico_ajudantes", "historico_ajudantes TEXT")
        except Exception:
            pass

        try:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS despesas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    descricao TEXT,
                    valor REAL DEFAULT 0,
                    data_registro TEXT,
                    tipo_despesa TEXT DEFAULT 'ROTA',
                    categoria TEXT,
                    motorista TEXT,
                    veiculo TEXT,
                    observacao TEXT
                )
            """)
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
            if cur.fetchone() is not None:
                cur.execute("PRAGMA table_info(despesas)")
                cols_desp = {row[1] for row in cur.fetchall() or []}

                def add_desp_col(name: str, ddl: str):
                    if name not in cols_desp:
                        cur.execute(f"ALTER TABLE despesas ADD COLUMN {ddl}")

                add_desp_col("id_local", "id_local TEXT")
                add_desp_col("forma_pagamento", "forma_pagamento TEXT")
                add_desp_col("comprovante_path", "comprovante_path TEXT")
                add_desp_col("estabelecimento", "estabelecimento TEXT")
                add_desp_col("documento", "documento TEXT")
                add_desp_col("litros", "litros REAL")
                add_desp_col("valor_litro", "valor_litro REAL")
                add_desp_col("desconto", "desconto REAL")
                add_desp_col("combustivel", "combustivel TEXT")
                add_desp_col("odometro", "odometro REAL")
                add_desp_col("lat", "lat REAL")
                add_desp_col("lon", "lon REAL")
                add_desp_col("accuracy", "accuracy REAL")
                add_desp_col("registrado_em", "registrado_em TEXT")
                add_desp_col("motorista_codigo", "motorista_codigo TEXT")
                add_desp_col("motorista_nome", "motorista_nome TEXT")
                add_desp_col("sync_key", "sync_key TEXT")
                add_desp_col("status_sync", "status_sync TEXT")
                add_desp_col("origem", "origem TEXT")
                add_desp_col("vinculo_prestacao_json", "vinculo_prestacao_json TEXT")
                add_desp_col("desktop_web_json", "desktop_web_json TEXT")
                add_desp_col("foto_despesa_ref_json", "foto_despesa_ref_json TEXT")
                add_desp_col("company_id", "company_id INTEGER")
                cur.execute(
                    "CREATE INDEX IF NOT EXISTS idx_despesas_prog_id_local ON despesas(codigo_programacao, id_local)"
                )
        except Exception:
            pass

        # controle de acesso ao app por motorista (desbloqueio admin)
        try:
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='motoristas'")
            if cur.fetchone() is not None:
                cur.execute("PRAGMA table_info(motoristas)")
                cols_m = {row[1] for row in cur.fetchall() or []}

                existing_ids: List[int] = []
                if "acesso_liberado" not in cols_m:
                    cur.execute("SELECT id FROM motoristas")
                    existing_ids = [int(r[0]) for r in (cur.fetchall() or []) if r[0] is not None]
                    # Novos cadastros devem entrar aptos ao app por padrão.
                    cur.execute("ALTER TABLE motoristas ADD COLUMN acesso_liberado INTEGER DEFAULT 1")

                if "acesso_liberado_por" not in cols_m:
                    cur.execute("ALTER TABLE motoristas ADD COLUMN acesso_liberado_por TEXT")
                if "acesso_liberado_em" not in cols_m:
                    cur.execute("ALTER TABLE motoristas ADD COLUMN acesso_liberado_em TEXT")
                if "acesso_obs" not in cols_m:
                    cur.execute("ALTER TABLE motoristas ADD COLUMN acesso_obs TEXT")

                # Nao derruba operacao atual: cadastros existentes ficam liberados.
                if existing_ids:
                    qmarks = ",".join(["?"] * len(existing_ids))
                    cur.execute(
                        f"UPDATE motoristas SET acesso_liberado=1 WHERE id IN ({qmarks})",
                        tuple(existing_ids),
                    )
        except Exception:
            pass

        # cadastro/login do app vendedor
        try:
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendedores'")
            if cur.fetchone() is not None:
                cur.execute("PRAGMA table_info(vendedores)")
                cols_v = {row[1] for row in cur.fetchall() or []}

                if "codigo" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN codigo TEXT")
                if "nome" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN nome TEXT")
                if "telefone" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN telefone TEXT")
                if "cidade_base" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN cidade_base TEXT")
                if "status" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN status TEXT DEFAULT 'ATIVO'")
                if "senha" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN senha TEXT")
                if "ultimo_login_em" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN ultimo_login_em TEXT")
                if "ultimo_login_ip" not in cols_v:
                    cur.execute("ALTER TABLE vendedores ADD COLUMN ultimo_login_ip TEXT")

                cur.execute(
                    """
                    UPDATE vendedores
                    SET status='ATIVO'
                    WHERE status IS NULL OR TRIM(status)=''
                    """
                )
                cur.execute(
                    "CREATE UNIQUE INDEX IF NOT EXISTS idx_vendedores_codigo ON vendedores(codigo)"
                )
        except Exception:
            pass

        # transferencias em banco (substituindo o JSON antigo)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transferencias (
                id TEXT PRIMARY KEY,
                codigo_origem TEXT NOT NULL,
                codigo_destino TEXT NOT NULL,
                cod_cliente TEXT NOT NULL,
                pedido TEXT NOT NULL,
                qtd_caixas INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'PENDENTE',
                obs TEXT DEFAULT NULL,
                snapshot TEXT DEFAULT NULL,
                motorista_origem TEXT DEFAULT NULL,
                motorista_destino TEXT DEFAULT NULL,
                qtd_convertida INTEGER DEFAULT 0,
                criado_em TEXT DEFAULT (datetime('now')),
                atualizado_em TEXT DEFAULT (datetime('now'))
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transferencias_conversoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transferencia_id TEXT NOT NULL,
                pedido_destino TEXT,
                cod_cliente_destino TEXT,
                qtd INTEGER NOT NULL,
                obs TEXT,
                nome_cliente_destino TEXT,
                novo_cliente INTEGER DEFAULT 0,
                criado_em TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(transferencia_id) REFERENCES transferencias(id) ON DELETE CASCADE
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS mobile_sync_idempotency (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                motorista_codigo TEXT NOT NULL,
                codigo_programacao TEXT NOT NULL,
                endpoint TEXT NOT NULL,
                idem_key TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now')),
                UNIQUE(motorista_codigo, codigo_programacao, endpoint, idem_key)
            )
        """)

        # substituicoes de motorista/veiculo em rota (handover)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS rota_substituicoes (
                id TEXT PRIMARY KEY,
                codigo_programacao TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'PENDENTE_ACEITE',
                motivo TEXT NOT NULL,
                km_evento INTEGER DEFAULT NULL,
                lat_evento REAL DEFAULT NULL,
                lon_evento REAL DEFAULT NULL,
                snapshot_json TEXT DEFAULT NULL,
                origem_motorista_nome TEXT DEFAULT NULL,
                origem_motorista_codigo TEXT DEFAULT NULL,
                origem_motorista_id INTEGER DEFAULT NULL,
                origem_veiculo TEXT DEFAULT NULL,
                destino_motorista_nome TEXT DEFAULT NULL,
                destino_motorista_codigo TEXT DEFAULT NULL,
                destino_motorista_id INTEGER DEFAULT NULL,
                destino_veiculo TEXT DEFAULT NULL,
                solicitado_em TEXT DEFAULT (datetime('now')),
                aceito_em TEXT DEFAULT NULL,
                atualizado_em TEXT DEFAULT (datetime('now'))
            )
        """)

        # programacao avulsa (vendedor) para uso em fim de semana
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacoes_avulsas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_avulsa TEXT NOT NULL UNIQUE,
                data_programada TEXT DEFAULT NULL,
                status TEXT NOT NULL DEFAULT 'AVULSA_ATIVA',
                motorista_id INTEGER DEFAULT NULL,
                motorista_codigo TEXT DEFAULT NULL,
                motorista_nome TEXT DEFAULT NULL,
                veiculo TEXT DEFAULT NULL,
                equipe TEXT DEFAULT NULL,
                local_rota TEXT DEFAULT NULL,
                observacao TEXT DEFAULT NULL,
                origem TEXT DEFAULT 'APP_VENDEDOR',
                criado_por TEXT DEFAULT NULL,
                criado_em TEXT DEFAULT (datetime('now')),
                conciliada_em TEXT DEFAULT NULL,
                programacao_oficial_codigo TEXT DEFAULT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacoes_avulsas_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                avulsa_id INTEGER NOT NULL,
                cod_cliente TEXT NOT NULL,
                nome_cliente TEXT NOT NULL,
                endereco TEXT DEFAULT NULL,
                cidade TEXT DEFAULT NULL,
                bairro TEXT DEFAULT NULL,
                ordem INTEGER DEFAULT 0,
                status_item TEXT DEFAULT 'PENDENTE',
                pedido TEXT DEFAULT NULL,
                nf TEXT DEFAULT NULL,
                caixas INTEGER DEFAULT 0,
                preco REAL DEFAULT NULL,
                observacao TEXT DEFAULT NULL,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(avulsa_id) REFERENCES programacoes_avulsas(id) ON DELETE CASCADE
            )
        """)
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_avulsas_status_data ON programacoes_avulsas(status, data_programada)"
        )
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_avulsas_itens_avulsa ON programacoes_avulsas_itens(avulsa_id, ordem, id)"
        )

        conn.commit()

def reconcile_transferencias_status() -> int:
    """
    Recalcula status do pedido de origem com base em transferencias ativas
    (PENDENTE/ACEITA):
    - saldo > 0  -> ALTERADO
    - saldo == 0 -> CANCELADO
    """
    fixed = 0
    now = datetime.now().isoformat(timespec="seconds")

    def _norm_pedido(v: Any) -> str:
        s = str(v or "").strip()
        if not s:
            return ""
        try:
            f = float(s.replace(",", "."))
            if abs(f - int(f)) < 1e-9:
                return str(int(f))
        except Exception:
            pass
        return s.upper()

    with get_conn() as conn:
        cur = conn.cursor()

        cur.execute(
            """
            SELECT codigo_origem, cod_cliente, pedido, COALESCE(SUM(qtd_caixas), 0) AS qtd
            FROM transferencias
            WHERE UPPER(TRIM(COALESCE(status, ''))) IN ('PENDENTE', 'ACEITA')
            GROUP BY codigo_origem, cod_cliente, pedido
            """
        )
        grupos = cur.fetchall() or []
        if not grupos:
            return 0

        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_itens = {row[1] for row in (cur.fetchall() or [])}
        has_pedido_col = "pedido" in cols_itens

        for g in grupos:
            codigo = str(g["codigo_origem"] or "").strip()
            cod_cli = str(g["cod_cliente"] or "").strip()
            pedido_src = g["pedido"]
            pedido_norm = _norm_pedido(pedido_src)
            qtd_transferida = int(g["qtd"] or 0)
            if not codigo or not cod_cli or qtd_transferida <= 0:
                continue

            if has_pedido_col:
                cur.execute(
                    """
                    SELECT rowid AS rid, codigo_programacao, cod_cliente, pedido, qnt_caixas
                    FROM programacao_itens
                    WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
                    """,
                    (codigo, cod_cli),
                )
                cands = cur.fetchall() or []
                base = None
                for r in cands:
                    if _norm_pedido(r["pedido"]) == pedido_norm:
                        base = r
                        break
            else:
                cur.execute(
                    """
                    SELECT rowid AS rid, codigo_programacao, cod_cliente, NULL AS pedido, qnt_caixas
                    FROM programacao_itens
                    WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
                    LIMIT 1
                    """,
                    (codigo, cod_cli),
                )
                base = cur.fetchone()

            if not base:
                continue

            base_qtd = int(base["qnt_caixas"] or 0)
            novo_caixas = max(base_qtd - qtd_transferida, 0)
            novo_status = "CANCELADO" if novo_caixas == 0 else "ALTERADO"
            detalhe = f"Transferencia de caixas (reconciliado): -{qtd_transferida} cx"

            sets = []
            params = []
            if "status_pedido" in cols_itens:
                sets.append("status_pedido=?")
                params.append(novo_status)
            if "alteracao_tipo" in cols_itens:
                sets.append("alteracao_tipo=?")
                params.append("QUANTIDADE")
            if "alteracao_detalhe" in cols_itens:
                sets.append("alteracao_detalhe=?")
                params.append(detalhe)
            if "caixas_atual" in cols_itens:
                sets.append("caixas_atual=?")
                params.append(novo_caixas)
            if "alterado_em" in cols_itens:
                sets.append("alterado_em=?")
                params.append(now)
            if "alterado_por" in cols_itens:
                sets.append("alterado_por=?")
                params.append("SISTEMA")

            if sets:
                params.append(base["rid"])
                cur.execute(f"UPDATE programacao_itens SET {', '.join(sets)} WHERE rowid=?", tuple(params))

            pedido_db = base["pedido"] if has_pedido_col else None
            cur.execute(
                """
                UPDATE programacao_itens_controle
                   SET status_pedido=?,
                       alteracao_tipo='QUANTIDADE',
                       alteracao_detalhe=?,
                       caixas_atual=?,
                       alterado_em=?,
                       alterado_por='SISTEMA',
                       updated_at=datetime('now')
                 WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
                   AND COALESCE(TRIM(pedido), '')=COALESCE(TRIM(?), '')
                """,
                (novo_status, detalhe, novo_caixas, now, codigo, cod_cli, pedido_db),
            )
            if cur.rowcount == 0:
                cur.execute(
                    """
                    INSERT INTO programacao_itens_controle
                        (codigo_programacao, cod_cliente, pedido, status_pedido,
                         alteracao_tipo, alteracao_detalhe, caixas_atual,
                         alterado_em, alterado_por, updated_at)
                    VALUES (?, ?, ?, ?, 'QUANTIDADE', ?, ?, ?, 'SISTEMA', datetime('now'))
                    """,
                    (codigo, cod_cli, pedido_db, novo_status, detalhe, novo_caixas, now),
                )

            fixed += 1

        conn.commit()
    return fixed


def sanitize_status_operacional_legado() -> int:
    """
    Limpa status_operacional terminal herdado em rotas que ainda estao ativas.
    Evita cenário: status=ATIVA e status_operacional=FINALIZADA.
    """
    fixed = 0
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if "status_operacional" not in cols:
            return 0

        sets = [
            "status_operacional=NULL",
        ]
        if "status_operacional_obs" in cols:
            sets.append("status_operacional_obs=NULL")
        if "status_operacional_em" in cols:
            sets.append("status_operacional_em=NULL")
        if "status_operacional_por" in cols:
            sets.append("status_operacional_por=NULL")
        if "finalizada_no_app" in cols:
            sets.append("finalizada_no_app=0")

        sql = f"""
            UPDATE programacoes
               SET {", ".join(sets)}
             WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
               AND UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
        """
        cur.execute(sql)
        fixed = int(cur.rowcount or 0)
        conn.commit()
    return fixed


def sanitize_status_finalizacao_inconsistente() -> int:
    """
    Corrige rotas marcadas como ativas/em entrega, mas com evidência de finalização.
    Regra: se data_chegada/hora_chegada/km_final já existem, status deve ser FINALIZADA.
    """
    fixed = 0
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            return 0
        has_data_chegada = "data_chegada" in cols
        has_hora_chegada = "hora_chegada" in cols
        has_km_final = "km_final" in cols
        if not (has_data_chegada or has_hora_chegada or has_km_final):
            return 0

        evidencias = []
        if has_data_chegada:
            evidencias.append("TRIM(COALESCE(data_chegada,''))<>''")
        if has_hora_chegada:
            evidencias.append("TRIM(COALESCE(hora_chegada,''))<>''")
        if has_km_final:
            evidencias.append("COALESCE(km_final,0)>0")
        evid_sql = " OR ".join(evidencias) or "0=1"

        set_cols = ["status='FINALIZADA'"]
        if "status_operacional" in cols:
            set_cols.append("status_operacional='FINALIZADA'")
        if "finalizada_no_app" in cols:
            set_cols.append("finalizada_no_app=1")

        sql = f"""
            UPDATE programacoes
               SET {", ".join(set_cols)}
             WHERE UPPER(TRIM(COALESCE(status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
               AND ({evid_sql})
        """
        cur.execute(sql)
        fixed = int(cur.rowcount or 0)
        conn.commit()
    return fixed


def reconcile_programacoes_motorista_links() -> int:
    """
    Preenche motorista_id/motorista_codigo/codigo_motorista em programacoes abertas,
    usando cadastro de motoristas. Evita roteamento incorreto no app mobile.
    """
    fixed = 0
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            return 0

        has_mid = "motorista_id" in cols
        has_mcod = "motorista_codigo" in cols
        has_cmot = "codigo_motorista" in cols
        if not (has_mid or has_mcod or has_cmot):
            return 0

        sets = []
        if has_mid:
            sets.append("motorista_id=?")
        if has_mcod:
            sets.append("motorista_codigo=?")
        if has_cmot:
            sets.append("codigo_motorista=?")
        set_sql = ", ".join(sets)
        if not set_sql:
            return 0

        select_parts = ["id", "COALESCE(motorista,'') AS motorista"]
        if has_mid:
            select_parts.append("COALESCE(motorista_id,0) AS motorista_id")
        else:
            select_parts.append("0 AS motorista_id")
        if has_mcod:
            select_parts.append("COALESCE(motorista_codigo,'') AS motorista_codigo")
        else:
            select_parts.append("'' AS motorista_codigo")
        if has_cmot:
            select_parts.append("COALESCE(codigo_motorista,'') AS codigo_motorista")
        else:
            select_parts.append("'' AS codigo_motorista")

        missing_predicates = []
        if has_mid:
            missing_predicates.append("COALESCE(motorista_id,0)=0")
        if has_mcod:
            missing_predicates.append("TRIM(COALESCE(motorista_codigo,''))=''")
        if has_cmot:
            missing_predicates.append("TRIM(COALESCE(codigo_motorista,''))=''")
        missing_sql = " OR ".join(missing_predicates) or "1=1"

        cur.execute(
            f"""
            SELECT {", ".join(select_parts)}
            FROM programacoes
            WHERE UPPER(TRIM(COALESCE(status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
              AND ({missing_sql})
            ORDER BY id DESC
            LIMIT 5000
            """
        )
        rows = cur.fetchall() or []
        for r in rows:
            pid = int(r["id"] or 0)
            mot_txt = str(r["motorista"] or "").strip()
            mid_cur = int(r["motorista_id"] or 0)
            mcod_cur = str(r["motorista_codigo"] or "").strip().upper()
            cmot_cur = str(r["codigo_motorista"] or "").strip().upper()

            mot_id = mid_cur if mid_cur > 0 else 0
            mot_codigo = mcod_cur or cmot_cur

            # 1) se tem ID, resolve codigo
            if mot_id > 0 and not mot_codigo:
                cur.execute("SELECT COALESCE(codigo,'') AS codigo FROM motoristas WHERE id=? LIMIT 1", (mot_id,))
                rr = cur.fetchone()
                mot_codigo = str((rr["codigo"] if rr else "") or "").strip().upper()

            # 2) se nao tem ID/codigo, tenta bater por codigo em texto e depois por nome
            if mot_id <= 0 or not mot_codigo:
                probe_code = mot_txt.upper()
                cur.execute(
                    "SELECT id, COALESCE(codigo,'') AS codigo FROM motoristas WHERE UPPER(TRIM(codigo))=UPPER(TRIM(?)) LIMIT 1",
                    (probe_code,),
                )
                rr = cur.fetchone()
                if not rr and "(" in mot_txt and ")" in mot_txt:
                    try:
                        code_in_name = mot_txt[mot_txt.rfind("(") + 1: mot_txt.rfind(")")].strip().upper()
                        cur.execute(
                            "SELECT id, COALESCE(codigo,'') AS codigo FROM motoristas WHERE UPPER(TRIM(codigo))=UPPER(TRIM(?)) LIMIT 1",
                            (code_in_name,),
                        )
                        rr = cur.fetchone()
                    except Exception:
                        rr = None
                if not rr and mot_txt:
                    cur.execute(
                        "SELECT id, COALESCE(codigo,'') AS codigo FROM motoristas WHERE UPPER(TRIM(nome))=UPPER(TRIM(?)) LIMIT 1",
                        (mot_txt,),
                    )
                    rr = cur.fetchone()
                if rr:
                    mot_id = int(rr["id"] or 0)
                    mot_codigo = str(rr["codigo"] or "").strip().upper()

            if mot_id <= 0 and not mot_codigo:
                continue

            params: List[Any] = []
            if has_mid:
                params.append(int(mot_id or 0))
            if has_mcod:
                params.append(mot_codigo)
            if has_cmot:
                params.append(mot_codigo)
            params.append(pid)

            cur.execute(f"UPDATE programacoes SET {set_sql} WHERE id=?", tuple(params))
            if int(cur.rowcount or 0) > 0:
                fixed += 1

        conn.commit()
    return fixed


def _resolve_motorista_vinculo(
    cur: sqlite3.Cursor,
    motorista_nome: str = "",
    motorista_id: int = 0,
    motorista_codigo: str = "",
) -> tuple[str, int, str]:
    nome = str(motorista_nome or "").strip().upper()
    mot_id = int(motorista_id or 0)
    codigo = str(motorista_codigo or "").strip().upper()
    row = None

    if mot_id > 0:
        cur.execute(
            """
            SELECT id, COALESCE(nome,'') AS nome, COALESCE(codigo,'') AS codigo
            FROM motoristas
            WHERE id=?
            LIMIT 1
            """,
            (mot_id,),
        )
        row = cur.fetchone()

    if not row and codigo:
        cur.execute(
            """
            SELECT id, COALESCE(nome,'') AS nome, COALESCE(codigo,'') AS codigo
            FROM motoristas
            WHERE UPPER(TRIM(codigo))=UPPER(TRIM(?))
            LIMIT 1
            """,
            (codigo,),
        )
        row = cur.fetchone()

    if not row and nome and "(" in nome and ")" in nome:
        try:
            code_in_name = nome[nome.rfind("(") + 1 : nome.rfind(")")].strip().upper()
        except Exception:
            code_in_name = ""
        if code_in_name:
            cur.execute(
                """
                SELECT id, COALESCE(nome,'') AS nome, COALESCE(codigo,'') AS codigo
                FROM motoristas
                WHERE UPPER(TRIM(codigo))=UPPER(TRIM(?))
                LIMIT 1
                """,
                (code_in_name,),
            )
            row = cur.fetchone()

    if not row and nome:
        cur.execute(
            """
            SELECT id, COALESCE(nome,'') AS nome, COALESCE(codigo,'') AS codigo
            FROM motoristas
            WHERE UPPER(TRIM(nome))=UPPER(TRIM(?))
            LIMIT 1
            """,
            (nome,),
        )
        row = cur.fetchone()

    if row:
        mot_id = int(row["id"] or 0)
        nome = str(row["nome"] or "").strip().upper()
        codigo = str(row["codigo"] or "").strip().upper()

    return nome, mot_id, codigo


@app.on_event("startup")
def _startup():
    ensure_tables()
    with sqlite3.connect(DB_PATH) as admin_conn:
        ensure_admin_user_bootstrap(
            admin_conn,
            os.environ.get("ROTA_ADMIN_PASS") or os.environ.get("ROTA_ADMIN_PASSWORD"),
        )
    logging.info("Banco publicado inicializado | env=%s | db=%s | admin=%s", APP_CONFIG.app_env, DB_PATH, "configured" if os.environ.get("ROTA_ADMIN_PASS") or os.environ.get("ROTA_ADMIN_PASSWORD") else "generated")
    log_startup_diagnostics(DB_PATH, APP_CONFIG)
    try:
        fixed = reconcile_transferencias_status()
        logging.info("Reconciliacao de transferencias concluida. Itens ajustados: %s", fixed)
    except Exception:
        logging.exception("Falha na reconciliacao de transferencias no startup")
    try:
        fixed_st = sanitize_status_operacional_legado()
        logging.info("Saneamento de status operacional legado concluido. Rotas ajustadas: %s", fixed_st)
    except Exception:
        logging.exception("Falha no saneamento de status operacional legado no startup")
    try:
        fixed_fin = sanitize_status_finalizacao_inconsistente()
        logging.info("Saneamento de finalizacao inconsistente concluido. Rotas ajustadas: %s", fixed_fin)
    except Exception:
        logging.exception("Falha no saneamento de finalizacao inconsistente no startup")
    try:
        fixed_links = reconcile_programacoes_motorista_links()
        logging.info("Reconciliacao de vinculos motorista/programacoes concluida. Rotas ajustadas: %s", fixed_links)
    except Exception:
        logging.exception("Falha na reconciliacao de vinculos motorista/programacoes no startup")

# =========================================================
# TOKEN HELPERS (HMAC + base64 urlsafe)
# =========================================================
def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("utf-8").rstrip("=")


def _b64d(s: str) -> bytes:
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode((s + pad).encode("utf-8"))


def create_token(
    codigo: str,
    perfil: str = "",
    *,
    company_id: int | None = None,
    user_id: int | None = None,
    username: str | None = None,
    role: str | None = None,
) -> str:
    payload = {
        "codigo": codigo,
        "exp": int(time.time()) + TOKEN_TTL_SECONDS,
    }
    perfil_n = str(perfil or "").strip().lower()
    if perfil_n:
        payload["perfil"] = perfil_n
    if company_id is not None:
        payload["company_id"] = int(company_id)
    if user_id is not None:
        payload["user_id"] = int(user_id)
    username_n = str(username or codigo or "").strip()
    if username_n:
        payload["username"] = username_n
    role_n = str(role or perfil_n or "").strip()
    if role_n:
        payload["role"] = role_n
    payload_bytes = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")

    sig = hmac.new(
        SECRET_KEY.encode("utf-8"),
        payload_bytes,
        hashlib.sha256
    ).digest()

    return f"{_b64e(payload_bytes)}.{_b64e(sig)}"


def verify_token(token: str) -> Optional[Dict[str, Any]]:
    try:
        parts = token.split(".")
        if len(parts) != 2:
            return None

        payload_bytes = _b64d(parts[0])
        sig_bytes = _b64d(parts[1])

        expected_sig = hmac.new(
            SECRET_KEY.encode("utf-8"),
            payload_bytes,
            hashlib.sha256
        ).digest()

        if not hmac.compare_digest(sig_bytes, expected_sig):
            return None

        payload_str = payload_bytes.decode("utf-8")
        try:
            payload = json.loads(payload_str)
        except ValueError:
            return None
        if not isinstance(payload, dict):
            return None

        exp = int(payload.get("exp", 0))
        if time.time() > exp:
            return None

        codigo = payload.get("codigo")
        if not codigo:
            return None

        perfil = str(payload.get("perfil") or "").strip().lower()
        out = {
            "codigo": codigo,
            "exp": exp,
            "perfil": perfil,
            "username": str(payload.get("username") or codigo or "").strip(),
            "role": str(payload.get("role") or perfil or "").strip(),
        }
        for int_key in ("company_id", "user_id"):
            if payload.get(int_key) not in (None, ""):
                out[int_key] = int(payload.get(int_key))
        return out
    except Exception:
        return None


def _billing_context_for_company(company_id: int) -> Dict[str, Any]:
    with get_conn() as conn:
        cur = conn.cursor()
        company_status = "active"
        subscription_status = "active"
        if table_exists(cur, "companies"):
            cur.execute("SELECT COALESCE(status, 'active') AS status FROM companies WHERE id=? LIMIT 1", (int(company_id),))
            row = cur.fetchone()
            if row:
                company_status = str(row["status"] or "active")
        if table_exists(cur, "subscriptions"):
            cur.execute(
                """
                SELECT COALESCE(status, 'active') AS status
                FROM subscriptions
                WHERE company_id=?
                ORDER BY id DESC
                LIMIT 1
                """,
                (int(company_id),),
            )
            row = cur.fetchone()
            if row:
                subscription_status = str(row["status"] or "active")
        return {
            "company_status": company_status,
            "subscription_status": subscription_status,
        }


def _audit_billing_block(company_id: int, request: Request, payload: Dict[str, Any]) -> None:
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "audit_logs"):
            return
        try:
            host = str(request.client.host or "").strip()
        except Exception:
            host = ""
        cur.execute(
            """
            INSERT INTO audit_logs (
                company_id, actor_type, action, entity_type, entity_id,
                severity, ip_address, metadata_json, created_at
            )
            VALUES (?, 'system', 'billing_operation_blocked', 'company', ?, 'warning', ?, ?, datetime('now'))
            """,
            (
                int(company_id),
                str(company_id),
                host,
                json.dumps(
                    {
                        "path": str(request.url.path or ""),
                        "method": str(request.method or ""),
                        "billing_status": payload.get("billing_status"),
                    },
                    ensure_ascii=True,
                    sort_keys=True,
                ),
            ),
        )


FEATURE_ENDPOINTS = {
    "/desktop/relatorios/km-veiculos": "advanced_reports",
    "/desktop/relatorios/despesas-categorias": "financial_reports",
    "/desktop/relatorios/mortalidade-motorista": "advanced_reports",
    "/desktop/monitoramento/rotas": "rotas",
    "/desktop/escala": "escala",
    "/desktop/centro-custos": "centro_custos",
    "/rotas/gps": "realtime_tracking",
}


def _can_use_feature_for_company(company_id: int, feature_name: str) -> bool:
    feature_key = str(feature_name or "").strip()
    if not feature_key:
        return True
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT p.features_json
                FROM subscriptions s
                JOIN plans p ON p.id = s.plan_id
                WHERE s.company_id=? AND s.status IN ('active', 'trialing', 'past_due', 'suspended')
                ORDER BY s.id DESC
                LIMIT 1
                """,
                (int(company_id),),
            )
            row = cur.fetchone()
            if not row:
                return False
            try:
                features = json.loads(str(row["features_json"] or "{}"))
            except Exception:
                features = {}
            return bool(isinstance(features, dict) and features.get(feature_key))
    except Exception:
        return False


app.add_middleware(BillingProtectionMiddleware, get_billing_context=_billing_context_for_company, audit_block=_audit_billing_block)
app.add_middleware(FeatureGateMiddleware, endpoint_features=FEATURE_ENDPOINTS, can_use_feature=_can_use_feature_for_company)
app.add_middleware(TenantContextMiddleware, verify_token=verify_token)

def _motorista_app_role(row: Any) -> str:
    if row is None:
        return "MOTORISTA"
    try:
        value = str(row["perfil_app"] or "").strip().upper()
    except Exception:
        try:
            value = str((row.get("perfil_app") or "")).strip().upper()
        except Exception:
            value = ""
    return "ADMIN" if value == "ADMIN" else "MOTORISTA"

# =========================================================
# SCHEMAS (Pydantic)
# =========================================================
class LoginIn(BaseModel):
    codigo: str
    senha: str


class LoginOut(BaseModel):
    token: str
    nome: str
    codigo: str
    company_id: Optional[int] = None
    perfil: Optional[str] = None
    role: Optional[str] = None
    is_admin: Optional[bool] = None


class VendedorRascunhoItemIn(BaseModel):
    id: Optional[str] = None
    cod_cliente: str
    nome_cliente: str
    cidade: Optional[str] = None
    bairro: Optional[str] = None
    endereco: Optional[str] = None
    vendedor_cadastro: Optional[str] = None
    vendedor_origem: str
    preco: float = 0.0
    caixas: int = 0
    status: Optional[str] = "PENDENTE"
    observacao: Optional[str] = ""
    alerta_codigo_programacao: Optional[str] = None
    alerta_status_rota: Optional[str] = None


class VendedorRascunhoCreateIn(BaseModel):
    itens: List[VendedorRascunhoItemIn] = Field(default_factory=list)


class VendedorRascunhoUpdateIn(BaseModel):
    caixas: Optional[int] = None
    preco: Optional[float] = None
    observacao: Optional[str] = None
    status: Optional[str] = None


class VendedorRascunhoDeleteBulkIn(BaseModel):
    ids: List[str] = Field(default_factory=list)


class VendedorPreProgramacaoUpsertIn(BaseModel):
    id: Optional[str] = None
    titulo: Optional[str] = None
    observacao: Optional[str] = ""
    status: Optional[str] = "ABERTA"
    item_ids: List[str] = Field(default_factory=list)


class MotoristaAcessoIn(BaseModel):
    liberado: bool
    admin: Optional[str] = None
    motivo: Optional[str] = None


class MotoristaSenhaIn(BaseModel):
    nova_senha: str
    admin: Optional[str] = None
    motivo: Optional[str] = None


class CompanyPlanChangeIn(BaseModel):
    plan_code: str
    reason: Optional[str] = None


class CompanyStatusIn(BaseModel):
    status: str
    reason: Optional[str] = None


class PaymentCreateIn(BaseModel):
    company_id: int
    subscription_id: Optional[int] = None
    amount: float = 0.0
    due_date: Optional[str] = None
    method: Optional[str] = None
    reference: Optional[str] = None
    notes: Optional[str] = None


class PaymentRegisterIn(BaseModel):
    method: Optional[str] = "manual"
    reference: Optional[str] = None
    notes: Optional[str] = None


class BillingAutomationIn(BaseModel):
    grace_days: int = 0


class RotaAtivaOut(BaseModel):
    codigo_programacao: str
    status: str = ""
    motorista: str = ""
    veiculo: str = ""
    equipe: str = ""
    local_rota: str = ""
    local_carregamento: str = ""
    data_criacao: str = ""
    tipo_estimativa: Optional[str] = None
    unidade_estimativa: Optional[str] = None
    tipo_operacao: Optional[str] = None
    operacao_tipo: Optional[str] = None
    transbordo: Optional[bool] = False
    transbordo_modalidade: Optional[str] = None
    transbordo_observacao: Optional[str] = None
    transbordo_grupo: Optional[str] = None
    transferencias_saida: Optional[int] = 0
    transferencias_entrada: Optional[int] = 0
    transferencias_pendentes: Optional[int] = 0
    estimativa_valor: Optional[float] = None
    caixas_estimado: Optional[int] = None
    usuario_criacao: Optional[str] = None
    usuario_ultima_edicao: Optional[str] = None
    capacidade_cx: Optional[int] = None
    total_caixas: Optional[int] = None
    caixas_saldo: Optional[int] = None
    media_carregada: Optional[float] = None
    kg_carregado: Optional[float] = None
    caixas_carregadas: Optional[int] = None
    caixa_final: Optional[int] = None
    ultimo_km_veiculo: Optional[float] = 0
    km_inicial_sugerido: Optional[float] = 0
    ultimo_km_programacao: Optional[str] = ""
    substituicao_pendente: Optional[int] = 0
    status_operacional: Optional[str] = None


class RotaDetalheOut(BaseModel):
    rota: Dict[str, Any]
    clientes: List[Dict[str, Any]]


class IniciarRotaIn(BaseModel):
    data_saida: str
    hora_saida: str
    km_inicial: int
    override_reason: Optional[str] = None
    idempotency_key: Optional[str] = None


class RotaGpsPingIn(BaseModel):
    lat: float
    lon: float
    speed: Optional[float] = None
    accuracy: Optional[float] = None
    timestamp: Optional[str] = None
    idempotency_key: Optional[str] = None


class FinalizarRotaIn(BaseModel):
    data_chegada: str
    hora_chegada: str
    km_final: int
    idempotency_key: Optional[str] = None


class RotaStatusOperacionalIn(BaseModel):
    status_operacional: str
    observacao: Optional[str] = None
    evento_em: Optional[str] = None
    idempotency_key: Optional[str] = None


class RotaReabrirIn(BaseModel):
    observacao: Optional[str] = None
    evento_em: Optional[str] = None
    idempotency_key: Optional[str] = None


class CarregamentoIn(BaseModel):
    # O app (Carregamento2Page) manda isso:
    nf_numero: str = Field(..., min_length=1)
    nf_kg: float = Field(0.0, ge=0)
    kg_carregado: Optional[float] = Field(default=None, ge=0)

    caixas_carregadas: int = Field(..., gt=0)

    # podem vir vazios, então não force min_length=1
    inicio_carregamento: Optional[str] = None
    fim_carregamento: Optional[str] = None

    nf_preco: float = 0.0
    local_carregado: str = ""

    # controle de media / aves por caixa
    media: Optional[float] = None
    media_1: Optional[float] = None
    media_2: Optional[float] = None
    media_3: Optional[float] = None
    qnt_aves_por_cx: Optional[int] = None
    aves_caixa_final: Optional[int] = None
    qnt_aves_caixa_final: Optional[int] = None
    mortalidade_aves: int = Field(0, ge=0)  # aves mortas no transbordo
    idempotency_key: Optional[str] = None


class ClienteControleIn(BaseModel):
    cod_cliente: str
    mortalidade_aves: int = 0
    media_aplicada: Optional[float] = None
    peso_previsto: Optional[float] = None

    # recebimentos (já fica pronto, mesmo que o app ainda não use)
    valor_recebido: Optional[float] = None
    forma_recebimento: Optional[str] = None
    obs_recebimento: Optional[str] = None

    # status do pedido / alteracoes
    status_pedido: Optional[str] = None
    pedido: Optional[str] = None
    caixas_atual: Optional[int] = None
    preco_atual: Optional[float] = None
    alterado_por: Optional[str] = None
    alteracao_tipo: Optional[str] = None
    alteracao_detalhe: Optional[str] = None
    lat_evento: Optional[float] = None
    lon_evento: Optional[float] = None
    lat_entrega: Optional[float] = None
    lon_entrega: Optional[float] = None
    accuracy_entrega: Optional[float] = None
    timestamp_entrega: Optional[str] = None
    endereco_evento: Optional[str] = None
    cidade_evento: Optional[str] = None
    bairro_evento: Optional[str] = None
    ordem_sugerida: Optional[int] = None
    eta: Optional[str] = None
    distancia: Optional[float] = None
    confianca_localizacao: Optional[float] = None
    mortalidade_foto_path: Optional[str] = None
    foto_mortalidade_path: Optional[str] = None
    foto_mortalidade: Optional[Dict[str, Any]] = None
    foto_registro: Optional[Dict[str, Any]] = None
    evento_em: Optional[str] = None
    idempotency_key: Optional[str] = None


class RotaTransbordoIn(BaseModel):
    aves_mortas_transbordo: int = Field(0, ge=0)
    mortalidade_transbordo_aves: Optional[int] = Field(default=None, ge=0)
    mortalidade_transbordo_kg: Optional[float] = Field(default=None, ge=0)
    obs_transbordo: Optional[str] = None
    mortalidade_transbordo_obs: Optional[str] = None
    foto_doa_path: Optional[str] = None
    doa_foto_path: Optional[str] = None
    mortalidade_transbordo_foto_path: Optional[str] = None
    foto_doa: Optional[Dict[str, Any]] = None
    foto_registro: Optional[Dict[str, Any]] = None
    idempotency_key: Optional[str] = None


class RotaDespesaIn(BaseModel):
    id_local: Optional[str] = None
    tipo: Optional[str] = None
    valor_total: float = Field(..., ge=0)
    descricao: Optional[str] = None
    forma_pagamento: Optional[str] = None
    comprovante_path: Optional[str] = None
    estabelecimento: Optional[str] = None
    documento: Optional[str] = None
    litros: Optional[float] = None
    valor_litro: Optional[float] = None
    desconto: Optional[float] = None
    combustivel: Optional[str] = None
    odometro: Optional[float] = None
    lat: Optional[float] = None
    lon: Optional[float] = None
    accuracy: Optional[float] = None
    timestamp: Optional[str] = None
    origem: Optional[str] = None
    registrado_em: Optional[str] = None
    motorista_codigo: Optional[str] = None
    motorista_nome: Optional[str] = None
    sync_key: Optional[str] = None
    status_sync: Optional[str] = None
    vinculo_prestacao: Optional[Dict[str, Any]] = None
    prestacao_contas: Optional[bool] = None
    desktop_web: Optional[Dict[str, Any]] = None
    foto_despesa: Optional[Dict[str, Any]] = None
    foto_registro: Optional[Dict[str, Any]] = None
    idempotency_key: Optional[str] = None


class RotaAjudantesIn(BaseModel):
    ajudantes_anteriores: Optional[str] = None
    ajudantes_novos: str
    motivo: Optional[str] = None
    alterado_em: Optional[str] = None
    origem: Optional[str] = None
    idempotency_key: Optional[str] = None


class ClienteReservaIn(BaseModel):
    cod_cliente: str
    nome_cliente: str
    pedido: Optional[str] = None
    qnt_caixas: int = Field(..., gt=0)
    preco: Optional[float] = None
    vendedor: Optional[str] = None
    cidade: Optional[str] = None
    observacao: Optional[str] = None
    status_pedido: Optional[str] = "PENDENTE"


class AvulsaItemIn(BaseModel):
    cod_cliente: str
    nome_cliente: str
    endereco: Optional[str] = None
    cidade: Optional[str] = None
    bairro: Optional[str] = None
    ordem: Optional[int] = 0
    observacao: Optional[str] = None


class ProgramacaoAvulsaIn(BaseModel):
    data_programada: Optional[str] = None
    motorista_id: Optional[int] = None
    motorista_codigo: Optional[str] = None
    motorista_nome: Optional[str] = None
    veiculo: Optional[str] = None
    equipe: Optional[str] = None
    local_rota: Optional[str] = None
    observacao: Optional[str] = None
    criado_por: Optional[str] = None
    itens: List[AvulsaItemIn] = Field(default_factory=list)


class ProgramacaoAvulsaConciliarIn(BaseModel):
    codigo_programacao_oficial: str
    usuario: Optional[str] = None


class DesktopRotaItemIn(BaseModel):
    cod_cliente: str
    nome_cliente: str
    qnt_caixas: Optional[int] = 0
    kg: Optional[float] = 0.0
    preco: Optional[float] = 0.0
    endereco: Optional[str] = None
    vendedor: Optional[str] = None
    pedido: Optional[str] = None
    produto: Optional[str] = None
    obs: Optional[str] = None
    ordem_sugerida: Optional[int] = None
    eta: Optional[str] = None
    distancia: Optional[float] = None
    confianca_localizacao: Optional[float] = None


class DesktopRotaUpsertIn(BaseModel):
    codigo_programacao: str
    data_criacao: Optional[str] = None
    motorista: Optional[str] = None
    motorista_id: Optional[int] = None
    motorista_codigo: Optional[str] = None
    codigo_motorista: Optional[str] = None
    veiculo: Optional[str] = None
    equipe: Optional[str] = None
    kg_estimado: Optional[float] = 0.0
    tipo_estimativa: Optional[str] = "KG"
    caixas_estimado: Optional[int] = 0
    operacao_tipo: Optional[str] = None
    transbordo_modalidade: Optional[str] = None
    transbordo_observacao: Optional[str] = None
    transbordo_grupo: Optional[str] = None
    status: Optional[str] = "ATIVA"
    local_rota: Optional[str] = None
    tipo_rota: Optional[str] = None
    local_carregamento: Optional[str] = None
    local_carregado: Optional[str] = None
    granja_carregada: Optional[str] = None
    local_carreg: Optional[str] = None
    adiantamento: Optional[float] = 0.0
    adiantamento_origem: Optional[str] = None
    pix_motorista: Optional[float] = 0.0
    total_caixas: Optional[int] = 0
    quilos: Optional[float] = 0.0
    nf_kg: Optional[float] = None
    nf_preco: Optional[float] = None
    nf_caixas: Optional[int] = None
    caixas_carregadas: Optional[int] = None
    usuario_criacao: Optional[str] = None
    usuario_ultima_edicao: Optional[str] = None
    linked_venda_ids: List[int] = Field(default_factory=list)
    vendas_usada_em: Optional[str] = None
    itens: List[DesktopRotaItemIn] = Field(default_factory=list)


class DesktopMotoristaUpsertIn(BaseModel):
    codigo: str
    nome: str
    telefone: Optional[str] = None
    cpf: Optional[str] = None
    status: Optional[str] = "ATIVO"
    perfil_app: Optional[str] = "MOTORISTA"
    senha: Optional[str] = None
    acesso_liberado: Optional[bool] = None
    acesso_liberado_por: Optional[str] = None
    acesso_obs: Optional[str] = None


class DesktopVendedorUpsertIn(BaseModel):
    codigo: str
    nome: str
    telefone: Optional[str] = None
    cidade_base: Optional[str] = None
    status: Optional[str] = "ATIVO"
    senha: Optional[str] = None


class DesktopVeiculoUpsertIn(BaseModel):
    placa: str
    modelo: str
    capacidade_cx: Optional[int] = 0
    status: Optional[str] = "ATIVO"


class DesktopAjudanteUpsertIn(BaseModel):
    nome: str
    sobrenome: str
    telefone: Optional[str] = None
    status: Optional[str] = "ATIVO"


class DesktopClienteUpsertIn(BaseModel):
    cod_cliente: str
    nome_cliente: str
    endereco: Optional[str] = None
    telefone: Optional[str] = None
    vendedor: Optional[str] = None


class DesktopClientesBulkUpsertIn(BaseModel):
    clientes: List[DesktopClienteUpsertIn] = Field(default_factory=list)


def _require_desktop_phone(value: Optional[str], detail: str) -> str:
    telefone = normalize_phone(value or "")
    if not is_valid_phone(telefone):
        raise HTTPException(status_code=400, detail=detail)
    return telefone


def _optional_desktop_phone(value: Optional[str], detail: str) -> str:
    telefone = normalize_phone(value or "")
    if telefone and not is_valid_phone(telefone):
        raise HTTPException(status_code=400, detail=detail)
    return telefone


def _optional_desktop_cpf(value: Optional[str]) -> str:
    cpf = normalize_cpf(value or "")
    if cpf and not is_valid_cpf(cpf):
        raise HTTPException(status_code=400, detail="cpf invalido.")
    return cpf


def _validate_desktop_password(value: str, detail: str) -> str:
    senha = _clean_text(value)
    if senha.startswith("pbkdf2_sha256$"):
        return senha
    if not is_valid_motorista_senha(senha):
        raise HTTPException(status_code=400, detail=detail)
    return senha


class DesktopRecebimentoIn(BaseModel):
    cod_cliente: str
    nome_cliente: str
    valor: float
    forma_pagamento: Optional[str] = "DINHEIRO"
    observacao: Optional[str] = ""
    num_nf: Optional[str] = ""


class DesktopDespesaIn(BaseModel):
    descricao: str
    valor: float
    categoria: Optional[str] = "OUTROS"
    observacao: Optional[str] = ""


class DesktopRotaCabecalhoIn(BaseModel):
    data_saida: Optional[str] = None
    hora_saida: Optional[str] = None
    data_chegada: Optional[str] = None
    hora_chegada: Optional[str] = None
    diaria_motorista_valor: Optional[float] = None
    qtd_diarias: Optional[float] = None
    qtd_ajudantes: Optional[int] = None
    total_motorista: Optional[float] = None
    total_ajudantes: Optional[float] = None
    observacao_motorista: Optional[str] = None
    observacao_ajudantes: Optional[str] = None


class DesktopRotaFinanceiroIn(BaseModel):
    nf_numero: Optional[str] = None
    nf_kg: Optional[float] = None
    nf_caixas: Optional[int] = None
    nf_kg_carregado: Optional[float] = None
    nf_kg_vendido: Optional[float] = None
    nf_saldo: Optional[float] = None
    nf_preco: Optional[float] = None
    media: Optional[float] = None
    nf_caixa_final: Optional[int] = None
    km_inicial: Optional[float] = None
    km_final: Optional[float] = None
    litros: Optional[float] = None
    km_rodado: Optional[float] = None
    media_km_l: Optional[float] = None
    custo_km: Optional[float] = None
    ced_200_qtd: Optional[int] = None
    ced_100_qtd: Optional[int] = None
    ced_50_qtd: Optional[int] = None
    ced_20_qtd: Optional[int] = None
    ced_10_qtd: Optional[int] = None
    ced_5_qtd: Optional[int] = None
    ced_2_qtd: Optional[int] = None
    valor_dinheiro: Optional[float] = None
    pix_motorista: Optional[float] = None
    adiantamento: Optional[float] = None
    adiantamento_origem: Optional[str] = None
    rota_observacao: Optional[str] = None


class DesktopDiariasSyncIn(BaseModel):
    qtd_diarias: float
    qtd_ajudantes: int = 0
    total_motorista: float
    total_ajudantes: float
    observacao_motorista: Optional[str] = ""
    observacao_ajudantes: Optional[str] = ""


def _validate_iso_date_optional(value: Optional[str], field_name: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        datetime.strptime(text, "%Y-%m-%d")
    except Exception:
        raise HTTPException(status_code=400, detail=f"{field_name} invalida. Use YYYY-MM-DD.")
    return text


def _validate_time_optional(value: Optional[str], field_name: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            datetime.strptime(text, fmt)
            return text if fmt == "%H:%M:%S" else f"{text}:00"
        except Exception:
            continue
    raise HTTPException(status_code=400, detail=f"{field_name} invalida. Use HH:MM ou HH:MM:SS.")


class DesktopProgramacaoClienteManualIn(BaseModel):
    cod_cliente: str
    nome_cliente: str


class DesktopProgramacaoStatusIn(BaseModel):
    status: Optional[str] = None
    prestacao_status: Optional[str] = None
    status_operacional: Optional[str] = None
    finalizada_no_app: Optional[int] = None


class DesktopSqlStatementIn(BaseModel):
    sql: str
    params: Optional[Any] = Field(default_factory=list)


class DesktopSqlMutateIn(BaseModel):
    statements: List[DesktopSqlStatementIn] = Field(default_factory=list)


class SubstituicaoRotaIn(BaseModel):
    motorista_destino_codigo: str
    veiculo_destino: str = ""
    motivo: str
    km_evento: Optional[int] = None
    lat_evento: Optional[float] = None
    lon_evento: Optional[float] = None


class SubstituicaoRotaDecisaoIn(BaseModel):
    motivo: Optional[str] = None

# =========================================================
# AUTH / DEPENDENCY
# =========================================================
def get_current_motorista(
    credentials: HTTPAuthorizationCredentials = Depends(security),
) -> Dict[str, Any]:
    token = credentials.credentials

    data = verify_token(token)
    if not data:
        raise HTTPException(status_code=401, detail="Token invÃ¡lido/expirado")
    perfil_token = str(data.get("perfil") or "").strip().lower()
    if perfil_token and perfil_token not in {"motorista", "admin"}:
        raise HTTPException(status_code=401, detail="Token invalido para o app do motorista")

    codigo = data.get("codigo")
    if not codigo:
        raise HTTPException(status_code=401, detail="Token sem cÃ³digo do motorista")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        default_company_id = _default_company_id(cur)
        select_parts = [
            "id",
            "nome",
            "codigo",
            "COALESCE(acesso_liberado, 0) AS acesso_liberado" if "acesso_liberado" in cols_m else "1 AS acesso_liberado",
            "UPPER(TRIM(COALESCE(perfil_app, 'MOTORISTA'))) AS perfil_app" if "perfil_app" in cols_m else "'MOTORISTA' AS perfil_app",
            f"COALESCE(company_id, {default_company_id}) AS company_id" if "company_id" in cols_m else f"{default_company_id} AS company_id",
        ]
        cur.execute(
            f"SELECT {', '.join(select_parts)} FROM motoristas WHERE codigo=?",
            (codigo,),
        )
        m = cur.fetchone()
        if not m:
            raise HTTPException(status_code=401, detail="Motorista nÃ£o encontrado")
        if "acesso_liberado" in cols_m and int(m["acesso_liberado"] or 0) != 1:
            raise HTTPException(status_code=403, detail="Acesso bloqueado. Solicite desbloqueio do administrador.")

        perfil_app = _motorista_app_role(m)
        row_company_id = _row_company_id(m, default_company_id)
        token_company_id = data.get("company_id")
        if token_company_id is not None and int(token_company_id) != row_company_id:
            raise HTTPException(status_code=403, detail="Token pertence a outra empresa")
        if perfil_token == "admin" and perfil_app != "ADMIN":
            raise HTTPException(status_code=403, detail="Usuario sem perfil admin para o app do motorista")
        is_admin = perfil_app == "ADMIN" or perfil_token == "admin"

    return {
        "codigo": m["codigo"],
        "nome": m["nome"],
        "id": m["id"],
        "company_id": row_company_id,
        "is_admin": is_admin,
        "perfil_app": perfil_app,
    }


def get_current_vendedor(
    credentials: HTTPAuthorizationCredentials = Depends(security),
) -> Dict[str, Any]:
    token = credentials.credentials

    data = verify_token(token)
    if not data:
        raise HTTPException(status_code=401, detail="Token invalido/expirado")
    perfil = str(data.get("perfil") or "").strip().lower()
    if perfil and perfil != "vendedor":
        raise HTTPException(status_code=401, detail="Token invalido para o app do vendedor")

    codigo = str(data.get("codigo") or "").strip().upper()
    if not codigo:
        raise HTTPException(status_code=401, detail="Token sem codigo do vendedor")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendedores'")
        if not cur.fetchone():
            raise HTTPException(status_code=401, detail="Cadastro de vendedores indisponivel")
        cur.execute("PRAGMA table_info(vendedores)")
        cols_v = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        default_company_id = _default_company_id(cur)
        company_expr = f"COALESCE(company_id, {default_company_id}) AS company_id" if "company_id" in cols_v else f"{default_company_id} AS company_id"
        cur.execute(
            f"""
            SELECT id, COALESCE(nome,'') AS nome, COALESCE(codigo,'') AS codigo, COALESCE(status,'ATIVO') AS status, {company_expr}
            FROM vendedores
            WHERE UPPER(TRIM(codigo))=?
            LIMIT 1
            """,
            (codigo,),
        )
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=401, detail="Vendedor nao encontrado")
        status = str(row["status"] or "ATIVO").strip().upper()
        if status not in {"ATIVO", ""}:
            raise HTTPException(status_code=403, detail="Acesso do vendedor bloqueado")
        row_company_id = _row_company_id(row, default_company_id)
        token_company_id = data.get("company_id")
        if token_company_id is not None and int(token_company_id) != row_company_id:
            raise HTTPException(status_code=403, detail="Token pertence a outra empresa")

    return {"codigo": row["codigo"], "nome": row["nome"], "id": row["id"], "company_id": row_company_id}


def _require_admin_user(admin: Dict[str, Any]) -> None:
    if not bool((admin or {}).get("is_admin")):
        raise HTTPException(status_code=403, detail="Apenas administradores podem executar esta operacao.")


@app.get("/admin/companies")
def admin_list_companies(
    status: str = Query(default=""),
    limit: int = Query(default=500, ge=1, le=5000),
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "companies"):
            return []
        params: List[Any] = []
        where = ""
        if str(status or "").strip():
            where = "WHERE status=?"
            params.append(str(status or "").strip())
        params.append(int(limit))
        cur.execute(
            f"""
            SELECT *
            FROM companies
            {where}
            ORDER BY id ASC
            LIMIT ?
            """,
            tuple(params),
        )
        return [row_to_dict(row) for row in (cur.fetchall() or [])]


@app.get("/admin/companies/{company_id}")
def admin_get_company(
    company_id: int,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "companies"):
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (int(company_id),))
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        return row_to_dict(row)


@app.put("/admin/companies/{company_id}/status")
def admin_update_company_status(
    company_id: int,
    payload: CompanyStatusIn,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    status = str(payload.status or "").strip().lower()
    if status not in {"active", "suspended", "cancelled", "inactive"}:
        raise HTTPException(status_code=400, detail="Status invalido.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE companies SET status=?, updated_at=datetime('now') WHERE id=?", (status, int(company_id)))
        if cur.rowcount == 0:
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        if table_exists(cur, "audit_logs"):
            cur.execute(
                """
                INSERT INTO audit_logs (
                    company_id, user_id, actor_type, action, entity_type, entity_id,
                    severity, metadata_json, created_at
                )
                VALUES (?, ?, 'admin', 'empresa_status_alterado', 'company', ?, 'warning', ?, datetime('now'))
                """,
                (
                    int(company_id),
                    int(admin.get("id") or 0),
                    str(company_id),
                    json.dumps({"status": status, "reason": str(payload.reason or "").strip()}, ensure_ascii=True, sort_keys=True),
                ),
            )
        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (int(company_id),))
        return {"ok": True, "company": row_to_dict(cur.fetchone())}


@app.get("/admin/companies/{company_id}/usage")
def admin_company_usage(
    company_id: int,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "companies"):
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        cur.execute("SELECT id FROM companies WHERE id=? LIMIT 1", (int(company_id),))
        if not cur.fetchone():
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        usage = vehicle_usage_snapshot(conn, int(company_id))
        return {
            "company_id": int(company_id),
            "vehicles": usage,
            "usuarios": _admin_count_table(cur, "usuarios", int(company_id)),
            "motoristas": _admin_count_table(cur, "motoristas", int(company_id)),
            "vendedores": _admin_count_table(cur, "vendedores", int(company_id)),
            "clientes": _admin_count_table(cur, "clientes", int(company_id)),
            "programacoes": _admin_count_table(cur, "programacoes", int(company_id)),
        }


@app.get("/admin/plans")
def admin_list_plans(
    include_inactive: bool = Query(default=False),
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        where = "" if include_inactive else "WHERE COALESCE(status,'active')='active'"
        cur.execute(f"SELECT * FROM plans {where} ORDER BY monthly_price ASC, id ASC")
        rows = []
        for row in cur.fetchall() or []:
            item = row_to_dict(row)
            try:
                item["features"] = json.loads(str(item.get("features_json") or "{}"))
            except Exception:
                item["features"] = {}
            rows.append(item)
        return rows


@app.get("/admin/subscriptions")
def admin_list_subscriptions(
    company_id: Optional[int] = Query(default=None),
    limit: int = Query(default=500, ge=1, le=5000),
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        params: List[Any] = []
        where = ""
        if company_id:
            where = "WHERE s.company_id=?"
            params.append(int(company_id))
        params.append(int(limit))
        cur.execute(
            f"""
            SELECT
                s.*,
                c.name AS company_name,
                p.code AS plan_code,
                p.name AS plan_name,
                p.vehicle_limit AS plan_vehicle_limit
            FROM subscriptions s
            JOIN companies c ON c.id = s.company_id
            JOIN plans p ON p.id = s.plan_id
            {where}
            ORDER BY s.id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return [row_to_dict(row) for row in (cur.fetchall() or [])]


@app.get("/admin/payments")
def admin_list_payments(
    company_id: Optional[int] = Query(default=None),
    status: str = Query(default=""),
    limit: int = Query(default=500, ge=1, le=5000),
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        clauses: List[str] = []
        params: List[Any] = []
        if company_id:
            clauses.append("p.company_id=?")
            params.append(int(company_id))
        if str(status or "").strip():
            clauses.append("p.status=?")
            params.append(str(status or "").strip())
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        params.append(int(limit))
        cur.execute(
            f"""
            SELECT p.*, c.name AS company_name
            FROM payments p
            JOIN companies c ON c.id = p.company_id
            {where}
            ORDER BY COALESCE(p.due_date, p.created_at) DESC, p.id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return [row_to_dict(row) for row in (cur.fetchall() or [])]


@app.post("/admin/payments")
def admin_create_payment(
    payload: PaymentCreateIn,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM companies WHERE id=? LIMIT 1", (int(payload.company_id),))
        if not cur.fetchone():
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
        subscription_id = payload.subscription_id
        if not subscription_id:
            cur.execute(
                """
                SELECT id
                FROM subscriptions
                WHERE company_id=? AND status IN ('active', 'trialing', 'past_due', 'suspended')
                ORDER BY id DESC
                LIMIT 1
                """,
                (int(payload.company_id),),
            )
            row = cur.fetchone()
            subscription_id = int(row["id"]) if row else None
        cur.execute(
            """
            INSERT INTO payments (
                subscription_id, company_id, amount, due_date, status, method,
                reference, notes, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, 'pending', ?, ?, ?, datetime('now'), datetime('now'))
            """,
            (
                subscription_id,
                int(payload.company_id),
                float(payload.amount or 0),
                payload.due_date,
                payload.method,
                payload.reference,
                payload.notes,
            ),
        )
        payment_id = int(cur.lastrowid)
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (payment_id,))
        return {"ok": True, "payment": row_to_dict(cur.fetchone())}


@app.post("/admin/payments/{payment_id}/registrar-pagamento")
def admin_register_payment(
    payment_id: int,
    payload: PaymentRegisterIn,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (int(payment_id),))
        payment = cur.fetchone()
        if not payment:
            raise HTTPException(status_code=404, detail="Pagamento nao encontrado.")
        cur.execute(
            """
            UPDATE payments
            SET
                status='paid',
                paid_at=datetime('now'),
                method=COALESCE(?, method),
                reference=COALESCE(?, reference),
                notes=COALESCE(?, notes),
                updated_at=datetime('now')
            WHERE id=?
            """,
            (payload.method, payload.reference, payload.notes, int(payment_id)),
        )
        subscription_id = payment["subscription_id"]
        if subscription_id:
            cur.execute(
                """
                UPDATE subscriptions
                SET
                    status='active',
                    current_period_start=date('now'),
                    current_period_end=date('now', '+30 day'),
                    next_due_date=date('now', '+30 day'),
                    updated_at=datetime('now')
                WHERE id=?
                """,
                (int(subscription_id),),
            )
        cur.execute("UPDATE companies SET status='active', updated_at=datetime('now') WHERE id=?", (int(payment["company_id"]),))
        if table_exists(cur, "audit_logs"):
            cur.execute(
                """
                INSERT INTO audit_logs (
                    company_id, user_id, actor_type, action, entity_type, entity_id,
                    severity, metadata_json, created_at
                )
                VALUES (?, ?, 'admin', 'pagamento_registrado', 'payment', ?, 'info', ?, datetime('now'))
                """,
                (
                    int(payment["company_id"]),
                    int(admin.get("id") or 0),
                    str(payment_id),
                    json.dumps({"method": payload.method, "reference": payload.reference}, ensure_ascii=True, sort_keys=True),
                ),
            )
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (int(payment_id),))
        return {"ok": True, "payment": row_to_dict(cur.fetchone())}


@app.get("/admin/audit-logs")
def admin_audit_logs(
    company_id: Optional[int] = Query(default=None),
    action: str = Query(default=""),
    limit: int = Query(default=500, ge=1, le=5000),
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        clauses: List[str] = []
        params: List[Any] = []
        if company_id:
            clauses.append("company_id=?")
            params.append(int(company_id))
        if str(action or "").strip():
            clauses.append("action=?")
            params.append(str(action or "").strip())
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        params.append(int(limit))
        cur.execute(
            f"""
            SELECT *
            FROM audit_logs
            {where}
            ORDER BY id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return [row_to_dict(row) for row in (cur.fetchall() or [])]


@app.get("/admin/companies/{company_id}/features")
def admin_company_features(
    company_id: int,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT p.code AS plan_code, p.name AS plan_name, p.features_json
            FROM subscriptions s
            JOIN plans p ON p.id = s.plan_id
            WHERE s.company_id=? AND s.status IN ('active', 'trialing', 'past_due', 'suspended')
            ORDER BY s.id DESC
            LIMIT 1
            """,
            (int(company_id),),
        )
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Assinatura ativa nao encontrada.")
        try:
            features = json.loads(str(row["features_json"] or "{}"))
        except Exception:
            features = {}
        return {
            "company_id": int(company_id),
            "plan_code": row["plan_code"],
            "plan_name": row["plan_name"],
            "features": features if isinstance(features, dict) else {},
        }


@app.post("/admin/billing/run-overdue-check")
def admin_run_overdue_check(
    payload: BillingAutomationIn,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    _require_admin_user(admin)
    with get_conn() as conn:
        summary = suspend_overdue_subscriptions_conn(conn, grace_days=int(payload.grace_days or 0))
    return {"ok": True, "summary": summary}


def _admin_count_table(cur: sqlite3.Cursor, table: str, company_id: int) -> int:
    if not table_exists(cur, table):
        return 0
    cur.execute(f"PRAGMA table_info({table})")
    cols = {str(row[1]).lower() for row in (cur.fetchall() or [])}
    if "company_id" in cols:
        cur.execute(f'SELECT COUNT(*) FROM "{table}" WHERE company_id=?', (int(company_id),))
    else:
        cur.execute(f'SELECT COUNT(*) FROM "{table}"')
    row = cur.fetchone()
    return int(row[0] if row else 0)


@app.put("/admin/companies/{company_id}/plan")
def admin_change_company_plan(
    company_id: int,
    payload: CompanyPlanChangeIn,
    admin: Dict[str, Any] = Depends(get_current_motorista),
):
    if not bool(admin.get("is_admin")):
        raise HTTPException(status_code=403, detail="Apenas administradores podem alterar plano.")

    plan_code = str(payload.plan_code or "").strip().lower()
    if not plan_code:
        raise HTTPException(status_code=400, detail="plan_code e obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "companies") or not table_exists(cur, "plans") or not table_exists(cur, "subscriptions"):
            raise HTTPException(status_code=500, detail="Schema SaaS indisponivel.")

        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (int(company_id),))
        company = cur.fetchone()
        if not company:
            raise HTTPException(status_code=404, detail="Empresa nao encontrada.")

        cur.execute("SELECT * FROM plans WHERE lower(code)=? AND COALESCE(status,'active')='active' LIMIT 1", (plan_code,))
        plan = cur.fetchone()
        if not plan:
            raise HTTPException(status_code=404, detail="Plano nao encontrado.")

        usage = vehicle_usage_snapshot(conn, int(company_id))
        vehicle_count = int(usage.get("vehicle_count") or 0)
        vehicle_limit = plan["vehicle_limit"]
        if vehicle_limit is not None and vehicle_count > int(vehicle_limit):
            raise HTTPException(
                status_code=409,
                detail=(
                    f"Downgrade bloqueado: empresa possui {vehicle_count} veiculos, "
                    f"mas o plano {plan['name']} permite {int(vehicle_limit)}."
                ),
            )

        cur.execute(
            """
            SELECT id, plan_id
            FROM subscriptions
            WHERE company_id=? AND status IN ('active', 'trialing', 'past_due')
            ORDER BY id DESC
            LIMIT 1
            """,
            (int(company_id),),
        )
        current = cur.fetchone()
        previous_plan_id = int(current["plan_id"]) if current else None
        if current:
            cur.execute(
                """
                UPDATE subscriptions
                SET plan_id=?, status='active', updated_at=datetime('now')
                WHERE id=?
                """,
                (int(plan["id"]), int(current["id"])),
            )
            subscription_id = int(current["id"])
        else:
            cur.execute(
                """
                INSERT INTO subscriptions (
                    company_id, plan_id, status, billing_cycle,
                    current_period_start, current_period_end, next_due_date,
                    created_at, updated_at
                )
                VALUES (
                    ?, ?, 'active', 'monthly',
                    date('now'), date('now', '+30 day'), date('now', '+30 day'),
                    datetime('now'), datetime('now')
                )
                """,
                (int(company_id), int(plan["id"])),
            )
            subscription_id = int(cur.lastrowid)

        if table_exists(cur, "audit_logs"):
            cur.execute(
                """
                INSERT INTO audit_logs (
                    company_id, user_id, actor_type, action, entity_type, entity_id,
                    severity, metadata_json, created_at
                )
                VALUES (?, ?, 'admin', 'plano_alterado', 'subscription', ?, 'info', ?, datetime('now'))
                """,
                (
                    int(company_id),
                    int(admin.get("id") or 0),
                    str(subscription_id),
                    json.dumps(
                        {
                            "new_plan_code": plan["code"],
                            "new_plan_id": int(plan["id"]),
                            "previous_plan_id": previous_plan_id,
                            "reason": str(payload.reason or "").strip(),
                            "vehicle_count": vehicle_count,
                        },
                        ensure_ascii=True,
                        sort_keys=True,
                    ),
                ),
            )

        cur.execute(
            """
            SELECT
                s.id AS subscription_id,
                s.company_id,
                s.status,
                p.id AS plan_id,
                p.code AS plan_code,
                p.name AS plan_name,
                p.vehicle_limit
            FROM subscriptions s
            JOIN plans p ON p.id = s.plan_id
            WHERE s.id=?
            LIMIT 1
            """,
            (subscription_id,),
        )
        updated = row_to_dict(cur.fetchone())

    return {"ok": True, "subscription": updated}


def _owner_filter_for_programacoes(
    conn: sqlite3.Connection,
    motorista: Dict[str, Any],
    alias: str = "p",
) -> tuple[str, tuple]:
    """
    Resolve filtro de posse por motorista com prioridade em chaves estáveis.
    Fallback por nome existe apenas para bancos legados sem coluna de vÃnculo.
    """
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(programacoes)")
    cols = {r[1] for r in cur.fetchall() or []}

    conds: List[str] = []
    params: List[Any] = []

    if bool(motorista.get("is_admin")):
        return "(1=1)", tuple()

    # Prioriza chaves estáveis.
    if "motorista_id" in cols:
        conds.append(f"{alias}.motorista_id=?")
        params.append(int(motorista["id"]))
    if "motorista_codigo" in cols:
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista_codigo,'')))=UPPER(TRIM(?))")
        params.append(str(motorista["codigo"]))
    if "codigo_motorista" in cols:
        conds.append(f"UPPER(TRIM(COALESCE({alias}.codigo_motorista,'')))=UPPER(TRIM(?))")
        params.append(str(motorista["codigo"]))

    # Fallback por nome/campo legado somente quando base NÃO possui colunas estáveis.
    if not conds:
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista,'')))=UPPER(TRIM(?))")
        params.append(str(motorista["nome"]))
        # Fallback adicional por codigo no campo motorista (bases legadas e dados mistos).
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista,'')))=UPPER(TRIM(?))")
        params.append(str(motorista["codigo"]))
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista,''))) LIKE UPPER(TRIM(?))")
        params.append(f"%({str(motorista['codigo'])})%")
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista,''))) LIKE UPPER(TRIM(?))")
        params.append(f"{str(motorista['codigo'])} -%")
        conds.append(f"UPPER(TRIM(COALESCE({alias}.motorista,''))) LIKE UPPER(TRIM(?))")
        params.append(f"{str(motorista['codigo'])}/%")

    return "(" + " OR ".join(conds) + ")", tuple(params)


def _fetch_programacao_owned(
    cur: sqlite3.Cursor,
    codigo_programacao: str,
    motorista: Dict[str, Any],
    select_cols: str = "p.id",
) -> Optional[sqlite3.Row]:
    owner_sql, owner_params = _owner_filter_for_programacoes(cur.connection, motorista, "p")
    sql = f"""
        SELECT {select_cols}
        FROM programacoes p
        WHERE p.codigo_programacao=?
          AND {owner_sql}
        ORDER BY p.id DESC
        LIMIT 1
    """
    cur.execute(sql, ((codigo_programacao or "").strip(), *owner_params))
    return cur.fetchone()


def _require_desktop_secret(
    x_desktop_secret: Optional[str] = Header(default=None, alias="X-Desktop-Secret"),
) -> bool:
    secret = (x_desktop_secret or "").strip()
    expected = str(SECRET_KEY or "")
    if not secret or not expected or not hmac.compare_digest(secret, expected):
        raise HTTPException(status_code=401, detail="Desktop secret inválido")
    return True


def _desktop_company_id(cur: sqlite3.Cursor, x_company_id: Optional[str] = None) -> int:
    try:
        requested = int(str(x_company_id or "").strip())
        if requested > 0:
            if table_exists(cur, "companies"):
                cur.execute("SELECT id FROM companies WHERE id=? LIMIT 1", (requested,))
                if not cur.fetchone():
                    raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
            return requested
    except HTTPException:
        raise
    except Exception:
        pass
    return _default_company_id(cur)


def _company_scope_condition(cur: sqlite3.Cursor, table: str, company_id: Optional[int], alias: str = "") -> tuple[str, List[Any]]:
    if not company_id or not table_exists(cur, table):
        return "", []
    cur.execute(f"PRAGMA table_info({table})")
    cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if "company_id" not in cols:
        return "", []
    prefix = f"{alias}." if alias else ""
    return f"{prefix}company_id=?", [int(company_id)]


def _next_codigo_avulsa(cur: sqlite3.Cursor, data_ref: Optional[str] = None) -> str:
    base = datetime.now().strftime("%Y%m%d")
    txt = str(data_ref or "").strip()
    if txt:
        # aceita YYYY-MM-DD ou DD/MM/YYYY apenas para compor prefixo do código
        try:
            if "-" in txt:
                d = datetime.fromisoformat(txt[:10])
                base = d.strftime("%Y%m%d")
            elif "/" in txt:
                dd, mm, yy = txt.split("/")[:3]
                yyyy = int(yy) + 2000 if len(yy) == 2 else int(yy)
                d = datetime(yyyy, int(mm), int(dd))
                base = d.strftime("%Y%m%d")
        except Exception:
            pass

    prefix = f"PGA{base}"
    cur.execute(
        """
        SELECT codigo_avulsa
        FROM programacoes_avulsas
        WHERE codigo_avulsa LIKE ?
        ORDER BY id DESC
        LIMIT 1
        """,
        (f"{prefix}-%",),
    )
    row = cur.fetchone()
    last_seq = 0
    if row and row[0]:
        s = str(row[0]).strip()
        if "-" in s:
            try:
                last_seq = int(s.split("-")[-1])
            except Exception:
                last_seq = 0
    return f"{prefix}-{last_seq + 1:03d}"

# =========================================================
# ENDPOINTS BÁSICOS
# =========================================================
@app.get("/ping")
def ping():
    return {"ok": True, "db": DB_PATH}


@app.get("/desktop/cadastros/motoristas")
def desktop_motoristas(
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='motoristas'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        company_id = _desktop_company_id(cur, x_company_id)
        clauses: List[str] = []
        params: List[Any] = []
        if "status" in cols:
            clauses.append("UPPER(COALESCE(status,'ATIVO')) IN ('ATIVO','DESATIVADO')")
        if "company_id" in cols:
            clauses.append("company_id=?")
            params.append(company_id)
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        perfil_expr = "UPPER(TRIM(COALESCE(perfil_app,'MOTORISTA')))" if "perfil_app" in cols else "'MOTORISTA'"
        cur.execute(
            f"""
            SELECT id, COALESCE(codigo,''), COALESCE(nome,''), COALESCE(status,'ATIVO'), {perfil_expr} AS perfil_app
            FROM motoristas
            {where}
            ORDER BY UPPER(COALESCE(nome,'')), id
            """,
            tuple(params),
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "codigo": str(r[1] or "").strip().upper(),
                    "nome": str(r[2] or "").strip().upper(),
                    "status": str(r[3] or "ATIVO").strip().upper(),
                    "perfil_app": str(r[4] or "MOTORISTA").strip().upper() or "MOTORISTA",
                }
            )
        return out


@app.get("/desktop/cadastros/vendedores")
def desktop_vendedores(
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendedores'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(vendedores)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        company_id = _desktop_company_id(cur, x_company_id)
        where = "WHERE company_id=?" if "company_id" in cols else ""
        params = (company_id,) if "company_id" in cols else ()
        status_expr = "COALESCE(status,'ATIVO')" if "status" in cols else "'ATIVO'"
        telefone_expr = "COALESCE(telefone,'')" if "telefone" in cols else "''"
        cidade_expr = "COALESCE(cidade_base,'')" if "cidade_base" in cols else "''"
        cur.execute(
            f"""
            SELECT
                id,
                COALESCE(codigo,''),
                COALESCE(nome,''),
                {telefone_expr} AS telefone,
                {cidade_expr} AS cidade_base,
                {status_expr} AS status
            FROM vendedores
            {where}
            ORDER BY UPPER(COALESCE(nome,'')), UPPER(COALESCE(codigo,'')), id
            """,
            params,
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "codigo": str(r[1] or "").strip().upper(),
                    "nome": str(r[2] or "").strip().upper(),
                    "telefone": str(r[3] or "").strip(),
                    "cidade_base": str(r[4] or "").strip().upper(),
                    "status": str(r[5] or "ATIVO").strip().upper(),
                }
            )
        return out


@app.get("/desktop/cadastros/veiculos")
def desktop_veiculos(
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='veiculos'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(veiculos)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        company_id = _desktop_company_id(cur, x_company_id)
        where = "WHERE company_id=?" if "company_id" in cols else ""
        params = (company_id,) if "company_id" in cols else ()
        status_expr = "COALESCE(status,'ATIVO')" if "status" in cols else "'ATIVO'"
        cur.execute(
            f"""
            SELECT id, COALESCE(placa,''), COALESCE(modelo,''), COALESCE(capacidade_cx, 0), {status_expr}
            FROM veiculos
            {where}
            ORDER BY UPPER(COALESCE(placa,'')), id
            """,
            params,
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "placa": str(r[1] or "").strip().upper(),
                    "modelo": str(r[2] or "").strip().upper(),
                    "capacidade_cx": int(r[3] or 0),
                    "status": str(r[4] or "ATIVO").strip().upper(),
                }
            )
        return out


@app.get("/desktop/cadastros/ajudantes")
def desktop_ajudantes(
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ajudantes'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(ajudantes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        company_id = _desktop_company_id(cur, x_company_id)
        clauses: List[str] = []
        params: List[Any] = []
        if "status" in cols:
            clauses.append("UPPER(COALESCE(status,'ATIVO'))='ATIVO'")
        if "company_id" in cols:
            clauses.append("company_id=?")
            params.append(company_id)
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        tel_expr = "COALESCE(telefone,'')" if "telefone" in cols else "''"
        cur.execute(
            f"""
            SELECT id, COALESCE(nome,''), COALESCE(sobrenome,''), {tel_expr} AS telefone, COALESCE(status,'ATIVO')
            FROM ajudantes
            {where}
            ORDER BY UPPER(COALESCE(nome,'')), UPPER(COALESCE(sobrenome,'')), id
            """,
            tuple(params),
        )
        out = []
        for r in cur.fetchall() or []:
            nome_base = str(r[1] or "").strip().upper()
            sobrenome = str(r[2] or "").strip().upper()
            nome = f"{nome_base} {sobrenome}".strip().upper()
            out.append(
                {
                    "id": int(r[0] or 0),
                    "nome": nome,
                    "nome_base": nome_base,
                    "sobrenome": sobrenome,
                    "telefone": str(r[3] or "").strip(),
                    "status": str(r[4] or "ATIVO").strip().upper(),
                }
            )
        return out


@app.post("/desktop/cadastros/motoristas/upsert")
def desktop_motoristas_upsert(
    payload: DesktopMotoristaUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    codigo = _clean_text(payload.codigo).upper()
    nome = _clean_text(payload.nome).upper()
    if not codigo or not nome:
        raise HTTPException(status_code=400, detail="codigo e nome sao obrigatorios.")
    if not is_valid_motorista_codigo(codigo):
        raise HTTPException(status_code=400, detail="codigo invalido. Use letras/numeros/._- e 3 a 24 caracteres.")
    if len(nome) < 3:
        raise HTTPException(status_code=400, detail="nome deve ter pelo menos 3 caracteres.")

    status = _clean_text(payload.status or "ATIVO").upper()
    if status == "DESATIVADO":
        status = "INATIVO"
    if status not in {"ATIVO", "INATIVO"}:
        raise HTTPException(status_code=400, detail="status invalido. Use ATIVO ou INATIVO.")
    perfil_app = _clean_text(payload.perfil_app or "MOTORISTA").upper()
    if perfil_app not in {"MOTORISTA", "ADMIN"}:
        raise HTTPException(status_code=400, detail="perfil_app invalido. Use MOTORISTA ou ADMIN.")

    telefone = _require_desktop_phone(payload.telefone, "telefone invalido. Informe DDD+numero.")
    cpf = _optional_desktop_cpf(payload.cpf)

    senha_in = _clean_text(payload.senha)
    senha_hash = ""
    if senha_in:
        senha_in = _validate_desktop_password(senha_in, "senha invalida. Use 4 a 24 caracteres.")
        senha_hash = senha_in if senha_in.startswith("pbkdf2_sha256$") else hash_password_pbkdf2(senha_in)

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela motoristas indisponivel.")

        company_id = _desktop_company_id(cur, x_company_id)
        if "company_id" in cols:
            cur.execute(
                "SELECT id, COALESCE(senha,'') AS senha FROM motoristas WHERE UPPER(TRIM(codigo))=? AND company_id=? LIMIT 1",
                (codigo, company_id),
            )
        else:
            cur.execute("SELECT id, COALESCE(senha,'') AS senha FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1", (codigo,))
        existing = cur.fetchone()
        if not existing and not senha_hash:
            raise HTTPException(status_code=400, detail="senha obrigatoria para novo motorista.")
        if cpf and "cpf" in cols:
            cpf_scope_sql, cpf_scope_params = _company_scope_condition(cur, "motoristas", company_id)
            cur.execute(
                f"""
                SELECT id FROM motoristas
                WHERE cpf=?
                {f' AND {cpf_scope_sql}' if cpf_scope_sql else ''}
                {'' if not existing else ' AND id<>?'}
                LIMIT 1
                """,
                (cpf, *cpf_scope_params, *([] if not existing else [int(existing["id"])])),
            )
            if cur.fetchone():
                raise HTTPException(status_code=409, detail="Ja existe motorista com este CPF.")

        set_parts: List[str] = []
        params: List[Any] = []

        if "nome" in cols:
            set_parts.append("nome=?")
            params.append(nome)
        if "telefone" in cols:
            set_parts.append("telefone=?")
            params.append(telefone)
        if "cpf" in cols:
            set_parts.append("cpf=?")
            params.append(cpf)
        if "status" in cols:
            set_parts.append("status=?")
            params.append(status)
        if "perfil_app" in cols:
            set_parts.append("perfil_app=?")
            params.append(perfil_app)
        # Atualização de acesso só quando vier explícito no payload.
        # Isso evita "desbloqueio automático" ao apenas editar cadastro.
        if payload.acesso_liberado is not None:
            if "acesso_liberado" in cols:
                set_parts.append("acesso_liberado=?")
                params.append(1 if bool(payload.acesso_liberado) else 0)
            if "acesso_liberado_por" in cols:
                set_parts.append("acesso_liberado_por=?")
                params.append(_clean_text(payload.acesso_liberado_por or "DESKTOP_SYNC").upper())
            if "acesso_liberado_em" in cols:
                set_parts.append("acesso_liberado_em=?")
                params.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            if "acesso_obs" in cols:
                set_parts.append("acesso_obs=?")
                params.append(_clean_text(payload.acesso_obs or "Sincronizado via Desktop"))

        if "senha" in cols:
            current_hash = _clean_text(existing["senha"] if existing else "")
            final_hash = senha_hash or current_hash
            if final_hash:
                set_parts.append("senha=?")
                params.append(final_hash)

        if existing:
            if not set_parts:
                return {"ok": True, "codigo": codigo, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE motoristas SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            return {"ok": True, "codigo": codigo, "updated": int(cur.rowcount or 0)}

        cols_ins: List[str] = ["codigo"]
        vals_ins: List[Any] = [codigo]
        if "nome" in cols:
            cols_ins.append("nome"); vals_ins.append(nome)
        if "telefone" in cols:
            cols_ins.append("telefone"); vals_ins.append(telefone)
        if "cpf" in cols:
            cols_ins.append("cpf"); vals_ins.append(cpf)
        if "status" in cols:
            cols_ins.append("status"); vals_ins.append(status)
        if "perfil_app" in cols:
            cols_ins.append("perfil_app"); vals_ins.append(perfil_app)
        if "senha" in cols:
            cols_ins.append("senha"); vals_ins.append(senha_hash or hash_password_pbkdf2("1234"))
        if "company_id" in cols:
            cols_ins.append("company_id"); vals_ins.append(company_id)
        acesso_novo = True if payload.acesso_liberado is None else bool(payload.acesso_liberado)
        if "acesso_liberado" in cols:
            cols_ins.append("acesso_liberado"); vals_ins.append(1 if acesso_novo else 0)
        if "acesso_liberado_por" in cols:
            cols_ins.append("acesso_liberado_por"); vals_ins.append(_clean_text(payload.acesso_liberado_por or "DESKTOP_SYNC").upper())
        if "acesso_liberado_em" in cols:
            cols_ins.append("acesso_liberado_em"); vals_ins.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        if "acesso_obs" in cols:
            cols_ins.append("acesso_obs"); vals_ins.append(_clean_text(payload.acesso_obs or "Sincronizado via Desktop"))

        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO motoristas ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "codigo": codigo, "created": 1}


@app.post("/desktop/cadastros/vendedores/upsert")
def desktop_vendedores_upsert(
    payload: DesktopVendedorUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    codigo = _clean_text(payload.codigo).upper()
    nome = _clean_text(payload.nome).upper()
    if not codigo or not nome:
        raise HTTPException(status_code=400, detail="codigo e nome sao obrigatorios.")
    if not is_valid_motorista_codigo(codigo):
        raise HTTPException(status_code=400, detail="codigo invalido. Use letras/numeros/._- e 3 a 24 caracteres.")
    if len(nome) < 3:
        raise HTTPException(status_code=400, detail="nome deve ter pelo menos 3 caracteres.")

    status = _clean_text(payload.status or "ATIVO").upper()
    if status not in {"ATIVO", "DESATIVADO"}:
        raise HTTPException(status_code=400, detail="status invalido. Use ATIVO ou DESATIVADO.")

    telefone = _optional_desktop_phone(payload.telefone, "telefone invalido. Informe DDD+numero ou deixe vazio.")
    cidade_base = _clean_text(payload.cidade_base).upper()
    senha_in = _clean_text(payload.senha)
    senha_hash = ""
    if senha_in:
        senha_in = _validate_desktop_password(senha_in, "senha invalida. Use 4 a 24 caracteres.")
        senha_hash = senha_in if senha_in.startswith("pbkdf2_sha256$") else hash_password_pbkdf2(senha_in)

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(vendedores)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela vendedores indisponivel.")

        company_id = _desktop_company_id(cur, x_company_id)
        if "company_id" in cols:
            cur.execute(
                "SELECT id, COALESCE(senha,'') AS senha FROM vendedores WHERE UPPER(TRIM(codigo))=? AND company_id=? LIMIT 1",
                (codigo, company_id),
            )
        else:
            cur.execute(
                "SELECT id, COALESCE(senha,'') AS senha FROM vendedores WHERE UPPER(TRIM(codigo))=? LIMIT 1",
                (codigo,),
            )
        existing = cur.fetchone()
        if not existing and not senha_hash:
            raise HTTPException(status_code=400, detail="senha obrigatoria para novo vendedor.")

        set_parts: List[str] = []
        params: List[Any] = []
        if "codigo" in cols:
            set_parts.append("codigo=?")
            params.append(codigo)
        if "nome" in cols:
            set_parts.append("nome=?")
            params.append(nome)
        if "telefone" in cols:
            set_parts.append("telefone=?")
            params.append(telefone)
        if "cidade_base" in cols:
            set_parts.append("cidade_base=?")
            params.append(cidade_base)
        if "status" in cols:
            set_parts.append("status=?")
            params.append(status)
        if "senha" in cols:
            current_hash = _clean_text(existing["senha"] if existing else "")
            final_hash = senha_hash or current_hash
            if final_hash:
                set_parts.append("senha=?")
                params.append(final_hash)

        if existing:
            if not set_parts:
                return {"ok": True, "codigo": codigo, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE vendedores SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            return {"ok": True, "codigo": codigo, "updated": int(cur.rowcount or 0)}

        cols_ins: List[str] = ["codigo"]
        vals_ins: List[Any] = [codigo]
        if "nome" in cols:
            cols_ins.append("nome"); vals_ins.append(nome)
        if "telefone" in cols:
            cols_ins.append("telefone"); vals_ins.append(telefone)
        if "cidade_base" in cols:
            cols_ins.append("cidade_base"); vals_ins.append(cidade_base)
        if "status" in cols:
            cols_ins.append("status"); vals_ins.append(status)
        if "senha" in cols:
            cols_ins.append("senha"); vals_ins.append(senha_hash or hash_password_pbkdf2("1234"))
        if "company_id" in cols:
            cols_ins.append("company_id"); vals_ins.append(company_id)
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO vendedores ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "codigo": codigo, "created": 1}


@app.post("/desktop/cadastros/veiculos/upsert")
def desktop_veiculos_upsert(
    payload: DesktopVeiculoUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    placa = _clean_text(payload.placa).upper()
    modelo = _clean_text(payload.modelo).upper()
    capacidade_cx = int(payload.capacidade_cx or 0)
    status = _clean_text(payload.status or "ATIVO").upper() or "ATIVO"
    if status not in {"ATIVO", "DESATIVADO"}:
        raise HTTPException(status_code=400, detail="status invalido. Use ATIVO ou DESATIVADO.")
    if not placa or not modelo:
        raise HTTPException(status_code=400, detail="placa e modelo sao obrigatorios.")
    placa_ok, placa_msg = validate_placa(placa)
    if not placa_ok:
        raise HTTPException(status_code=400, detail=placa_msg or "placa invalida.")
    if capacidade_cx < 0:
        raise HTTPException(status_code=400, detail="capacidade_cx deve ser inteiro >= 0.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(veiculos)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela veiculos indisponivel.")

        company_id = _desktop_company_id(cur, x_company_id)
        if "company_id" in cols:
            cur.execute(
                "SELECT id FROM veiculos WHERE UPPER(TRIM(placa))=? AND company_id=? LIMIT 1",
                (placa, company_id),
            )
        else:
            cur.execute("SELECT id FROM veiculos WHERE UPPER(TRIM(placa))=? LIMIT 1", (placa,))
        existing = cur.fetchone()

        set_parts: List[str] = []
        params: List[Any] = []
        if "placa" in cols:
            set_parts.append("placa=?"); params.append(placa)
        if "modelo" in cols:
            set_parts.append("modelo=?"); params.append(modelo)
        if "capacidade_cx" in cols:
            set_parts.append("capacidade_cx=?"); params.append(capacidade_cx)
        if "status" in cols:
            set_parts.append("status=?"); params.append(status)

        if existing:
            if not set_parts:
                return {"ok": True, "placa": placa, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE veiculos SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            return {"ok": True, "placa": placa, "updated": int(cur.rowcount or 0)}

        limit_result = check_vehicle_limit(conn, company_id, exclude_placa=placa)
        if not bool(limit_result.get("ok", False)):
            conn.commit()
            raise HTTPException(status_code=403, detail=limit_result.get("error") or "Limite de veiculos atingido.")

        cols_ins: List[str] = []
        vals_ins: List[Any] = []
        if "placa" in cols:
            cols_ins.append("placa"); vals_ins.append(placa)
        if "modelo" in cols:
            cols_ins.append("modelo"); vals_ins.append(modelo)
        if "capacidade_cx" in cols:
            cols_ins.append("capacidade_cx"); vals_ins.append(capacidade_cx)
        if "status" in cols:
            cols_ins.append("status"); vals_ins.append(status)
        if "company_id" in cols:
            cols_ins.append("company_id"); vals_ins.append(company_id)
        if not cols_ins:
            raise HTTPException(status_code=500, detail="Colunas de veiculos indisponiveis.")
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO veiculos ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "placa": placa, "created": 1}


@app.post("/desktop/cadastros/ajudantes/upsert")
def desktop_ajudantes_upsert(
    payload: DesktopAjudanteUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    nome = _clean_text(payload.nome).upper()
    sobrenome = _clean_text(payload.sobrenome).upper()
    telefone = _clean_text(payload.telefone)
    status = _clean_text(payload.status or "ATIVO").upper()
    if status not in {"ATIVO", "DESATIVADO"}:
        raise HTTPException(status_code=400, detail="status invalido. Use ATIVO ou DESATIVADO.")
    if not nome or not sobrenome:
        raise HTTPException(status_code=400, detail="nome e sobrenome sao obrigatorios.")
    if len(nome) < 2:
        raise HTTPException(status_code=400, detail="nome deve ter pelo menos 2 caracteres.")
    if len(sobrenome) < 2:
        raise HTTPException(status_code=400, detail="sobrenome deve ter pelo menos 2 caracteres.")
    telefone = _require_desktop_phone(telefone, "telefone invalido. Informe DDD+numero.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(ajudantes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela ajudantes indisponivel.")

        company_id = _desktop_company_id(cur, x_company_id)
        if "company_id" in cols:
            cur.execute(
                """
                SELECT id FROM ajudantes
                WHERE UPPER(TRIM(COALESCE(nome,'')))=? AND UPPER(TRIM(COALESCE(sobrenome,'')))=? AND company_id=?
                LIMIT 1
                """,
                (nome, sobrenome, company_id),
            )
        else:
            cur.execute(
                """
                SELECT id FROM ajudantes
                WHERE UPPER(TRIM(COALESCE(nome,'')))=? AND UPPER(TRIM(COALESCE(sobrenome,'')))=?
                LIMIT 1
                """,
                (nome, sobrenome),
            )
        existing = cur.fetchone()

        set_parts: List[str] = []
        params: List[Any] = []
        if "nome" in cols:
            set_parts.append("nome=?"); params.append(nome)
        if "sobrenome" in cols:
            set_parts.append("sobrenome=?"); params.append(sobrenome)
        if "telefone" in cols:
            set_parts.append("telefone=?"); params.append(telefone)
        if "status" in cols:
            set_parts.append("status=?"); params.append(status)

        if existing:
            if not set_parts:
                return {"ok": True, "nome": nome, "sobrenome": sobrenome, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE ajudantes SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            return {"ok": True, "nome": nome, "sobrenome": sobrenome, "updated": int(cur.rowcount or 0)}

        cols_ins: List[str] = []
        vals_ins: List[Any] = []
        if "nome" in cols:
            cols_ins.append("nome"); vals_ins.append(nome)
        if "sobrenome" in cols:
            cols_ins.append("sobrenome"); vals_ins.append(sobrenome)
        if "telefone" in cols:
            cols_ins.append("telefone"); vals_ins.append(telefone)
        if "status" in cols:
            cols_ins.append("status"); vals_ins.append(status)
        if "company_id" in cols:
            cols_ins.append("company_id"); vals_ins.append(company_id)
        if not cols_ins:
            raise HTTPException(status_code=500, detail="Colunas de ajudantes indisponiveis.")
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO ajudantes ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "nome": nome, "sobrenome": sobrenome, "created": 1}


def _desktop_cliente_upsert_cur(
    cur: sqlite3.Cursor,
    payload: DesktopClienteUpsertIn,
    cols: set[str] | None = None,
    company_id: Optional[int] = None,
) -> Dict[str, Any]:
    cod_cliente = _clean_text(payload.cod_cliente).upper()
    nome_cliente = _clean_text(payload.nome_cliente).upper()
    if not cod_cliente or not nome_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente e nome_cliente sao obrigatorios.")

    cols = cols or _ensure_clientes_table(cur)
    if not cols:
        raise HTTPException(status_code=500, detail="Tabela clientes indisponivel.")

    endereco = _clean_text(payload.endereco).upper()
    telefone = _clean_text(payload.telefone).upper()
    vendedor = _clean_text(payload.vendedor).upper()

    scope_sql, scope_params = _company_scope_condition(cur, "clientes", company_id)
    scope_extra = f" AND {scope_sql}" if scope_sql else ""
    cur.execute(
        f"""
        SELECT
            rowid AS row_ref,
            COALESCE(nome_cliente,'') AS nome_cliente,
            COALESCE(endereco,'') AS endereco,
            COALESCE(telefone,'') AS telefone,
            COALESCE(vendedor,'') AS vendedor
        FROM clientes
        WHERE UPPER(TRIM(cod_cliente))=?
        {scope_extra}
        LIMIT 1
        """,
        (cod_cliente, *scope_params),
    )
    existing = cur.fetchone()

    if existing:
        endereco_final = endereco or str(existing["endereco"] or "").strip().upper()
        telefone_final = telefone or str(existing["telefone"] or "").strip().upper()
        vendedor_final = vendedor or str(existing["vendedor"] or "").strip().upper()
        set_parts: List[str] = []
        params: List[Any] = []
        if "cod_cliente" in cols:
            set_parts.append("cod_cliente=?"); params.append(cod_cliente)
        if "nome_cliente" in cols:
            set_parts.append("nome_cliente=?"); params.append(nome_cliente)
        if "endereco" in cols:
            set_parts.append("endereco=?"); params.append(endereco_final)
        if "telefone" in cols:
            set_parts.append("telefone=?"); params.append(telefone_final)
        if "vendedor" in cols:
            set_parts.append("vendedor=?"); params.append(vendedor_final)
        if not set_parts:
            return {"ok": True, "cod_cliente": cod_cliente, "updated": 0}
        params.append(int(existing["row_ref"]))
        cur.execute(f"UPDATE clientes SET {', '.join(set_parts)} WHERE rowid=?", tuple(params))
        return {"ok": True, "cod_cliente": cod_cliente, "updated": int(cur.rowcount or 0)}

    cols_ins: List[str] = []
    vals_ins: List[Any] = []
    if "cod_cliente" in cols:
        cols_ins.append("cod_cliente"); vals_ins.append(cod_cliente)
    if "nome_cliente" in cols:
        cols_ins.append("nome_cliente"); vals_ins.append(nome_cliente)
    if "endereco" in cols:
        cols_ins.append("endereco"); vals_ins.append(endereco)
    if "telefone" in cols:
        cols_ins.append("telefone"); vals_ins.append(telefone)
    if "vendedor" in cols:
        cols_ins.append("vendedor"); vals_ins.append(vendedor)
    if "company_id" in cols:
        cols_ins.append("company_id"); vals_ins.append(company_id or _default_company_id(cur))
    if not cols_ins:
        raise HTTPException(status_code=500, detail="Colunas de clientes indisponiveis.")
    ph = ", ".join(["?"] * len(cols_ins))
    cur.execute(f"INSERT INTO clientes ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
    return {"ok": True, "cod_cliente": cod_cliente, "created": 1}


@app.post("/desktop/cadastros/clientes/upsert")
def desktop_clientes_upsert(
    payload: DesktopClienteUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        return _desktop_cliente_upsert_cur(cur, payload, company_id=company_id)


@app.post("/desktop/cadastros/clientes/bulk-upsert")
def desktop_clientes_bulk_upsert(
    payload: DesktopClientesBulkUpsertIn,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    itens = payload.clientes or []
    if not itens:
        return {"ok": True, "total": 0, "created": 0, "updated": 0, "falhas": []}

    total = 0
    created = 0
    updated = 0
    falhas: List[Dict[str, Any]] = []
    with get_conn() as conn:
        cur = conn.cursor()
        cols = _ensure_clientes_table(cur)
        company_id = _desktop_company_id(cur, x_company_id)
        for idx, item in enumerate(itens, start=1):
            try:
                result = _desktop_cliente_upsert_cur(cur, item, cols=cols, company_id=company_id)
                total += 1
                created += int(result.get("created") or 0)
                updated += int(result.get("updated") or 0)
            except HTTPException as exc:
                falhas.append(
                    {
                        "index": idx,
                        "cod_cliente": _clean_text(getattr(item, "cod_cliente", "")),
                        "detail": exc.detail,
                    }
                )
            except Exception as exc:
                logging.exception("Falha ao importar cliente em lote")
                falhas.append(
                    {
                        "index": idx,
                        "cod_cliente": _clean_text(getattr(item, "cod_cliente", "")),
                        "detail": str(exc or "Falha inesperada."),
                    }
                )

        if falhas:
            raise HTTPException(
                status_code=400,
                detail={
                    "message": "Falha ao salvar alguns clientes.",
                    "salvos": total,
                    "total_falhas": len(falhas),
                    "falhas": falhas[:20],
                },
            )
    return {"ok": True, "total": total, "created": created, "updated": updated, "falhas": []}


@app.delete("/desktop/cadastros/motoristas/{codigo}")
def desktop_motoristas_delete(
    codigo: str,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    cod = _clean_text(codigo).upper()
    if not cod:
        raise HTTPException(status_code=400, detail="codigo obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        bloqueio = _cadastro_delete_block_reason(cur, "motoristas", codigo=cod, company_id=company_id)
        if bloqueio:
            raise HTTPException(status_code=409, detail=f"{bloqueio} Use status DESATIVADO.")
        scope_sql, scope_params = _company_scope_condition(cur, "motoristas", company_id)
        cur.execute(
            f"DELETE FROM motoristas WHERE UPPER(TRIM(COALESCE(codigo,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
            (cod, *scope_params),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "codigo": cod, "deleted": deleted}


@app.delete("/desktop/cadastros/vendedores/{codigo}")
def desktop_vendedores_delete(
    codigo: str,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    cod = _clean_text(codigo).upper()
    if not cod:
        raise HTTPException(status_code=400, detail="codigo obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        bloqueio = _cadastro_delete_block_reason(cur, "vendedores", codigo=cod, company_id=company_id)
        if bloqueio:
            raise HTTPException(status_code=409, detail=f"{bloqueio} Use status DESATIVADO.")
        scope_sql, scope_params = _company_scope_condition(cur, "vendedores", company_id)
        cur.execute(
            f"DELETE FROM vendedores WHERE UPPER(TRIM(COALESCE(codigo,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
            (cod, *scope_params),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "codigo": cod, "deleted": deleted}


@app.delete("/desktop/cadastros/veiculos/{placa}")
def desktop_veiculos_delete(
    placa: str,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    plc = _clean_text(placa).upper()
    if not plc:
        raise HTTPException(status_code=400, detail="placa obrigatoria.")
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        bloqueio = _cadastro_delete_block_reason(cur, "veiculos", placa=plc, company_id=company_id)
        if bloqueio:
            raise HTTPException(status_code=409, detail=f"{bloqueio} Use status DESATIVADO.")
        scope_sql, scope_params = _company_scope_condition(cur, "veiculos", company_id)
        cur.execute(
            f"DELETE FROM veiculos WHERE UPPER(TRIM(COALESCE(placa,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
            (plc, *scope_params),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "placa": plc, "deleted": deleted}


@app.delete("/desktop/cadastros/ajudantes/{ajudante_id}")
def desktop_ajudantes_delete(
    ajudante_id: int,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    aid = int(ajudante_id or 0)
    if aid <= 0:
        raise HTTPException(status_code=400, detail="ajudante_id invalido.")
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        bloqueio = _cadastro_delete_block_reason(cur, "ajudantes", ajudante_id=aid, company_id=company_id)
        if bloqueio:
            raise HTTPException(status_code=409, detail=f"{bloqueio} Use status DESATIVADO.")
        scope_sql, scope_params = _company_scope_condition(cur, "ajudantes", company_id)
        cur.execute(
            f"DELETE FROM ajudantes WHERE id=?{f' AND {scope_sql}' if scope_sql else ''}",
            (aid, *scope_params),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "ajudante_id": aid, "deleted": deleted}


@app.delete("/desktop/cadastros/clientes/{cod_cliente}")
def desktop_clientes_delete(
    cod_cliente: str,
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    cod = _clean_text(cod_cliente).upper()
    if not cod:
        raise HTTPException(status_code=400, detail="cod_cliente obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        company_id = _desktop_company_id(cur, x_company_id)
        bloqueio = _cadastro_delete_block_reason(cur, "clientes", cod_cliente=cod, company_id=company_id)
        if bloqueio:
            raise HTTPException(status_code=409, detail=f"{bloqueio} Use cadastro ativo/inativo em vez de excluir.")
        scope_sql, scope_params = _company_scope_condition(cur, "clientes", company_id)
        cur.execute(
            f"DELETE FROM clientes WHERE UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
            (cod, *scope_params),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "cod_cliente": cod, "deleted": deleted}


@app.get("/desktop/clientes/base")
def desktop_clientes_base(
    q: str = Query("", description="Busca por codigo/nome/cidade"),
    vendedor: str = Query("", description="Filtro por vendedor"),
    cidade: str = Query("", description="Filtro por cidade"),
    ordem: str = Query("nome", description="nome|codigo"),
    limit: int = Query(300, ge=1, le=1000),
    _ok: bool = Depends(_require_desktop_secret),
    x_company_id: Optional[str] = Header(default=None, alias="X-Company-ID"),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
        if not cur.fetchone():
            return []
        cols = _ensure_clientes_columns(cur)
        company_id = _desktop_company_id(cur, x_company_id)
        scope_sql, scope_params = _company_scope_condition(cur, "clientes", company_id)

        col_cod = "cod_cliente" if "cod_cliente" in cols else "''"
        col_nome = "nome_cliente" if "nome_cliente" in cols else "''"
        col_end = "endereco" if "endereco" in cols else "''"
        col_tel = "telefone" if "telefone" in cols else "''"
        col_cid = "cidade" if "cidade" in cols else "''"
        col_bai = "bairro" if "bairro" in cols else "''"
        col_vend = "vendedor" if "vendedor" in cols else "''"
        col_preco = "preco" if "preco" in cols else "''"
        col_caixas = "caixas" if "caixas" in cols else "''"

        term = (q or "").strip().upper()
        vend_f = (vendedor or "").strip().upper()
        cid_f = (cidade or "").strip().upper()
        ordem_sql = (
            "UPPER(TRIM(COALESCE(cod_cliente, ''))), UPPER(TRIM(COALESCE(nome_cliente, '')))"
            if (ordem or "").strip().lower() == "codigo"
            else "UPPER(TRIM(COALESCE(nome_cliente, ''))), UPPER(TRIM(COALESCE(cod_cliente, '')))"
        )
        like = f"%{term}%"
        like_vend = f"%{vend_f}%"
        like_cid = f"%{cid_f}%"
        cur.execute(
            f"""
            SELECT
                TRIM(COALESCE({col_cod}, '')) AS cod_cliente,
                TRIM(COALESCE({col_nome}, '')) AS nome_cliente,
                TRIM(COALESCE({col_end}, '')) AS endereco,
                TRIM(COALESCE({col_tel}, '')) AS telefone,
                TRIM(COALESCE({col_cid}, '')) AS cidade,
                TRIM(COALESCE({col_bai}, '')) AS bairro,
                TRIM(COALESCE({col_vend}, '')) AS vendedor,
                TRIM(COALESCE({col_preco}, '')) AS preco,
                TRIM(COALESCE({col_caixas}, '')) AS caixas
            FROM clientes
            WHERE
                (
                    (? = '')
                    OR UPPER(TRIM(COALESCE({col_cod}, ''))) LIKE ?
                    OR UPPER(TRIM(COALESCE({col_nome}, ''))) LIKE ?
                    OR UPPER(TRIM(COALESCE({col_cid}, ''))) LIKE ?
                )
            AND ((? = '') OR UPPER(TRIM(COALESCE({col_vend}, ''))) LIKE ?)
            AND ((? = '') OR UPPER(TRIM(COALESCE({col_cid}, ''))) LIKE ?)
            {f'AND {scope_sql}' if scope_sql else ''}
            ORDER BY {ordem_sql}
            LIMIT ?
            """,
            (term, like, like, like, vend_f, like_vend, cid_f, like_cid, *scope_params, int(limit)),
        )
        out: List[Dict[str, Any]] = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "cod_cliente": str(r["cod_cliente"] or "").strip(),
                    "nome_cliente": str(r["nome_cliente"] or "").strip(),
                    "endereco": str(r["endereco"] or "").strip(),
                    "telefone": str(r["telefone"] or "").strip(),
                    "cidade": str(r["cidade"] or "").strip(),
                    "bairro": str(r["bairro"] or "").strip(),
                    "vendedor": str(r["vendedor"] or "").strip(),
                    "preco": str(r["preco"] or "").strip(),
                    "caixas": str(r["caixas"] or "").strip(),
                }
            )
        return out


@app.post("/desktop/programacoes/reconciliar-vinculos")
def desktop_reconciliar_vinculos_motorista(_ok: bool = Depends(_require_desktop_secret)):
    fixed = reconcile_programacoes_motorista_links()
    return {"ok": True, "rotas_ajustadas": int(fixed)}


@app.post("/desktop/rotas/upsert")
def desktop_rotas_upsert(payload: DesktopRotaUpsertIn, _ok: bool = Depends(_require_desktop_secret)):
    codigo = str(payload.codigo_programacao or "").strip().upper()
    if not codigo:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        data_criacao = (payload.data_criacao or datetime.now().strftime("%Y-%m-%d %H:%M:%S")).strip()
        motorista = str(payload.motorista or "").strip().upper()
        motorista_id = int(payload.motorista_id or 0)
        motorista_codigo = str(payload.motorista_codigo or payload.codigo_motorista or "").strip().upper()
        veiculo = str(payload.veiculo or "").strip().upper()
        equipe = str(payload.equipe or "").strip().upper()
        status = str(payload.status or "ATIVA").strip().upper() or "ATIVA"
        tipo_estimativa = str(payload.tipo_estimativa or "KG").strip().upper()
        local_rota = str(payload.local_rota or payload.tipo_rota or "").strip().upper()
        local_carregamento = str(
            payload.local_carregamento
            or payload.local_carregado
            or payload.granja_carregada
            or payload.local_carreg
            or ""
        ).strip().upper()
        adiantamento_origem = str(payload.adiantamento_origem or "").strip().upper()
        nf_kg = float(payload.nf_kg or 0.0)
        nf_preco = float(payload.nf_preco or 0.0)
        caixas_carregadas = int(
            payload.caixas_carregadas
            if payload.caixas_carregadas is not None
            else (payload.nf_caixas or 0)
        )
        linked_venda_ids = []
        for raw_id in (payload.linked_venda_ids or []):
            try:
                rid = int(raw_id)
            except Exception:
                continue
            if rid > 0 and rid not in linked_venda_ids:
                linked_venda_ids.append(rid)
        vendas_usada_em = str(payload.vendas_usada_em or data_criacao or "").strip() or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if tipo_estimativa not in ("KG", "CX"):
            tipo_estimativa = "KG"
        operacao_tipo = str(payload.operacao_tipo or "").strip().upper().replace("-", "_").replace(" ", "_")
        if operacao_tipo not in ("TRANSBORDO", "VENDA"):
            operacao_tipo = "TRANSBORDO" if tipo_estimativa == "CX" else "VENDA"
        transbordo_modalidade = "EMPRESA_BUSCA" if operacao_tipo == "TRANSBORDO" else "CIF"
        transbordo_grupo = (
            str(payload.transbordo_grupo or codigo).strip().upper()
            if operacao_tipo == "TRANSBORDO"
            else ""
        )

        select_parts = [
            "id",
            "COALESCE(status,'') AS status",
            "COALESCE(status_operacional,'') AS status_operacional",
            "COALESCE(motorista,'') AS motorista",
        ]
        if "motorista_id" in cols_prog:
            select_parts.append("COALESCE(motorista_id,0) AS motorista_id")
        else:
            select_parts.append("0 AS motorista_id")
        if "motorista_codigo" in cols_prog:
            select_parts.append("COALESCE(motorista_codigo,'') AS motorista_codigo")
        else:
            select_parts.append("'' AS motorista_codigo")
        if "codigo_motorista" in cols_prog:
            select_parts.append("COALESCE(codigo_motorista,'') AS codigo_motorista")
        else:
            select_parts.append("'' AS codigo_motorista")

        cur.execute(
            f"""
            SELECT {", ".join(select_parts)}
            FROM programacoes
            WHERE codigo_programacao=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        )
        row = cur.fetchone()
        pid = int(row["id"] or 0) if row else 0
        status_atual = str((row["status"] if row else "") or "").strip().upper()
        status_op_atual = str((row["status_operacional"] if row else "") or "").strip().upper()
        motorista_atual = str((row["motorista"] if row else "") or "").strip().upper()
        motorista_id_atual = int((row["motorista_id"] if row else 0) or 0)
        motorista_codigo_atual = str(
            ((row["motorista_codigo"] if row else "") or (row["codigo_motorista"] if row else "")) or ""
        ).strip().upper()
        state_atual = _programacao_state(cur, codigo) if pid > 0 else None

        status_execucao = {"EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"}
        status_fechado = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}
        status_edicao_desktop = {"", "ATIVA", "ABERTA", "PENDENTE", "PROGRAMADA"}

        # Regra de negocio: desktop edita somente programação ainda não iniciada.
        # Se já está em execução, substituição deve ser via APK.
        if pid > 0:
            if state_atual and str(state_atual.get("prestacao_status") or "").upper() == "FECHADA":
                raise HTTPException(
                    status_code=409,
                    detail=f"Programacao {codigo} esta com a prestacao FECHADA e nao pode ser alterada pelo desktop.",
                )
            st_eff = status_op_atual if status_op_atual else status_atual
            if st_eff in status_execucao:
                raise HTTPException(
                    status_code=409,
                    detail=f"Programacao {codigo} ja esta em execucao ({st_eff}); altere via APK (substituir rota/transferir caixas).",
                )
            if st_eff in status_fechado:
                raise HTTPException(
                    status_code=409,
                    detail=f"Programacao {codigo} ja esta encerrada ({st_eff}) e nao pode ser reaberta pelo desktop.",
                )

            # Não deixa payload legacy degradar status para ATIVA indevidamente.
            if status not in status_edicao_desktop:
                status = status_atual or "ATIVA"

        if pid > 0:
            if not motorista and motorista_atual:
                motorista = motorista_atual
            if motorista_id <= 0 and motorista_id_atual > 0:
                motorista_id = motorista_id_atual
            if not motorista_codigo and motorista_codigo_atual:
                motorista_codigo = motorista_codigo_atual

        motorista, motorista_id, motorista_codigo = _resolve_motorista_vinculo(
            cur,
            motorista_nome=motorista,
            motorista_id=motorista_id,
            motorista_codigo=motorista_codigo,
        )

        has_stable_motorista_link = any(
            col in cols_prog for col in ("motorista_id", "motorista_codigo", "codigo_motorista")
        )
        if has_stable_motorista_link:
            if motorista_id <= 0 or not motorista_codigo:
                raise HTTPException(
                    status_code=400,
                    detail="Programacao oficial exige motorista valido para vinculo com o app do motorista.",
                )
        elif not motorista:
            raise HTTPException(
                status_code=400,
                detail="Programacao oficial exige motorista informado.",
            )

        if pid > 0:
            sets = [
                "motorista=?",
                "veiculo=?",
                "equipe=?",
                "kg_estimado=?",
                "status=?",
            ]
            vals: List[Any] = [
                motorista,
                veiculo,
                equipe,
                float(payload.kg_estimado or 0.0),
                status,
            ]
            if "motorista_id" in cols_prog:
                sets.append("motorista_id=?")
                vals.append(int(motorista_id or 0))
            if "motorista_codigo" in cols_prog:
                sets.append("motorista_codigo=?")
                vals.append(motorista_codigo)
            if "codigo_motorista" in cols_prog:
                sets.append("codigo_motorista=?")
                vals.append(motorista_codigo)
            if "tipo_estimativa" in cols_prog:
                sets.append("tipo_estimativa=?")
                vals.append(tipo_estimativa)
            if "caixas_estimado" in cols_prog:
                sets.append("caixas_estimado=?")
                vals.append(int(payload.caixas_estimado or 0))
            if "codigo" in cols_prog:
                sets.append("codigo=COALESCE(NULLIF(TRIM(codigo), ''), ?)")
                vals.append(codigo)
            if "data" in cols_prog:
                sets.append("data=COALESCE(NULLIF(TRIM(data), ''), ?)")
                vals.append(data_criacao)
            if "operacao_tipo" in cols_prog:
                sets.append("operacao_tipo=?")
                vals.append(operacao_tipo)
            if "transbordo_modalidade" in cols_prog:
                sets.append("transbordo_modalidade=?")
                vals.append(transbordo_modalidade)
            if "transbordo_grupo" in cols_prog:
                sets.append("transbordo_grupo=?")
                vals.append(transbordo_grupo)
            if "local_rota" in cols_prog:
                sets.append("local_rota=?")
                vals.append(local_rota)
            if "tipo_rota" in cols_prog:
                sets.append("tipo_rota=?")
                vals.append(local_rota)
            if "local_carregamento" in cols_prog:
                sets.append("local_carregamento=?")
                vals.append(local_carregamento)
            if "granja_carregada" in cols_prog:
                sets.append("granja_carregada=?")
                vals.append(local_carregamento)
            if "local_carregado" in cols_prog:
                sets.append("local_carregado=?")
                vals.append(local_carregamento)
            if "local_carreg" in cols_prog:
                sets.append("local_carreg=?")
                vals.append(local_carregamento)
            if "adiantamento" in cols_prog:
                sets.append("adiantamento=?")
                vals.append(float(payload.adiantamento or 0.0))
            if "adiantamento_rota" in cols_prog:
                sets.append("adiantamento_rota=?")
                vals.append(float(payload.adiantamento or 0.0))
            if "adiantamento_origem" in cols_prog:
                sets.append("adiantamento_origem=?")
                vals.append(adiantamento_origem)
            if "pix_motorista" in cols_prog:
                sets.append("pix_motorista=?")
                vals.append(float(payload.pix_motorista or 0.0))
            if "total_caixas" in cols_prog:
                sets.append("total_caixas=?")
                vals.append(int(payload.total_caixas or 0))
            if "quilos" in cols_prog:
                sets.append("quilos=?")
                vals.append(float(payload.quilos or 0.0))
            if "nf_kg" in cols_prog and nf_kg > 0:
                sets.append("nf_kg=?")
                vals.append(nf_kg)
            if "kg_nf" in cols_prog and nf_kg > 0:
                sets.append("kg_nf=?")
                vals.append(nf_kg)
            if "nf_preco" in cols_prog and nf_preco > 0:
                sets.append("nf_preco=?")
                vals.append(nf_preco)
            if "preco_nf" in cols_prog and nf_preco > 0:
                sets.append("preco_nf=?")
                vals.append(nf_preco)
            if "caixas_carregadas" in cols_prog and caixas_carregadas > 0:
                sets.append("caixas_carregadas=?")
                vals.append(caixas_carregadas)
            if "qnt_cx_carregada" in cols_prog and caixas_carregadas > 0:
                sets.append("qnt_cx_carregada=?")
                vals.append(caixas_carregadas)
            if "nf_caixas" in cols_prog and caixas_carregadas > 0:
                sets.append("nf_caixas=?")
                vals.append(caixas_carregadas)
            if "usuario_ultima_edicao" in cols_prog:
                sets.append("usuario_ultima_edicao=?")
                vals.append(str(payload.usuario_ultima_edicao or payload.usuario_criacao or "").strip().upper())

            vals.append(pid)
            cur.execute(f"UPDATE programacoes SET {', '.join(sets)} WHERE id=?", tuple(vals))
        else:
            col_names = ["codigo_programacao", "data_criacao", "motorista", "veiculo", "equipe", "kg_estimado", "status"]
            values: List[Any] = [codigo, data_criacao, motorista, veiculo, equipe, float(payload.kg_estimado or 0.0), status]
            if "motorista_id" in cols_prog:
                col_names.append("motorista_id")
                values.append(int(motorista_id or 0))
            if "motorista_codigo" in cols_prog:
                col_names.append("motorista_codigo")
                values.append(motorista_codigo)
            if "codigo_motorista" in cols_prog:
                col_names.append("codigo_motorista")
                values.append(motorista_codigo)
            if "tipo_estimativa" in cols_prog:
                col_names.append("tipo_estimativa")
                values.append(tipo_estimativa)
            if "caixas_estimado" in cols_prog:
                col_names.append("caixas_estimado")
                values.append(int(payload.caixas_estimado or 0))
            if "codigo" in cols_prog:
                col_names.append("codigo")
                values.append(codigo)
            if "data" in cols_prog:
                col_names.append("data")
                values.append(data_criacao)
            if "operacao_tipo" in cols_prog:
                col_names.append("operacao_tipo")
                values.append(operacao_tipo)
            if "transbordo_modalidade" in cols_prog:
                col_names.append("transbordo_modalidade")
                values.append(transbordo_modalidade)
            if "transbordo_grupo" in cols_prog:
                col_names.append("transbordo_grupo")
                values.append(transbordo_grupo)
            if "local_rota" in cols_prog:
                col_names.append("local_rota")
                values.append(local_rota)
            if "tipo_rota" in cols_prog:
                col_names.append("tipo_rota")
                values.append(local_rota)
            if "local_carregamento" in cols_prog:
                col_names.append("local_carregamento")
                values.append(local_carregamento)
            if "granja_carregada" in cols_prog:
                col_names.append("granja_carregada")
                values.append(local_carregamento)
            if "local_carregado" in cols_prog:
                col_names.append("local_carregado")
                values.append(local_carregamento)
            if "local_carreg" in cols_prog:
                col_names.append("local_carreg")
                values.append(local_carregamento)
            if "adiantamento" in cols_prog:
                col_names.append("adiantamento")
                values.append(float(payload.adiantamento or 0.0))
            if "adiantamento_rota" in cols_prog:
                col_names.append("adiantamento_rota")
                values.append(float(payload.adiantamento or 0.0))
            if "adiantamento_origem" in cols_prog:
                col_names.append("adiantamento_origem")
                values.append(adiantamento_origem)
            if "pix_motorista" in cols_prog:
                col_names.append("pix_motorista")
                values.append(float(payload.pix_motorista or 0.0))
            if "total_caixas" in cols_prog:
                col_names.append("total_caixas")
                values.append(int(payload.total_caixas or 0))
            if "quilos" in cols_prog:
                col_names.append("quilos")
                values.append(float(payload.quilos or 0.0))
            if "nf_kg" in cols_prog:
                col_names.append("nf_kg")
                values.append(nf_kg if nf_kg > 0 else None)
            if "kg_nf" in cols_prog:
                col_names.append("kg_nf")
                values.append(nf_kg if nf_kg > 0 else None)
            if "nf_preco" in cols_prog:
                col_names.append("nf_preco")
                values.append(nf_preco if nf_preco > 0 else None)
            if "preco_nf" in cols_prog:
                col_names.append("preco_nf")
                values.append(nf_preco if nf_preco > 0 else None)
            if "caixas_carregadas" in cols_prog:
                col_names.append("caixas_carregadas")
                values.append(caixas_carregadas if caixas_carregadas > 0 else 0)
            if "qnt_cx_carregada" in cols_prog:
                col_names.append("qnt_cx_carregada")
                values.append(caixas_carregadas if caixas_carregadas > 0 else 0)
            if "nf_caixas" in cols_prog:
                col_names.append("nf_caixas")
                values.append(caixas_carregadas if caixas_carregadas > 0 else 0)
            if "usuario_criacao" in cols_prog:
                col_names.append("usuario_criacao")
                values.append(str(payload.usuario_criacao or "").strip().upper())
            if "usuario_ultima_edicao" in cols_prog:
                col_names.append("usuario_ultima_edicao")
                values.append(str(payload.usuario_ultima_edicao or payload.usuario_criacao or "").strip().upper())

            ph = ", ".join(["?"] * len(col_names))
            cur.execute(
                f"INSERT INTO programacoes ({', '.join(col_names)}) VALUES ({ph})",
                tuple(values),
            )

        # Itens da programação: substitui snapshot no servidor.
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_itens = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        has_item_obs = "observacao" in cols_itens
        cur.execute("DELETE FROM programacao_itens WHERE codigo_programacao=?", (codigo,))
        for it in (payload.itens or []):
            cod_cli = str(it.cod_cliente or "").strip().upper()
            nome_cli = str(it.nome_cliente or "").strip().upper()
            if not cod_cli or not nome_cli:
                continue
            item_data: Dict[str, Any] = {
                "codigo_programacao": codigo,
                "cod_cliente": cod_cli,
                "nome_cliente": nome_cli,
                "qnt_caixas": int(it.qnt_caixas or 0),
                "kg": float(it.kg or 0.0),
                "preco": float(it.preco or 0.0),
                "endereco": str(it.endereco or "").strip().upper(),
                "vendedor": str(it.vendedor or "").strip().upper(),
                "pedido": str(it.pedido or "").strip().upper(),
                "produto": str(it.produto or "").strip().upper(),
            }
            if has_item_obs:
                item_data["observacao"] = str(it.obs or "").strip().upper()
            if "ordem_sugerida" in cols_itens and it.ordem_sugerida is not None:
                item_data["ordem_sugerida"] = int(it.ordem_sugerida)
            eta_txt = str(it.eta or "").strip()
            if "eta" in cols_itens and eta_txt:
                item_data["eta"] = eta_txt
            if "distancia" in cols_itens and it.distancia is not None:
                item_data["distancia"] = float(it.distancia)
            if "confianca_localizacao" in cols_itens and it.confianca_localizacao is not None:
                item_data["confianca_localizacao"] = float(it.confianca_localizacao)

            keys = list(item_data.keys())
            vals = [item_data[k] for k in keys]
            cur.execute(
                f"INSERT INTO programacao_itens ({', '.join(keys)}) VALUES ({', '.join(['?'] * len(keys))})",
                tuple(vals),
            )

        if linked_venda_ids:
            placeholders = ",".join(["?"] * len(linked_venda_ids))
            cur.execute(
                f"""
                SELECT id,
                       IFNULL(usada,0) AS usada,
                       UPPER(TRIM(COALESCE(codigo_programacao,''))) AS codigo_programacao
                FROM vendas_importadas
                WHERE id IN ({placeholders})
                """,
                tuple(linked_venda_ids),
            )
            venda_rows = cur.fetchall() or []
            venda_map = {int(r["id"] or 0): r for r in venda_rows}
            missing_ids = [rid for rid in linked_venda_ids if rid not in venda_map]
            if missing_ids:
                if pid > 0:
                    logging.warning(
                        "desktop_rotas_upsert(%s): ignorando %s venda(s) antiga(s) ausente(s) na fila de importacao: %s",
                        codigo,
                        len(missing_ids),
                        missing_ids,
                    )
                    linked_venda_ids = [rid for rid in linked_venda_ids if rid in venda_map]
                else:
                    raise HTTPException(
                        status_code=409,
                        detail="existem vendas vinculadas informadas que nao existem mais na fila de importacao.",
                    )
            for rid in linked_venda_ids:
                row_v = venda_map.get(rid)
                venda_prog = str((row_v["codigo_programacao"] if row_v else "") or "").strip().upper()
                venda_usada = int((row_v["usada"] if row_v else 0) or 0)
                if venda_prog and venda_prog != codigo:
                    raise HTTPException(
                        status_code=409,
                        detail="existem vendas vinculadas a outra programacao; desvincule antes de salvar.",
                    )
                if venda_usada == 1 and venda_prog != codigo:
                    raise HTTPException(
                        status_code=409,
                        detail="existem vendas ja consumidas por outra programacao; recarregue a lista antes de salvar.",
                    )
            cur.executemany(
                """
                UPDATE vendas_importadas
                   SET usada=1,
                       usada_em=?,
                       codigo_programacao=?,
                       selecionada=0
                 WHERE id=?
                """,
                [(vendas_usada_em, codigo, rid) for rid in linked_venda_ids],
            )

        conn.commit()

    return {"ok": True, "codigo_programacao": codigo}


@app.post("/desktop/avulsas")
def desktop_avulsa_criar(payload: ProgramacaoAvulsaIn, _ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        codigo = _next_codigo_avulsa(cur, payload.data_programada)
        cur.execute(
            """
            INSERT INTO programacoes_avulsas
                (codigo_avulsa, data_programada, status, motorista_id, motorista_codigo, motorista_nome,
                 veiculo, equipe, local_rota, observacao, origem, criado_por, criado_em)
            VALUES (?, ?, 'AVULSA_ATIVA', ?, ?, ?, ?, ?, ?, ?, 'APP_VENDEDOR', ?, datetime('now'))
            """,
            (
                codigo,
                (payload.data_programada or None),
                payload.motorista_id,
                (payload.motorista_codigo or "").strip().upper() or None,
                (payload.motorista_nome or "").strip().upper() or None,
                (payload.veiculo or "").strip().upper() or None,
                (payload.equipe or "").strip().upper() or None,
                (payload.local_rota or "").strip().upper() or None,
                (payload.observacao or "").strip() or None,
                (payload.criado_por or "VENDEDOR_APP").strip().upper(),
            ),
        )
        avulsa_id = int(cur.lastrowid or 0)
        ordem_base = 1
        for i, it in enumerate(payload.itens or [], start=ordem_base):
            cod = (it.cod_cliente or "").strip()
            nome = (it.nome_cliente or "").strip()
            if not cod or not nome:
                continue
            cur.execute(
                """
                INSERT INTO programacoes_avulsas_itens
                    (avulsa_id, cod_cliente, nome_cliente, endereco, cidade, bairro, ordem, observacao, status_item, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'PENDENTE', datetime('now'))
                """,
                (
                    avulsa_id,
                    cod,
                    nome,
                    (it.endereco or "").strip() or None,
                    (it.cidade or "").strip() or None,
                    (it.bairro or "").strip() or None,
                    int(it.ordem if it.ordem is not None else i),
                    (it.observacao or "").strip() or None,
                ),
            )
        conn.commit()
        return {"ok": True, "codigo_avulsa": codigo, "id": avulsa_id}


@app.get("/desktop/avulsas")
def desktop_avulsas_listar(
    status: str = Query("", description="Filtro de status"),
    data_de: str = Query("", description="YYYY-MM-DD"),
    data_ate: str = Query("", description="YYYY-MM-DD"),
    limit: int = Query(200, ge=1, le=1000),
    _ok: bool = Depends(_require_desktop_secret),
):
    with get_conn() as conn:
        cur = conn.cursor()
        where = []
        params: List[Any] = []
        st = (status or "").strip().upper()
        if st:
            where.append("UPPER(COALESCE(status,''))=?")
            params.append(st)
        if (data_de or "").strip():
            where.append("COALESCE(data_programada,'') >= ?")
            params.append((data_de or "").strip())
        if (data_ate or "").strip():
            where.append("COALESCE(data_programada,'') <= ?")
            params.append((data_ate or "").strip())
        where_sql = ("WHERE " + " AND ".join(where)) if where else ""
        cur.execute(
            f"""
            SELECT
                id, codigo_avulsa, COALESCE(data_programada,''), COALESCE(status,''),
                COALESCE(motorista_codigo,''), COALESCE(motorista_nome,''),
                COALESCE(veiculo,''), COALESCE(equipe,''), COALESCE(local_rota,''),
                COALESCE(programacao_oficial_codigo,''), COALESCE(criado_por,''), COALESCE(criado_em,'')
            FROM programacoes_avulsas
            {where_sql}
            ORDER BY id DESC
            LIMIT ?
            """,
            (*params, int(limit)),
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "codigo_avulsa": str(r[1] or "").strip(),
                    "data_programada": str(r[2] or "").strip(),
                    "status": str(r[3] or "").strip().upper(),
                    "motorista_codigo": str(r[4] or "").strip().upper(),
                    "motorista_nome": str(r[5] or "").strip().upper(),
                    "veiculo": str(r[6] or "").strip().upper(),
                    "equipe": str(r[7] or "").strip().upper(),
                    "local_rota": str(r[8] or "").strip().upper(),
                    "programacao_oficial_codigo": str(r[9] or "").strip().upper(),
                    "criado_por": str(r[10] or "").strip().upper(),
                    "criado_em": str(r[11] or "").strip(),
                }
            )
        return out


@app.get("/desktop/avulsas/{codigo_avulsa}")
def desktop_avulsa_detalhe(codigo_avulsa: str, _ok: bool = Depends(_require_desktop_secret)):
    codigo = (codigo_avulsa or "").strip().upper()
    if not codigo:
        raise HTTPException(status_code=400, detail="codigo_avulsa invalido.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT *
            FROM programacoes_avulsas
            WHERE UPPER(COALESCE(codigo_avulsa,''))=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        )
        head = cur.fetchone()
        if not head:
            raise HTTPException(status_code=404, detail="Programacao avulsa nao encontrada.")
        av = row_to_dict(head)
        av_id = int(av.get("id") or 0)
        cur.execute(
            """
            SELECT *
            FROM programacoes_avulsas_itens
            WHERE avulsa_id=?
            ORDER BY ordem, id
            """,
            (av_id,),
        )
        itens = [row_to_dict(r) for r in (cur.fetchall() or [])]
        return {"avulsa": av, "itens": itens}


@app.post("/desktop/avulsas/{codigo_avulsa}/conciliar")
def desktop_avulsa_conciliar(
    codigo_avulsa: str,
    payload: ProgramacaoAvulsaConciliarIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    codigo = (codigo_avulsa or "").strip().upper()
    oficial = (payload.codigo_programacao_oficial or "").strip().upper()
    usuario = (payload.usuario or "DESKTOP").strip().upper()
    if not codigo or not oficial:
        raise HTTPException(status_code=400, detail="codigo_avulsa e codigo_programacao_oficial sao obrigatorios.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, COALESCE(status,'') AS status
            FROM programacoes_avulsas
            WHERE UPPER(COALESCE(codigo_avulsa,''))=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        )
        av = cur.fetchone()
        if not av:
            raise HTTPException(status_code=404, detail="Programacao avulsa nao encontrada.")
        st = str(av["status"] or "").strip().upper()
        if st in ("CONCILIADA", "CANCELADA"):
            raise HTTPException(status_code=409, detail=f"Programacao avulsa bloqueada para conciliacao (status={st}).")

        cur.execute(
            """
            SELECT id, COALESCE(num_nf, COALESCE(nf_numero, '')) AS nf_numero
            FROM programacoes
            WHERE UPPER(COALESCE(codigo_programacao,''))=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (oficial,),
        )
        row_prog = cur.fetchone()
        if not row_prog:
            raise HTTPException(status_code=404, detail="Programacao oficial nao encontrada.")

        prog_id = int(row_prog["id"])
        nf_oficial = str(row_prog["nf_numero"] or "").strip()

        # Itens da programação oficial (fonte de pedido/caixas/preço).
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_pi = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        has_obs = "observacao" in cols_pi
        select_obs = ", COALESCE(observacao,'') AS observacao" if has_obs else ", '' AS observacao"
        cur.execute(
            f"""
            SELECT
                id,
                UPPER(TRIM(COALESCE(cod_cliente,''))) AS cod_cliente_u,
                COALESCE(cod_cliente,'') AS cod_cliente,
                COALESCE(nome_cliente,'') AS nome_cliente,
                COALESCE(pedido,'') AS pedido,
                COALESCE(qnt_caixas,0) AS qnt_caixas,
                COALESCE(preco,0) AS preco
                {select_obs}
            FROM programacao_itens
            WHERE codigo_programacao=?
            ORDER BY id ASC
            """,
            (oficial,),
        )
        oficial_itens = cur.fetchall() or []

        # Itens avulsos a conciliar.
        cur.execute(
            """
            SELECT id, UPPER(TRIM(COALESCE(cod_cliente,''))) AS cod_cliente_u, COALESCE(cod_cliente,'') AS cod_cliente
            FROM programacoes_avulsas_itens
            WHERE avulsa_id=?
            ORDER BY ordem, id
            """,
            (int(av["id"]),),
        )
        avulsos = cur.fetchall() or []

        # Mapa por cliente (com fila para suportar pedidos repetidos do mesmo cliente).
        fila_por_cliente: Dict[str, List[sqlite3.Row]] = {}
        for it in oficial_itens:
            key = str(it["cod_cliente_u"] or "").strip()
            if not key:
                continue
            fila_por_cliente.setdefault(key, []).append(it)

        matched = 0
        pendentes = 0
        pendentes_codigos: List[str] = []
        for av_item in avulsos:
            key = str(av_item["cod_cliente_u"] or "").strip()
            fila = fila_por_cliente.get(key) or []
            if fila:
                src = fila.pop(0)
                cur.execute(
                    """
                    UPDATE programacoes_avulsas_itens
                    SET pedido=?,
                        nf=?,
                        caixas=?,
                        preco=?,
                        status_item='CONCILIADO',
                        updated_at=datetime('now')
                    WHERE id=?
                    """,
                    (
                        str(src["pedido"] or "").strip() or None,
                        nf_oficial or None,
                        int(src["qnt_caixas"] or 0),
                        float(src["preco"] or 0.0),
                        int(av_item["id"]),
                    ),
                )
                matched += 1
            else:
                cur.execute(
                    """
                    UPDATE programacoes_avulsas_itens
                    SET status_item='PENDENTE_CONCILIACAO',
                        updated_at=datetime('now')
                    WHERE id=?
                    """,
                    (int(av_item["id"]),),
                )
                pendentes += 1
                pendentes_codigos.append(str(av_item["cod_cliente"] or "").strip())

        # Status final da avulsa conforme resultado.
        novo_status = "CONCILIADA" if pendentes == 0 else "CONCILIADA_PARCIAL"
        cur.execute(
            """
            UPDATE programacoes_avulsas
            SET status=?,
                conciliada_em=datetime('now'),
                programacao_oficial_codigo=?,
                observacao=COALESCE(observacao,'') || CASE WHEN COALESCE(observacao,'')='' THEN '' ELSE ' | ' END || ?
            WHERE id=?
            """,
            (
                novo_status,
                oficial,
                f"CONCILIADA POR {usuario} (MATCH={matched} PEND={pendentes})",
                int(av["id"]),
            ),
        )
        conn.commit()
        return {
            "ok": True,
            "codigo_avulsa": codigo,
            "codigo_programacao_oficial": oficial,
            "status_avulsa": novo_status,
            "itens_total": len(avulsos),
            "itens_conciliados": matched,
            "itens_pendentes": pendentes,
            "pendentes_cod_cliente": pendentes_codigos[:100],
            "programacao_id": prog_id,
        }


@app.get("/clientes/base")
def listar_clientes_base(
    q: str = Query("", description="Busca por código/nome/cidade"),
    limit: int = Query(200, ge=1, le=1000),
    m=Depends(get_current_motorista),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
        if not cur.fetchone():
            return []
        _ensure_clientes_columns(cur)
        company_id = int((m or {}).get("company_id") or _default_company_id(cur))
        scope_sql, scope_params = _company_scope_condition(cur, "clientes", company_id)

        term = (q or "").strip().upper()
        like = f"%{term}%"
        cur.execute(
            f"""
            SELECT
                TRIM(COALESCE(cod_cliente, '')) AS cod_cliente,
                TRIM(COALESCE(nome_cliente, '')) AS nome_cliente,
                TRIM(COALESCE(cidade, '')) AS cidade,
                TRIM(COALESCE(vendedor, '')) AS vendedor
            FROM clientes
            WHERE
                (
                    (? = '')
                    OR UPPER(TRIM(COALESCE(cod_cliente, ''))) LIKE ?
                    OR UPPER(TRIM(COALESCE(nome_cliente, ''))) LIKE ?
                    OR UPPER(TRIM(COALESCE(cidade, ''))) LIKE ?
                )
                {f'AND {scope_sql}' if scope_sql else ''}
            ORDER BY UPPER(TRIM(COALESCE(nome_cliente, ''))), UPPER(TRIM(COALESCE(cod_cliente, '')))
            LIMIT ?
            """,
            (term, like, like, like, *scope_params, int(limit)),
        )
        out: List[Dict[str, Any]] = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "cod_cliente": (r["cod_cliente"] or "").strip(),
                    "nome_cliente": (r["nome_cliente"] or "").strip(),
                    "cidade": (r["cidade"] or "").strip(),
                    "vendedor": (r["vendedor"] or "").strip(),
                }
            )
        return out


@app.post("/rotas/{codigo_programacao}/clientes/reserva")
def criar_cliente_reserva(
    codigo_programacao: str,
    payload: ClienteReservaIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    cod_cliente = (payload.cod_cliente or "").strip()
    nome_cliente = (payload.nome_cliente or "").strip()
    pedido_in = (payload.pedido or "").strip()
    qnt_caixas = int(payload.qnt_caixas or 0)
    status_in = (payload.status_pedido or "PENDENTE").strip().upper() or "PENDENTE"
    vendedor = (payload.vendedor or "").strip()
    cidade = (payload.cidade or "").strip()
    observacao = (payload.observacao or "").strip()

    if not codigo_programacao:
        raise HTTPException(status_code=400, detail="Código da programação é obrigatório.")
    if not cod_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente é obrigatório.")
    if not nome_cliente:
        raise HTTPException(status_code=400, detail="nome_cliente é obrigatório.")
    if qnt_caixas <= 0:
        raise HTTPException(status_code=400, detail="qnt_caixas deve ser maior que zero.")

    pedido = pedido_in or f"RES-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    status_final = status_in if status_in in ("PENDENTE", "ALTERADO", "CANCELADO", "ENTREGUE") else "PENDENTE"

    with get_conn() as conn:
        cur = conn.cursor()
        row_prog = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.codigo_programacao")
        if not row_prog:
            raise HTTPException(status_code=404, detail="Programação não encontrada para este motorista.")

        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
        if not cur.fetchone():
            raise HTTPException(status_code=500, detail="Tabela programacao_itens não encontrada.")

        cur.execute("PRAGMA table_info(programacao_itens)")
        cols = {r[1] for r in cur.fetchall() or []}

        has_pedido = "pedido" in cols
        has_status = "status_pedido" in cols
        has_caixas_atual = "caixas_atual" in cols

        existing = None
        if has_pedido:
            cur.execute(
                """
                SELECT *
                FROM programacao_itens
                WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(TRIM(pedido), '')=COALESCE(TRIM(?), '')
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente, pedido),
            )
            existing = cur.fetchone()

        if existing:
            novo_qtd = int((existing["qnt_caixas"] or 0)) + qnt_caixas
            params: List[Any] = [novo_qtd]
            sets = ["qnt_caixas=?"]
            if has_caixas_atual:
                caixa_atual_base = int((existing["caixas_atual"] or existing["qnt_caixas"] or 0))
                sets.append("caixas_atual=?")
                params.append(caixa_atual_base + qnt_caixas)
            if has_status:
                sets.append("status_pedido=?")
                params.append(status_final)
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE programacao_itens SET {', '.join(sets)} WHERE id=?", params)
        else:
            data: Dict[str, Any] = {
                "codigo_programacao": codigo_programacao,
                "cod_cliente": cod_cliente,
                "nome_cliente": nome_cliente,
                "qnt_caixas": qnt_caixas,
            }
            if has_pedido:
                data["pedido"] = pedido
            if "kg" in cols:
                data["kg"] = 0
            if "preco" in cols:
                data["preco"] = payload.preco
            if "preco_atual" in cols:
                data["preco_atual"] = payload.preco
            if "vendedor" in cols:
                data["vendedor"] = vendedor
            if "observacao" in cols:
                data["observacao"] = observacao or f"Pedido reserva via app ({cidade})"
            if "produto" in cols:
                data["produto"] = "RESERVA"
            if has_status:
                data["status_pedido"] = status_final
            if has_caixas_atual:
                data["caixas_atual"] = qnt_caixas

            keys = list(data.keys())
            vals = [data[k] for k in keys]
            cur.execute(
                f"INSERT INTO programacao_itens ({', '.join(keys)}) VALUES ({', '.join(['?'] * len(keys))})",
                vals,
            )

        itens_select_expr = _programacao_itens_select_expr(conn, "pi")
        if has_pedido:
            cur.execute(
                f"""
                SELECT {itens_select_expr}
                FROM programacao_itens pi
                WHERE pi.codigo_programacao=? AND pi.cod_cliente=? AND COALESCE(TRIM(pi.pedido), '')=COALESCE(TRIM(?), '')
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente, pedido),
            )
        else:
            cur.execute(
                f"""
                SELECT {itens_select_expr}
                FROM programacao_itens pi
                WHERE pi.codigo_programacao=? AND pi.cod_cliente=?
                ORDER BY pi.id DESC
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente),
            )
        item = cur.fetchone()
        return {"ok": True, "item": row_to_dict(item), "pedido": pedido}


@app.post("/auth/motorista/login", response_model=LoginOut)
def autenticar_motorista(payload: LoginIn):
    codigo = (payload.codigo or "").strip().upper()
    senha = (payload.senha or "").strip()

    if not codigo or not senha:
        raise HTTPException(status_code=400, detail="Codigo e senha sao obrigatorios")

    with get_conn() as conn:
        cur = conn.cursor()
        m, auth_err = authenticate_motorista(cur, codigo, senha)

    if auth_err == "blocked":
        raise HTTPException(status_code=403, detail="Acesso bloqueado. Solicite desbloqueio do administrador.")
    if not m:
        raise HTTPException(status_code=401, detail="Codigo ou senha invalidos")

    perfil_app = _motorista_app_role(m)
    is_admin = perfil_app == "ADMIN"
    company_id = _row_company_id(m)
    token = create_token(
        m["codigo"],
        perfil="admin" if is_admin else "motorista",
        company_id=company_id,
        user_id=int(m["id"]),
        username=str(m["nome"] or ""),
        role=perfil_app,
    )
    return {
        "token": token,
        "nome": m["nome"],
        "codigo": m["codigo"],
        "company_id": company_id,
        "perfil": "admin" if is_admin else "motorista",
        "role": perfil_app,
        "is_admin": is_admin,
    }


@app.post("/auth/admin/login", response_model=LoginOut)
def autenticar_admin(payload: LoginIn):
    codigo = (payload.codigo or "").strip().upper()
    senha = (payload.senha or "").strip()

    if not codigo or not senha:
        raise HTTPException(status_code=400, detail="Codigo e senha sao obrigatorios")

    with get_conn() as conn:
        cur = conn.cursor()
        m, auth_err = authenticate_motorista(cur, codigo, senha)

    if auth_err == "blocked":
        raise HTTPException(status_code=403, detail="Acesso bloqueado. Solicite desbloqueio do administrador.")
    if not m:
        raise HTTPException(status_code=401, detail="Codigo ou senha invalidos")

    perfil_app = _motorista_app_role(m)
    if perfil_app != "ADMIN":
        raise HTTPException(status_code=403, detail="Este usuario nao possui perfil admin para acesso total ao app.")

    company_id = _row_company_id(m)
    token = create_token(
        m["codigo"],
        perfil="admin",
        company_id=company_id,
        user_id=int(m["id"]),
        username=str(m["nome"] or ""),
        role="ADMIN",
    )
    return {
        "token": token,
        "nome": m["nome"],
        "codigo": m["codigo"],
        "company_id": company_id,
        "perfil": "admin",
        "role": "ADMIN",
        "is_admin": True,
    }


@app.post("/auth/vendedor/login", response_model=LoginOut)
def autenticar_vendedor(payload: LoginIn, request: Request):
    codigo = (payload.codigo or "").strip()
    senha = (payload.senha or "").strip()

    if not codigo or not senha:
        raise HTTPException(status_code=400, detail="Nome/codigo e senha sao obrigatorios")

    with get_conn() as conn:
        cur = conn.cursor()
        v, auth_err = authenticate_vendedor(cur, codigo, senha)
        if v:
            try:
                cur.execute("PRAGMA table_info(vendedores)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                set_parts: List[str] = []
                params: List[Any] = []
                if "ultimo_login_em" in cols:
                    set_parts.append("ultimo_login_em=?")
                    params.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                if "ultimo_login_ip" in cols:
                    host = ""
                    try:
                        host = str(request.client.host or "").strip()
                    except Exception:
                        host = ""
                    set_parts.append("ultimo_login_ip=?")
                    params.append(host)
                if set_parts:
                    params.append(int(v["id"]))
                    cur.execute(f"UPDATE vendedores SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            except Exception:
                logging.debug("Falha ao registrar ultimo login do vendedor", exc_info=True)

    if auth_err == "blocked":
        raise HTTPException(status_code=403, detail="Acesso bloqueado. Procure o administrador.")
    if not v:
        raise HTTPException(status_code=401, detail="Nome/codigo ou senha invalidos")

    company_id = _row_company_id(v)
    token = create_token(
        v["codigo"],
        perfil="vendedor",
        company_id=company_id,
        user_id=int(v["id"]),
        username=str(v["nome"] or ""),
        role="vendedor",
    )
    return {"token": token, "nome": v["nome"], "codigo": v["codigo"], "company_id": company_id}


@app.get("/vendedor/rascunho")
def vendedor_rascunho_listar(
    limit: int = Query(500, ge=1, le=2000),
    vendedor=Depends(get_current_vendedor),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendedor_rascunho_itens'")
        if not cur.fetchone():
            return []
        cur.execute(
            """
            SELECT
                id,
                COALESCE(cod_cliente,'') AS cod_cliente,
                COALESCE(nome_cliente,'') AS nome_cliente,
                COALESCE(cidade,'') AS cidade,
                COALESCE(bairro,'') AS bairro,
                COALESCE(endereco,'') AS endereco,
                COALESCE(vendedor_cadastro,'') AS vendedor_cadastro,
                COALESCE(vendedor_origem,'') AS vendedor_origem,
                COALESCE(preco, 0) AS preco,
                COALESCE(caixas, 0) AS caixas,
                COALESCE(status,'PENDENTE') AS status,
                COALESCE(observacao,'') AS observacao,
                COALESCE(alerta_codigo_programacao,'') AS alerta_codigo_programacao,
                COALESCE(alerta_status_rota,'') AS alerta_status_rota,
                COALESCE(criado_em,'') AS criado_em,
                COALESCE(atualizado_em,'') AS atualizado_em,
                COALESCE(criado_por_codigo,'') AS criado_por_codigo,
                COALESCE(atualizado_por_codigo,'') AS atualizado_por_codigo
            FROM vendedor_rascunho_itens
            ORDER BY COALESCE(atualizado_em, criado_em) DESC, id DESC
            LIMIT ?
            """,
            (int(limit),),
        )
        out: List[Dict[str, Any]] = []
        for row in cur.fetchall() or []:
            out.append(
                {
                    "id": str(row["id"] or "").strip(),
                    "cod_cliente": str(row["cod_cliente"] or "").strip().upper(),
                    "nome_cliente": str(row["nome_cliente"] or "").strip().upper(),
                    "cidade": str(row["cidade"] or "").strip().upper(),
                    "bairro": str(row["bairro"] or "").strip().upper(),
                    "endereco": str(row["endereco"] or "").strip().upper(),
                    "vendedor_cadastro": str(row["vendedor_cadastro"] or "").strip().upper(),
                    "vendedor_origem": str(row["vendedor_origem"] or "").strip().upper(),
                    "preco": float(row["preco"] or 0.0),
                    "caixas": int(row["caixas"] or 0),
                    "status": str(row["status"] or "PENDENTE").strip().upper(),
                    "observacao": str(row["observacao"] or "").strip(),
                    "alerta_codigo_programacao": str(row["alerta_codigo_programacao"] or "").strip().upper(),
                    "alerta_status_rota": str(row["alerta_status_rota"] or "").strip().upper(),
                    "criado_em": str(row["criado_em"] or "").strip(),
                    "updated_at": str(row["atualizado_em"] or "").strip(),
                    "criado_por_codigo": str(row["criado_por_codigo"] or "").strip().upper(),
                    "atualizado_por_codigo": str(row["atualizado_por_codigo"] or "").strip().upper(),
                }
            )
        return out


@app.post("/vendedor/rascunho/itens")
def vendedor_rascunho_criar(
    payload: VendedorRascunhoCreateIn,
    vendedor=Depends(get_current_vendedor),
):
    itens = payload.itens or []
    if not itens:
        raise HTTPException(status_code=400, detail="itens obrigatorios.")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ids: List[str] = []
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendedor_rascunho_itens'")
        if not cur.fetchone():
            raise HTTPException(status_code=500, detail="Tabela de rascunho indisponivel.")

        for item in itens:
            cod_cliente = _clean_text(item.cod_cliente).upper()
            nome_cliente = _clean_text(item.nome_cliente).upper()
            vendedor_origem = _clean_text(item.vendedor_origem or vendedor["codigo"]).upper() or str(vendedor["codigo"]).upper()
            if not cod_cliente or not nome_cliente or not vendedor_origem:
                continue
            item_id = _clean_text(item.id) or uuid4().hex.upper()
            status = _clean_text(item.status or "PENDENTE").upper()
            if status not in {"PENDENTE", "FINALIZADA"}:
                status = "PENDENTE"
            ids.append(item_id)
            cur.execute(
                """
                INSERT INTO vendedor_rascunho_itens (
                    id, cod_cliente, nome_cliente, cidade, bairro, endereco,
                    vendedor_cadastro, vendedor_origem, preco, caixas, status,
                    observacao, alerta_codigo_programacao, alerta_status_rota,
                    criado_em, atualizado_em, criado_por_codigo, atualizado_por_codigo
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    cod_cliente=excluded.cod_cliente,
                    nome_cliente=excluded.nome_cliente,
                    cidade=excluded.cidade,
                    bairro=excluded.bairro,
                    endereco=excluded.endereco,
                    vendedor_cadastro=excluded.vendedor_cadastro,
                    vendedor_origem=excluded.vendedor_origem,
                    preco=excluded.preco,
                    caixas=excluded.caixas,
                    status=excluded.status,
                    observacao=excluded.observacao,
                    alerta_codigo_programacao=excluded.alerta_codigo_programacao,
                    alerta_status_rota=excluded.alerta_status_rota,
                    atualizado_em=excluded.atualizado_em,
                    atualizado_por_codigo=excluded.atualizado_por_codigo
                """,
                (
                    item_id,
                    cod_cliente,
                    nome_cliente,
                    _clean_text(item.cidade).upper(),
                    _clean_text(item.bairro).upper(),
                    _clean_text(item.endereco).upper(),
                    _clean_text(item.vendedor_cadastro).upper(),
                    vendedor_origem,
                    float(item.preco or 0.0),
                    max(int(item.caixas or 0), 0),
                    status,
                    _clean_text(item.observacao),
                    _clean_text(item.alerta_codigo_programacao).upper(),
                    _clean_text(item.alerta_status_rota).upper(),
                    now,
                    now,
                    str(vendedor["codigo"]).upper(),
                    str(vendedor["codigo"]).upper(),
                ),
            )

    return {"ok": True, "ids": ids, "count": len(ids)}


@app.patch("/vendedor/rascunho/{item_id}")
def vendedor_rascunho_atualizar(
    item_id: str,
    payload: VendedorRascunhoUpdateIn,
    vendedor=Depends(get_current_vendedor),
):
    target = _clean_text(item_id)
    if not target:
        raise HTTPException(status_code=400, detail="item_id obrigatorio.")

    sets: List[str] = []
    params: List[Any] = []
    if payload.caixas is not None:
        sets.append("caixas=?")
        params.append(max(int(payload.caixas or 0), 0))
    if payload.preco is not None:
        sets.append("preco=?")
        params.append(float(payload.preco or 0.0))
    if payload.observacao is not None:
        sets.append("observacao=?")
        params.append(_clean_text(payload.observacao))
    if payload.status is not None:
        status = _clean_text(payload.status).upper()
        if status not in {"PENDENTE", "FINALIZADA"}:
            raise HTTPException(status_code=400, detail="status invalido.")
        sets.append("status=?")
        params.append(status)

    if not sets:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")

    sets.append("atualizado_em=?")
    params.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    sets.append("atualizado_por_codigo=?")
    params.append(str(vendedor["codigo"]).upper())
    params.append(target)

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            f"UPDATE vendedor_rascunho_itens SET {', '.join(sets)} WHERE id=?",
            tuple(params),
        )
        if int(cur.rowcount or 0) <= 0:
            raise HTTPException(status_code=404, detail="Item de rascunho nao encontrado.")
    return {"ok": True, "id": target}


@app.delete("/vendedor/rascunho/{item_id}")
def vendedor_rascunho_remover(
    item_id: str,
    vendedor=Depends(get_current_vendedor),
):
    target = _clean_text(item_id)
    if not target:
        raise HTTPException(status_code=400, detail="item_id obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM vendedor_rascunho_itens WHERE id=?", (target,))
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "id": target, "deleted": deleted}


@app.post("/vendedor/rascunho/remover-em-lote")
def vendedor_rascunho_remover_em_lote(
    payload: VendedorRascunhoDeleteBulkIn,
    vendedor=Depends(get_current_vendedor),
):
    ids = [str(item or "").strip() for item in (payload.ids or []) if str(item or "").strip()]
    if not ids:
        return {"ok": True, "deleted": 0}
    with get_conn() as conn:
        cur = conn.cursor()
        qmarks = ",".join(["?"] * len(ids))
        cur.execute(f"DELETE FROM vendedor_rascunho_itens WHERE id IN ({qmarks})", tuple(ids))
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "deleted": deleted}


@app.get("/vendedor/pre-programacoes")
def vendedor_pre_programacoes_listar(
    status: str = Query("ABERTA", description="ABERTA|FECHADA|TODAS"),
    limit: int = Query(100, ge=1, le=500),
    vendedor=Depends(get_current_vendedor),
):
    status_norm = _clean_text(status).upper() or "ABERTA"
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "vendedor_pre_programacoes"):
            return []

        params: List[Any] = []
        where_sql = ""
        if status_norm not in {"", "TODAS"}:
            where_sql = "WHERE UPPER(COALESCE(status,'ABERTA'))=?"
            params.append(status_norm)

        params.append(int(limit))
        cur.execute(
            f"""
            SELECT
                id,
                COALESCE(titulo,'') AS titulo,
                COALESCE(observacao,'') AS observacao,
                COALESCE(status,'ABERTA') AS status,
                COALESCE(criado_em,'') AS criado_em,
                COALESCE(atualizado_em,'') AS atualizado_em,
                COALESCE(criado_por_codigo,'') AS criado_por_codigo,
                COALESCE(atualizado_por_codigo,'') AS atualizado_por_codigo
            FROM vendedor_pre_programacoes
            {where_sql}
            ORDER BY COALESCE(atualizado_em, criado_em) DESC, id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        headers = cur.fetchall() or []
        if not headers:
            return []

        agrupado: Dict[str, Dict[str, Any]] = {}
        ordered_ids: List[str] = []
        for row in headers:
            pid = str(row["id"] or "").strip()
            if not pid:
                continue
            ordered_ids.append(pid)
            agrupado[pid] = {
                "id": pid,
                "titulo": str(row["titulo"] or "").strip(),
                "observacao": str(row["observacao"] or "").strip(),
                "status": str(row["status"] or "ABERTA").strip().upper(),
                "criado_em": str(row["criado_em"] or "").strip(),
                "updated_at": str(row["atualizado_em"] or "").strip(),
                "criado_por_codigo": str(row["criado_por_codigo"] or "").strip().upper(),
                "atualizado_por_codigo": str(row["atualizado_por_codigo"] or "").strip().upper(),
                "itens_total": 0,
                "itens_pendentes": 0,
                "itens_finalizadas": 0,
                "vendedores": [],
            }

        qmarks = ",".join(["?"] * len(ordered_ids))
        cur.execute(
            f"""
            SELECT
                ppi.pre_programacao_id,
                COALESCE(ppi.rascunho_item_id,'') AS rascunho_item_id,
                COALESCE(vri.status,'') AS item_status,
                UPPER(TRIM(COALESCE(vri.vendedor_origem, vri.vendedor_cadastro, ''))) AS vendedor_item
            FROM vendedor_pre_programacao_itens ppi
            LEFT JOIN vendedor_rascunho_itens vri
              ON vri.id = ppi.rascunho_item_id
            WHERE ppi.pre_programacao_id IN ({qmarks})
            ORDER BY COALESCE(ppi.ordem,0), ppi.id
            """,
            tuple(ordered_ids),
        )
        for row in cur.fetchall() or []:
            pid = str(row["pre_programacao_id"] or "").strip()
            if not pid or pid not in agrupado:
                continue
            base = agrupado[pid]
            rid = str(row["rascunho_item_id"] or "").strip()
            if rid:
                base["itens_total"] = int(base["itens_total"] or 0) + 1
                item_status = str(row["item_status"] or "").strip().upper()
                if item_status == "FINALIZADA":
                    base["itens_finalizadas"] = int(base["itens_finalizadas"] or 0) + 1
                else:
                    base["itens_pendentes"] = int(base["itens_pendentes"] or 0) + 1
                vendedor_item = str(row["vendedor_item"] or "").strip().upper()
                if vendedor_item and vendedor_item not in base["vendedores"]:
                    base["vendedores"].append(vendedor_item)

        out = [agrupado[pid] for pid in ordered_ids if pid in agrupado]
        for item in out:
            item["vendedores"] = sorted(item["vendedores"])
        return out


@app.get("/vendedor/pre-programacoes/{pre_programacao_id}")
def vendedor_pre_programacao_detalhe(
    pre_programacao_id: str,
    vendedor=Depends(get_current_vendedor),
):
    target = _clean_text(pre_programacao_id)
    if not target:
        raise HTTPException(status_code=400, detail="pre_programacao_id obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "vendedor_pre_programacoes"):
            raise HTTPException(status_code=404, detail="Pre-programacao nao encontrada.")
        cur.execute(
            """
            SELECT
                id,
                COALESCE(titulo,'') AS titulo,
                COALESCE(observacao,'') AS observacao,
                COALESCE(status,'ABERTA') AS status,
                COALESCE(criado_em,'') AS criado_em,
                COALESCE(atualizado_em,'') AS atualizado_em,
                COALESCE(criado_por_codigo,'') AS criado_por_codigo,
                COALESCE(atualizado_por_codigo,'') AS atualizado_por_codigo
            FROM vendedor_pre_programacoes
            WHERE id=?
            """,
            (target,),
        )
        header = cur.fetchone()
        if not header:
            raise HTTPException(status_code=404, detail="Pre-programacao nao encontrada.")

        cur.execute(
            """
            SELECT
                COALESCE(ppi.ordem,0) AS ordem,
                COALESCE(vri.id,'') AS id,
                COALESCE(vri.cod_cliente,'') AS cod_cliente,
                COALESCE(vri.nome_cliente,'') AS nome_cliente,
                COALESCE(vri.cidade,'') AS cidade,
                COALESCE(vri.bairro,'') AS bairro,
                COALESCE(vri.endereco,'') AS endereco,
                COALESCE(vri.vendedor_cadastro,'') AS vendedor_cadastro,
                COALESCE(vri.vendedor_origem,'') AS vendedor_origem,
                COALESCE(vri.preco,0) AS preco,
                COALESCE(vri.caixas,0) AS caixas,
                COALESCE(vri.status,'PENDENTE') AS status,
                COALESCE(vri.observacao,'') AS observacao,
                COALESCE(vri.alerta_codigo_programacao,'') AS alerta_codigo_programacao,
                COALESCE(vri.alerta_status_rota,'') AS alerta_status_rota,
                COALESCE(vri.criado_em,'') AS criado_em,
                COALESCE(vri.atualizado_em,'') AS atualizado_em,
                COALESCE(vri.criado_por_codigo,'') AS criado_por_codigo,
                COALESCE(vri.atualizado_por_codigo,'') AS atualizado_por_codigo
            FROM vendedor_pre_programacao_itens ppi
            LEFT JOIN vendedor_rascunho_itens vri
              ON vri.id = ppi.rascunho_item_id
            WHERE ppi.pre_programacao_id=?
            ORDER BY COALESCE(ppi.ordem,0), ppi.id
            """,
            (target,),
        )
        itens: List[Dict[str, Any]] = []
        item_ids: List[str] = []
        for row in cur.fetchall() or []:
            item_id = str(row["id"] or "").strip()
            if not item_id:
                continue
            item_ids.append(item_id)
            itens.append(
                {
                    "id": item_id,
                    "ordem": int(row["ordem"] or 0),
                    "cod_cliente": str(row["cod_cliente"] or "").strip().upper(),
                    "nome_cliente": str(row["nome_cliente"] or "").strip().upper(),
                    "cidade": str(row["cidade"] or "").strip().upper(),
                    "bairro": str(row["bairro"] or "").strip().upper(),
                    "endereco": str(row["endereco"] or "").strip().upper(),
                    "vendedor_cadastro": str(row["vendedor_cadastro"] or "").strip().upper(),
                    "vendedor_origem": str(row["vendedor_origem"] or "").strip().upper(),
                    "preco": float(row["preco"] or 0.0),
                    "caixas": int(row["caixas"] or 0),
                    "status": str(row["status"] or "PENDENTE").strip().upper(),
                    "observacao": str(row["observacao"] or "").strip(),
                    "alerta_codigo_programacao": str(row["alerta_codigo_programacao"] or "").strip().upper(),
                    "alerta_status_rota": str(row["alerta_status_rota"] or "").strip().upper(),
                    "criado_em": str(row["criado_em"] or "").strip(),
                    "updated_at": str(row["atualizado_em"] or "").strip(),
                    "criado_por_codigo": str(row["criado_por_codigo"] or "").strip().upper(),
                    "atualizado_por_codigo": str(row["atualizado_por_codigo"] or "").strip().upper(),
                }
            )

        return {
            "pre_programacao": {
                "id": str(header["id"] or "").strip(),
                "titulo": str(header["titulo"] or "").strip(),
                "observacao": str(header["observacao"] or "").strip(),
                "status": str(header["status"] or "ABERTA").strip().upper(),
                "criado_em": str(header["criado_em"] or "").strip(),
                "updated_at": str(header["atualizado_em"] or "").strip(),
                "criado_por_codigo": str(header["criado_por_codigo"] or "").strip().upper(),
                "atualizado_por_codigo": str(header["atualizado_por_codigo"] or "").strip().upper(),
            },
            "item_ids": item_ids,
            "itens": itens,
        }


@app.post("/vendedor/pre-programacoes/upsert")
def vendedor_pre_programacao_upsert(
    payload: VendedorPreProgramacaoUpsertIn,
    vendedor=Depends(get_current_vendedor),
):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pid = _clean_text(payload.id) or uuid4().hex.upper()
    titulo = _clean_text(payload.titulo)
    if not titulo:
        titulo = f"PRE-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    status = _clean_text(payload.status or "ABERTA").upper() or "ABERTA"
    if status not in {"ABERTA", "FECHADA"}:
        raise HTTPException(status_code=400, detail="status invalido.")

    cleaned_ids: List[str] = []
    for item_id in payload.item_ids or []:
        target = str(item_id or "").strip()
        if target and target not in cleaned_ids:
            cleaned_ids.append(target)

    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "vendedor_pre_programacoes"):
            raise HTTPException(status_code=500, detail="Tabela de pre-programacao indisponivel.")
        cur.execute(
            """
            INSERT INTO vendedor_pre_programacoes (
                id, titulo, observacao, status,
                criado_em, atualizado_em, criado_por_codigo, atualizado_por_codigo
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                titulo=excluded.titulo,
                observacao=excluded.observacao,
                status=excluded.status,
                atualizado_em=excluded.atualizado_em,
                atualizado_por_codigo=excluded.atualizado_por_codigo
            """,
            (
                pid,
                titulo,
                _clean_text(payload.observacao),
                status,
                now,
                now,
                str(vendedor["codigo"]).upper(),
                str(vendedor["codigo"]).upper(),
            ),
        )

        valid_ids = cleaned_ids
        if cleaned_ids:
            qmarks = ",".join(["?"] * len(cleaned_ids))
            cur.execute(
                f"SELECT id FROM vendedor_rascunho_itens WHERE id IN ({qmarks})",
                tuple(cleaned_ids),
            )
            existentes = {
                str(row["id"] or "").strip()
                for row in (cur.fetchall() or [])
                if str(row["id"] or "").strip()
            }
            valid_ids = [item_id for item_id in cleaned_ids if item_id in existentes]

        cur.execute(
            "DELETE FROM vendedor_pre_programacao_itens WHERE pre_programacao_id=?",
            (pid,),
        )
        for ordem, item_id in enumerate(valid_ids, start=1):
            cur.execute(
                """
                INSERT INTO vendedor_pre_programacao_itens (
                    pre_programacao_id, rascunho_item_id, ordem, criado_em, atualizado_em
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (pid, item_id, int(ordem), now, now),
            )

    return {"ok": True, "id": pid, "count": len(valid_ids)}


@app.delete("/vendedor/pre-programacoes/{pre_programacao_id}")
def vendedor_pre_programacao_remover(
    pre_programacao_id: str,
    vendedor=Depends(get_current_vendedor),
):
    target = _clean_text(pre_programacao_id)
    if not target:
        raise HTTPException(status_code=400, detail="pre_programacao_id obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM vendedor_pre_programacao_itens WHERE pre_programacao_id=?",
            (target,),
        )
        cur.execute("DELETE FROM vendedor_pre_programacoes WHERE id=?", (target,))
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "id": target, "deleted": deleted}


@app.get("/admin/motoristas/acesso")
def admin_listar_acesso_motoristas(_ok=Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {r[1] for r in (cur.fetchall() or [])}

        has_acesso = "acesso_liberado" in cols
        has_por = "acesso_liberado_por" in cols
        has_em = "acesso_liberado_em" in cols
        has_obs = "acesso_obs" in cols

        cur.execute(
            f"""
            SELECT
                id,
                TRIM(COALESCE(nome,'')) AS nome,
                UPPER(TRIM(COALESCE(codigo,''))) AS codigo,
                {("COALESCE(acesso_liberado, 1)" if has_acesso else "1")} AS acesso_liberado,
                {("TRIM(COALESCE(acesso_liberado_por,''))" if has_por else "''")} AS acesso_liberado_por,
                {("TRIM(COALESCE(acesso_liberado_em,''))" if has_em else "''")} AS acesso_liberado_em,
                {("TRIM(COALESCE(acesso_obs,''))" if has_obs else "''")} AS acesso_obs
            FROM motoristas
            WHERE TRIM(COALESCE(codigo,'')) <> ''
            ORDER BY nome
            """
        )
        out = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "id": int(r["id"]) if r["id"] is not None else None,
                    "nome": str(r["nome"] or ""),
                    "codigo": str(r["codigo"] or ""),
                    "acesso_liberado": int(r["acesso_liberado"] or 0),
                    "acesso_liberado_por": str(r["acesso_liberado_por"] or ""),
                    "acesso_liberado_em": str(r["acesso_liberado_em"] or ""),
                    "acesso_obs": str(r["acesso_obs"] or ""),
                }
            )
        return {"ok": True, "motoristas": out}


@app.post("/admin/motoristas/acesso/{codigo_motorista}")
def admin_set_acesso_motorista(
    codigo_motorista: str,
    payload: MotoristaAcessoIn,
    _ok=Depends(_require_desktop_secret),
):
    codigo = (codigo_motorista or "").strip().upper()
    if not codigo:
        raise HTTPException(status_code=400, detail="Codigo do motorista obrigatorio.")

    admin_nome = (payload.admin or "").strip() or "ADMIN"
    motivo = (payload.motivo or "").strip()
    liberado = 1 if bool(payload.liberado) else 0
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {r[1] for r in (cur.fetchall() or [])}
        if "acesso_liberado" not in cols:
            raise HTTPException(status_code=500, detail="Controle de acesso ainda não inicializado no banco.")

        cur.execute(
            "SELECT id, nome, codigo FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1",
            (codigo,),
        )
        m = cur.fetchone()
        if not m:
            raise HTTPException(status_code=404, detail="Motorista nao encontrado.")

        sets = ["acesso_liberado=?"]
        params: List[Any] = [liberado]
        if "acesso_liberado_por" in cols:
            sets.append("acesso_liberado_por=?")
            params.append(admin_nome)
        if "acesso_liberado_em" in cols:
            sets.append("acesso_liberado_em=?")
            params.append(ts)
        if "acesso_obs" in cols:
            sets.append("acesso_obs=?")
            params.append(motivo)
        params.append(int(m["id"]))
        cur.execute(f"UPDATE motoristas SET {', '.join(sets)} WHERE id=?", tuple(params))

        return {
            "ok": True,
            "codigo": str(m["codigo"] or "").strip().upper(),
            "nome": str(m["nome"] or "").strip().upper(),
            "acesso_liberado": liberado,
            "acesso_liberado_por": admin_nome,
            "acesso_liberado_em": ts,
            "acesso_obs": motivo,
        }


@app.post("/admin/motoristas/senha/{codigo_motorista}")
def admin_set_senha_motorista(
    codigo_motorista: str,
    payload: MotoristaSenhaIn,
    _ok=Depends(_require_desktop_secret),
):
    codigo = (codigo_motorista or "").strip().upper()
    if not codigo:
        raise HTTPException(status_code=400, detail="Codigo do motorista obrigatorio.")

    senha = (payload.nova_senha or "").strip()
    if len(senha) < 4 or len(senha) > 24:
        raise HTTPException(status_code=400, detail="Senha invalida. Use 4 a 24 caracteres.")

    admin_nome = (payload.admin or "").strip() or "ADMIN"
    motivo = (payload.motivo or "").strip()

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, nome, codigo FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1",
            (codigo,),
        )
        m = cur.fetchone()
        if not m:
            raise HTTPException(status_code=404, detail="Motorista nao encontrado.")

        senha_hash = hash_password_pbkdf2(senha)
        cur.execute("UPDATE motoristas SET senha=? WHERE id=?", (senha_hash, int(m["id"])))

        return {
            "ok": True,
            "codigo": str(m["codigo"] or "").strip().upper(),
            "nome": str(m["nome"] or "").strip().upper(),
            "senha_atualizada_por": admin_nome,
            "motivo": motivo,
        }


# =========================================================
# âœ… ROTAS ATIVAS (TODAS) - SEM FILTRAR POR MOTORISTA
# =========================================================
@app.get("/rotas/ativas_todas", response_model=List[RotaAtivaOut])
def listar_rotas_ativas_todas(m=Depends(get_current_motorista)):
    """
    Lista programações ativas de TODOS os motoristas.
    Isso é necessário para a tela de TRANSFERÊNCIA (destino).
    """
    if not ENABLE_ROTAS_ATIVAS_TODAS:
        raise HTTPException(status_code=403, detail="Endpoint desabilitado por configuração.")

    with get_conn() as conn:
        cur = conn.cursor()
        local_expr = _local_rota_expr(conn)
        local_carreg_expr = _local_carregamento_expr(conn)
        media_expr = _media_carregada_expr(conn)
        kg_carregado_expr = _kg_carregado_expr(conn)
        caixas_carregadas_expr = _caixas_carregadas_expr(conn)
        caixa_final_expr = _caixa_final_expr(conn)
        caixas_saldo_expr = _caixas_saldo_subquery(conn, "p")
        not_finalized_sql = _rotas_not_finalizadas_clause(conn, "p")
        equipe_cols_expr = _equipe_cols_expr(conn, "e")
        has_equipe_id = col_exists(conn, "programacoes", "equipe_id")
        equipe_id_select = "p.equipe_id," if has_equipe_id else "NULL AS equipe_id,"
        if has_equipe_id:
            equipe_join_on = """
              ON (
                (p.equipe IS NOT NULL AND TRIM(p.equipe) != '' AND UPPER(TRIM(e.codigo)) = UPPER(TRIM(p.equipe)))
                OR (p.equipe_id IS NOT NULL AND e.id = p.equipe_id)
              )
            """
        else:
            equipe_join_on = """
              ON (
                p.equipe IS NOT NULL AND TRIM(p.equipe) != '' AND UPPER(TRIM(e.codigo)) = UPPER(TRIM(p.equipe))
              )
            """
        cur.execute(
            """
            SELECT
                p.codigo_programacao,
                p.status,
                p.status_operacional,
                p.motorista,
                p.veiculo,
                p.equipe,
                """ + equipe_id_select + """
                """ + equipe_cols_expr + """,
                """ + local_expr + """,
                """ + local_carreg_expr + """,
                """ + media_expr + """,
                """ + kg_carregado_expr + """,
                """ + caixas_carregadas_expr + """,
                """ + caixa_final_expr + """,
                p.data_criacao,
                COALESCE(p.tipo_estimativa, 'KG') AS tipo_estimativa,
                COALESCE(p.caixas_estimado, 0) AS caixas_estimado,
                COALESCE(p.operacao_tipo, CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END) AS operacao_tipo,
                COALESCE(p.transbordo_modalidade, '') AS transbordo_modalidade,
                COALESCE(p.transbordo_observacao, '') AS transbordo_observacao,
                COALESCE(p.transbordo_grupo, '') AS transbordo_grupo,
                COALESCE(p.usuario_criacao, '') AS usuario_criacao,
                COALESCE(p.usuario_ultima_edicao, '') AS usuario_ultima_edicao,
                p.total_caixas,
                (
                    SELECT CAST(NULLIF(TRIM(v.capacidade_cx), '') AS INTEGER)
                    FROM veiculos v
                    WHERE UPPER(TRIM(v.placa)) = UPPER(TRIM(p.veiculo))
                       OR UPPER(TRIM(v.modelo)) = UPPER(TRIM(p.veiculo))
                    LIMIT 1
                ) AS capacidade_cx,
                """ + caixas_saldo_expr + """
            FROM programacoes p
            LEFT JOIN equipes e
            """ + equipe_join_on + """
            WHERE TRIM(COALESCE(p.codigo_programacao,''))<>''
              AND """ + not_finalized_sql + """
            ORDER BY p.id DESC
            LIMIT 200
            """
        )
        rows = cur.fetchall()
        response_rows = []
        for r in rows:
            d = _decorate_rota_row(row_to_dict(r), cur)
            codigo = str(d.get("codigo_programacao") or "").strip()
            pend_sub = _has_pending_substituicao(cur, codigo) if codigo else False
            d["substituicao_pendente"] = 1 if pend_sub else 0
            d["status_operacional"] = _status_operacional_especial(d, pend_substituicao=pend_sub)
            d = _attach_ultimo_km_veiculo(cur, d)
            d.update(_transferencias_resumo(cur, codigo))
            response_rows.append(d)

    return response_rows


@app.get("/rotas/ativas", response_model=List[RotaAtivaOut])
def rotas_ativas(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        local_expr = _local_rota_expr(conn)
        local_carreg_expr = _local_carregamento_expr(conn)
        media_expr = _media_carregada_expr(conn)
        kg_carregado_expr = _kg_carregado_expr(conn)
        caixas_carregadas_expr = _caixas_carregadas_expr(conn)
        caixa_final_expr = _caixa_final_expr(conn)
        caixas_saldo_expr = _caixas_saldo_subquery(conn, "p")
        not_finalized_sql = _rotas_not_finalizadas_clause(conn, "p")
        equipe_cols_expr = _equipe_cols_expr(conn, "e")
        has_equipe_id = col_exists(conn, "programacoes", "equipe_id")
        equipe_id_select = "p.equipe_id," if has_equipe_id else "NULL AS equipe_id,"
        if has_equipe_id:
            equipe_join_on = """
              ON (
                (p.equipe IS NOT NULL AND TRIM(p.equipe) != '' AND UPPER(TRIM(e.codigo)) = UPPER(TRIM(p.equipe)))
                OR (p.equipe_id IS NOT NULL AND e.id = p.equipe_id)
              )
            """
        else:
            equipe_join_on = """
              ON (
                p.equipe IS NOT NULL AND TRIM(p.equipe) != '' AND UPPER(TRIM(e.codigo)) = UPPER(TRIM(p.equipe))
              )
            """
        owner_sql, owner_params = _owner_filter_for_programacoes(conn, m, "p")
        cur.execute(
            """
            SELECT
                p.codigo_programacao,
                p.status,
                p.status_operacional,
                p.motorista,
                p.veiculo,
                p.equipe,
                """ + equipe_id_select + """
                """ + equipe_cols_expr + """,
                """ + local_expr + """,
                """ + local_carreg_expr + """,
                """ + media_expr + """,
                """ + kg_carregado_expr + """,
                """ + caixas_carregadas_expr + """,
                """ + caixa_final_expr + """,
                p.data_criacao,
                COALESCE(p.tipo_estimativa, 'KG') AS tipo_estimativa,
                COALESCE(p.caixas_estimado, 0) AS caixas_estimado,
                COALESCE(p.operacao_tipo, CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END) AS operacao_tipo,
                COALESCE(p.transbordo_modalidade, '') AS transbordo_modalidade,
                COALESCE(p.transbordo_observacao, '') AS transbordo_observacao,
                COALESCE(p.transbordo_grupo, '') AS transbordo_grupo,
                COALESCE(p.usuario_criacao, '') AS usuario_criacao,
                COALESCE(p.usuario_ultima_edicao, '') AS usuario_ultima_edicao,
                p.total_caixas,
                (
                    SELECT CAST(NULLIF(TRIM(v.capacidade_cx), '') AS INTEGER)
                    FROM veiculos v
                    WHERE UPPER(TRIM(v.placa)) = UPPER(TRIM(p.veiculo))
                       OR UPPER(TRIM(v.modelo)) = UPPER(TRIM(p.veiculo))
                    LIMIT 1
                ) AS capacidade_cx,
                """ + caixas_saldo_expr + """
            FROM programacoes p
            LEFT JOIN equipes e
            """ + equipe_join_on + """
            WHERE """ + owner_sql + """
              AND TRIM(COALESCE(p.codigo_programacao,''))<>''
              AND """ + not_finalized_sql + """
            ORDER BY p.id DESC
            LIMIT 200
            """,
            owner_params,
        )
        rows = cur.fetchall()
        response_rows = []
        for r in rows:
            d = _decorate_rota_row(row_to_dict(r), cur)
            codigo = str(d.get("codigo_programacao") or "").strip()
            pend_sub = _has_pending_substituicao(cur, codigo) if codigo else False
            d["substituicao_pendente"] = 1 if pend_sub else 0
            d["status_operacional"] = _status_operacional_especial(d, pend_substituicao=pend_sub)
            d = _attach_ultimo_km_veiculo(cur, d)
            d.update(_transferencias_resumo(cur, codigo))
            response_rows.append(d)

    return response_rows


@app.get("/rotas/{codigo_programacao}", response_model=RotaDetalheOut)
def rota_detalhe(codigo_programacao: str, m=Depends(get_current_motorista)):
    codigo_programacao = (codigo_programacao or "").strip()

    with get_conn() as conn:
        cur = conn.cursor()
        owner_sql, owner_params = _owner_filter_for_programacoes(conn, m, "p")
        caixas_saldo_expr = _caixas_saldo_subquery(conn, "p")

        cur.execute(
            """
            SELECT
                p.*,
                (
                    SELECT CAST(NULLIF(TRIM(v.capacidade_cx), '') AS INTEGER)
                    FROM veiculos v
                    WHERE UPPER(TRIM(v.placa)) = UPPER(TRIM(p.veiculo))
                       OR UPPER(TRIM(v.modelo)) = UPPER(TRIM(p.veiculo))
                    LIMIT 1
                ) AS capacidade_cx,
                """ + caixas_saldo_expr + """
            FROM programacoes p
            WHERE p.codigo_programacao=?
              AND """ + owner_sql + """
            ORDER BY p.id DESC
            LIMIT 1
            """,
            (codigo_programacao, *owner_params),
        )
        pr = cur.fetchone()

        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")

        equipes_map = _load_equipes_map(cur)

        itens_select_expr = _programacao_itens_select_expr(conn, "pi")
        cur.execute(
            """
            SELECT """ + itens_select_expr + """
            FROM programacao_itens pi
            WHERE pi.codigo_programacao=?
            ORDER BY id ASC
            LIMIT 2000
            """,
            (codigo_programacao,),
        )
        itens = cur.fetchall()

        cur.execute(
            """
            SELECT *
            FROM programacao_itens_controle
            WHERE codigo_programacao=?
            """,
            (codigo_programacao,),
        )
        controles = cur.fetchall()

        rota = row_to_dict(pr)
        try:
            rota = _apply_equipe_nome(rota, equipes_map, cur)
            rota = _decorate_rota_row(rota, cur)
            pend_sub = _has_pending_substituicao(cur, codigo_programacao)
            rota["substituicao_pendente"] = 1 if pend_sub else 0
            rota["status_operacional"] = _status_operacional_especial(rota, pend_substituicao=pend_sub)
            rota["substituicoes"] = _list_substituicoes_por_rota(cur, codigo_programacao, limit=20)
            rota.update(_transferencias_resumo(cur, codigo_programacao))
            rota["transferencias_saida_lista"] = _list_transferencias_por_origem(conn, codigo_programacao)
            rota["transferencias_entrada_lista"] = _list_transferencias_por_destino(conn, codigo_programacao)
            rota = _attach_ultimo_km_veiculo(cur, rota)
        except Exception:
            logging.debug("Falha ao decorar rota desktop; retornando cabecalho bruto.", exc_info=True)
            rota.setdefault("substituicao_pendente", 0)
            rota.setdefault("substituicoes", [])
            rota.update(_transferencias_resumo(cur, codigo_programacao))
            rota = _attach_ultimo_km_veiculo(cur, rota)

        controle_map = {}
        for row in controles:
            rd = row_to_dict(row)
            key = (
                str(rd.get("cod_cliente") or "").strip().upper(),
                _norm_pedido_key(rd.get("pedido")),
            )
            controle_map[key] = rd

        clientes = []
        for i in itens:
            d = row_to_dict(i)
            cod = str(d.get("cod_cliente") or "").strip().upper()
            ped = _norm_pedido_key(d.get("pedido"))
            c = controle_map.get((cod, ped))
            if c:
                d["mortalidade_aves"] = c.get("mortalidade_aves")
                d["media_aplicada"] = c.get("media_aplicada")
                d["peso_previsto"] = c.get("peso_previsto")
                d["recebido_valor"] = c.get("valor_recebido")
                d["valor_recebido"] = c.get("valor_recebido")
                d["recebido_forma"] = c.get("forma_recebimento")
                d["recebido_obs"] = c.get("obs_recebimento")
                d["status_pedido"] = c.get("status_pedido")
                d["alteracao_tipo"] = c.get("alteracao_tipo")
                d["alteracao_detalhe"] = c.get("alteracao_detalhe")
                d["caixas_atual"] = c.get("caixas_atual")
                d["preco_atual"] = c.get("preco_atual")
                d["alterado_em"] = c.get("alterado_em")
                d["alterado_por"] = c.get("alterado_por")
                d["lat_evento"] = c.get("lat_evento")
                d["lon_evento"] = c.get("lon_evento")
                d["endereco_evento"] = c.get("endereco_evento")
                d["cidade_evento"] = c.get("cidade_evento")
                d["bairro_evento"] = c.get("bairro_evento")
                if c.get("ordem_sugerida") is not None:
                    d["ordem_sugerida"] = c.get("ordem_sugerida")
                if c.get("eta") not in (None, ""):
                    d["eta"] = c.get("eta")
                if c.get("distancia") is not None:
                    d["distancia"] = c.get("distancia")
                if c.get("confianca_localizacao") is not None:
                    d["confianca_localizacao"] = c.get("confianca_localizacao")
            clientes.append(d)

        return {"rota": rota, "clientes": clientes}


@app.get("/desktop/rotas/{codigo_programacao}", response_model=RotaDetalheOut)
def rota_detalhe_desktop(codigo_programacao: str, _ok: bool = Depends(_require_desktop_secret)):
    codigo_programacao = (codigo_programacao or "").strip()

    with get_conn() as conn:
        cur = conn.cursor()
        caixas_saldo_expr = _caixas_saldo_subquery(conn, "p")

        cur.execute(
            """
            SELECT
                p.*,
                (
                    SELECT CAST(NULLIF(TRIM(v.capacidade_cx), '') AS INTEGER)
                    FROM veiculos v
                    WHERE UPPER(TRIM(v.placa)) = UPPER(TRIM(p.veiculo))
                       OR UPPER(TRIM(v.modelo)) = UPPER(TRIM(p.veiculo))
                    LIMIT 1
                ) AS capacidade_cx,
                """ + caixas_saldo_expr + """
            FROM programacoes p
            WHERE p.codigo_programacao=?
            ORDER BY p.id DESC
            LIMIT 1
            """,
            (codigo_programacao,),
        )
        pr = cur.fetchone()

        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada")

        try:
            equipes_map = _load_equipes_map(cur)
        except Exception:
            logging.debug("Falha ao carregar equipes para rota desktop; usando fallback vazio.", exc_info=True)
            equipes_map = {}

        itens = []
        try:
            itens_select_expr = _programacao_itens_select_expr(conn, "pi")
            cur.execute(
                """
                SELECT """ + itens_select_expr + """
                FROM programacao_itens pi
                WHERE pi.codigo_programacao=?
                ORDER BY id ASC
                LIMIT 2000
                """,
                (codigo_programacao,),
            )
            itens = cur.fetchall() or []
        except Exception:
            logging.debug("Falha ao carregar itens da programacao no detalhe desktop; retornando lista vazia.", exc_info=True)
            itens = []

        controles = []
        try:
            cur.execute(
                """
                SELECT *
                FROM programacao_itens_controle
                WHERE codigo_programacao=?
                """,
                (codigo_programacao,),
            )
            controles = cur.fetchall() or []
        except Exception:
            logging.debug("Falha ao carregar controles da programacao no detalhe desktop; retornando lista vazia.", exc_info=True)
            controles = []

        # Último log por cliente/pedido para enriquecer rastreabilidade
        log_map = {}
        try:
            cur.execute(
                """
                SELECT cod_cliente, COALESCE(pedido, '') AS pedido, payload_json, created_at, id
                FROM programacao_itens_log
                WHERE codigo_programacao=?
                ORDER BY id DESC
                LIMIT 5000
                """,
                (codigo_programacao,),
            )
            for lr in (cur.fetchall() or []):
                cod_l = str(lr["cod_cliente"] or "").strip().upper()
                ped_l = _norm_pedido_key(lr["pedido"])
                key_l = (cod_l, ped_l)
                if key_l in log_map:
                    continue
                payload_raw = lr["payload_json"] or ""
                payload_obj = {}
                if payload_raw:
                    try:
                        tmp = json.loads(payload_raw)
                        if isinstance(tmp, dict):
                            payload_obj = tmp
                    except Exception:
                        payload_obj = {}
                log_map[key_l] = {
                    "created_at": str(lr["created_at"] or ""),
                    "payload": payload_obj,
                }
        except Exception:
            log_map = {}

        rota = row_to_dict(pr)
        try:
            rota = _apply_equipe_nome(rota, equipes_map, cur)
            rota = _decorate_rota_row(rota, cur)
            pend_sub = _has_pending_substituicao(cur, codigo_programacao)
            rota["substituicao_pendente"] = 1 if pend_sub else 0
            rota["status_operacional"] = _status_operacional_especial(rota, pend_substituicao=pend_sub)
            rota["substituicoes"] = _list_substituicoes_por_rota(cur, codigo_programacao, limit=20)
            rota.update(_transferencias_resumo(cur, codigo_programacao))
            rota["transferencias_saida_lista"] = _list_transferencias_por_origem(conn, codigo_programacao)
            rota["transferencias_entrada_lista"] = _list_transferencias_por_destino(conn, codigo_programacao)
            rota = _attach_ultimo_km_veiculo(cur, rota)
        except Exception:
            logging.debug("Falha ao decorar rota desktop; retornando cabecalho bruto.", exc_info=True)
            rota.setdefault("substituicao_pendente", 0)
            rota.setdefault("substituicoes", [])
            rota.update(_transferencias_resumo(cur, codigo_programacao))
            rota = _attach_ultimo_km_veiculo(cur, rota)

        controle_map = {}
        for row in controles:
            rd = row_to_dict(row)
            key = (
                str(rd.get("cod_cliente") or "").strip().upper(),
                _norm_pedido_key(rd.get("pedido")),
            )
            controle_map[key] = rd

        clientes = []
        for i in itens:
            d = row_to_dict(i)
            cod = str(d.get("cod_cliente") or "").strip().upper()
            ped = _norm_pedido_key(d.get("pedido"))
            c = controle_map.get((cod, ped))
            if c:
                d["mortalidade_aves"] = c.get("mortalidade_aves")
                d["media_aplicada"] = c.get("media_aplicada")
                d["peso_previsto"] = c.get("peso_previsto")
                d["recebido_valor"] = c.get("valor_recebido")
                d["valor_recebido"] = c.get("valor_recebido")
                d["recebido_forma"] = c.get("forma_recebimento")
                d["recebido_obs"] = c.get("obs_recebimento")
                d["status_pedido"] = c.get("status_pedido")
                d["alteracao_tipo"] = c.get("alteracao_tipo")
                d["alteracao_detalhe"] = c.get("alteracao_detalhe")
                d["caixas_atual"] = c.get("caixas_atual")
                d["preco_atual"] = c.get("preco_atual")
                d["alterado_em"] = c.get("alterado_em") or c.get("updated_at")
                d["alterado_por"] = c.get("alterado_por")
                d["lat_evento"] = c.get("lat_evento")
                d["lon_evento"] = c.get("lon_evento")
                d["endereco_evento"] = c.get("endereco_evento")
                d["cidade_evento"] = c.get("cidade_evento")
                d["bairro_evento"] = c.get("bairro_evento")
                if c.get("ordem_sugerida") is not None:
                    d["ordem_sugerida"] = c.get("ordem_sugerida")
                if c.get("eta") not in (None, ""):
                    d["eta"] = c.get("eta")
                if c.get("distancia") is not None:
                    d["distancia"] = c.get("distancia")
                if c.get("confianca_localizacao") is not None:
                    d["confianca_localizacao"] = c.get("confianca_localizacao")

            # Fallback por logs do item (quando controle não tiver horário/local completos)
            lg = log_map.get((cod, ped))
            if lg:
                payload_obj = lg.get("payload") if isinstance(lg, dict) else {}
                if not isinstance(payload_obj, dict):
                    payload_obj = {}
                created_at = str(lg.get("created_at") or "")
                if not d.get("alterado_em"):
                    d["alterado_em"] = created_at
                d["evento_datahora"] = created_at
                d["lat_evento"] = (
                    payload_obj.get("lat_evento")
                    if payload_obj.get("lat_evento") not in (None, "")
                    else (payload_obj.get("latitude") if payload_obj.get("latitude") not in (None, "") else payload_obj.get("lat"))
                )
                d["lon_evento"] = (
                    payload_obj.get("lon_evento")
                    if payload_obj.get("lon_evento") not in (None, "")
                    else (
                        payload_obj.get("longitude")
                        if payload_obj.get("longitude") not in (None, "")
                        else (payload_obj.get("lon") if payload_obj.get("lon") not in (None, "") else payload_obj.get("lng"))
                    )
                )
                d["endereco_evento"] = payload_obj.get("endereco") or payload_obj.get("endereco_cliente") or ""
                d["cidade_evento"] = payload_obj.get("cidade") or payload_obj.get("cidade_cliente") or ""
                d["bairro_evento"] = payload_obj.get("bairro") or payload_obj.get("bairro_cliente") or ""
                if d.get("ordem_sugerida") in (None, ""):
                    d["ordem_sugerida"] = (
                        payload_obj.get("ordem_sugerida")
                        if payload_obj.get("ordem_sugerida") not in (None, "")
                        else payload_obj.get("ordem")
                    )
                if d.get("eta") in (None, "") and payload_obj.get("eta") not in (None, ""):
                    d["eta"] = payload_obj.get("eta")
                if d.get("distancia") in (None, "") and payload_obj.get("distancia") not in (None, ""):
                    d["distancia"] = payload_obj.get("distancia")
                if d.get("confianca_localizacao") in (None, "") and payload_obj.get("confianca_localizacao") not in (None, ""):
                    d["confianca_localizacao"] = payload_obj.get("confianca_localizacao")

            clientes.append(d)

        return {"rota": rota, "clientes": clientes}


@app.get("/desktop/rotas/{codigo_programacao}/recebimentos")
def desktop_listar_recebimentos(
    codigo_programacao: str,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                id,
                UPPER(TRIM(COALESCE(cod_cliente,''))) AS cod_cliente,
                TRIM(COALESCE(nome_cliente,'')) AS nome_cliente,
                COALESCE(valor,0) AS valor,
                UPPER(TRIM(COALESCE(forma_pagamento,''))) AS forma_pagamento,
                TRIM(COALESCE(observacao,'')) AS observacao,
                TRIM(COALESCE(num_nf,'')) AS num_nf,
                TRIM(COALESCE(data_registro,'')) AS data_registro
            FROM recebimentos
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
            ORDER BY id DESC
            """,
            (prog,),
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "id": int(r["id"] or 0),
                    "cod_cliente": str(r["cod_cliente"] or ""),
                    "nome_cliente": str(r["nome_cliente"] or ""),
                    "valor": float(r["valor"] or 0.0),
                    "forma_pagamento": str(r["forma_pagamento"] or ""),
                    "observacao": str(r["observacao"] or ""),
                    "num_nf": str(r["num_nf"] or ""),
                    "data_registro": str(r["data_registro"] or ""),
                }
            )
    return {"ok": True, "codigo_programacao": prog, "recebimentos": out}


@app.post("/desktop/rotas/{codigo_programacao}/recebimentos")
def desktop_criar_recebimento(
    codigo_programacao: str,
    payload: DesktopRecebimentoIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    cod = str(payload.cod_cliente or "").strip().upper()
    nome = str(payload.nome_cliente or "").strip().upper()
    forma = str(payload.forma_pagamento or "DINHEIRO").strip().upper() or "DINHEIRO"
    formas_validas = {"DINHEIRO", "PIX", "CARTAO", "BOLETO", "OUTRO"}
    obs = str(payload.observacao or "").strip().upper()
    num_nf = str(payload.num_nf or "").strip().upper()
    valor = float(payload.valor or 0.0)
    if not cod or not nome:
        raise HTTPException(status_code=400, detail="cod_cliente e nome_cliente obrigatorios.")
    if valor <= 0:
        raise HTTPException(status_code=400, detail="valor deve ser maior que zero.")
    if forma not in formas_validas:
        raise HTTPException(status_code=400, detail="forma_pagamento invalida.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        _ensure_programacao_has_cliente(cur, prog, cod)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            """
            INSERT INTO recebimentos
                (codigo_programacao, cod_cliente, nome_cliente, valor, forma_pagamento, observacao, num_nf, data_registro)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (prog, cod, nome, valor, forma, obs, num_nf, ts),
        )
        rid = int(cur.lastrowid or 0)
    return {"ok": True, "id": rid, "codigo_programacao": prog}


@app.delete("/desktop/rotas/{codigo_programacao}/recebimentos/{cod_cliente}")
def desktop_zerar_recebimentos_cliente(
    codigo_programacao: str,
    cod_cliente: str,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    cod = (cod_cliente or "").strip().upper()
    if not prog or not cod:
        raise HTTPException(status_code=400, detail="codigo_programacao e cod_cliente obrigatorios.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        cur.execute(
            "DELETE FROM recebimentos WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?)) AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))",
            (prog, cod),
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "deleted": deleted, "codigo_programacao": prog, "cod_cliente": cod}


def _upsert_diarias_despesas_desktop(
    cur: sqlite3.Cursor,
    prog: str,
    *,
    total_motorista: float,
    total_ajudantes: float,
    observacao_motorista: str,
    observacao_ajudantes: str,
) -> None:
    cur.execute("PRAGMA table_info(despesas)")
    cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if not cols:
        return

    if "id_local" in cols:
        cur.execute(
            """
            DELETE FROM despesas
             WHERE codigo_programacao=?
               AND UPPER(TRIM(COALESCE(categoria, '')))=?
               AND (
                    UPPER(TRIM(COALESCE(descricao, ''))) IN (?, ?)
                 OR UPPER(TRIM(COALESCE(id_local, ''))) IN (?, ?)
               )
            """,
            (
                prog,
                "DIARIAS",
                "DIARIAS MOTORISTA",
                "DIARIAS AJUDANTES",
                "AUTO_DIARIA_MOTORISTA",
                "AUTO_DIARIA_AJUDANTES",
            ),
        )
    else:
        cur.execute(
            """
            DELETE FROM despesas
             WHERE codigo_programacao=?
               AND UPPER(TRIM(COALESCE(categoria, '')))=?
               AND UPPER(TRIM(COALESCE(descricao, ''))) IN (?, ?)
            """,
            (prog, "DIARIAS", "DIARIAS MOTORISTA", "DIARIAS AJUDANTES"),
        )

    now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for descricao, valor, observacao, id_local in (
        ("DIARIAS MOTORISTA", round(float(total_motorista or 0.0), 2), observacao_motorista, "AUTO_DIARIA_MOTORISTA"),
        ("DIARIAS AJUDANTES", round(float(total_ajudantes or 0.0), 2), observacao_ajudantes, "AUTO_DIARIA_AJUDANTES"),
    ):
        if valor <= 0:
            continue
        data = {
            "codigo_programacao": prog,
            "descricao": descricao,
            "valor": valor,
            "data_registro": now_s,
            "tipo_despesa": "DIARIAS",
            "categoria": "DIARIAS",
            "observacao": observacao,
            "id_local": id_local,
            "forma_pagamento": "PAGO",
            "origem": "RECEBIMENTOS",
            "registrado_em": now_s,
        }
        ins_cols = [k for k in data if k.lower() in cols]
        cur.execute(
            f"INSERT INTO despesas ({', '.join(ins_cols)}) VALUES ({', '.join('?' for _ in ins_cols)})",
            tuple(data[k] for k in ins_cols),
        )


@app.put("/desktop/rotas/{codigo_programacao}/cabecalho")
def desktop_atualizar_cabecalho_rota(
    codigo_programacao: str,
    payload: DesktopRotaCabecalhoIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    data_saida = _validate_iso_date_optional(payload.data_saida, "data_saida")
    hora_saida = _validate_time_optional(payload.hora_saida, "hora_saida")
    data_chegada = _validate_iso_date_optional(payload.data_chegada, "data_chegada")
    hora_chegada = _validate_time_optional(payload.hora_chegada, "hora_chegada")
    diaria_motorista_valor = float(payload.diaria_motorista_valor or 0.0) if payload.diaria_motorista_valor is not None else None
    if diaria_motorista_valor is not None and diaria_motorista_valor < 0:
        raise HTTPException(status_code=400, detail="diaria_motorista_valor invalida.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        sets: List[str] = []
        vals: List[Any] = []
        nf_saldo_payload = payload.nf_saldo
        if (
            nf_saldo_payload is None
            and payload.nf_kg is not None
            and payload.nf_kg_carregado is not None
            and float(payload.nf_kg) > 0
            and float(payload.nf_kg_carregado) > 0
        ):
            nf_saldo_payload = round(max(float(payload.nf_kg) - float(payload.nf_kg_carregado), 0.0), 2)
        if "data_saida" in cols_prog:
            sets.append("data_saida=?")
            vals.append(data_saida)
        if "hora_saida" in cols_prog:
            sets.append("hora_saida=?")
            vals.append(hora_saida)
        if "data_chegada" in cols_prog:
            sets.append("data_chegada=?")
            vals.append(data_chegada)
        if "hora_chegada" in cols_prog:
            sets.append("hora_chegada=?")
            vals.append(hora_chegada)
        if ("diaria_motorista_valor" in cols_prog) and (diaria_motorista_valor is not None):
            sets.append("diaria_motorista_valor=?")
            vals.append(diaria_motorista_valor)

        if not sets:
            raise HTTPException(status_code=500, detail="Colunas de cabecalho indisponiveis em programacoes.")

        vals.append(prog)
        cur.execute(
            f"UPDATE programacoes SET {', '.join(sets)} WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))",
            tuple(vals),
        )
        updated = int(cur.rowcount or 0)
        if updated <= 0:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")

        has_diarias_payload = any(
            value is not None
            for value in (
                payload.qtd_diarias,
                payload.qtd_ajudantes,
                payload.total_motorista,
                payload.total_ajudantes,
                payload.observacao_motorista,
                payload.observacao_ajudantes,
            )
        )
        if has_diarias_payload:
            qtd = float(payload.qtd_diarias or 0.0)
            qtd_ajudantes = int(payload.qtd_ajudantes or 0)
            total_mot = float(payload.total_motorista or 0.0)
            total_ajud = float(payload.total_ajudantes or 0.0)
            if qtd < 0 or qtd_ajudantes < 0 or total_mot < 0 or total_ajud < 0:
                raise HTTPException(status_code=400, detail="Valores de diarias invalidos.")
            obs_motorista = str(payload.observacao_motorista or "").strip().upper()
            obs_ajudantes = str(payload.observacao_ajudantes or "").strip().upper()
            _upsert_diarias_despesas_desktop(
                cur,
                prog,
                total_motorista=total_mot,
                total_ajudantes=total_ajud,
                observacao_motorista=obs_motorista,
                observacao_ajudantes=obs_ajudantes,
            )
    return {"ok": True, "codigo_programacao": prog, "updated": updated}


@app.put("/desktop/rotas/{codigo_programacao}/financeiro")
def desktop_atualizar_financeiro_rota(
    codigo_programacao: str,
    payload: DesktopRotaFinanceiroIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    def _add_num(sets: List[str], vals: List[Any], cols: set, col: str, value: Optional[Any], caster):
        if value is None:
            return
        if col in cols:
            sets.append(f"{col}=?")
            vals.append(caster(value))

    def _ensure_non_negative(name: str, value: Optional[Any]):
        if value is None:
            return
        if float(value) < 0:
            raise HTTPException(status_code=400, detail=f"{name} invalido.")

    for field_name in (
        "nf_kg",
        "nf_caixas",
        "nf_kg_carregado",
        "nf_kg_vendido",
        "nf_preco",
        "media",
        "nf_caixa_final",
        "km_inicial",
        "km_final",
        "litros",
        "km_rodado",
        "media_km_l",
        "custo_km",
        "ced_200_qtd",
        "ced_100_qtd",
        "ced_50_qtd",
        "ced_20_qtd",
        "ced_10_qtd",
        "ced_5_qtd",
        "ced_2_qtd",
        "valor_dinheiro",
        "pix_motorista",
        "adiantamento",
    ):
        _ensure_non_negative(field_name, getattr(payload, field_name))

    if (
        payload.km_inicial is not None
        and payload.km_final is not None
        and float(payload.km_inicial) > 0
        and float(payload.km_final) > 0
        and float(payload.km_final) < float(payload.km_inicial)
    ):
        raise HTTPException(status_code=400, detail="km_final nao pode ser menor que km_inicial.")
    if (
        payload.nf_kg is not None
        and payload.nf_kg_carregado is not None
        and float(payload.nf_kg) > 0
        and float(payload.nf_kg_carregado) > float(payload.nf_kg)
    ):
        logging.warning(
            "Financeiro com KG carregado maior que NF | prog=%s | nf_kg=%s | carregado=%s",
            prog,
            payload.nf_kg,
            payload.nf_kg_carregado,
        )
    if (
        payload.nf_kg_carregado is not None
        and payload.nf_kg_vendido is not None
        and float(payload.nf_kg_carregado) > 0
        and float(payload.nf_kg_vendido) > float(payload.nf_kg_carregado)
    ):
        logging.warning(
            "Financeiro com KG vendido maior que carregado | prog=%s | carregado=%s | vendido=%s",
            prog,
            payload.nf_kg_carregado,
            payload.nf_kg_vendido,
        )

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        sets: List[str] = []
        vals: List[Any] = []

        if payload.nf_numero is not None:
            nf_numero = str(payload.nf_numero or "").strip()
            if "nf_numero" in cols_prog:
                sets.append("nf_numero=?")
                vals.append(nf_numero)
            if "num_nf" in cols_prog:
                sets.append("num_nf=?")
                vals.append(nf_numero)
        _add_num(sets, vals, cols_prog, "nf_kg", payload.nf_kg, float)
        _add_num(sets, vals, cols_prog, "nf_caixas", payload.nf_caixas, int)
        _add_num(sets, vals, cols_prog, "nf_kg_carregado", payload.nf_kg_carregado, float)
        _add_num(sets, vals, cols_prog, "nf_kg_vendido", payload.nf_kg_vendido, float)
        _add_num(sets, vals, cols_prog, "nf_saldo", nf_saldo_payload, float)
        _add_num(sets, vals, cols_prog, "nf_preco", payload.nf_preco, float)
        if (payload.nf_preco is not None) and ("preco_nf" in cols_prog):
            sets.append("preco_nf=?")
            vals.append(float(payload.nf_preco))
        _add_num(sets, vals, cols_prog, "media", payload.media, float)

        if payload.nf_caixa_final is not None:
            caixa_final = int(payload.nf_caixa_final or 0)
            if "aves_caixa_final" in cols_prog:
                sets.append("aves_caixa_final=?")
                vals.append(caixa_final)
            if "qnt_aves_caixa_final" in cols_prog:
                sets.append("qnt_aves_caixa_final=?")
                vals.append(caixa_final)

        _add_num(sets, vals, cols_prog, "km_inicial", payload.km_inicial, float)
        _add_num(sets, vals, cols_prog, "km_final", payload.km_final, float)
        _add_num(sets, vals, cols_prog, "litros", payload.litros, float)
        _add_num(sets, vals, cols_prog, "km_rodado", payload.km_rodado, float)
        _add_num(sets, vals, cols_prog, "media_km_l", payload.media_km_l, float)
        _add_num(sets, vals, cols_prog, "custo_km", payload.custo_km, float)

        _add_num(sets, vals, cols_prog, "ced_200_qtd", payload.ced_200_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_100_qtd", payload.ced_100_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_50_qtd", payload.ced_50_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_20_qtd", payload.ced_20_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_10_qtd", payload.ced_10_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_5_qtd", payload.ced_5_qtd, int)
        _add_num(sets, vals, cols_prog, "ced_2_qtd", payload.ced_2_qtd, int)
        _add_num(sets, vals, cols_prog, "valor_dinheiro", payload.valor_dinheiro, float)
        _add_num(sets, vals, cols_prog, "pix_motorista", payload.pix_motorista, float)

        if payload.adiantamento is not None:
            adiant = float(payload.adiantamento or 0.0)
            if "adiantamento" in cols_prog:
                sets.append("adiantamento=?")
                vals.append(adiant)
            if "adiantamento_rota" in cols_prog:
                sets.append("adiantamento_rota=?")
                vals.append(adiant)

        if payload.adiantamento_origem is not None and "adiantamento_origem" in cols_prog:
            sets.append("adiantamento_origem=?")
            vals.append(str(payload.adiantamento_origem or "").strip().upper())

        if payload.rota_observacao is not None and "rota_observacao" in cols_prog:
            sets.append("rota_observacao=?")
            vals.append(str(payload.rota_observacao or "").strip())

        if not sets:
            raise HTTPException(status_code=500, detail="Colunas financeiras indisponiveis em programacoes.")

        vals.append(prog)
        cur.execute(
            f"UPDATE programacoes SET {', '.join(sets)} WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))",
            tuple(vals),
        )
        updated = int(cur.rowcount or 0)
        if updated <= 0:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")

    return {"ok": True, "codigo_programacao": prog, "updated": updated}


@app.post("/desktop/rotas/{codigo_programacao}/diarias/sync")
def desktop_sync_diarias_despesas(
    codigo_programacao: str,
    payload: DesktopDiariasSyncIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    qtd = float(payload.qtd_diarias or 0.0)
    qtd_ajudantes = int(payload.qtd_ajudantes or 0)
    total_mot = float(payload.total_motorista or 0.0)
    total_ajud = float(payload.total_ajudantes or 0.0)
    if qtd < 0 or qtd_ajudantes < 0 or total_mot < 0 or total_ajud < 0:
        raise HTTPException(status_code=400, detail="Valores de diarias invalidos.")

    obs_motorista = str(payload.observacao_motorista or "").strip().upper()
    obs_ajudantes = str(payload.observacao_ajudantes or "").strip().upper()

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        _upsert_diarias_despesas_desktop(
            cur,
            prog,
            total_motorista=total_mot,
            total_ajudantes=total_ajud,
            observacao_motorista=obs_motorista,
            observacao_ajudantes=obs_ajudantes,
        )
    return {
        "ok": True,
        "codigo_programacao": prog,
        "qtd_diarias": qtd,
        "qtd_ajudantes": qtd_ajudantes,
        "total_motorista": total_mot,
        "total_ajudantes": total_ajud,
        "total_geral": round(total_mot + total_ajud, 2),
    }


@app.post("/desktop/rotas/{codigo_programacao}/clientes/manual")
def desktop_upsert_cliente_manual_programacao(
    codigo_programacao: str,
    payload: DesktopProgramacaoClienteManualIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    cod = str(payload.cod_cliente or "").strip().upper()
    nome = str(payload.nome_cliente or "").strip().upper()
    if not prog or not cod or not nome:
        raise HTTPException(status_code=400, detail="codigo_programacao, cod_cliente e nome_cliente sao obrigatorios.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        cur.execute(
            """
            SELECT COUNT(1) AS n
            FROM programacao_itens
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
              AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))
            """,
            (prog, cod),
        )
        row = cur.fetchone()
        exists = int((row["n"] if row else 0) or 0)
        if exists:
            cur.execute(
                """
                UPDATE programacao_itens
                   SET nome_cliente=?
                 WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
                   AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))
                """,
                (nome, prog, cod),
            )
            action = "updated"
        else:
            cur.execute(
                """
                INSERT INTO programacao_itens
                    (codigo_programacao, cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido)
                VALUES (?, ?, ?, 0, 0, 0, '', '', 'MANUAL')
                """,
                (prog, cod, nome),
            )
            action = "created"
    return {"ok": True, "codigo_programacao": prog, "cod_cliente": cod, "nome_cliente": nome, "action": action}


@app.get("/desktop/rotas/{codigo_programacao}/despesas")
def desktop_listar_despesas(
    codigo_programacao: str,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(despesas)")
        cols_desp = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        id_local_expr = "TRIM(COALESCE(id_local,''))" if "id_local" in cols_desp else "''"
        forma_expr = "TRIM(COALESCE(forma_pagamento,''))" if "forma_pagamento" in cols_desp else "''"
        origem_expr = "TRIM(COALESCE(origem,''))" if "origem" in cols_desp else "''"
        cur.execute(
            f"""
            SELECT
                id,
                UPPER(TRIM(COALESCE(codigo_programacao,''))) AS codigo_programacao,
                TRIM(COALESCE(descricao,'')) AS descricao,
                COALESCE(valor,0) AS valor,
                UPPER(TRIM(COALESCE(categoria,'OUTROS'))) AS categoria,
                TRIM(COALESCE(observacao,'')) AS observacao,
                TRIM(COALESCE(data_registro,'')) AS data_registro,
                {id_local_expr} AS id_local,
                {forma_expr} AS forma_pagamento,
                {origem_expr} AS origem
            FROM despesas
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
            ORDER BY id DESC
            """,
            (prog,),
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "id": int(r["id"] or 0),
                    "codigo_programacao": str(r["codigo_programacao"] or ""),
                    "descricao": str(r["descricao"] or ""),
                    "valor": float(r["valor"] or 0.0),
                    "categoria": str(r["categoria"] or "OUTROS"),
                    "observacao": str(r["observacao"] or ""),
                    "data_registro": str(r["data_registro"] or ""),
                    "id_local": str(r["id_local"] or ""),
                    "forma_pagamento": str(r["forma_pagamento"] or ""),
                    "origem": str(r["origem"] or ""),
                }
            )
    return {"ok": True, "codigo_programacao": prog, "despesas": out}


@app.get("/desktop/rotas/{codigo_programacao}/bundle")
def desktop_rota_bundle(
    codigo_programacao: str,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    detalhe = rota_detalhe_desktop(prog, _ok=True)
    receb = desktop_listar_recebimentos(prog, _ok=True)
    desp = desktop_listar_despesas(prog, _ok=True)
    logistica = desktop_rota_logistica(prog, _ok=True)
    return {
        "ok": True,
        "codigo_programacao": prog,
        "rota": detalhe.get("rota") if isinstance(detalhe, dict) else None,
        "clientes": detalhe.get("clientes") if isinstance(detalhe, dict) else [],
        "recebimentos": receb.get("recebimentos") if isinstance(receb, dict) else [],
        "despesas": desp.get("despesas") if isinstance(desp, dict) else [],
        "logistica": logistica.get("logistica") if isinstance(logistica, dict) else {},
    }


@app.post("/desktop/rotas/{codigo_programacao}/despesas")
def desktop_criar_despesa(
    codigo_programacao: str,
    payload: DesktopDespesaIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    desc = str(payload.descricao or "").strip().upper()
    cat = str(payload.categoria or "OUTROS").strip().upper() or "OUTROS"
    obs = str(payload.observacao or "").strip().upper()
    val = float(payload.valor or 0.0)
    if not desc:
        raise HTTPException(status_code=400, detail="descricao obrigatoria.")
    if val <= 0:
        raise HTTPException(status_code=400, detail="valor deve ser maior que zero.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            """
            INSERT INTO despesas (codigo_programacao, descricao, valor, categoria, observacao, data_registro)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (prog, desc, val, cat, obs, ts),
        )
        did = int(cur.lastrowid or 0)
    return {"ok": True, "id": did, "codigo_programacao": prog}


@app.get("/desktop/despesas/{despesa_id}")
def desktop_obter_despesa(
    despesa_id: int,
    _ok: bool = Depends(_require_desktop_secret),
):
    did = int(despesa_id or 0)
    if did <= 0:
        raise HTTPException(status_code=400, detail="despesa_id invalido.")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                id,
                UPPER(TRIM(COALESCE(codigo_programacao,''))) AS codigo_programacao,
                TRIM(COALESCE(descricao,'')) AS descricao,
                COALESCE(valor,0) AS valor,
                UPPER(TRIM(COALESCE(categoria,'OUTROS'))) AS categoria,
                TRIM(COALESCE(observacao,'')) AS observacao,
                TRIM(COALESCE(data_registro,'')) AS data_registro
            FROM despesas
            WHERE id=?
            LIMIT 1
            """,
            (did,),
        )
        row = cur.fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="despesa nao encontrada.")
    return {
        "ok": True,
        "despesa": {
            "id": int(row["id"] or 0),
            "codigo_programacao": str(row["codigo_programacao"] or ""),
            "descricao": str(row["descricao"] or ""),
            "valor": float(row["valor"] or 0.0),
            "categoria": str(row["categoria"] or "OUTROS"),
            "observacao": str(row["observacao"] or ""),
            "data_registro": str(row["data_registro"] or ""),
        },
    }


@app.put("/desktop/despesas/{despesa_id}")
def desktop_atualizar_despesa(
    despesa_id: int,
    payload: DesktopDespesaIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    did = int(despesa_id or 0)
    if did <= 0:
        raise HTTPException(status_code=400, detail="despesa_id invalido.")

    desc = str(payload.descricao or "").strip().upper()
    cat = str(payload.categoria or "OUTROS").strip().upper() or "OUTROS"
    obs = str(payload.observacao or "").strip().upper()
    val = float(payload.valor or 0.0)
    if not desc:
        raise HTTPException(status_code=400, detail="descricao obrigatoria.")
    if val <= 0:
        raise HTTPException(status_code=400, detail="valor deve ser maior que zero.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, codigo_programacao FROM despesas WHERE id=? LIMIT 1", (did,))
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="despesa nao encontrada.")
        _ensure_programacao_mutable(cur, str(row["codigo_programacao"] or ""))
        cur.execute(
            "UPDATE despesas SET descricao=?, valor=?, categoria=?, observacao=? WHERE id=?",
            (desc, val, cat, obs, did),
        )
        prog = str(row["codigo_programacao"] or "").strip().upper()
    return {"ok": True, "id": did, "codigo_programacao": prog}


@app.delete("/desktop/despesas/{despesa_id}")
def desktop_excluir_despesa(
    despesa_id: int,
    codigo_programacao: str = Query(""),
    _ok: bool = Depends(_require_desktop_secret),
):
    did = int(despesa_id or 0)
    if did <= 0:
        raise HTTPException(status_code=400, detail="despesa_id invalido.")
    prog = (codigo_programacao or "").strip().upper()

    with get_conn() as conn:
        cur = conn.cursor()
        if prog:
            _ensure_programacao_mutable(cur, prog)
            cur.execute(
                "DELETE FROM despesas WHERE id=? AND UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))",
                (did, prog),
            )
        else:
            cur.execute("SELECT codigo_programacao FROM despesas WHERE id=? LIMIT 1", (did,))
            row = cur.fetchone()
            if not row:
                raise HTTPException(status_code=404, detail="despesa nao encontrada.")
            _ensure_programacao_mutable(cur, str(row["codigo_programacao"] or ""))
            cur.execute("DELETE FROM despesas WHERE id=?", (did,))
        deleted = int(cur.rowcount or 0)
    if deleted <= 0:
        raise HTTPException(status_code=404, detail="despesa nao encontrada.")
    return {"ok": True, "id": did, "deleted": deleted}


@app.put("/desktop/rotas/{codigo_programacao}/status")
def desktop_atualizar_status_rota(
    codigo_programacao: str,
    payload: DesktopProgramacaoStatusIn,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, prog)
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        sets: List[str] = []
        vals: List[Any] = []
        if ("status" in cols_prog) and (payload.status is not None):
            sets.append("status=?")
            vals.append(str(payload.status or "").strip().upper())
        if ("prestacao_status" in cols_prog) and (payload.prestacao_status is not None):
            sets.append("prestacao_status=?")
            vals.append(str(payload.prestacao_status or "").strip().upper())
        if ("status_operacional" in cols_prog) and (payload.status_operacional is not None):
            status_op = str(payload.status_operacional or "").strip().upper()
            sets.append("status_operacional=?")
            vals.append(status_op or None)
        if ("finalizada_no_app" in cols_prog) and (payload.finalizada_no_app is not None):
            sets.append("finalizada_no_app=?")
            vals.append(1 if int(payload.finalizada_no_app or 0) == 1 else 0)

        if not sets:
            raise HTTPException(status_code=400, detail="Nenhum campo de status informado para atualizacao.")

        vals.append(prog)
        cur.execute(
            f"UPDATE programacoes SET {', '.join(sets)} WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))",
            tuple(vals),
        )
        updated = int(cur.rowcount or 0)
        if updated <= 0:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")
    return {"ok": True, "codigo_programacao": prog, "updated": updated}


@app.delete("/desktop/rotas/{codigo_programacao}")
def desktop_excluir_programacao(
    codigo_programacao: str,
    delete_vendas: int = Query(0, description="1 para apagar vendas vinculadas; 0 para devolver a Importar Vendas"),
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        state = _programacao_state(cur, prog)
        if not state:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")
        if state["prestacao_status"] == "FECHADA":
            raise HTTPException(status_code=409, detail="prestacao fechada; exclusao da programacao esta bloqueada.")

        st_eff = str(state.get("status_operacional") or state.get("status") or "").strip().upper()
        if st_eff not in {"", "ATIVA", "ABERTA", "PENDENTE", "PROGRAMADA"}:
            raise HTTPException(
                status_code=409,
                detail=f"programacao {prog} esta em estado {st_eff or '-'} e nao pode ser excluida.",
            )

        if int(delete_vendas or 0) == 1:
            cur.execute("DELETE FROM vendas_importadas WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)", (prog,))
        else:
            cur.execute(
                """
                UPDATE vendas_importadas
                   SET usada=0,
                       usada_em='',
                       codigo_programacao='',
                       selecionada=0
                 WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)
                """,
                (prog,),
            )

        for table_name in (
            "programacao_itens_log",
            "programacao_itens_controle",
            "programacao_itens",
            "recebimentos",
            "despesas",
            "rota_gps_pings",
            "rota_substituicoes",
            "cliente_localizacao_amostras",
        ):
            if table_exists(cur, table_name):
                cur.execute(f"DELETE FROM {table_name} WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)", (prog,))

        if table_exists(cur, "transferencias"):
            cur.execute(
                """
                DELETE FROM transferencias
                WHERE UPPER(COALESCE(codigo_origem,''))=UPPER(?)
                   OR UPPER(COALESCE(codigo_destino,''))=UPPER(?)
                """,
                (prog, prog),
            )

        cur.execute("DELETE FROM programacoes WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)", (prog,))
        deleted = int(cur.rowcount or 0)
        if deleted <= 0:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")

    return {"ok": True, "codigo_programacao": prog, "deleted": deleted}


def _collect_desktop_logistica(cur: sqlite3.Cursor, prog: str) -> Dict[str, Any]:
    out: Dict[str, Any] = {
        "pend_substituicao": 0,
        "pend_transferencia": 0,
        "transf_out": 0,
        "transf_in": 0,
        "base_cx": 0,
        "atual_cx": 0,
        "esperado_cx": 0,
        "delta_cx": 0,
        "itens_ok": True,
        "resumo": [],
    }

    if table_exists(cur, "rota_substituicoes"):
        cur.execute(
            """
            SELECT COUNT(1) AS n
            FROM rota_substituicoes
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
              AND UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
            """,
            (prog,),
        )
        row = cur.fetchone()
        out["pend_substituicao"] = int((row["n"] if row else 0) or 0)

    if table_exists(cur, "transferencias"):
        cur.execute(
            """
            SELECT
                SUM(CASE
                        WHEN UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
                         AND (UPPER(TRIM(COALESCE(codigo_origem,'')))=UPPER(TRIM(?))
                           OR UPPER(TRIM(COALESCE(codigo_destino,'')))=UPPER(TRIM(?)))
                        THEN 1 ELSE 0
                    END) AS pend,
                SUM(CASE
                        WHEN UPPER(TRIM(COALESCE(codigo_origem,'')))=UPPER(TRIM(?))
                         AND UPPER(TRIM(COALESCE(status,''))) IN ('PENDENTE','ACEITA','CONVERTIDA')
                        THEN COALESCE(qtd_caixas,0) ELSE 0
                    END) AS out_cx,
                SUM(CASE
                        WHEN UPPER(TRIM(COALESCE(codigo_destino,'')))=UPPER(TRIM(?))
                         AND UPPER(TRIM(COALESCE(status,''))) IN ('PENDENTE','ACEITA','CONVERTIDA')
                        THEN COALESCE(qtd_caixas,0) ELSE 0
                    END) AS in_cx
            FROM transferencias
            """,
            (prog, prog, prog, prog),
        )
        row = cur.fetchone()
        out["pend_transferencia"] = int((row["pend"] if row else 0) or 0)
        out["transf_out"] = int((row["out_cx"] if row else 0) or 0)
        out["transf_in"] = int((row["in_cx"] if row else 0) or 0)

    if table_exists(cur, "programacao_itens"):
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_it = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        has_cx_atual = "caixas_atual" in cols_it
        cur.execute(
            f"""
            SELECT
                COALESCE(SUM(COALESCE(qnt_caixas,0)),0) AS base_cx,
                COALESCE(SUM(COALESCE({"caixas_atual" if has_cx_atual else "qnt_caixas"},0)),0) AS atual_cx
            FROM programacao_itens
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
            """,
            (prog,),
        )
        row = cur.fetchone()
        out["base_cx"] = int((row["base_cx"] if row else 0) or 0)
        atual_calc = int((row["atual_cx"] if row else 0) or 0)

        if (not has_cx_atual) and table_exists(cur, "programacao_itens_controle"):
            try:
                cur.execute("PRAGMA table_info(programacao_itens_controle)")
                cols_ctl = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                if "caixas_atual" in cols_ctl:
                    cur.execute(
                        """
                        SELECT COALESCE(SUM(COALESCE(caixas_atual,0)),0) AS cx
                        FROM programacao_itens_controle
                        WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
                        """,
                        (prog,),
                    )
                    rr = cur.fetchone()
                    atual_calc = int((rr["cx"] if rr else 0) or 0)
            except Exception:
                logging.debug("Falha ao calcular caixas_atual via controle", exc_info=True)

        out["atual_cx"] = max(int(atual_calc), 0)
        out["esperado_cx"] = 0
        out["delta_cx"] = int(out["atual_cx"]) - int(out["esperado_cx"])
        out["itens_ok"] = bool(out["atual_cx"] == 0)

    out["resumo"] = [
        f"Substituicoes pendentes: {out['pend_substituicao']}",
        f"Transferencias pendentes: {out['pend_transferencia']}",
        f"Transferencia caixas (origem): {out['transf_out']} cx",
        f"Transferencia caixas (destino): {out['transf_in']} cx",
        f"Caixas base: {out['base_cx']} cx",
        f"Caixas atuais: {out['atual_cx']} cx",
        f"Caixas esperadas no fechamento: {out['esperado_cx']} cx",
        f"Delta caixas: {out['delta_cx']} cx",
    ]
    return out


def _is_allowed_desktop_sql_mutation(sql: str) -> bool:
    s = str(sql or "").strip()
    if not s:
        return False
    s_up = s.upper()
    if s_up.startswith("INSERT ") or s_up.startswith("UPDATE ") or s_up.startswith("DELETE ") or s_up.startswith("REPLACE "):
        banned = ("PRAGMA ", "ATTACH ", "DETACH ", "VACUUM ", "ALTER ", "DROP ", "CREATE ", "sqlite_master")
        if any(tok in s_up for tok in banned):
            return False
        allowed_tables = {
            "USUARIOS",
            "MOTORISTAS",
            "VEICULOS",
            "AJUDANTES",
            "EQUIPES",
            "CLIENTES",
            "PROGRAMACOES",
            "PROGRAMACAO_ITENS",
            "PROGRAMACAO_ITENS_CONTROLE",
            "PROGRAMACAO_ITENS_LOG",
            "RECEBIMENTOS",
            "DESPESAS",
            "VENDAS_IMPORTADAS",
            "CENTRO_CUSTOS",
            "ROTAS",
            "TRANSFERENCIAS",
            "ROTA_SUBSTITUICOES",
            "ROTA_GPS_PINGS",
            "CLIENTE_LOCALIZACAO_AMOSTRAS",
        }
        match = (
            re.match(r"^(?:INSERT\s+INTO|REPLACE\s+INTO)\s+([A-Z0-9_]+)\b", s_up)
            or re.match(r"^UPDATE\s+([A-Z0-9_]+)\b", s_up)
            or re.match(r"^DELETE\s+FROM\s+([A-Z0-9_]+)\b", s_up)
        )
        table_name = str(match.group(1) if match else "").upper()
        return table_name in allowed_tables
    return False


def _programacao_state(cur, codigo_programacao: str) -> Optional[Dict[str, Any]]:
    prog = (codigo_programacao or "").strip().upper()
    if not prog:
        return None

    cur.execute("PRAGMA table_info(programacoes)")
    cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if not cols_prog:
        return None

    prest_expr = "UPPER(TRIM(COALESCE(prestacao_status,'PENDENTE')))" if "prestacao_status" in cols_prog else "'PENDENTE'"
    status_expr = "UPPER(TRIM(COALESCE(status,'')))" if "status" in cols_prog else "''"
    status_op_expr = "UPPER(TRIM(COALESCE(status_operacional,'')))" if "status_operacional" in cols_prog else "''"
    cur.execute(
        f"""
        SELECT
            codigo_programacao,
            {prest_expr} AS prestacao_status,
            {status_expr} AS status,
            {status_op_expr} AS status_operacional
        FROM programacoes
        WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
        ORDER BY id DESC
        LIMIT 1
        """,
        (prog,),
    )
    row = cur.fetchone()
    if not row:
        return None
    return {
        "codigo_programacao": str(row["codigo_programacao"] or "").strip().upper(),
        "prestacao_status": str(row["prestacao_status"] or "PENDENTE").strip().upper(),
        "status": str(row["status"] or "").strip().upper(),
        "status_operacional": str(row["status_operacional"] or "").strip().upper(),
    }


def _ensure_programacao_mutable(cur, codigo_programacao: str) -> Dict[str, Any]:
    state = _programacao_state(cur, codigo_programacao)
    if not state:
        raise HTTPException(status_code=404, detail="programacao nao encontrada.")
    if state["prestacao_status"] == "FECHADA":
        raise HTTPException(
            status_code=409,
            detail="prestacao fechada; alteracoes de recebimentos, despesas e cabecalho estao bloqueadas.",
        )
    return state


def _ensure_programacao_has_cliente(cur: sqlite3.Cursor, codigo_programacao: str, cod_cliente: str) -> None:
    prog = (codigo_programacao or "").strip().upper()
    cod = (cod_cliente or "").strip().upper()
    if not prog or not cod:
        raise HTTPException(status_code=400, detail="codigo_programacao e cod_cliente obrigatorios.")
    if not table_exists(cur, "programacao_itens"):
        return
    cur.execute(
        """
        SELECT 1
        FROM programacao_itens
        WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
          AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))
        LIMIT 1
        """,
        (prog, cod),
    )
    if not cur.fetchone():
        raise HTTPException(
            status_code=409,
            detail="cliente nao vinculado a programacao; inclua o cliente na rota antes de lancar recebimento.",
        )


def _cadastro_delete_block_reason(
    cur: sqlite3.Cursor,
    table: str,
    *,
    codigo: str = "",
    placa: str = "",
    ajudante_id: int = 0,
    cod_cliente: str = "",
    company_id: Optional[int] = None,
) -> str:
    try:
        if table == "motoristas":
            cod = (codigo or "").strip().upper()
            if not cod or not table_exists(cur, "programacoes"):
                return ""
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            conds: List[str] = []
            params: List[Any] = []
            for col in ("motorista_codigo", "codigo_motorista", "motorista"):
                if col in cols:
                    conds.append(f"UPPER(TRIM(COALESCE({col},'')))=UPPER(TRIM(?))")
                    params.append(cod)
            if not conds:
                return ""
            scope_sql, scope_params = _company_scope_condition(cur, "programacoes", company_id)
            where = f"({' OR '.join(conds)})"
            if scope_sql:
                where += f" AND {scope_sql}"
                params.extend(scope_params)
            cur.execute(f"SELECT COUNT(*) AS qtd FROM programacoes WHERE {where}", tuple(params))
            if int((cur.fetchone() or [0])[0] or 0) > 0:
                return "Motorista vinculado a programacao/rota."
            return ""

        if table == "vendedores":
            cod = (codigo or "").strip().upper()
            if not cod:
                return ""
            if table_exists(cur, "clientes"):
                scope_sql, scope_params = _company_scope_condition(cur, "clientes", company_id)
                cur.execute(
                    f"""
                    SELECT COUNT(*) AS qtd
                    FROM clientes
                    WHERE UPPER(TRIM(COALESCE(vendedor,'')))=UPPER(TRIM(?))
                    {f'AND {scope_sql}' if scope_sql else ''}
                    """,
                    (cod, *scope_params),
                )
                if int((cur.fetchone() or [0])[0] or 0) > 0:
                    return "Vendedor vinculado ao cadastro de clientes."
            if table_exists(cur, "programacoes"):
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                conds: List[str] = []
                params: List[Any] = []
                for col in ("usuario_criacao", "usuario_ultima_edicao"):
                    if col in cols:
                        conds.append(f"UPPER(TRIM(COALESCE({col},'')))=UPPER(TRIM(?))")
                        params.append(cod)
                if conds:
                    scope_sql, scope_params = _company_scope_condition(cur, "programacoes", company_id)
                    where = f"({' OR '.join(conds)})"
                    if scope_sql:
                        where += f" AND {scope_sql}"
                        params.extend(scope_params)
                    cur.execute(
                        f"SELECT COUNT(*) AS qtd FROM programacoes WHERE {where}",
                        tuple(params),
                    )
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Vendedor vinculado a programacao."
            return ""

        if table == "veiculos":
            plc = (placa or "").strip().upper()
            if not plc or not table_exists(cur, "programacoes"):
                return ""
            scope_sql, scope_params = _company_scope_condition(cur, "programacoes", company_id)
            cur.execute(
                f"SELECT COUNT(*) AS qtd FROM programacoes WHERE UPPER(TRIM(COALESCE(veiculo,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
                (plc, *scope_params),
            )
            if int((cur.fetchone() or [0])[0] or 0) > 0:
                return "Veiculo vinculado a programacao/rota."
            return ""

        if table == "clientes":
            cod = (cod_cliente or "").strip().upper()
            if not cod:
                return ""
            if table_exists(cur, "programacao_itens"):
                scope_sql, scope_params = _company_scope_condition(cur, "programacao_itens", company_id)
                cur.execute(
                    f"SELECT COUNT(*) AS qtd FROM programacao_itens WHERE UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
                    (cod, *scope_params),
                )
                if int((cur.fetchone() or [0])[0] or 0) > 0:
                    return "Cliente vinculado a programacao."
            if table_exists(cur, "recebimentos"):
                scope_sql, scope_params = _company_scope_condition(cur, "recebimentos", company_id)
                cur.execute(
                    f"SELECT COUNT(*) AS qtd FROM recebimentos WHERE UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?)){f' AND {scope_sql}' if scope_sql else ''}",
                    (cod, *scope_params),
                )
                if int((cur.fetchone() or [0])[0] or 0) > 0:
                    return "Cliente vinculado a recebimentos."
            return ""

        if table == "ajudantes":
            aid = int(ajudante_id or 0)
            if aid <= 0 or not table_exists(cur, "ajudantes"):
                return ""
            scope_sql, scope_params = _company_scope_condition(cur, "ajudantes", company_id)
            cur.execute(
                f"SELECT COALESCE(nome,'') AS nome, COALESCE(sobrenome,'') AS sobrenome FROM ajudantes WHERE id=?{f' AND {scope_sql}' if scope_sql else ''} LIMIT 1",
                (aid, *scope_params),
            )
            row = cur.fetchone()
            if not row:
                return ""
            nome = str(row["nome"] or "").strip().upper() if row else ""
            sobrenome = str(row["sobrenome"] or "").strip().upper() if row else ""
            alvo = f"{nome} {sobrenome}".strip()
            if table_exists(cur, "equipes"):
                cur.execute("PRAGMA table_info(equipes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                conds = []
                params = []
                for col in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2"):
                    if col in cols:
                        conds.append(f"UPPER(TRIM(COALESCE({col},'')))=UPPER(TRIM(?))")
                        params.append(str(aid))
                        if nome:
                            conds.append(f"UPPER(TRIM(COALESCE({col},'')))=UPPER(TRIM(?))")
                            params.append(nome)
                        if alvo:
                            conds.append(f"UPPER(TRIM(COALESCE({col},'')))=UPPER(TRIM(?))")
                            params.append(alvo)
                if conds:
                    scope_sql, scope_params = _company_scope_condition(cur, "equipes", company_id)
                    where = f"({' OR '.join(conds)})"
                    if scope_sql:
                        where += f" AND {scope_sql}"
                        params.extend(scope_params)
                    cur.execute(f"SELECT COUNT(*) AS qtd FROM equipes WHERE {where}", tuple(params))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Ajudante vinculado a equipe."
            if nome and table_exists(cur, "programacoes"):
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                if "equipe" in cols:
                    expr = "UPPER(TRIM(COALESCE(equipe,'')))"
                    scope_sql, scope_params = _company_scope_condition(cur, "programacoes", company_id)
                    cur.execute(
                        f"""
                        SELECT COUNT(*) AS qtd
                        FROM programacoes
                        WHERE
                            (
                                {expr}=UPPER(TRIM(?))
                                OR {expr}=UPPER(TRIM(?))
                                OR {expr} LIKE UPPER(TRIM(?))
                                OR {expr} LIKE UPPER(TRIM(?))
                                OR {expr} LIKE UPPER(TRIM(?))
                            )
                            {f'AND {scope_sql}' if scope_sql else ''}
                        """,
                        (nome, alvo, f"{nome}|%", f"%|{nome}", f"%|{nome}|%", *scope_params),
                    )
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Ajudante vinculado a programacao."
            return ""
    except Exception:
        logging.debug("Falha ao validar vinculos antes da exclusao de cadastro", exc_info=True)
    return ""


def _normalize_desktop_sql_params(params: Any):
    if params is None:
        return ()
    if isinstance(params, dict):
        out = {}
        for k, v in params.items():
            out[str(k)] = _normalize_desktop_sql_scalar(v)
        return out
    if isinstance(params, (list, tuple)):
        return tuple(_normalize_desktop_sql_scalar(v) for v in params)
    return (_normalize_desktop_sql_scalar(params),)


def _normalize_desktop_sql_scalar(v: Any):
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    if isinstance(v, (bytes, bytearray)):
        try:
            return bytes(v).decode("utf-8", errors="replace")
        except Exception:
            return str(v)
    return str(v)


@app.post("/desktop/sql/mutate")
def desktop_sql_mutate(payload: DesktopSqlMutateIn, _ok: bool = Depends(_require_desktop_secret)):
    stmts = payload.statements or []
    if not stmts:
        return {"ok": True, "executed": 0}
    if len(stmts) > 5000:
        raise HTTPException(status_code=400, detail="Quantidade de statements excede o limite (5000).")

    with get_conn() as conn:
        cur = conn.cursor()
        executed = 0
        for st in stmts:
            sql = str(st.sql or "").strip()
            if not _is_allowed_desktop_sql_mutation(sql):
                raise HTTPException(status_code=400, detail=f"SQL nao permitido para mutate: {sql[:80]}")
            params = _normalize_desktop_sql_params(st.params)
            try:
                cur.execute(sql, params)
            except Exception as exc:
                raise HTTPException(status_code=400, detail=f"Falha ao executar SQL mutate: {exc}")
            executed += 1
    return {"ok": True, "executed": executed}


@app.get("/desktop/overview")
def desktop_overview(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        where_not_finalizadas = _rotas_not_finalizadas_clause(conn, "p")
        try:
            cur.execute("PRAGMA table_info(programacoes)")
            cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        except Exception:
            cols_prog = set()

        total_prog = 0
        total_vendas = 0
        total_clientes_ativos = 0
        prestacao_pendente = 0
        sem_despesa = 0
        rotas: List[Dict[str, Any]] = []

        try:
            cur.execute(
                "SELECT COUNT(*) AS n FROM programacoes p WHERE " + where_not_finalizadas
            )
            row = cur.fetchone()
            total_prog = int((row["n"] if row else 0) or 0)
        except Exception:
            total_prog = 0

        try:
            cur.execute("SELECT COUNT(*) AS n FROM vendas_importadas")
            row = cur.fetchone()
            total_vendas = int((row["n"] if row else 0) or 0)
        except Exception:
            total_vendas = 0

        try:
            cur.execute(
                """
                SELECT COUNT(DISTINCT i.cod_cliente) AS n
                FROM programacao_itens i
                WHERE i.codigo_programacao IN (
                    SELECT p.codigo_programacao
                    FROM programacoes p
                    WHERE """ + where_not_finalizadas + """
                )
                """
            )
            row = cur.fetchone()
            total_clientes_ativos = int((row["n"] if row else 0) or 0)
        except Exception:
            total_clientes_ativos = 0

        try:
            status_col = "status_operacional" if "status_operacional" in cols_prog else "status"
            status_expr = f"UPPER(TRIM(COALESCE({status_col}, '')))"
            where_final = f"{status_expr} IN ('FINALIZADA', 'FINALIZADO')"
            if "prestacao_status" in cols_prog:
                where_prest = "UPPER(TRIM(COALESCE(prestacao_status, 'PENDENTE'))) <> 'FECHADA'"
            else:
                where_prest = "1=1"
            where_base = f"{where_final} AND {where_prest}"
            cur.execute(f"SELECT COUNT(*) AS n FROM programacoes WHERE {where_base}")
            row = cur.fetchone()
            prestacao_pendente = int((row["n"] if row else 0) or 0)
        except Exception:
            prestacao_pendente = 0

        try:
            cur.execute(
                f"""
                SELECT COUNT(*) AS n
                FROM programacoes p
                WHERE {where_base}
                  AND NOT EXISTS (
                      SELECT 1
                      FROM despesas d
                      WHERE UPPER(TRIM(COALESCE(d.codigo_programacao, '')))
                          = UPPER(TRIM(COALESCE(p.codigo_programacao, '')))
                  )
                """
            )
            row = cur.fetchone()
            sem_despesa = int((row["n"] if row else 0) or 0)
        except Exception:
            sem_despesa = 0

        try:
            cur.execute(
                """
                SELECT
                    p.codigo_programacao,
                    COALESCE(p.motorista, '') AS motorista,
                    COALESCE(p.veiculo, '') AS veiculo,
                    COALESCE(p.data_criacao, '') AS data_criacao,
                    COALESCE(p.status, '') AS status,
                    COALESCE(p.status_operacional, '') AS status_operacional
                FROM programacoes p
                WHERE """ + where_not_finalizadas + """
                ORDER BY p.id DESC
                LIMIT 120
                """
            )
            for r in (cur.fetchall() or []):
                rotas.append(
                    {
                        "codigo_programacao": str(r["codigo_programacao"] or "").strip().upper(),
                        "motorista": str(r["motorista"] or "").strip(),
                        "veiculo": str(r["veiculo"] or "").strip(),
                        "data_criacao": str(r["data_criacao"] or "").strip(),
                        "status": str(r["status"] or "").strip().upper(),
                        "status_operacional": str(r["status_operacional"] or "").strip().upper(),
                    }
                )
        except Exception:
            rotas = []

    return {
        "ok": True,
        "total_programacoes_ativas": total_prog,
        "total_vendas_importadas": total_vendas,
        "total_clientes_ativos": total_clientes_ativos,
        "pendencias": {
            "rotas_abertas": total_prog,
            "prestacao_pendente": prestacao_pendente,
            "sem_despesa": sem_despesa,
        },
        "rotas": rotas,
    }


@app.get("/desktop/programacoes")
def desktop_programacoes_listar(
    modo: str = Query("todas", description="todas|ativas|finalizadas_pendentes|finalizadas_prestacao"),
    limit: int = Query(400, ge=1, le=5000),
    _ok: bool = Depends(_require_desktop_secret),
):
    modo_n = (modo or "todas").strip().lower()
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        has_status = "status" in cols
        has_status_op = "status_operacional" in cols
        has_finalizada_app = "finalizada_no_app" in cols
        has_prest = "prestacao_status" in cols
        tipo_estimativa_expr = "COALESCE(p.tipo_estimativa, 'KG')" if "tipo_estimativa" in cols else "'KG'"
        operacao_tipo_expr = (
            "COALESCE(p.operacao_tipo, CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END)"
            if "operacao_tipo" in cols and "tipo_estimativa" in cols
            else ("CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END" if "tipo_estimativa" in cols else "'VENDA'")
        )
        transbordo_modalidade_expr = "COALESCE(p.transbordo_modalidade, '')" if "transbordo_modalidade" in cols else "''"
        transbordo_grupo_expr = "COALESCE(p.transbordo_grupo, '')" if "transbordo_grupo" in cols else "''"

        where = []
        params: List[Any] = []
        where.append("TRIM(COALESCE(p.codigo_programacao,''))<>''")

        if modo_n == "ativas":
            where_not_finalizadas = _rotas_not_finalizadas_clause(conn, "p")
            where.append(where_not_finalizadas)
        elif modo_n == "finalizadas_pendentes":
            if has_status_op:
                st_cond = "UPPER(TRIM(COALESCE(p.status_operacional,''))) IN ('FINALIZADA','FINALIZADO')"
            elif has_status:
                st_cond = "UPPER(TRIM(COALESCE(p.status,''))) IN ('FINALIZADA','FINALIZADO')"
            else:
                st_cond = "1=0"
            if has_finalizada_app:
                st_cond = f"({st_cond} AND COALESCE(p.finalizada_no_app,0)=1)"
            where.append(st_cond)
            if has_prest:
                where.append("COALESCE(p.prestacao_status,'PENDENTE')='PENDENTE'")
        elif modo_n == "finalizadas_prestacao":
            parts = []
            if has_status_op:
                parts.append("UPPER(TRIM(COALESCE(p.status_operacional,''))) IN ('FINALIZADA','FINALIZADO')")
            if has_status:
                parts.append("UPPER(TRIM(COALESCE(p.status,''))) IN ('FINALIZADA','FINALIZADO')")
            where.append("(" + (" OR ".join(parts) if parts else "1=0") + ")")
            if has_prest:
                where.append("UPPER(TRIM(COALESCE(p.prestacao_status,'PENDENTE'))) IN ('PENDENTE','FECHADA')")

        where_sql = ("WHERE " + " AND ".join(where)) if where else ""
        cur.execute(
            f"""
            SELECT
                p.codigo_programacao,
                COALESCE(p.data_criacao, '') AS data_criacao,
                COALESCE(p.motorista, '') AS motorista,
                COALESCE(p.veiculo, '') AS veiculo,
                COALESCE(p.status, '') AS status,
                COALESCE(p.status_operacional, '') AS status_operacional,
                COALESCE(p.prestacao_status, 'PENDENTE') AS prestacao_status,
                {tipo_estimativa_expr} AS tipo_estimativa,
                {operacao_tipo_expr} AS operacao_tipo,
                {transbordo_modalidade_expr} AS transbordo_modalidade,
                {transbordo_grupo_expr} AS transbordo_grupo
            FROM programacoes p
            {where_sql}
            ORDER BY p.id DESC
            LIMIT ?
            """,
            tuple(params + [int(limit)]),
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "codigo_programacao": str(r["codigo_programacao"] or "").strip().upper(),
                    "data_criacao": str(r["data_criacao"] or "").strip(),
                    "data_referencia": str(r["data_criacao"] or "").strip(),
                    "motorista": str(r["motorista"] or "").strip().upper(),
                    "veiculo": str(r["veiculo"] or "").strip().upper(),
                    "status": str(r["status"] or "").strip().upper(),
                    "status_operacional": str(r["status_operacional"] or "").strip().upper(),
                    "prestacao_status": str(r["prestacao_status"] or "").strip().upper(),
                    "tipo_estimativa": str(r["tipo_estimativa"] or "KG").strip().upper(),
                    "operacao_tipo": str(r["operacao_tipo"] or "VENDA").strip().upper(),
                    "transbordo_modalidade": str(r["transbordo_modalidade"] or "").strip().upper(),
                    "transbordo_grupo": str(r["transbordo_grupo"] or "").strip().upper(),
                }
            )
    return {"ok": True, "modo": modo_n, "programacoes": out}


@app.get("/desktop/centro-custos/rows")
def desktop_centro_custos_rows(
    limit: int = Query(5000, ge=1, le=20000),
    _ok: bool = Depends(_require_desktop_secret),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        data_expr = "COALESCE(p.data_saida,p.data_criacao,'')"
        km_expr = "COALESCE(p.km_rodado,0)" if "km_rodado" in cols else "0"
        if "nf_kg_carregado" in cols:
            kg_expr = "COALESCE(p.nf_kg_carregado, p.kg_carregado, 0)"
        elif "kg_carregado" in cols:
            kg_expr = "COALESCE(p.kg_carregado,0)"
        else:
            kg_expr = "0"
        tipo_estimativa_expr = "COALESCE(p.tipo_estimativa, 'KG')" if "tipo_estimativa" in cols else "'KG'"
        operacao_tipo_expr = (
            "COALESCE(p.operacao_tipo, CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END)"
            if "operacao_tipo" in cols and "tipo_estimativa" in cols
            else ("CASE WHEN COALESCE(p.tipo_estimativa, 'KG')='CX' THEN 'TRANSBORDO' ELSE 'VENDA' END" if "tipo_estimativa" in cols else "'VENDA'")
        )
        transbordo_modalidade_expr = "COALESCE(p.transbordo_modalidade, '')" if "transbordo_modalidade" in cols else "''"
        transbordo_grupo_expr = "COALESCE(p.transbordo_grupo, '')" if "transbordo_grupo" in cols else "''"

        cur.execute(
            f"""
            SELECT p.codigo_programacao,
                   UPPER(TRIM(COALESCE(p.veiculo,''))) AS veiculo,
                   {data_expr} AS data_ref,
                   {km_expr} AS km_rodado,
                   {kg_expr} AS kg_carregado,
                   {tipo_estimativa_expr} AS tipo_estimativa,
                   {operacao_tipo_expr} AS operacao_tipo,
                   {transbordo_modalidade_expr} AS transbordo_modalidade,
                   {transbordo_grupo_expr} AS transbordo_grupo,
                   COALESCE(d.total_desp,0) AS total_desp
              FROM programacoes p
         LEFT JOIN (
                SELECT codigo_programacao, COALESCE(SUM(valor),0) AS total_desp
                  FROM despesas
                 GROUP BY codigo_programacao
            ) d ON d.codigo_programacao = p.codigo_programacao
             WHERE TRIM(COALESCE(p.veiculo,'')) <> ''
             ORDER BY p.id DESC
             LIMIT ?
            """,
            (int(limit),),
        )
        rows = []
        for r in (cur.fetchall() or []):
            rows.append(
                {
                    "codigo_programacao": str(r["codigo_programacao"] or "").strip(),
                    "veiculo": str(r["veiculo"] or "").strip().upper(),
                    "data_ref": str(r["data_ref"] or "").strip(),
                    "km_rodado": float(r["km_rodado"] or 0.0),
                    "kg_carregado": float(r["kg_carregado"] or 0.0),
                    "tipo_estimativa": str(r["tipo_estimativa"] or "KG").strip().upper(),
                    "operacao_tipo": str(r["operacao_tipo"] or "VENDA").strip().upper(),
                    "transbordo_modalidade": str(r["transbordo_modalidade"] or "").strip().upper(),
                    "transbordo_grupo": str(r["transbordo_grupo"] or "").strip().upper(),
                    "total_desp": float(r["total_desp"] or 0.0),
                }
            )
    return {"ok": True, "rows": rows}


@app.get("/desktop/relatorios/km-veiculos")
def desktop_relatorio_km_veiculos(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
        media_expr = "COALESCE(media_km_l,0)" if "media_km_l" in cols else "0"
        cur.execute(
            f"""
            SELECT
                UPPER(TRIM(COALESCE(veiculo,''))) AS veiculo,
                COUNT(1) AS viagens,
                COALESCE(SUM({km_expr}),0) AS km_rodado,
                COALESCE(AVG(NULLIF({media_expr},0)),0) AS media_km_l
            FROM programacoes
            GROUP BY UPPER(TRIM(COALESCE(veiculo,'')))
            ORDER BY km_rodado DESC, veiculo ASC
            """
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "veiculo": str(r["veiculo"] or "").strip().upper(),
                    "viagens": int(r["viagens"] or 0),
                    "km_rodado": float(r["km_rodado"] or 0.0),
                    "media_km_l": float(r["media_km_l"] or 0.0),
                }
            )
    return {"ok": True, "rows": out}


@app.get("/desktop/relatorios/despesas-categorias")
def desktop_relatorio_despesas_categorias(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                UPPER(TRIM(COALESCE(categoria,'OUTROS'))) AS categoria,
                COUNT(1) AS qtd,
                COALESCE(SUM(valor),0) AS total
            FROM despesas
            GROUP BY UPPER(TRIM(COALESCE(categoria,'OUTROS')))
            ORDER BY total DESC, categoria ASC
            """
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "categoria": str(r["categoria"] or "OUTROS").strip().upper() or "OUTROS",
                    "qtd": int(r["qtd"] or 0),
                    "total": float(r["total"] or 0.0),
                }
            )
    return {"ok": True, "rows": out}


@app.get("/desktop/relatorios/rotina-motoristas")
def desktop_relatorio_rotina_motoristas(
    motorista_like: str = Query("", description="Filtro parcial por motorista."),
    _ok: bool = Depends(_require_desktop_secret),
):
    filtro_mot = str(motorista_like or "").strip().upper()
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
        if "nf_kg_vendido" in cols and "kg_vendido" in cols:
            kg_expr = "COALESCE(nf_kg_vendido, kg_vendido, 0)"
        elif "nf_kg_vendido" in cols:
            kg_expr = "COALESCE(nf_kg_vendido,0)"
        elif "kg_vendido" in cols:
            kg_expr = "COALESCE(kg_vendido,0)"
        else:
            kg_expr = "0"
        sql = f"""
            SELECT
                COALESCE(codigo_programacao,'') AS codigo_programacao,
                COALESCE(motorista,'') AS motorista,
                COALESCE(equipe,'') AS equipe,
                COALESCE(status,'') AS status,
                {kg_expr} AS kg_vendido,
                {km_expr} AS km_rodado
            FROM programacoes
            WHERE 1=1
        """
        params: List[Any] = []
        if filtro_mot:
            sql += " AND UPPER(COALESCE(motorista,'')) LIKE ?"
            params.append(f"%{filtro_mot}%")
        sql += " ORDER BY id DESC"
        cur.execute(sql, tuple(params))
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "codigo_programacao": str(r["codigo_programacao"] or ""),
                    "motorista": str(r["motorista"] or ""),
                    "equipe": str(r["equipe"] or ""),
                    "status": str(r["status"] or ""),
                    "kg_vendido": float(r["kg_vendido"] or 0.0),
                    "km_rodado": float(r["km_rodado"] or 0.0),
                }
            )
    return {"ok": True, "rows": out}


@app.get("/desktop/relatorios/mortalidade-motorista")
def desktop_relatorio_mortalidade_motorista(
    codigo_like: str = Query("", description="Filtro parcial por codigo da programacao."),
    motorista_like: str = Query("", description="Filtro parcial por motorista."),
    data_like: str = Query("", description="Padroes de data separados por '|'."),
    _ok: bool = Depends(_require_desktop_secret),
):
    filtro_cod = str(codigo_like or "").strip().upper()
    filtro_mot = str(motorista_like or "").strip().upper()
    data_patterns = [p.strip() for p in str(data_like or "").split("|") if str(p or "").strip()]

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
        has_status = "status" in cols_p
        has_data_criacao = "data_criacao" in cols_p
        has_data = "data" in cols_p
        data_expr = (
            "COALESCE(p.data_criacao,'')"
            if has_data_criacao
            else ("COALESCE(p.data,'')" if has_data else "''")
        )
        status_expr = "COALESCE(p.status,'')" if has_status else "''"

        sql = f"""
            SELECT
                COALESCE(p.codigo_programacao,'') as codigo_programacao,
                COALESCE(p.motorista,'') as motorista,
                {data_expr} as data_ref,
                {status_expr} as status_ref,
                COALESCE(SUM(COALESCE(pc.mortalidade_aves, 0)), 0) as mortalidade_total,
                COUNT(CASE WHEN COALESCE(pc.mortalidade_aves,0) > 0 THEN 1 END) as clientes_com_mortalidade
            FROM programacoes p
            LEFT JOIN programacao_itens_controle pc
              ON UPPER(COALESCE(pc.codigo_programacao,'')) = UPPER(COALESCE(p.codigo_programacao,''))
            WHERE 1=1
        """
        params: List[Any] = []

        if filtro_cod:
            sql += " AND UPPER(COALESCE(p.codigo_programacao,'')) LIKE ?"
            params.append(f"%{filtro_cod}%")
        if filtro_mot:
            sql += " AND UPPER(COALESCE(p.motorista,'')) LIKE ?"
            params.append(f"%{filtro_mot}%")
        if data_patterns:
            clauses = []
            for pat in data_patterns:
                clauses.append(f"{data_expr} LIKE ?")
                params.append(f"%{pat}%")
            sql += " AND (" + " OR ".join(clauses) + ")"

        sql += """
            GROUP BY p.codigo_programacao, p.motorista, data_ref, status_ref
            ORDER BY mortalidade_total ASC, p.codigo_programacao DESC
        """
        cur.execute(sql, tuple(params))
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "codigo_programacao": str(r["codigo_programacao"] or ""),
                    "motorista": str(r["motorista"] or ""),
                    "data_ref": str(r["data_ref"] or ""),
                    "status_ref": str(r["status_ref"] or ""),
                    "mortalidade_total": int(r["mortalidade_total"] or 0),
                    "clientes_com_mortalidade": int(r["clientes_com_mortalidade"] or 0),
                }
            )
    return {"ok": True, "rows": out}


@app.get("/desktop/escala/rows")
def desktop_escala_rows(
    limit: int = Query(5000, ge=1, le=20000),
    _ok: bool = Depends(_require_desktop_secret),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        local_expr = (
            "COALESCE(local_rota,'')"
            if "local_rota" in cols
            else ("COALESCE(local,'')" if "local" in cols else "''")
        )
        km_rodado_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
        status_op_expr = "COALESCE(status_operacional,'')" if "status_operacional" in cols else "''"
        finalizada_app_expr = "COALESCE(finalizada_no_app,0)" if "finalizada_no_app" in cols else "0"
        kg_estimado_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols else "0"
        data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols else "''"
        hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols else "''"
        data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols else "''"
        hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols else "''"
        data_ref_expr = (
            "COALESCE(data_saida,data_criacao,data,'')"
            if ("data_saida" in cols or "data_criacao" in cols or "data" in cols)
            else "''"
        )

        cur.execute(
            f"""
            SELECT
                COALESCE(codigo_programacao,'') AS codigo_programacao,
                {data_ref_expr} AS data_ref,
                COALESCE(motorista,'') AS motorista,
                COALESCE(equipe,'') AS equipe,
                COALESCE(status,'') AS status,
                {status_op_expr} AS status_operacional,
                {finalizada_app_expr} AS finalizada_no_app,
                {kg_estimado_expr} AS kg_estimado,
                {data_saida_expr} AS data_saida,
                {hora_saida_expr} AS hora_saida,
                {data_chegada_expr} AS data_chegada,
                {hora_chegada_expr} AS hora_chegada,
                {local_expr} AS local_rota,
                {km_rodado_expr} AS km_rodado
            FROM programacoes
            ORDER BY id DESC
            LIMIT ?
            """,
            (int(limit),),
        )
        out: List[Dict[str, Any]] = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "codigo_programacao": str(r["codigo_programacao"] or "").strip().upper(),
                    "data_ref": str(r["data_ref"] or "").strip(),
                    "motorista": str(r["motorista"] or "").strip(),
                    "equipe": str(r["equipe"] or "").strip(),
                    "status": str(r["status"] or "").strip(),
                    "status_operacional": str(r["status_operacional"] or "").strip(),
                    "finalizada_no_app": int(r["finalizada_no_app"] or 0),
                    "kg_estimado": float(r["kg_estimado"] or 0.0),
                    "data_saida": str(r["data_saida"] or "").strip(),
                    "hora_saida": str(r["hora_saida"] or "").strip(),
                    "data_chegada": str(r["data_chegada"] or "").strip(),
                    "hora_chegada": str(r["hora_chegada"] or "").strip(),
                    "local_rota": str(r["local_rota"] or "").strip(),
                    "km_rodado": float(r["km_rodado"] or 0.0),
                }
            )
    return {"ok": True, "rows": out}


@app.get("/desktop/veiculos/{veiculo}/ultimo-km-final")
def desktop_ultimo_km_final_veiculo(
    veiculo: str,
    exclude_programacao: str = Query("", description="Codigo da programacao a excluir da busca."),
    _ok: bool = Depends(_require_desktop_secret),
):
    v = str(veiculo or "").strip().upper()
    ex = str(exclude_programacao or "").strip().upper()
    if not v:
        raise HTTPException(status_code=400, detail="veiculo obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        ultimo = _ultimo_km_final_veiculo(cur, v, exclude_programacao=ex)
    return {
        "ok": True,
        "veiculo": v,
        "km_final": safe_float(ultimo.get("km_final"), 0.0),
        "km_inicial_sugerido": safe_float(ultimo.get("km_final"), 0.0),
        "codigo_programacao": str(ultimo.get("codigo_programacao") or ""),
    }


@app.get("/veiculos/{veiculo}/ultimo-km-final")
def ultimo_km_final_veiculo_mobile(
    veiculo: str,
    exclude_programacao: str = Query("", description="Codigo da programacao a excluir da busca."),
    m=Depends(get_current_motorista),
):
    v = str(veiculo or "").strip().upper()
    ex = str(exclude_programacao or "").strip().upper()
    if not v:
        raise HTTPException(status_code=400, detail="veiculo obrigatorio.")
    with get_conn() as conn:
        cur = conn.cursor()
        ultimo = _ultimo_km_final_veiculo(cur, v, exclude_programacao=ex)
    return {
        "ok": True,
        "veiculo": v,
        "km_final": safe_float(ultimo.get("km_final"), 0.0),
        "km_inicial_sugerido": safe_float(ultimo.get("km_final"), 0.0),
        "codigo_programacao": str(ultimo.get("codigo_programacao") or ""),
    }


@app.get("/desktop/rotas/{codigo_programacao}/logistica")
def desktop_rota_logistica(
    codigo_programacao: str,
    _ok: bool = Depends(_require_desktop_secret),
):
    prog = str(codigo_programacao or "").strip().upper()
    if not prog:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        out = _collect_desktop_logistica(cur, prog)
    return {"ok": True, "codigo_programacao": prog, "logistica": out}


@app.get("/desktop/monitoramento/rotas")
def desktop_rotas_monitoramento(_ok: bool = Depends(_require_desktop_secret)):
    """
    Retorna monitoramento consolidado para o desktop:
    - rotas ativas
    - último ping GPS (lat/lon/velocidade/precisão/horário)
    """
    with get_conn() as conn:
        cur = conn.cursor()
        where_not_finalizadas = _rotas_not_finalizadas_clause(conn, "p")
        cur.execute(
            """
            SELECT
                p.codigo_programacao,
                COALESCE(p.motorista, '') AS motorista,
                COALESCE(p.veiculo, '') AS veiculo,
                COALESCE(p.status, '') AS status,
                COALESCE(p.status_operacional, '') AS status_operacional,
                g.lat,
                g.lon,
                g.speed,
                g.accuracy,
                g.recorded_at
            FROM programacoes p
            LEFT JOIN (
                SELECT r1.codigo_programacao, r1.lat, r1.lon, r1.speed, r1.accuracy, r1.recorded_at
                FROM rota_gps_pings r1
                INNER JOIN (
                    SELECT codigo_programacao, MAX(id) AS max_id
                    FROM rota_gps_pings
                    GROUP BY codigo_programacao
                ) r2 ON r2.max_id = r1.id
            ) g ON g.codigo_programacao = p.codigo_programacao
            WHERE """ + where_not_finalizadas + """
            ORDER BY p.id DESC
            LIMIT 500
            """
        )
        rows = cur.fetchall() or []

        out = []
        for r in rows:
            codigo = str(r["codigo_programacao"] or "").strip()
            pend_sub = _has_pending_substituicao(cur, codigo) if codigo else False
            status_operacional = _status_operacional_especial(dict(r), pend_substituicao=pend_sub)
            status_base = str(r["status"] or "").strip()
            out.append(
                {
                    "codigo_programacao": codigo,
                    "motorista": str(r["motorista"] or "").strip(),
                    "veiculo": str(r["veiculo"] or "").strip(),
                    "status": status_operacional or status_base,
                    "status_base": status_base,
                    "status_operacional": status_operacional,
                    "lat": r["lat"],
                    "lon": r["lon"],
                    "speed": r["speed"],
                    "accuracy": r["accuracy"],
                    "recorded_at": str(r["recorded_at"] or "").strip(),
                }
            )
        return {"rotas": out}


@app.post("/rotas/{codigo_programacao}/clientes/controle")
def salvar_controle_cliente(
    codigo_programacao: str,
    payload: ClienteControleIn,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    company_id = int(m.get("company_id") or 1)
    is_admin = bool(m.get("is_admin"))
    codigo_programacao = (codigo_programacao or "").strip()

    cod_cliente = (payload.cod_cliente or "").strip()
    if not cod_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente é obrigatório")

    with get_conn() as conn:
        cur = conn.cursor()

        # garante que a rota pertence ao motorista
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")
        status_atual = str(pr["status"] or "").strip().upper()
        if (not is_admin) and status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota encerrada. Alteracoes bloqueadas para status {status_atual}.",
            )
        if (not is_admin) and status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota ainda nao iniciada (status={status_atual or 'N/D'}). Inicie a rota para alterar pedidos.",
            )
        if _has_pending_substituicao(cur, codigo_programacao):
            raise HTTPException(
                status_code=409,
                detail="Rota em transferencia de motorista. Alteracoes bloqueadas ate concluir aceite/recusa.",
            )

        # normaliza
        mort = int(payload.mortalidade_aves or 0)
        media_aplicada = payload.media_aplicada
        peso_previsto = payload.peso_previsto
        valor_recebido = payload.valor_recebido
        forma_recebimento = (payload.forma_recebimento or None)
        obs_recebimento = (payload.obs_recebimento or None)

        status_in = (payload.status_pedido or "").strip().upper() or None
        pedido = (payload.pedido or "").strip() or None
        if not pedido:
            raise HTTPException(status_code=400, detail="pedido é obrigatório para controle do cliente.")
        caixas_atual = payload.caixas_atual
        preco_atual = payload.preco_atual
        alterado_por = (payload.alterado_por or nome_motorista or None)
        alteracao_tipo = (payload.alteracao_tipo or None)
        alteracao_detalhe = (payload.alteracao_detalhe or None)
        lat_evento = payload.lat_evento if payload.lat_evento is not None else payload.lat_entrega
        lon_evento = payload.lon_evento if payload.lon_evento is not None else payload.lon_entrega
        endereco_evento = (payload.endereco_evento or None)
        cidade_evento = (payload.cidade_evento or None)
        bairro_evento = (payload.bairro_evento or None)
        ordem_sugerida = payload.ordem_sugerida
        eta = (payload.eta or None)
        distancia = payload.distancia
        confianca_localizacao = payload.confianca_localizacao

        # busca item base (para status/valores), priorizando o pedido informado
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_prog_itens = {row[1] for row in cur.fetchall()}
        has_pedido_col = "pedido" in cols_prog_itens

        item_base = None
        if has_pedido_col:
            cur.execute(
                """
                SELECT cod_cliente, nome_cliente, qnt_caixas, preco, pedido
                FROM programacao_itens
                WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente, pedido),
            )
            item_base = cur.fetchone()
        else:
            cur.execute(
                """
                SELECT cod_cliente, nome_cliente, qnt_caixas, preco, pedido
                FROM programacao_itens
                WHERE codigo_programacao=? AND cod_cliente=?
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente),
            )
            item_base = cur.fetchone()
        if not item_base:
            raise HTTPException(status_code=404, detail="Item de cliente/pedido não encontrado na programação.")

        base_caixas = item_base["qnt_caixas"] if item_base else None
        base_preco = item_base["preco"] if item_base else None
        nome_cliente = item_base["nome_cliente"] if item_base else ""
        if pedido is None and item_base:
            pedido = item_base["pedido"]

        allowed_status = {"PENDENTE", "ENTREGUE", "CANCELADO", "ALTERADO"}
        if status_in and status_in not in allowed_status:
            raise HTTPException(status_code=400, detail=f"status_pedido inválido: {status_in}.")

        # regra: pedido ENTREGUE não pode mais ser alterado
        status_atual = None
        if has_pedido_col and pedido:
            cur.execute(
                """
                SELECT
                    COALESCE(
                        NULLIF(TRIM(pi.status_pedido), ''),
                        NULLIF(TRIM(pc.status_pedido), ''),
                        ''
                    ) AS status_atual
                FROM programacao_itens pi
                LEFT JOIN programacao_itens_controle pc
                  ON pc.codigo_programacao = pi.codigo_programacao
                 AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))
                 AND COALESCE(TRIM(pc.pedido), '') = COALESCE(TRIM(pi.pedido), '')
                WHERE pi.codigo_programacao=?
                  AND UPPER(TRIM(pi.cod_cliente))=UPPER(TRIM(?))
                  AND COALESCE(TRIM(pi.pedido), '')=COALESCE(TRIM(?), '')
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente, pedido),
            )
            row_status = cur.fetchone()
            if row_status:
                status_atual = (row_status["status_atual"] or "").strip().upper()

        if not status_atual and not has_pedido_col:
            cur.execute(
                """
                SELECT
                    COALESCE(
                        NULLIF(TRIM(pi.status_pedido), ''),
                        NULLIF(TRIM(pc.status_pedido), ''),
                        ''
                    ) AS status_atual
                FROM programacao_itens pi
                LEFT JOIN programacao_itens_controle pc
                  ON pc.codigo_programacao = pi.codigo_programacao
                 AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))
                 AND COALESCE(TRIM(pc.pedido), '') = COALESCE(TRIM(pi.pedido), '')
                WHERE pi.codigo_programacao=?
                  AND UPPER(TRIM(pi.cod_cliente))=UPPER(TRIM(?))
                LIMIT 1
                """,
                (codigo_programacao, cod_cliente),
            )
            row_status = cur.fetchone()
            if row_status:
                status_atual = (row_status["status_atual"] or "").strip().upper()

        if (not is_admin) and status_atual == "ENTREGUE":
            raise HTTPException(
                status_code=409,
                detail="Pedido jÃ¡ estÃ¡ ENTREGUE e estÃ¡ bloqueado para alteraÃ§Ãµes.",
            )

        # resolve status se nao veio do app
        status = status_in
        if not status:
            alterado = False
            if caixas_atual is not None and base_caixas is not None:
                try:
                    alterado = int(caixas_atual) != int(base_caixas)
                except Exception:
                    alterado = True
            if not alterado and preco_atual is not None and base_preco is not None:
                try:
                    alterado = float(preco_atual) != float(base_preco)
                except Exception:
                    alterado = True

            if alterado:
                status = "ALTERADO"
            elif valor_recebido is not None and float(valor_recebido) > 0:
                status = "ENTREGUE"
            else:
                status = "PENDENTE"

        evento_em = _evento_iso_or_now(payload.evento_em)
        alterado_em = evento_em if status in {"ALTERADO", "ENTREGUE", "CANCELADO"} else None

        # valida faixa de caixas para evitar manipulação indevida
        if caixas_atual is not None:
            try:
                caixas_atual = int(caixas_atual)
            except Exception:
                raise HTTPException(status_code=400, detail="caixas_atual inválido.")
            if caixas_atual < 0:
                raise HTTPException(status_code=400, detail="caixas_atual não pode ser negativo.")
            if base_caixas is not None:
                try:
                    base_caixas_int = int(base_caixas)
                except Exception:
                    base_caixas_int = None
                if base_caixas_int is not None and caixas_atual > base_caixas_int:
                    raise HTTPException(
                        status_code=400,
                        detail=f"caixas_atual ({caixas_atual}) não pode ser maior que caixas do pedido ({base_caixas_int}).",
                    )

        # atualiza controle por cliente (compatÃvel com bases sem UNIQUE)
        if ordem_sugerida is not None:
            try:
                ordem_sugerida = int(ordem_sugerida)
            except Exception:
                raise HTTPException(status_code=400, detail="ordem_sugerida invalida.")
            if ordem_sugerida < 0:
                raise HTTPException(status_code=400, detail="ordem_sugerida nao pode ser negativa.")

        if distancia is not None:
            try:
                distancia = float(distancia)
            except Exception:
                raise HTTPException(status_code=400, detail="distancia invalida.")
            if distancia < 0:
                raise HTTPException(status_code=400, detail="distancia nao pode ser negativa.")

        if confianca_localizacao is not None:
            try:
                confianca_localizacao = float(confianca_localizacao)
            except Exception:
                raise HTTPException(status_code=400, detail="confianca_localizacao invalida.")
            if confianca_localizacao < 0:
                raise HTTPException(status_code=400, detail="confianca_localizacao nao pode ser negativa.")

        caixas_eff = caixas_atual if caixas_atual is not None else base_caixas
        if status in ("ENTREGUE", "CANCELADO"):
            caixas_eff = 0
        preco_eff = preco_atual if preco_atual is not None else base_preco

        cur.execute(
            """
            UPDATE programacao_itens_controle
               SET mortalidade_aves=?,
                   media_aplicada=?,
                   peso_previsto=?,
                   valor_recebido=?,
                   forma_recebimento=?,
                   obs_recebimento=?,
                   status_pedido=?,
                   alteracao_tipo=?,
                   alteracao_detalhe=?,
                   pedido=?,
                   caixas_atual=?,
                   preco_atual=?,
                   alterado_em=COALESCE(?, alterado_em),
                   alterado_por=?,
                   lat_evento=COALESCE(?, lat_evento),
                   lon_evento=COALESCE(?, lon_evento),
                   endereco_evento=COALESCE(NULLIF(?, ''), endereco_evento),
                   cidade_evento=COALESCE(NULLIF(?, ''), cidade_evento),
                   bairro_evento=COALESCE(NULLIF(?, ''), bairro_evento),
                   ordem_sugerida=COALESCE(?, ordem_sugerida),
                   eta=COALESCE(NULLIF(?, ''), eta),
                   distancia=COALESCE(?, distancia),
                   confianca_localizacao=COALESCE(?, confianca_localizacao),
                   updated_at=datetime('now')
             WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
            """,
            (
                mort,
                media_aplicada,
                peso_previsto,
                valor_recebido,
                forma_recebimento,
                obs_recebimento,
                status,
                alteracao_tipo,
                alteracao_detalhe,
                pedido,
                caixas_eff,
                preco_eff,
                alterado_em,
                alterado_por,
                lat_evento,
                lon_evento,
                endereco_evento,
                cidade_evento,
                bairro_evento,
                ordem_sugerida,
                eta,
                distancia,
                confianca_localizacao,
                codigo_programacao,
                cod_cliente,
                pedido,
            ),
        )

        if cur.rowcount == 0:
            cur.execute(
                """
                INSERT INTO programacao_itens_controle
                    (codigo_programacao, cod_cliente, mortalidade_aves, media_aplicada, peso_previsto,
                     valor_recebido, forma_recebimento, obs_recebimento,
                     status_pedido, alteracao_tipo, alteracao_detalhe, pedido,
                     caixas_atual, preco_atual, alterado_em, alterado_por,
                     lat_evento, lon_evento, endereco_evento, cidade_evento, bairro_evento,
                     ordem_sugerida, eta, distancia, confianca_localizacao, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                """,
                (
                    codigo_programacao,
                    cod_cliente,
                    mort,
                    media_aplicada,
                    peso_previsto,
                    valor_recebido,
                    forma_recebimento,
                    obs_recebimento,
                    status,
                    alteracao_tipo,
                    alteracao_detalhe,
                    pedido,
                    caixas_eff,
                    preco_eff,
                    alterado_em,
                    alterado_por,
                    lat_evento,
                    lon_evento,
                    endereco_evento,
                    cidade_evento,
                    bairro_evento,
                    ordem_sugerida,
                    eta,
                    distancia,
                    confianca_localizacao,
                ),
            )
        try:
            cur.execute("PRAGMA table_info(programacao_itens_controle)")
            cols_ctrl_company = {row[1] for row in cur.fetchall() or []}
            if "company_id" in cols_ctrl_company:
                cur.execute(
                    """
                    UPDATE programacao_itens_controle
                       SET company_id=COALESCE(company_id, ?)
                     WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
                    """,
                    (company_id, codigo_programacao, cod_cliente, pedido),
                )
        except Exception:
            pass

        foto_mortalidade_ref = _store_mobile_photo(
            cur,
            codigo_programacao,
            payload.foto_mortalidade or payload.foto_registro,
            "MORTALIDADE_CLIENTE",
            "MORTALIDADE_CLIENTE",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            cod_cliente=cod_cliente,
            cliente_nome=str(nome_cliente or ""),
            pedido=str(pedido or ""),
            id_vinculo=str(pedido or cod_cliente or ""),
            path_hint=payload.foto_mortalidade_path or payload.mortalidade_foto_path or "",
            company_id=company_id,
        )
        try:
            cur.execute("PRAGMA table_info(programacao_itens_controle)")
            cols_ctrl_extra = {row[1] for row in cur.fetchall() or []}
            extra_sets = []
            extra_params: List[Any] = []

            def add_ctrl_extra(col: str, value: Any):
                if col in cols_ctrl_extra and value not in (None, ""):
                    extra_sets.append(f"{col}=?")
                    extra_params.append(value)

            add_ctrl_extra("lat_entrega", payload.lat_entrega if payload.lat_entrega is not None else lat_evento)
            add_ctrl_extra("lon_entrega", payload.lon_entrega if payload.lon_entrega is not None else lon_evento)
            add_ctrl_extra("accuracy_entrega", payload.accuracy_entrega)
            add_ctrl_extra("timestamp_entrega", payload.timestamp_entrega or (evento_em if status == "ENTREGUE" else None))
            add_ctrl_extra("foto_mortalidade_path", payload.foto_mortalidade_path)
            add_ctrl_extra("mortalidade_foto_path", payload.mortalidade_foto_path)
            if foto_mortalidade_ref is not None:
                add_ctrl_extra("foto_mortalidade_ref_json", json.dumps(foto_mortalidade_ref, ensure_ascii=False))
            if extra_sets:
                extra_params.extend([codigo_programacao, cod_cliente, pedido])
                cur.execute(
                    f"""
                    UPDATE programacao_itens_controle
                       SET {', '.join(extra_sets)}
                     WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
                    """,
                    tuple(extra_params),
                )
        except Exception:
            pass

        # atualiza tabela base de itens (se colunas existirem)
        cols = cols_prog_itens

        sets = []
        params = []
        if "status_pedido" in cols:
            sets.append("status_pedido=?"); params.append(status)
        if "alteracao_tipo" in cols:
            params.append(alteracao_tipo)
            sets.append("alteracao_tipo=?")
        if "alteracao_detalhe" in cols:
            params.append(alteracao_detalhe)
            sets.append("alteracao_detalhe=?")
        if "caixas_atual" in cols:
            params.append(caixas_eff)
            sets.append("caixas_atual=?")
        if "preco_atual" in cols:
            params.append(preco_atual if preco_atual is not None else base_preco)
            sets.append("preco_atual=?")
        if "alterado_em" in cols:
            params.append(alterado_em)
            sets.append("alterado_em=?")
        if "alterado_por" in cols:
            params.append(alterado_por)
            sets.append("alterado_por=?")
        if "ordem_sugerida" in cols and ordem_sugerida is not None:
            params.append(ordem_sugerida)
            sets.append("ordem_sugerida=?")
        if "eta" in cols and eta not in (None, ""):
            params.append(eta)
            sets.append("eta=?")
        if "distancia" in cols and distancia is not None:
            params.append(distancia)
            sets.append("distancia=?")
        if "confianca_localizacao" in cols and confianca_localizacao is not None:
            params.append(confianca_localizacao)
            sets.append("confianca_localizacao=?")

        if sets:
            if has_pedido_col:
                params.extend([codigo_programacao, cod_cliente, pedido])
                cur.execute(
                    f"""
                    UPDATE programacao_itens
                    SET {', '.join(sets)}
                    WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
                    """,
                    tuple(params),
                )
            else:
                params.extend([codigo_programacao, cod_cliente])
                cur.execute(
                    f"UPDATE programacao_itens SET {', '.join(sets)} WHERE codigo_programacao=? AND cod_cliente=?",
                    tuple(params),
                )

        # sincroniza recebimentos (se tabela existir)
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
        if cur.fetchone() is not None:
            cur.execute("PRAGMA table_info(recebimentos)")
            cols_receb = {row[1] for row in cur.fetchall() or []}
            has_receb_pedido = "pedido" in cols_receb
            if valor_recebido is not None and float(valor_recebido) > 0:
                if has_receb_pedido:
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido,'')=COALESCE(?, '')",
                        (codigo_programacao, cod_cliente, pedido),
                    )
                    receb_cols = [
                        "codigo_programacao", "cod_cliente", "pedido", "nome_cliente", "valor",
                        "forma_pagamento", "observacao", "data_registro",
                    ]
                    receb_vals = [
                        codigo_programacao, cod_cliente, pedido, nome_cliente, float(valor_recebido),
                        (forma_recebimento or "DINHEIRO"), (obs_recebimento or None), _now_iso(),
                    ]
                    if "company_id" in cols_receb:
                        receb_cols.append("company_id")
                        receb_vals.append(company_id)
                    placeholders = ", ".join(["?"] * len(receb_cols))
                    cur.execute(
                        f"INSERT INTO recebimentos ({', '.join(receb_cols)}) VALUES ({placeholders})",
                        tuple(receb_vals),
                    )
                else:
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=?",
                        (codigo_programacao, cod_cliente),
                    )
                    receb_cols = [
                        "codigo_programacao", "cod_cliente", "nome_cliente", "valor",
                        "forma_pagamento", "observacao", "data_registro",
                    ]
                    receb_vals = [
                        codigo_programacao, cod_cliente, nome_cliente, float(valor_recebido),
                        (forma_recebimento or "DINHEIRO"), (obs_recebimento or None), _now_iso(),
                    ]
                    if "company_id" in cols_receb:
                        receb_cols.append("company_id")
                        receb_vals.append(company_id)
                    placeholders = ", ".join(["?"] * len(receb_cols))
                    cur.execute(
                        f"INSERT INTO recebimentos ({', '.join(receb_cols)}) VALUES ({placeholders})",
                        tuple(receb_vals),
                    )
            else:
                if has_receb_pedido:
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido,'')=COALESCE(?, '')",
                        (codigo_programacao, cod_cliente, pedido),
                    )
                else:
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=?",
                        (codigo_programacao, cod_cliente),
                    )

        # log de sincronizacao
        try:
            log_payload = payload.model_dump() if hasattr(payload, "model_dump") else payload.dict()
            log_payload.update(
                {
                    "motorista": nome_motorista,
                    "status_pedido": status,
                    "pedido": pedido,
                    "caixas_atual": caixas_atual if caixas_atual is not None else base_caixas,
                    "preco_atual": preco_atual if preco_atual is not None else base_preco,
                    "alterado_por": alterado_por,
                    "alterado_em": alterado_em,
                    "foto_mortalidade_ref": foto_mortalidade_ref,
                }
            )
            payload_json = json.dumps(log_payload, ensure_ascii=False)
            cur.execute(
                """
                INSERT INTO programacao_itens_log
                    (codigo_programacao, cod_cliente, evento, payload_json, company_id)
                VALUES (?, ?, ?, ?, ?)
                """,
                (codigo_programacao, cod_cliente, "cliente_controle", payload_json, company_id),
            )
        except Exception:
            pass

        try:
            _registrar_amostra_localizacao_cliente(
                cur,
                cod_cliente=cod_cliente,
                codigo_programacao=codigo_programacao,
                pedido=pedido,
                lat_evento=lat_evento,
                lon_evento=lon_evento,
                endereco_evento=endereco_evento,
                cidade_evento=cidade_evento,
                bairro_evento=bairro_evento,
                status_pedido=status,
                motorista_codigo=codigo_motorista,
                motorista_nome=nome_motorista,
                origem="APP",
                company_id=company_id,
            )
        except Exception:
            pass

        _registrar_roteiro_operacional(
            cur,
            tipo_evento=status or "CLIENTE_CONTROLE",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino="CLIENTE",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            pedido=pedido or "",
            cod_cliente=cod_cliente,
            cliente_nome=str(nome_cliente or ""),
            caixas=caixas_eff or 0,
            kg=peso_previsto or 0,
            media=media_aplicada or 0,
            aves_por_caixa=0,
            data_hora=evento_em or payload.timestamp_entrega or _now_iso(),
            observacao=alteracao_detalhe or obs_recebimento or "",
            payload={
                **_payload_dict(payload),
                "status_pedido": status,
                "foto_mortalidade_ref": foto_mortalidade_ref,
                "valor_recebido": valor_recebido,
            },
        )

        conn.commit()

    return {"ok": True}


@app.get("/rotas/{codigo_programacao}/clientes/{cod_cliente}/logs")
def listar_logs_cliente(
    codigo_programacao: str,
    cod_cliente: str,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    cod_cliente = (cod_cliente or "").strip()

    if not codigo_programacao or not cod_cliente:
        raise HTTPException(status_code=400, detail="Código e cliente são obrigatórios")

    with get_conn() as conn:
        cur = conn.cursor()

        # garante que a rota pertence ao motorista
        pr = _fetch_programacao_owned(
            cur,
            codigo_programacao,
            m,
            "p.id, p.status, p.carregamento_fechado, COALESCE(p.tipo_estimativa, 'KG') AS tipo_estimativa, COALESCE(p.caixas_estimado, 0) AS caixas_estimado",
        )
        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")

        cur.execute(
            """
            SELECT evento, payload_json, created_at
            FROM programacao_itens_log
            WHERE codigo_programacao=? AND cod_cliente=?
            ORDER BY id DESC
            LIMIT 200
            """,
            (codigo_programacao, cod_cliente),
        )
        rows = cur.fetchall()

    out = []
    for r in rows:
        payload_raw = r["payload_json"] or ""
        try:
            payload = json.loads(payload_raw) if payload_raw else {}
        except Exception:
            payload = {"raw": payload_raw}
        out.append(
            {
                "evento": r["evento"],
                "payload": payload,
                "created_at": r["created_at"],
            }
        )
    return out


@app.post("/rotas/{codigo_programacao}/gps")
def salvar_gps(
    codigo_programacao: str,
    payload: RotaGpsPingIn,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    company_id = int(m.get("company_id") or 1)
    codigo_programacao = (codigo_programacao or "").strip()

    if not codigo_programacao:
        raise HTTPException(status_code=400, detail="Codigo de programacao invalido.")

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "gps", payload.idempotency_key):
            return {"ok": True, "deduplicado": True}

        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.carregamento_fechado")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")

        ts = None
        if payload.timestamp:
            try:
                ts = datetime.fromisoformat(payload.timestamp)
            except Exception:
                ts = None
        if ts is None:
            ts = datetime.now()

        cur.execute("PRAGMA table_info(rota_gps_pings)")
        cols_gps = {row[1] for row in cur.fetchall() or []}
        gps_cols = ["codigo_programacao", "motorista", "lat", "lon", "speed", "accuracy", "recorded_at"]
        gps_vals = [
            codigo_programacao,
            nome_motorista,
            float(payload.lat),
            float(payload.lon),
            (float(payload.speed) if payload.speed is not None else None),
            (float(payload.accuracy) if payload.accuracy is not None else None),
            ts.isoformat(timespec="seconds"),
        ]
        if "company_id" in cols_gps:
            gps_cols.append("company_id")
            gps_vals.append(company_id)
        placeholders = ", ".join(["?"] * len(gps_cols))
        cur.execute(
            f"INSERT INTO rota_gps_pings ({', '.join(gps_cols)}) VALUES ({placeholders})",
            tuple(gps_vals),
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "gps", payload.idempotency_key)
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "clientes_controle", payload.idempotency_key)
        conn.commit()

    return {"ok": True}


@app.post("/rotas/{codigo_programacao}/status-operacional")
def atualizar_status_operacional_rota(
    codigo_programacao: str,
    payload: RotaStatusOperacionalIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    nome_motorista = str(m.get("nome") or "").strip().upper()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    status_in = _status_operacional_normalizado(payload.status_operacional)
    observacao = str(payload.observacao or "").strip()

    allowed = {
        "NORMAL",
        "CAMINHAO_QUEBROU",
        "PROBLEMA_NA_ROTA",
        "PNEU_FURADO",
        "ACIDENTE",
        "ATRASO",
    }
    if status_in not in allowed:
        raise HTTPException(status_code=400, detail=f"status_operacional invalido: {status_in or '-'}")

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "status_operacional", payload.idempotency_key):
            return {"ok": True, "deduplicado": True, "codigo_programacao": codigo_programacao}
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")

        status_rota = str(pr["status"] or "").strip().upper()
        if status_rota in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_rota}).")

        now_iso = _evento_iso_or_now(payload.evento_em)
        if status_in == "NORMAL":
            cur.execute(
                """
                UPDATE programacoes
                   SET status_operacional=NULL,
                       status_operacional_obs=NULL,
                       status_operacional_em=?,
                       status_operacional_por=?
                 WHERE id=?
                """,
                (now_iso, nome_motorista, pr["id"]),
            )
            status_out = None
        else:
            cur.execute(
                """
                UPDATE programacoes
                   SET status_operacional=?,
                       status_operacional_obs=?,
                       status_operacional_em=?,
                       status_operacional_por=?
                 WHERE id=?
                """,
                (status_in, (observacao or None), now_iso, nome_motorista, pr["id"]),
            )
            status_out = status_in
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="STATUS_OPERACIONAL",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino=status_in,
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            data_hora=now_iso,
            observacao=observacao,
            payload=_payload_dict(payload),
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "status_operacional", payload.idempotency_key)
        conn.commit()

    return {
        "ok": True,
        "codigo_programacao": codigo_programacao,
        "status_rota": status_rota,
        "status_operacional": status_out,
    }


@app.post("/rotas/{codigo_programacao}/reabrir")
def reabrir_rota_app(
    codigo_programacao: str,
    payload: RotaReabrirIn,
    m=Depends(get_current_motorista),
):
    if not bool(m.get("is_admin")):
        raise HTTPException(status_code=403, detail="Somente admin pode reabrir rota.")

    codigo_programacao = (codigo_programacao or "").strip().upper()
    if not codigo_programacao:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")

    codigo_actor = str(m.get("codigo") or "").strip().upper()
    nome_actor = str(m.get("nome") or "").strip().upper() or codigo_actor or "ADMIN"
    evento_em = _evento_iso_or_now(payload.evento_em)

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_actor, codigo_programacao, "reabrir_rota", payload.idempotency_key):
            return {"ok": True, "deduplicado": True, "codigo_programacao": codigo_programacao}

        _ensure_programacao_mutable(cur, codigo_programacao)
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        sets: List[str] = []
        vals: List[Any] = []
        if "status" in cols_prog:
            sets.append("status=?")
            vals.append("ATIVA")
        if "status_operacional" in cols_prog:
            sets.append("status_operacional=?")
            vals.append(None)
        if "status_operacional_obs" in cols_prog:
            sets.append("status_operacional_obs=?")
            vals.append(None)
        if "status_operacional_em" in cols_prog:
            sets.append("status_operacional_em=?")
            vals.append(evento_em)
        if "status_operacional_por" in cols_prog:
            sets.append("status_operacional_por=?")
            vals.append(nome_actor)
        if "finalizada_no_app" in cols_prog:
            sets.append("finalizada_no_app=?")
            vals.append(0)

        if not sets:
            raise HTTPException(status_code=500, detail="Nenhum campo elegivel para reabrir a rota.")

        vals.append(codigo_programacao)
        cur.execute(
            f"UPDATE programacoes SET {', '.join(sets)} WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))",
            tuple(vals),
        )
        if int(cur.rowcount or 0) <= 0:
            raise HTTPException(status_code=404, detail="programacao nao encontrada.")

        _registrar_roteiro_operacional(
            cur,
            tipo_evento="REABRIR_ROTA",
            codigo_programacao=codigo_programacao,
            origem="APP_ADMIN",
            destino="ATIVA",
            motorista_codigo=codigo_actor,
            motorista_nome=nome_actor,
            data_hora=evento_em,
            observacao=payload.observacao or "",
            payload=_payload_dict(payload),
        )
        _idempotency_mark(cur, codigo_actor, codigo_programacao, "reabrir_rota", payload.idempotency_key)
        conn.commit()

    return {
        "ok": True,
        "codigo_programacao": codigo_programacao,
        "status": "ATIVA",
        "status_operacional": None,
    }


@app.post("/rotas/{codigo_programacao}/iniciar")
def iniciar_rota(
    codigo_programacao: str,
    payload: IniciarRotaIn,
    m=Depends(get_current_motorista),
    override_token_hdr: Optional[str] = Header(default=None, alias="X-Override-Token"),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    codigo_programacao = (codigo_programacao or "").strip()

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "iniciar_rota", payload.idempotency_key):
            return {"ok": True, "status": "EM_ROTA", "deduplicado": True}
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")
        if int(payload.km_inicial or 0) <= 0:
            raise HTTPException(status_code=400, detail="KM inicial deve ser maior que 0.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(status_code=409, detail="Rota já está em andamento.")
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_atual}).")

        # Bloqueia iniciar uma nova rota se já houver outra em andamento para o mesmo motorista.
        owner_sql, owner_params = _owner_filter_for_programacoes(conn, m, "p")
        cur.execute(
            """
            SELECT p.codigo_programacao, COALESCE(p.status,'') AS status
            FROM programacoes p
            WHERE p.codigo_programacao <> ?
              AND """ + owner_sql + """
              AND UPPER(TRIM(COALESCE(p.status,''))) IN ('EM_ROTA', 'EM ROTA', 'INICIADA', 'EM_ENTREGAS', 'EM ENTREGAS', 'CARREGADA')
            ORDER BY p.id DESC
            LIMIT 1
            """,
            (codigo_programacao, *owner_params),
        )
        rota_em_andamento = cur.fetchone()
        if rota_em_andamento:
            cod_busy = str(rota_em_andamento["codigo_programacao"] or "").strip()
            st_busy = str(rota_em_andamento["status"] or "").strip().upper()
            raise HTTPException(
                status_code=409,
                detail=f"Motorista já possui rota em andamento ({cod_busy}, status={st_busy}). Finalize a anterior antes de iniciar nova rota.",
            )

        # valida GPS: minimo 5 km nos ultimos 15 minutos (opcional por flag)
        override_token = os.environ.get("ROTA_OVERRIDE_TOKEN")
        can_override = bool(override_token) and payload.override_reason and override_token_hdr == override_token

        if ENABLE_START_GPS_GATE and not can_override:
            distancia_m = _gps_distance_last_minutes(cur, codigo_programacao, minutes=15)
            if distancia_m < 5000:
                raise HTTPException(
                    status_code=409,
                    detail="GPS insuficiente: mova pelo menos 5 km nos ultimos 15 minutos.",
                )
        elif can_override:
            # registra override manual (somente se ROTA_OVERRIDE_TOKEN estiver configurado)
            cur.execute(
                """
                INSERT INTO rota_gps_override_log
                    (codigo_programacao, motorista, motivo, created_at)
                VALUES (?, ?, ?, datetime('now'))
                """,
                (codigo_programacao, nome_motorista, payload.override_reason),
            )

        cur.execute(
            """
            UPDATE programacoes
               SET status='EM_ROTA',
                   data_saida=?,
                   hora_saida=?,
                   km_inicial=?
             WHERE id=?
            """,
            (payload.data_saida, payload.hora_saida, payload.km_inicial, pr["id"]),
        )
        cur.execute("SELECT COALESCE(veiculo, '') AS veiculo FROM programacoes WHERE id=? LIMIT 1", (pr["id"],))
        veiculo_row = cur.fetchone()
        veiculo_rota = str((veiculo_row["veiculo"] if veiculo_row else "") or "").strip().upper()
        ultimo = _ultimo_km_final_veiculo(cur, veiculo_rota, exclude_programacao=codigo_programacao)
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="INICIAR_ROTA",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino=veiculo_rota,
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            data_hora=f"{payload.data_saida} {payload.hora_saida}".strip(),
            observacao=payload.override_reason or "",
            payload={**_payload_dict(payload), "km_inicial": payload.km_inicial, "ultimo_km_veiculo": ultimo},
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "iniciar_rota", payload.idempotency_key)
        conn.commit()

    return {
        "ok": True,
        "status": "EM_ROTA",
        "veiculo": veiculo_rota,
        "km_inicial": safe_float(payload.km_inicial, 0.0),
        "ultimo_km_veiculo": safe_float(ultimo.get("km_final"), 0.0),
        "ultimo_km_programacao": str(ultimo.get("codigo_programacao") or ""),
    }


@app.post("/rotas/{codigo_programacao}/finalizar")
def finalizar_rota(codigo_programacao: str, payload: FinalizarRotaIn, m=Depends(get_current_motorista)):
    codigo_programacao = (codigo_programacao or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "finalizar_rota", payload.idempotency_key):
            return {"ok": True, "status": "FINALIZADA", "deduplicado": True}
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.km_inicial")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")
        if int(payload.km_final or 0) < 0:
            raise HTTPException(status_code=400, detail="KM final invalido.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_atual}).")
        if status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(status_code=409, detail=f"Transicao invalida para finalizar (status={status_atual or 'N/D'}).")
        if _has_pending_substituicao(cur, codigo_programacao):
            raise HTTPException(
                status_code=409,
                detail="Nao e possivel finalizar: existe substituicao de motorista pendente de aceite.",
            )

        try:
            km_ini = int(pr["km_inicial"]) if pr["km_inicial"] is not None else None
        except Exception:
            km_ini = None
        if km_ini is not None and km_ini > 0 and int(payload.km_final or 0) > 0 and int(payload.km_final) < km_ini:
            raise HTTPException(status_code=409, detail="KM final menor que KM inicial.")

        # ============================
        # RECONCILIACAO ANTIFRAUDE
        # ============================
        cur.execute("PRAGMA table_info(programacoes)")
        cols_prog = {row[1] for row in cur.fetchall() or []}
        cand_cols = [c for c in ("caixas_carregadas", "qnt_cx_carregada", "nf_caixas", "total_caixas") if c in cols_prog]
        caixas_carregadas = 0
        def _to_int_db(v: Any) -> int:
            try:
                if v is None:
                    return 0
                if isinstance(v, (int, float)):
                    return int(float(v))
                s = str(v).strip()
                if not s:
                    return 0
                s = s.replace(" ", "")
                if "," in s:
                    s = s.replace(".", "").replace(",", ".")
                return int(float(s))
            except Exception:
                return 0

        if cand_cols:
            cur.execute(
                f"SELECT {', '.join([f'COALESCE({c},0)' for c in cand_cols])} FROM programacoes WHERE id=? LIMIT 1",
                (pr["id"],),
            )
            rw = cur.fetchone()
            if rw:
                for i in range(len(cand_cols)):
                    v = _to_int_db(rw[i])
                    if v > 0:
                        caixas_carregadas = v
                        break
        if caixas_carregadas <= 0:
            try:
                cur.execute(
                    """
                    SELECT COALESCE(SUM(COALESCE(qnt_caixas, 0)), 0) AS qtd
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                    """,
                    (codigo_programacao,),
                )
                row_itens = cur.fetchone()
                caixas_carregadas = int(float((row_itens["qtd"] if row_itens else 0) or 0))
            except Exception:
                caixas_carregadas = 0
            if caixas_carregadas <= 0:
                raise HTTPException(
                    status_code=409,
                    detail="Nao e possivel finalizar: caixas carregadas nao informadas no carregamento.",
                )

        cur.execute(
            """
            SELECT COALESCE(COUNT(*),0)
            FROM transferencias
            WHERE (codigo_origem=? OR codigo_destino=?)
              AND (
                    UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
                 OR (
                        UPPER(TRIM(COALESCE(status,'')))='ACEITA'
                    AND COALESCE(qtd_convertida, 0) < COALESCE(qtd_caixas, 0)
                    )
              )
            """,
            (codigo_programacao, codigo_programacao),
        )
        pend_transfer = int((cur.fetchone() or [0])[0] or 0)
        if pend_transfer > 0:
            raise HTTPException(
                status_code=409,
                detail=f"Nao e possivel finalizar: existem {pend_transfer} transferencia(s) pendente(s) ou nao convertida(s).",
            )

        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_pi = {row[1] for row in cur.fetchall() or []}
        cur.execute("PRAGMA table_info(programacao_itens_controle)")
        cols_pc = {row[1] for row in cur.fetchall() or []}

        has_pi_pedido = "pedido" in cols_pi
        has_pc_pedido = "pedido" in cols_pc
        has_pi_status = "status_pedido" in cols_pi
        has_pc_status = "status_pedido" in cols_pc

        join_on = "pc.codigo_programacao = pi.codigo_programacao AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))"
        if has_pi_pedido and has_pc_pedido:
            join_on += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"

        st_pi_expr = "COALESCE(NULLIF(TRIM(pi.status_pedido),''), 'PENDENTE')" if has_pi_status else "'PENDENTE'"
        st_pc_expr = "NULLIF(TRIM(pc.status_pedido),'')" if has_pc_status else "NULL"
        pedido_expr = "COALESCE(pi.pedido, '')" if has_pi_pedido else "''"

        cur.execute(
            f"""
            SELECT
                COALESCE(pi.cod_cliente, '') AS cod_cliente,
                {pedido_expr} AS pedido,
                COALESCE({st_pc_expr}, {st_pi_expr}, 'PENDENTE') AS status_eff
            FROM programacao_itens pi
            LEFT JOIN programacao_itens_controle pc
              ON {join_on}
            WHERE pi.codigo_programacao=?
            """,
            (codigo_programacao,),
        )
        pendentes = []
        for it in (cur.fetchall() or []):
            status_eff = str(it["status_eff"] or "PENDENTE").strip().upper()
            if status_eff == "" or status_eff == "PENDENTE":
                pendentes.append(f"{it['cod_cliente']} / {it['pedido'] or '-'} [{status_eff}]")

        if pendentes:
            raise HTTPException(
                status_code=409,
                detail=f"Nao e possivel finalizar: {len(pendentes)} pedido(s) pendente(s).",
            )

        total_em_aberto = _total_caixas_ativas_programacao(cur, codigo_programacao)
        if total_em_aberto != 0:
            raise HTTPException(
                status_code=409,
                detail=(
                    "Reconciliacao nao fechou. "
                    f"Carregadas={caixas_carregadas}, saldo em aberto={total_em_aberto}. "
                    "Revise cancelamentos/redirecionamentos/entregas."
                ),
            )
        cur.execute(
            """
            UPDATE programacoes
               SET status='FINALIZADA',
                   data_chegada=?,
                   hora_chegada=?,
                   km_final=?
             WHERE id=?
            """,
            (payload.data_chegada, payload.hora_chegada, payload.km_final, pr["id"]),
        )
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="FINALIZAR_ROTA",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino="FINALIZADA",
            motorista_codigo=codigo_motorista,
            data_hora=f"{payload.data_chegada} {payload.hora_chegada}".strip(),
            caixas=caixas_carregadas,
            observacao=f"KM final {payload.km_final}",
            payload={**_payload_dict(payload), "km_inicial": km_ini, "caixas_carregadas": caixas_carregadas},
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "finalizar_rota", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "status": "FINALIZADA"}


@app.get("/rotas/{codigo_programacao}/nfs-disponiveis")
def listar_nfs_disponiveis_carregamento(
    codigo_programacao: str,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_compras_app_schema(cur)
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.veiculo")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")
        cur.execute(
            f"""
            SELECT c.*,
                   COALESCE(f.perfil_fornecedor, '') AS perfil_fornecedor,
                   COALESCE(f.razao_social, c.fornecedor_razao, '') AS fornecedor_nome
              FROM compras_nfe c
              LEFT JOIN fornecedores f
                ON (
                    (c.fornecedor_id IS NOT NULL AND f.id=c.fornecedor_id)
                    OR (TRIM(COALESCE(c.fornecedor_documento,''))<>'' AND TRIM(COALESCE(f.documento,''))=TRIM(COALESCE(c.fornecedor_documento,'')))
                )
             WHERE TRIM(COALESCE(c.numero,''))<>''
               AND (c.codigo_programacao IS NULL OR TRIM(c.codigo_programacao)='' OR UPPER(TRIM(c.codigo_programacao))=UPPER(TRIM(?)))
               AND COALESCE(c.estoque_kg_entrada, 0) > 0
               AND UPPER(TRIM(COALESCE(c.situacao_nfe, 'AUTORIZADO'))) NOT IN ('CANCELADA', 'CANCELADO', 'DENEGADA')
               AND {_nf_frango_clause()}
             ORDER BY COALESCE(c.emissao, c.created_at, '') DESC, c.id DESC
             LIMIT 200
            """,
            (codigo_programacao,),
        )
        rows = [_nf_compra_payload(r) for r in (cur.fetchall() or [])]
    return {"ok": True, "codigo_programacao": codigo_programacao, "total": len(rows), "notas": rows}


@app.get("/notas-fiscais/frango/disponiveis")
def listar_nfs_frango_disponiveis(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        _ensure_compras_app_schema(cur)
        cur.execute(
            f"""
            SELECT c.*,
                   COALESCE(f.perfil_fornecedor, '') AS perfil_fornecedor,
                   COALESCE(f.razao_social, c.fornecedor_razao, '') AS fornecedor_nome
              FROM compras_nfe c
              LEFT JOIN fornecedores f
                ON (
                    (c.fornecedor_id IS NOT NULL AND f.id=c.fornecedor_id)
                    OR (TRIM(COALESCE(c.fornecedor_documento,''))<>'' AND TRIM(COALESCE(f.documento,''))=TRIM(COALESCE(c.fornecedor_documento,'')))
                )
             WHERE TRIM(COALESCE(c.numero,''))<>''
               AND (c.codigo_programacao IS NULL OR TRIM(c.codigo_programacao)='')
               AND COALESCE(c.estoque_kg_entrada, 0) > 0
               AND UPPER(TRIM(COALESCE(c.situacao_nfe, 'AUTORIZADO'))) NOT IN ('CANCELADA', 'CANCELADO', 'DENEGADA')
               AND {_nf_frango_clause()}
             ORDER BY COALESCE(c.emissao, c.created_at, '') DESC, c.id DESC
             LIMIT 200
            """
        )
        rows = [_nf_compra_payload(r) for r in (cur.fetchall() or [])]
    return {"ok": True, "total": len(rows), "notas": rows}


@app.post("/rotas/{codigo_programacao}/carregamento")
def salvar_carregamento(
    codigo_programacao: str,
    payload: CarregamentoIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "carregamento", payload.idempotency_key):
            return {"ok": True, "deduplicado": True}

        cur.execute("PRAGMA table_info(programacoes)")
        cols = {row[1] for row in cur.fetchall()}

        def has(col: str) -> bool:
            return col in cols

        carregamento_fechado_sel = "p.carregamento_fechado" if has("carregamento_fechado") else "0 AS carregamento_fechado"
        tipo_estimativa_sel = "p.tipo_estimativa" if has("tipo_estimativa") else "'KG' AS tipo_estimativa"

        # garante que a programação é do motorista logado
        pr = _fetch_programacao_owned(
            cur,
            codigo_programacao,
            m,
            f"p.id, p.status, {carregamento_fechado_sel}, {tipo_estimativa_sel}",
        )
        if not pr:
            raise HTTPException(status_code=404, detail="Rota não encontrada para este motorista")
        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota encerrada (status={status_atual}).",
            )
        if status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota ainda nao iniciada (status={status_atual or 'N/D'}). Inicie a rota antes do carregamento.",
            )

        # apos o primeiro salvamento, o carregamento fica fechado para edicoes
        if has("carregamento_fechado"):
            ja_fechado_raw = 0
            try:
                if hasattr(pr, "keys") and "carregamento_fechado" in pr.keys():
                    ja_fechado_raw = pr["carregamento_fechado"]
            except Exception:
                ja_fechado_raw = 0
            ja_fechado = int(ja_fechado_raw or 0)
            if ja_fechado > 0:
                raise HTTPException(
                    status_code=409,
                    detail="Carregamento ja foi salvo e esta bloqueado para alteracoes.",
                )

        # normaliza
        nf_numero = (payload.nf_numero or "").strip()
        nf_kg = float(payload.nf_kg or 0.0)
        nf_preco = float(payload.nf_preco or 0.0)
        local_carregado = (payload.local_carregado or "").strip()
        kg_carregado = float(payload.kg_carregado or 0.0) if payload.kg_carregado is not None else 0.0
        caixas = int(payload.caixas_carregadas or 0)
        inicio = (payload.inicio_carregamento or "").strip()
        fim = (payload.fim_carregamento or "").strip()

        aves_por_caixa = int(payload.qnt_aves_por_cx or 0)
        if aves_por_caixa <= 0:
            aves_por_caixa = 6

        media = payload.media
        media = float(media) if media is not None else None
        media_1 = payload.media_1
        media_2 = payload.media_2
        media_3 = payload.media_3
        media_1 = float(media_1) if media_1 is not None else None
        media_2 = float(media_2) if media_2 is not None else None
        media_3 = float(media_3) if media_3 is not None else None

        mortalidade = int(payload.mortalidade_aves or 0)
        if mortalidade < 0:
            mortalidade = 0

        caixa_final_raw = payload.aves_caixa_final
        if caixa_final_raw is None:
            caixa_final_raw = payload.qnt_aves_caixa_final
        caixa_final = int(caixa_final_raw or 0)
        if caixa_final < 0:
            caixa_final = 0

        nf_compra = _fetch_nf_compra_by_numero(cur, nf_numero)
        if nf_compra is not None:
            vinculo_atual = str(nf_compra["codigo_programacao"] or "").strip().upper()
            if vinculo_atual and vinculo_atual != codigo_programacao.strip().upper():
                raise HTTPException(
                    status_code=409,
                    detail=f"NF {nf_numero} ja esta vinculada a programacao {vinculo_atual}.",
                )
            if not _nf_compra_is_frango(nf_compra):
                raise HTTPException(
                    status_code=409,
                    detail=f"NF {nf_numero} nao pertence a fornecedor/produto de frango vivo.",
                )
            nf_payload = _nf_compra_payload(nf_compra)
            if nf_kg <= 0:
                nf_kg = float(nf_payload["nf_kg"] or 0.0)
            if nf_preco <= 0:
                nf_preco = float(nf_payload["nf_preco"] or 0.0)
            if caixas <= 0:
                caixas = int(nf_payload["nf_caixas"] or 0)
            if not local_carregado:
                local_carregado = str(nf_payload["fornecedor"] or "").strip()

        if kg_carregado <= 0 and caixas > 0 and aves_por_caixa > 0 and media is not None and float(media or 0) > 0:
            kg_carregado = round(float(media or 0) * float(caixas) * float(aves_por_caixa), 3)
        nf_saldo = round(max(nf_kg - kg_carregado, 0.0), 2) if nf_kg > 0 and kg_carregado > 0 else 0.0

        tipo_estimativa = str(pr["tipo_estimativa"] or "KG").strip().upper()
        if tipo_estimativa not in ("KG", "CX"):
            tipo_estimativa = "KG"
        if tipo_estimativa == "KG" and nf_kg <= 0 and kg_carregado <= 0:
            raise HTTPException(
                status_code=400,
                detail="CIF exige peso informado (nf_kg ou kg_carregado maior que zero).",
            )

        sets = []
        params: List[Any] = []

        # status: apos carregamento com rota em andamento, avancar para EM_ENTREGAS
        status_result = "CARREGADA"
        if has("status"):
            st_atual_raw = ""
            try:
                if hasattr(pr, "keys") and "status" in pr.keys():
                    st_atual_raw = pr["status"] or ""
            except Exception:
                st_atual_raw = ""
            st_atual = str(st_atual_raw).strip().upper()
            if st_atual in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
                sets.append("status=?")
                params.append("EM_ENTREGAS")
                status_result = "EM_ENTREGAS"
            else:
                sets.append("status=?")
                params.append("CARREGADA")
                status_result = "CARREGADA"

        # âœ… campos que o app usa (se existirem)
        if has("nf_numero"):
            sets.append("nf_numero=?"); params.append(nf_numero)
        if has("nf_kg"):
            sets.append("nf_kg=?"); params.append(nf_kg)
        if has("nf_preco"):
            sets.append("nf_preco=?"); params.append(nf_preco)
        if has("local_carregado"):
            sets.append("local_carregado=?"); params.append(local_carregado)
        if has("kg_carregado"):
            sets.append("kg_carregado=?"); params.append(kg_carregado)
        if has("nf_kg_carregado"):
            sets.append("nf_kg_carregado=?"); params.append(kg_carregado)
        if has("nf_saldo"):
            sets.append("nf_saldo=?"); params.append(nf_saldo)
        if has("caixas_carregadas"):
            sets.append("caixas_carregadas=?"); params.append(caixas)
        if has("inicio_carregamento"):
            sets.append("inicio_carregamento=?"); params.append(inicio or None)
        if has("fim_carregamento"):
            sets.append("fim_carregamento=?"); params.append(fim or None)
        if has("carregamento_fechado"):
            sets.append("carregamento_fechado=?"); params.append(1)
        if has("carregamento_salvo_em"):
            sets.append("carregamento_salvo_em=?"); params.append(datetime.now().isoformat(timespec="seconds"))

        # ✅ colunas “desktop”/alternativas (se existirem)
        if has("num_nf"):
            sets.append("num_nf=?"); params.append(nf_numero)
        if has("kg_nf"):
            sets.append("kg_nf=?"); params.append(nf_kg)
        if has("preco_nf"):
            sets.append("preco_nf=?"); params.append(nf_preco)
        if has("granja_carregada"):
            sets.append("granja_carregada=?"); params.append(local_carregado)
        if has("qnt_cx_carregada"):
            sets.append("qnt_cx_carregada=?"); params.append(caixas)
        if has("qnt_aves_por_cx"):
            sets.append("qnt_aves_por_cx=?"); params.append(aves_por_caixa)
        if has("aves_caixa_final"):
            sets.append("aves_caixa_final=?"); params.append(caixa_final)
        if has("qnt_aves_caixa_final"):
            sets.append("qnt_aves_caixa_final=?"); params.append(caixa_final)
        if has("mortalidade_aves"):
            sets.append("mortalidade_aves=?"); params.append(mortalidade)

        # média (só se vier)
        if media is not None:
            if has("media"):
                sets.append("media=?"); params.append(media)
        if media_1 is not None and has("media_1"):
            sets.append("media_1=?"); params.append(media_1)
        if media_2 is not None and has("media_2"):
            sets.append("media_2=?"); params.append(media_2)
        if media_3 is not None and has("media_3"):
            sets.append("media_3=?"); params.append(media_3)

        if not sets:
            return {"ok": True, "status": status_atual, "warning": "Nenhuma coluna compatível encontrada para atualizar."}

        sql = f"UPDATE programacoes SET {', '.join(sets)} WHERE id=?"
        params.append(pr["id"])
        cur.execute(sql, tuple(params))
        if nf_compra is not None:
            cur.execute(
                """
                UPDATE compras_nfe
                   SET codigo_programacao=?,
                       vinculada_em=?,
                       vinculada_por=?,
                       updated_at=?
                 WHERE id=?
                   AND (codigo_programacao IS NULL OR TRIM(codigo_programacao)='' OR UPPER(TRIM(codigo_programacao))=UPPER(TRIM(?)))
                """,
                (
                    codigo_programacao.strip().upper(),
                    datetime.now().isoformat(timespec="seconds"),
                    codigo_motorista,
                    datetime.now().isoformat(timespec="seconds"),
                    int(nf_compra["id"] or 0),
                    codigo_programacao.strip().upper(),
                ),
            )
            if int(cur.rowcount or 0) <= 0:
                raise HTTPException(status_code=409, detail=f"NF {nf_numero} foi vinculada por outra rota. Recarregue a lista.")
        caixas_ativas = _total_caixas_ativas_programacao(cur, codigo_programacao)
        deficit_caixas = max(0, int(caixas_ativas) - int(caixas))
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="CARREGAMENTO",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino=local_carregado,
            motorista_codigo=codigo_motorista,
            caixas=caixas,
            kg=kg_carregado or nf_kg,
            media=media or 0,
            aves_por_caixa=aves_por_caixa,
            nf_numero=nf_numero,
            nf_preco=nf_preco,
            data_hora=fim or inicio or _now_iso(),
            observacao=f"Saldo NF KG: {nf_saldo:.2f}" if nf_saldo else "",
            payload={**_payload_dict(payload), "status_result": status_result, "nf_saldo": nf_saldo},
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "carregamento", payload.idempotency_key)
        conn.commit()

    return {
        "ok": True,
        "status": status_result,
        "reconciliar": deficit_caixas > 0,
        "caixas_carregadas": int(caixas),
        "caixas_ativas": int(caixas_ativas),
        "deficit_caixas": int(deficit_caixas),
    }


@app.post("/rotas/{codigo_programacao}/CARREGAR")
def salvar_carregamento_alias(
    codigo_programacao: str,
    payload: CarregamentoIn,
    m=Depends(get_current_motorista),
):
    return salvar_carregamento(codigo_programacao, payload, m)


@app.post("/rotas/{codigo_programacao}/transbordo")
def salvar_transbordo_rota(
    codigo_programacao: str,
    payload: RotaTransbordoIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    nome_motorista = (m.get("nome") or "").strip()
    company_id = int(m.get("company_id") or 1)
    is_admin = bool(m.get("is_admin"))

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "transbordo", payload.idempotency_key):
            return {"ok": True, "deduplicado": True}

        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.veiculo")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")
        status_atual = str(pr["status"] or "").strip().upper()
        if (not is_admin) and status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_atual}).")

        aves = payload.mortalidade_transbordo_aves
        if aves is None:
            aves = payload.aves_mortas_transbordo
        aves = max(0, int(aves or 0))
        kg = payload.mortalidade_transbordo_kg
        kg = float(kg or 0.0)
        obs = (payload.obs_transbordo or payload.mortalidade_transbordo_obs or "").strip() or None
        path_foto = (
            payload.foto_doa_path
            or payload.doa_foto_path
            or payload.mortalidade_transbordo_foto_path
            or ""
        )
        foto_ref = _store_mobile_photo(
            cur,
            codigo_programacao,
            payload.foto_doa or payload.foto_registro,
            "MORTALIDADE_DOA",
            "DOA_TRANSBORDO",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            id_vinculo="TRANSBORDO",
            path_hint=path_foto,
            company_id=company_id,
        )

        cur.execute("PRAGMA table_info(programacoes)")
        cols = {row[1] for row in cur.fetchall() or []}
        sets = []
        params: List[Any] = []

        def add_prog(col: str, value: Any):
            if col in cols:
                sets.append(f"{col}=?")
                params.append(value)

        add_prog("mortalidade_transbordo_aves", aves)
        add_prog("mortalidade_transbordo_kg", kg)
        add_prog("obs_transbordo", obs)
        add_prog("foto_doa_path", payload.foto_doa_path or path_foto or None)
        add_prog("doa_foto_path", payload.doa_foto_path or path_foto or None)
        add_prog("mortalidade_transbordo_foto_path", payload.mortalidade_transbordo_foto_path or path_foto or None)
        if foto_ref is not None:
            add_prog("foto_doa_ref_json", json.dumps(foto_ref, ensure_ascii=False))
        if sets:
            params.append(pr["id"])
            cur.execute(f"UPDATE programacoes SET {', '.join(sets)} WHERE id=?", tuple(params))

        try:
            cur.execute(
                """
                INSERT INTO programacao_itens_log (codigo_programacao, cod_cliente, pedido, evento, payload_json, registrado_em, company_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    codigo_programacao,
                    "__TRANSBORDO__",
                    "",
                    "transbordo",
                    json.dumps({**_payload_dict(payload), "foto_ref": foto_ref}, ensure_ascii=False),
                    _now_iso(),
                    company_id,
                ),
            )
        except Exception:
            pass
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="TRANSBORDO",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino="TRANSBORDO",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            kg=kg,
            observacao=obs or "",
            payload={**_payload_dict(payload), "aves": aves, "foto_ref": foto_ref},
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "transbordo", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "aves": aves, "kg": kg, "foto": foto_ref}


@app.post("/rotas/{codigo_programacao}/despesas")
def salvar_despesa_rota_mobile(
    codigo_programacao: str,
    payload: RotaDespesaIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    codigo_motorista = (payload.motorista_codigo or m.get("codigo") or "").strip().upper()
    nome_motorista = (payload.motorista_nome or m.get("nome") or "").strip()
    company_id = int(m.get("company_id") or 1)
    idem = payload.idempotency_key or payload.sync_key or payload.id_local

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "despesas", idem):
            return {"ok": True, "deduplicado": True}
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.veiculo")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")

        valor = float(payload.valor_total or 0.0)
        if valor <= 0:
            raise HTTPException(status_code=400, detail="valor_total deve ser maior que zero.")
        tipo = (payload.tipo or "OUTRAS").strip().upper()
        registrado_em = (payload.registrado_em or payload.timestamp or _now_iso()).strip()
        id_local = (payload.id_local or payload.sync_key or "").strip() or None
        foto_ref = _store_mobile_photo(
            cur,
            codigo_programacao,
            payload.foto_despesa or payload.foto_registro,
            "DESPESA",
            f"DESPESA_{tipo}",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            id_vinculo=id_local or tipo,
            path_hint=payload.comprovante_path or "",
            company_id=company_id,
        )

        cur.execute("PRAGMA table_info(despesas)")
        cols = {row[1] for row in cur.fetchall() or []}
        data = {
            "codigo_programacao": codigo_programacao,
            "descricao": payload.descricao or tipo,
            "valor": valor,
            "data_registro": registrado_em,
            "tipo_despesa": "ROTA",
            "categoria": tipo,
            "motorista": nome_motorista,
            "veiculo": pr["veiculo"] if "veiculo" in pr.keys() else None,
            "observacao": payload.descricao,
            "id_local": id_local,
            "forma_pagamento": payload.forma_pagamento,
            "comprovante_path": payload.comprovante_path,
            "estabelecimento": payload.estabelecimento,
            "documento": payload.documento,
            "litros": payload.litros,
            "valor_litro": payload.valor_litro,
            "desconto": payload.desconto,
            "combustivel": payload.combustivel,
            "odometro": payload.odometro,
            "lat": payload.lat,
            "lon": payload.lon,
            "accuracy": payload.accuracy,
            "registrado_em": registrado_em,
            "motorista_codigo": codigo_motorista,
            "motorista_nome": nome_motorista,
            "sync_key": payload.sync_key,
            "status_sync": payload.status_sync or "SINCRONIZADO",
            "origem": payload.origem or "APP_MOTORISTA",
            "vinculo_prestacao_json": json.dumps(payload.vinculo_prestacao or {}, ensure_ascii=False),
            "desktop_web_json": json.dumps(payload.desktop_web or {}, ensure_ascii=False),
            "foto_despesa_ref_json": json.dumps(foto_ref or {}, ensure_ascii=False),
            "company_id": company_id,
        }
        insert_cols = [col for col in data.keys() if col in cols]
        if not insert_cols:
            raise HTTPException(status_code=500, detail="Tabela de despesas sem colunas compativeis.")

        updated = False
        if id_local and "id_local" in cols:
            cur.execute(
                "SELECT id FROM despesas WHERE codigo_programacao=? AND id_local=? LIMIT 1",
                (codigo_programacao, id_local),
            )
            old = cur.fetchone()
            if old:
                set_cols = [col for col in insert_cols if col not in ("codigo_programacao", "id_local")]
                params = [data[col] for col in set_cols]
                params.append(old["id"])
                cur.execute(f"UPDATE despesas SET {', '.join([c + '=?' for c in set_cols])} WHERE id=?", tuple(params))
                updated = True
        if not updated:
            placeholders = ", ".join(["?"] * len(insert_cols))
            cur.execute(
                f"INSERT INTO despesas ({', '.join(insert_cols)}) VALUES ({placeholders})",
                tuple(data[col] for col in insert_cols),
            )
        despesa_id = cur.lastrowid
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="DESPESA",
            codigo_programacao=codigo_programacao,
            origem=payload.origem or "APP_MOTORISTA",
            destino=payload.estabelecimento or tipo,
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            kg=payload.litros or 0,
            data_hora=registrado_em,
            observacao=payload.descricao or tipo,
            payload={**_payload_dict(payload), "despesa_id": despesa_id, "foto_ref": foto_ref},
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "despesas", idem)
        conn.commit()

    return {"ok": True, "id": despesa_id, "id_local": id_local, "foto": foto_ref}


@app.post("/rotas/{codigo_programacao}/ajudantes")
def alterar_ajudantes_rota(
    codigo_programacao: str,
    payload: RotaAjudantesIn,
    m=Depends(get_current_motorista),
):
    codigo_programacao = (codigo_programacao or "").strip()
    codigo_motorista = (m.get("codigo") or "").strip().upper()
    nome_motorista = (m.get("nome") or "").strip()
    company_id = int(m.get("company_id") or 1)
    novos = (payload.ajudantes_novos or "").strip()
    if not novos:
        raise HTTPException(status_code=400, detail="ajudantes_novos e obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        if _idempotency_seen(cur, codigo_motorista, codigo_programacao, "ajudantes", payload.idempotency_key):
            return {"ok": True, "deduplicado": True}
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status, p.equipe")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista")

        alterado_em = (payload.alterado_em or _now_iso()).strip()
        historico_item = {
            "anteriores": payload.ajudantes_anteriores if payload.ajudantes_anteriores is not None else pr["equipe"],
            "novos": novos,
            "motivo": payload.motivo,
            "origem": payload.origem or "APP_MOTORISTA",
            "alterado_em": alterado_em,
            "alterado_por": nome_motorista or codigo_motorista,
        }
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {row[1] for row in cur.fetchall() or []}
        sets = []
        params: List[Any] = []
        if "equipe" in cols:
            sets.append("equipe=?"); params.append(novos)
        if "ajudantes_alteracao_motivo" in cols:
            sets.append("ajudantes_alteracao_motivo=?"); params.append(payload.motivo)
        if "ajudantes_alterado_em" in cols:
            sets.append("ajudantes_alterado_em=?"); params.append(alterado_em)
        if "historico_ajudantes" in cols:
            anterior_json = ""
            try:
                cur.execute("SELECT historico_ajudantes FROM programacoes WHERE id=? LIMIT 1", (pr["id"],))
                row_hist = cur.fetchone()
                anterior_json = str(row_hist["historico_ajudantes"] or "") if row_hist else ""
            except Exception:
                anterior_json = ""
            try:
                historico = json.loads(anterior_json) if anterior_json else []
                if not isinstance(historico, list):
                    historico = []
            except Exception:
                historico = []
            historico.append(historico_item)
            sets.append("historico_ajudantes=?"); params.append(json.dumps(historico, ensure_ascii=False))
        if sets:
            params.append(pr["id"])
            cur.execute(f"UPDATE programacoes SET {', '.join(sets)} WHERE id=?", tuple(params))
        try:
            cur.execute(
                """
                INSERT INTO programacao_itens_log (codigo_programacao, cod_cliente, pedido, evento, payload_json, registrado_em, company_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    codigo_programacao,
                    "__AJUDANTES__",
                    "",
                    "ajudantes",
                    json.dumps(historico_item, ensure_ascii=False),
                    alterado_em,
                    company_id,
                ),
            )
        except Exception:
            pass
        _registrar_roteiro_operacional(
            cur,
            tipo_evento="AJUDANTES",
            codigo_programacao=codigo_programacao,
            origem="APP_MOTORISTA",
            destino="EQUIPE",
            motorista_codigo=codigo_motorista,
            motorista_nome=nome_motorista,
            data_hora=alterado_em,
            observacao=payload.motivo or "",
            payload=historico_item,
        )
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "ajudantes", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "ajudantes": novos, "alterado_em": alterado_em}


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _evento_iso_or_now(raw: Optional[str]) -> str:
    txt = str(raw or "").strip()
    if not txt:
        return _now_iso()
    try:
        return datetime.fromisoformat(txt).isoformat(timespec="seconds")
    except Exception:
        return _now_iso()


def _idempotency_seen(
    cur: sqlite3.Cursor,
    motorista_codigo: str,
    codigo_programacao: str,
    endpoint: str,
    idem_key: Optional[str],
) -> bool:
    key = str(idem_key or "").strip()
    if not key:
        return False
    cur.execute(
        """
        SELECT 1
        FROM mobile_sync_idempotency
        WHERE motorista_codigo=?
          AND codigo_programacao=?
          AND endpoint=?
          AND idem_key=?
        LIMIT 1
        """,
        (
            str(motorista_codigo or "").strip().upper(),
            str(codigo_programacao or "").strip(),
            str(endpoint or "").strip(),
            key,
        ),
    )
    return cur.fetchone() is not None


def _idempotency_mark(
    cur: sqlite3.Cursor,
    motorista_codigo: str,
    codigo_programacao: str,
    endpoint: str,
    idem_key: Optional[str],
) -> None:
    key = str(idem_key or "").strip()
    if not key:
        return
    cur.execute(
        """
        INSERT OR IGNORE INTO mobile_sync_idempotency
            (motorista_codigo, codigo_programacao, endpoint, idem_key, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            str(motorista_codigo or "").strip().upper(),
            str(codigo_programacao or "").strip(),
            str(endpoint or "").strip(),
            key,
            _now_iso(),
        ),
    )


def _safe_file_segment(value: Any, fallback: str = "item") -> str:
    txt = str(value or "").strip()
    txt = re.sub(r"[^A-Za-z0-9_.-]+", "_", txt)
    txt = txt.strip("._-")
    return txt[:80] or fallback


def _mobile_photo_storage_root() -> str:
    root = os.environ.get("ROTA_MOBILE_PHOTOS_DIR") or os.path.join(BASE_DIR, ".rotahub_runtime", "fotos_rotas")
    os.makedirs(root, exist_ok=True)
    return root


def _store_mobile_photo(
    cur: sqlite3.Cursor,
    codigo_programacao: str,
    photo: Optional[Dict[str, Any]],
    default_categoria: str,
    default_tipo: str,
    *,
    motorista_codigo: str = "",
    motorista_nome: str = "",
    cod_cliente: str = "",
    cliente_nome: str = "",
    pedido: str = "",
    id_vinculo: str = "",
    path_hint: str = "",
    company_id: Optional[int] = None,
) -> Optional[Dict[str, Any]]:
    photo_in = dict(photo or {})
    path_local = str(photo_in.get("path_local") or photo_in.get("arquivo_path") or path_hint or "").strip()
    imagem_base64 = str(photo_in.get("imagem_base64") or photo_in.get("base64") or "").strip()
    if not photo_in and not path_local and not imagem_base64:
        return None

    categoria = str(photo_in.get("categoria") or default_categoria or "GERAL").strip().upper()
    tipo_registro = str(photo_in.get("tipo_registro") or photo_in.get("tipo") or default_tipo or categoria).strip().upper()
    registrado_em = str(photo_in.get("registrado_em") or photo_in.get("created_at") or _now_iso()).strip()
    id_foto = str(photo_in.get("id_foto") or photo_in.get("id") or "").strip()
    if not id_foto:
        id_foto = f"{_safe_file_segment(codigo_programacao, 'rota')}_{_safe_file_segment(tipo_registro, 'foto')}_{uuid4().hex[:12]}"

    mime_type = str(photo_in.get("mime_type") or photo_in.get("content_type") or "image/jpeg").strip()
    arquivo_nome = str(photo_in.get("arquivo_nome_sugerido") or photo_in.get("arquivo_nome") or "").strip()
    ext = ".png" if "png" in mime_type.lower() else ".jpg"
    if not arquivo_nome:
        arquivo_nome = f"{_safe_file_segment(id_foto, 'foto')}{ext}"
    safe_name = _safe_file_segment(arquivo_nome, f"{_safe_file_segment(id_foto, 'foto')}{ext}")
    if "." not in os.path.basename(safe_name):
        safe_name += ext

    storage_path = str(photo_in.get("storage_path") or "").strip()
    tamanho_bytes = 0
    if imagem_base64:
        try:
            if "," in imagem_base64[:80]:
                imagem_base64 = imagem_base64.split(",", 1)[1]
            data = base64.b64decode(imagem_base64, validate=False)
            destino_dir = os.path.join(
                _mobile_photo_storage_root(),
                _safe_file_segment(codigo_programacao, "rota"),
                _safe_file_segment(categoria, "categoria"),
            )
            os.makedirs(destino_dir, exist_ok=True)
            destino = os.path.join(destino_dir, safe_name)
            with open(destino, "wb") as fh:
                fh.write(data)
            storage_path = destino
            tamanho_bytes = len(data)
        except Exception:
            storage_path = storage_path or ""

    ref = {
        "id_foto": id_foto,
        "categoria": categoria,
        "tipo_registro": tipo_registro,
        "path_local": path_local,
        "storage_path": storage_path,
        "arquivo_nome": safe_name,
        "mime_type": mime_type,
        "tamanho_bytes": tamanho_bytes,
        "registrado_em": registrado_em,
    }
    payload_to_store = dict(photo_in)
    payload_to_store.pop("imagem_base64", None)
    payload_to_store.pop("base64", None)
    payload_to_store.update(ref)

    try:
        cur.execute(
            """
            INSERT INTO rota_fotos (
                id_foto, codigo_programacao, categoria, tipo_registro, cod_cliente, cliente_nome, pedido,
                id_vinculo, path_local, storage_path, arquivo_nome, mime_type, tamanho_bytes,
                motorista_codigo, motorista_nome, registrado_em, payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id_foto) DO UPDATE SET
                codigo_programacao=excluded.codigo_programacao,
                categoria=excluded.categoria,
                tipo_registro=excluded.tipo_registro,
                cod_cliente=excluded.cod_cliente,
                cliente_nome=excluded.cliente_nome,
                pedido=excluded.pedido,
                id_vinculo=excluded.id_vinculo,
                path_local=excluded.path_local,
                storage_path=excluded.storage_path,
                arquivo_nome=excluded.arquivo_nome,
                mime_type=excluded.mime_type,
                tamanho_bytes=excluded.tamanho_bytes,
                motorista_codigo=excluded.motorista_codigo,
                motorista_nome=excluded.motorista_nome,
                registrado_em=excluded.registrado_em,
                payload_json=excluded.payload_json
            """,
            (
                id_foto,
                codigo_programacao,
                categoria,
                tipo_registro,
                cod_cliente or None,
                cliente_nome or None,
                pedido or None,
                id_vinculo or None,
                path_local or None,
                storage_path or None,
                safe_name,
                mime_type,
                tamanho_bytes,
                motorista_codigo or None,
                motorista_nome or None,
                registrado_em,
                json.dumps(payload_to_store, ensure_ascii=False),
            ),
        )
        try:
            cur.execute("PRAGMA table_info(rota_fotos)")
            cols_fotos = {row[1] for row in cur.fetchall() or []}
            if "company_id" in cols_fotos:
                cur.execute(
                    "UPDATE rota_fotos SET company_id=COALESCE(company_id, ?) WHERE id_foto=?",
                    (int(company_id or 1), id_foto),
                )
        except Exception:
            pass
    except Exception:
        return ref
    return ref


def _payload_dict(payload: Any) -> Dict[str, Any]:
    if payload is None:
        return {}
    try:
        if hasattr(payload, "model_dump"):
            return dict(payload.model_dump())
    except Exception:
        pass
    try:
        if hasattr(payload, "dict"):
            return dict(payload.dict())
    except Exception:
        pass
    if isinstance(payload, dict):
        return dict(payload)
    return {"value": str(payload)}


def _registrar_roteiro_operacional(
    cur: sqlite3.Cursor,
    *,
    tipo_evento: str,
    codigo_programacao: str,
    origem: str = "",
    destino: str = "",
    motorista_codigo: str = "",
    motorista_nome: str = "",
    pedido: str = "",
    cod_cliente: str = "",
    cliente_nome: str = "",
    caixas: Any = 0,
    kg: Any = 0.0,
    media: Any = 0.0,
    aves_por_caixa: Any = 0,
    nf_numero: str = "",
    nf_preco: Any = 0.0,
    lotes: Any = "",
    data_hora: str = "",
    observacao: str = "",
    payload: Optional[Dict[str, Any]] = None,
    company_id: Optional[int] = None,
) -> None:
    try:
        codigo = str(codigo_programacao or "").strip().upper()
        cur.execute("PRAGMA table_info(roteiro_operacional)")
        cols = {row[1] for row in cur.fetchall() or []}
        if company_id is None and "company_id" in cols and codigo:
            try:
                cur.execute(
                    """
                    SELECT company_id
                      FROM programacoes
                     WHERE UPPER(TRIM(COALESCE(codigo_programacao, '')))=UPPER(TRIM(?))
                       AND company_id IS NOT NULL
                     LIMIT 1
                    """,
                    (codigo,),
                )
                row_company = cur.fetchone()
                if row_company:
                    company_id = int(row_company["company_id"] or 1)
            except Exception:
                company_id = None
        insert_cols = [
            "tipo_evento", "codigo_programacao", "origem", "destino", "motorista_codigo", "motorista_nome",
            "pedido", "cod_cliente", "cliente_nome", "caixas", "kg", "media", "aves_por_caixa",
            "nf_numero", "nf_preco", "lotes", "data_hora", "observacao", "payload_json", "created_at",
        ]
        insert_vals = [
            str(tipo_evento or "").strip().upper(),
            codigo,
            str(origem or "").strip(),
            str(destino or "").strip(),
            str(motorista_codigo or "").strip().upper(),
            str(motorista_nome or "").strip(),
            str(pedido or "").strip(),
            str(cod_cliente or "").strip().upper(),
            str(cliente_nome or "").strip(),
            int(float(caixas or 0)),
            safe_float(kg, 0.0),
            safe_float(media, 0.0),
            int(float(aves_por_caixa or 0)),
            str(nf_numero or "").strip(),
            safe_float(nf_preco, 0.0),
            json.dumps(lotes, ensure_ascii=False) if isinstance(lotes, (dict, list)) else str(lotes or "").strip(),
            str(data_hora or "").strip() or _now_iso(),
            str(observacao or "").strip(),
            json.dumps(payload or {}, ensure_ascii=False, sort_keys=True),
            _now_iso(),
        ]
        if "company_id" in cols:
            insert_cols.append("company_id")
            insert_vals.append(int(company_id or 1))
        placeholders = ", ".join(["?"] * len(insert_cols))
        cur.execute(
            f"INSERT INTO roteiro_operacional ({', '.join(insert_cols)}) VALUES ({placeholders})",
            tuple(insert_vals),
        )
    except Exception:
        logging.debug("Falha ao registrar roteiro operacional.", exc_info=True)


def _parse_snapshot(snapshot_raw: Optional[str]) -> Dict[str, Any]:
    if not snapshot_raw:
        return {}
    try:
        return json.loads(snapshot_raw)
    except Exception:
        return {}


def _list_transferencia_conversoes(cur: sqlite3.Cursor, transferencia_id: str) -> List[Dict[str, Any]]:
    cur.execute(
        """
        SELECT pedido_destino, cod_cliente_destino, qtd, obs, nome_cliente_destino, novo_cliente, criado_em
        FROM transferencias_conversoes
        WHERE transferencia_id=?
        ORDER BY id ASC
        """,
        (transferencia_id,),
    )
    rows = cur.fetchall()
    conv = []
    for row in rows:
        item = {
            "pedido_destino": row["pedido_destino"],
            "cod_cliente_destino": row["cod_cliente_destino"],
            "qtd": row["qtd"],
            "obs": row["obs"],
            "criado_em": row["criado_em"],
        }
        if row["nome_cliente_destino"]:
            item["nome_cliente_destino"] = row["nome_cliente_destino"]
        if row["novo_cliente"]:
            item["novo_cliente"] = bool(row["novo_cliente"])
        conv.append(item)
    return conv


def _serialize_transferencia_row(row: sqlite3.Row, cur: sqlite3.Cursor) -> Dict[str, Any]:
    qtd_total = int(row["qtd_caixas"] or 0)
    qtd_convertida = int(row["qtd_convertida"] or 0)
    cod_cliente = str(row["cod_cliente"] or "").strip().upper()
    pedido = str(row["pedido"] or "").strip().upper()
    snapshot = _parse_snapshot(row["snapshot"])
    carga_sem_cliente = cod_cliente == "TRANSBORDO" and pedido == "TRANSBORDO"
    if isinstance(snapshot, dict):
        carga_sem_cliente = bool(snapshot.get("carga_sem_cliente") or snapshot.get("transbordo") or carga_sem_cliente)
    carga_raiz = ""
    carga_origem_imediata = ""
    if isinstance(snapshot, dict):
        carga_raiz = str(snapshot.get("carga_raiz_programacao") or snapshot.get("carga_origem_programacao") or "").strip().upper()
        carga_origem_imediata = str(snapshot.get("carga_origem_imediata") or "").strip().upper()
    return {
        "id": row["id"],
        "status": row["status"],
        "codigo_origem": row["codigo_origem"],
        "codigo_destino": row["codigo_destino"],
        "cod_cliente": cod_cliente,
        "pedido": pedido,
        "qtd_caixas": qtd_total,
        "snapshot": snapshot,
        "obs": row["obs"],
        "motorista_origem": row["motorista_origem"],
        "motorista_destino": row["motorista_destino"],
        "qtd_convertida": qtd_convertida,
        "qtd_saldo": max(qtd_total - qtd_convertida, 0),
        "carga_sem_cliente": carga_sem_cliente,
        "transbordo": carga_sem_cliente,
        "tipo_transferencia": "TRANSBORDO_CARGA" if carga_sem_cliente else "PEDIDO_CLIENTE",
        "carga_raiz_programacao": carga_raiz,
        "carga_origem_imediata": carga_origem_imediata,
        "criado_em": row["criado_em"],
        "atualizado_em": row["atualizado_em"],
        "conversoes": _list_transferencia_conversoes(cur, row["id"]),
    }


def _fetch_transferencia_by_id(conn: sqlite3.Connection, transferencia_id: str) -> Optional[Dict[str, Any]]:
    cur = conn.cursor()
    cur.execute("SELECT * FROM transferencias WHERE id=? LIMIT 1", (transferencia_id,))
    row = cur.fetchone()
    if not row:
        return None
    return _serialize_transferencia_row(row, cur)


def _list_transferencias_por_destino(
    conn: sqlite3.Connection, codigo_destino: str, status: Optional[str] = None
) -> List[Dict[str, Any]]:
    cur = conn.cursor()
    sql = "SELECT * FROM transferencias WHERE codigo_destino=?"
    params = [codigo_destino]
    if status:
        sql += " AND UPPER(status)=?"
        params.append(status.strip().upper())
    sql += " ORDER BY criado_em DESC"
    cur.execute(sql, tuple(params))
    return [_serialize_transferencia_row(row, cur) for row in cur.fetchall()]


def _list_transferencias_por_origem(
    conn: sqlite3.Connection, codigo_origem: str, status: Optional[str] = None
) -> List[Dict[str, Any]]:
    cur = conn.cursor()
    sql = "SELECT * FROM transferencias WHERE codigo_origem=?"
    params = [codigo_origem]
    if status:
        sql += " AND UPPER(status)=?"
        params.append(status.strip().upper())
    sql += " ORDER BY criado_em DESC"
    cur.execute(sql, tuple(params))
    return [_serialize_transferencia_row(row, cur) for row in cur.fetchall()]


def _resolve_transferencia_destino(item: Dict[str, Any]):
    codigo_destino = (item.get("codigo_destino") or "").strip()
    pedido_dest = (item.get("pedido") or "").strip()
    cod_cli_dest = (item.get("cod_cliente") or "").strip()
    snapshot = item.get("snapshot") or {}
    novo = None
    nome_novo = ""
    status_pedido = None
    if isinstance(snapshot, dict):
        pedido_dest = (
            snapshot.get("pedido_destino")
            or snapshot.get("pedido")
            or pedido_dest
        )
        cod_cli_dest = (
            snapshot.get("cod_cliente_destino")
            or snapshot.get("cod_cliente")
            or cod_cli_dest
        )
        novo_candidate = snapshot.get("novo_cliente")
        if isinstance(novo_candidate, dict):
            novo = novo_candidate
            nome_novo = (
                novo.get("nome_cliente_destino")
                or novo.get("nome_cliente")
                or nome_novo
            )
            status_pedido = (
                (novo.get("status_pedido") or "").strip().upper()
            )
            if not status_pedido:
                status_pedido = None
        if not nome_novo:
            nome_novo = (
                snapshot.get("nome_cliente_destino")
                or snapshot.get("nome_cliente")
                or nome_novo
            )
    return (
        codigo_destino,
        pedido_dest,
        cod_cli_dest,
        novo,
        nome_novo.strip(),
        status_pedido,
    )


def _resolve_carga_raiz_programacao(cur: sqlite3.Cursor, codigo_programacao: str) -> str:
    codigo = (codigo_programacao or "").strip().upper()
    if not codigo:
        return ""
    try:
        cur.execute(
            """
            SELECT codigo_programacao, tipo_estimativa, operacao_tipo, transbordo_grupo
            FROM programacoes
            WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        )
        row = cur.fetchone()
        if row:
            operacao = str(row["operacao_tipo"] or "").strip().upper().replace("-", "_").replace(" ", "_")
            tipo = str(row["tipo_estimativa"] or "").strip().upper()
            grupo = str(row["transbordo_grupo"] or "").strip().upper()
            if operacao == "TRANSBORDO" or tipo == "CX":
                return grupo or codigo
    except Exception:
        pass
    try:
        cur.execute(
            """
            SELECT codigo_origem, snapshot, atualizado_em, criado_em
            FROM transferencias
            WHERE UPPER(TRIM(COALESCE(codigo_destino,'')))=UPPER(TRIM(?))
              AND UPPER(TRIM(COALESCE(status,''))) NOT IN ('CANCELADA','CANCELADO','RECUSADA','RECUSADO')
            ORDER BY COALESCE(atualizado_em, criado_em, '') DESC
            LIMIT 1
            """,
            (codigo,),
        )
        tr = cur.fetchone()
        if tr:
            snapshot = _parse_snapshot(tr["snapshot"])
            raiz = str(snapshot.get("carga_raiz_programacao") or snapshot.get("carga_origem_programacao") or "").strip().upper()
            return raiz or str(tr["codigo_origem"] or "").strip().upper()
    except Exception:
        pass
    return codigo


def _upsert_item_destino_transferencia(
    cur: sqlite3.Cursor,
    *,
    codigo_destino: str,
    cod_cliente: str,
    pedido: str,
    nome_cliente: str,
    qtd_caixas: int,
    obs: str,
    alterado_por: str,
    carga_raiz_programacao: str = "",
    carga_origem_imediata: str = "",
    transferencia_origem_id: str = "",
) -> None:
    codigo_destino = (codigo_destino or "").strip()
    cod_cliente = (cod_cliente or "").strip().upper()
    pedido = (pedido or "").strip().upper()
    nome_cliente = (nome_cliente or cod_cliente or "CLIENTE TRANSBORDO").strip().upper()
    carga_raiz_programacao = (carga_raiz_programacao or "").strip().upper()
    carga_origem_imediata = (carga_origem_imediata or "").strip().upper()
    transferencia_origem_id = (transferencia_origem_id or "").strip()
    qtd_caixas = max(int(qtd_caixas or 0), 0)
    if not codigo_destino or not cod_cliente or not pedido or qtd_caixas <= 0:
        return

    cur.execute("PRAGMA table_info(programacao_itens)")
    cols_itens = {row[1] for row in (cur.fetchall() or [])}
    has_pedido_col = "pedido" in cols_itens
    existing = None
    if has_pedido_col:
        cur.execute(
            """
            SELECT rowid AS rid, qnt_caixas, caixas_atual
            FROM programacao_itens
            WHERE codigo_programacao=?
              AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
              AND COALESCE(TRIM(pedido),'')=COALESCE(TRIM(?),'')
            LIMIT 1
            """,
            (codigo_destino, cod_cliente, pedido),
        )
    else:
        cur.execute(
            """
            SELECT rowid AS rid, qnt_caixas, caixas_atual
            FROM programacao_itens
            WHERE codigo_programacao=?
              AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
            LIMIT 1
            """,
            (codigo_destino, cod_cliente),
        )
    existing = cur.fetchone()
    now = _now_iso()
    detalhe = f"Conversao de transbordo: +{qtd_caixas} cx"
    if obs:
        detalhe = f"{detalhe}. Obs: {obs}"
    if existing:
        nova_qtd = int(existing["qnt_caixas"] or 0) + qtd_caixas
        sets = []
        params: List[Any] = []
        if "qnt_caixas" in cols_itens:
            sets.append("qnt_caixas=?"); params.append(nova_qtd)
        if "caixas_atual" in cols_itens:
            atual = existing["caixas_atual"]
            atual_int = int(atual or existing["qnt_caixas"] or 0)
            sets.append("caixas_atual=?"); params.append(atual_int + qtd_caixas)
        if "status_pedido" in cols_itens:
            sets.append("status_pedido=?"); params.append("PENDENTE")
        if "alteracao_tipo" in cols_itens:
            sets.append("alteracao_tipo=?"); params.append("TRANSBORDO")
        if "alteracao_detalhe" in cols_itens:
            sets.append("alteracao_detalhe=?"); params.append(detalhe)
        if "alterado_em" in cols_itens:
            sets.append("alterado_em=?"); params.append(now)
        if "alterado_por" in cols_itens:
            sets.append("alterado_por=?"); params.append(alterado_por or "SISTEMA")
        if "carga_raiz_programacao" in cols_itens and carga_raiz_programacao:
            sets.append("carga_raiz_programacao=COALESCE(NULLIF(carga_raiz_programacao, ''), ?)")
            params.append(carga_raiz_programacao)
        if "carga_origem_imediata" in cols_itens and carga_origem_imediata:
            sets.append("carga_origem_imediata=?"); params.append(carga_origem_imediata)
        if "transferencia_origem_id" in cols_itens and transferencia_origem_id:
            sets.append("transferencia_origem_id=?"); params.append(transferencia_origem_id)
        if sets:
            params.append(existing["rid"])
            cur.execute(f"UPDATE programacao_itens SET {', '.join(sets)} WHERE rowid=?", tuple(params))
    else:
        values: Dict[str, Any] = {
            "codigo_programacao": codigo_destino,
            "cod_cliente": cod_cliente,
            "nome_cliente": nome_cliente,
            "produto": "TRANSBORDO",
            "endereco": "",
            "qnt_caixas": qtd_caixas,
            "kg": 0,
            "preco": 0,
            "vendedor": "",
            "pedido": pedido,
            "observacao": detalhe,
            "status_pedido": "PENDENTE",
            "alteracao_tipo": "TRANSBORDO",
            "alteracao_detalhe": detalhe,
            "caixas_atual": qtd_caixas,
            "preco_atual": 0,
            "alterado_em": now,
            "alterado_por": alterado_por or "SISTEMA",
            "carga_raiz_programacao": carga_raiz_programacao,
            "carga_origem_imediata": carga_origem_imediata,
            "transferencia_origem_id": transferencia_origem_id,
        }
        keys = [key for key in values if key in cols_itens]
        if keys:
            cur.execute(
                f"INSERT INTO programacao_itens ({', '.join(keys)}) VALUES ({', '.join(['?'] * len(keys))})",
                tuple(values[key] for key in keys),
            )

    cur.execute(
        """
        UPDATE programacao_itens_controle
           SET status_pedido='PENDENTE',
               alteracao_tipo='TRANSBORDO',
               alteracao_detalhe=?,
               caixas_atual=COALESCE(caixas_atual, 0) + ?,
               alterado_em=?,
               alterado_por=?,
               updated_at=datetime('now')
         WHERE codigo_programacao=?
           AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
           AND COALESCE(TRIM(pedido),'')=COALESCE(TRIM(?),'')
        """,
        (detalhe, qtd_caixas, now, alterado_por or "SISTEMA", codigo_destino, cod_cliente, pedido),
    )
    if cur.rowcount == 0:
        cur.execute(
            """
            INSERT INTO programacao_itens_controle
                (codigo_programacao, cod_cliente, pedido, status_pedido,
                 alteracao_tipo, alteracao_detalhe, caixas_atual,
                 alterado_em, alterado_por, updated_at)
            VALUES (?, ?, ?, 'PENDENTE', 'TRANSBORDO', ?, ?, ?, ?, datetime('now'))
            """,
            (codigo_destino, cod_cliente, pedido, detalhe, qtd_caixas, now, alterado_por or "SISTEMA"),
        )


class TransferenciaCreateIn(BaseModel):
    codigo_destino: str
    pedido: Optional[str] = ""
    cod_cliente: Optional[str] = ""
    qtd_caixas: int
    snapshot: Optional[Dict[str, Any]] = None
    obs: Optional[str] = None


def _recalcular_origem_transferencia(
    cur: sqlite3.Cursor,
    codigo_origem: str,
    cod_cliente: str,
    pedido_ref: str,
    alterado_por: str,
    evento: str,
) -> None:
    codigo_origem = (codigo_origem or "").strip()
    cod_cliente = (cod_cliente or "").strip()
    pedido_norm = _norm_pedido_key(pedido_ref)
    if not codigo_origem or not cod_cliente:
        return

    cur.execute("PRAGMA table_info(programacao_itens)")
    cols_itens = {row[1] for row in (cur.fetchall() or [])}
    has_pedido_col = "pedido" in cols_itens

    base = None
    if has_pedido_col:
        cur.execute(
            """
            SELECT rowid AS rid, codigo_programacao, cod_cliente, pedido, qnt_caixas
            FROM programacao_itens
            WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
            """,
            (codigo_origem, cod_cliente),
        )
        cands = cur.fetchall() or []
        for r in cands:
            if _norm_pedido_key(r["pedido"]) == pedido_norm:
                base = r
                break
    else:
        cur.execute(
            """
            SELECT rowid AS rid, codigo_programacao, cod_cliente, NULL AS pedido, qnt_caixas
            FROM programacao_itens
            WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
            LIMIT 1
            """,
            (codigo_origem, cod_cliente),
        )
        base = cur.fetchone()

    if not base:
        return

    # Soma transferências ainda ativas (pendentes/aceitas) para este pedido.
    qtd_ativa = 0
    cur.execute(
        """
        SELECT pedido, qtd_caixas
        FROM transferencias
        WHERE codigo_origem=?
          AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
          AND UPPER(TRIM(COALESCE(status, ''))) IN ('PENDENTE', 'ACEITA')
        """,
        (codigo_origem, cod_cliente),
    )
    for tr in (cur.fetchall() or []):
        if _norm_pedido_key(tr["pedido"]) == pedido_norm:
            qtd_ativa += int(tr["qtd_caixas"] or 0)

    base_qtd = int(base["qnt_caixas"] or 0)
    novo_caixas = max(base_qtd - qtd_ativa, 0)
    if qtd_ativa <= 0:
        novo_status = "PENDENTE"
        detalhe = f"Transferencia {evento}: saldo restaurado ({novo_caixas} cx)."
    else:
        novo_status = "CANCELADO" if novo_caixas == 0 else "ALTERADO"
        detalhe = f"Transferencia {evento}: saldo em transferencia {qtd_ativa} cx."

    now = _now_iso()
    sets = []
    params: List[Any] = []
    if "status_pedido" in cols_itens:
        sets.append("status_pedido=?"); params.append(novo_status)
    if "alteracao_tipo" in cols_itens:
        sets.append("alteracao_tipo=?"); params.append("QUANTIDADE")
    if "alteracao_detalhe" in cols_itens:
        sets.append("alteracao_detalhe=?"); params.append(detalhe)
    if "caixas_atual" in cols_itens:
        sets.append("caixas_atual=?"); params.append(novo_caixas)
    if "alterado_em" in cols_itens:
        sets.append("alterado_em=?"); params.append(now)
    if "alterado_por" in cols_itens:
        sets.append("alterado_por=?"); params.append(alterado_por or "SISTEMA")

    if sets:
        params.append(base["rid"])
        cur.execute(f"UPDATE programacao_itens SET {', '.join(sets)} WHERE rowid=?", tuple(params))

    pedido_db = base["pedido"] if has_pedido_col else None
    cur.execute(
        """
        UPDATE programacao_itens_controle
           SET status_pedido=?,
               alteracao_tipo='QUANTIDADE',
               alteracao_detalhe=?,
               caixas_atual=?,
               alterado_em=?,
               alterado_por=?,
               updated_at=datetime('now')
         WHERE codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))
           AND COALESCE(TRIM(pedido), '')=COALESCE(TRIM(?), '')
        """,
        (novo_status, detalhe, novo_caixas, now, (alterado_por or "SISTEMA"), codigo_origem, cod_cliente, pedido_db),
    )
    if cur.rowcount == 0:
        cur.execute(
            """
            INSERT INTO programacao_itens_controle
                (codigo_programacao, cod_cliente, pedido, status_pedido,
                 alteracao_tipo, alteracao_detalhe, caixas_atual,
                 alterado_em, alterado_por, updated_at)
            VALUES (?, ?, ?, ?, 'QUANTIDADE', ?, ?, ?, ?, datetime('now'))
            """,
            (
                codigo_origem,
                cod_cliente,
                pedido_db,
                novo_status,
                detalhe,
                novo_caixas,
                now,
                (alterado_por or "SISTEMA"),
            ),
        )


def _rota_pertence_ao_motorista(codigo_programacao: str, motorista: Dict[str, Any]) -> bool:
    with get_conn() as conn:
        cur = conn.cursor()
        return _fetch_programacao_owned(cur, (codigo_programacao or "").strip(), motorista, "p.id") is not None


def _has_pending_substituicao(cur: sqlite3.Cursor, codigo_programacao: str) -> bool:
    cur.execute(
        """
        SELECT COUNT(*)
        FROM rota_substituicoes
        WHERE codigo_programacao=?
          AND UPPER(TRIM(COALESCE(status,'')))='PENDENTE_ACEITE'
        """,
        ((codigo_programacao or "").strip(),),
    )
    return int((cur.fetchone() or [0])[0] or 0) > 0


def _list_substituicoes_por_rota(cur: sqlite3.Cursor, codigo_programacao: str, limit: int = 20) -> List[Dict[str, Any]]:
    cur.execute(
        """
        SELECT *
        FROM rota_substituicoes
        WHERE codigo_programacao=?
        ORDER BY solicitado_em DESC
        LIMIT ?
        """,
        ((codigo_programacao or "").strip(), int(limit)),
    )
    return [_serialize_substituicao_row(r) for r in (cur.fetchall() or [])]


def _get_motorista_by_codigo(cur: sqlite3.Cursor, codigo: str) -> Optional[sqlite3.Row]:
    cod = (codigo or "").strip().upper()
    if not cod:
        return None
    cur.execute(
        "SELECT id, nome, codigo FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1",
        (cod,),
    )
    return cur.fetchone()


def _registrar_amostra_localizacao_cliente(
    cur: sqlite3.Cursor,
    *,
    cod_cliente: str,
    codigo_programacao: str,
    pedido: Optional[str],
    lat_evento: Optional[float],
    lon_evento: Optional[float],
    endereco_evento: Optional[str],
    cidade_evento: Optional[str],
    bairro_evento: Optional[str],
    status_pedido: Optional[str],
    motorista_codigo: Optional[str],
    motorista_nome: Optional[str],
    origem: str = "APP",
    company_id: Optional[int] = None,
) -> None:
    has_geo = lat_evento not in (None, "") and lon_evento not in (None, "")
    has_addr = any(str(v or "").strip() for v in (endereco_evento, cidade_evento, bairro_evento))
    if not has_geo and not has_addr:
        return
    cur.execute("PRAGMA table_info(cliente_localizacao_amostras)")
    cols = {row[1] for row in cur.fetchall() or []}
    insert_cols = [
        "cod_cliente", "codigo_programacao", "pedido", "latitude", "longitude", "endereco", "cidade", "bairro",
        "status_pedido", "motorista_codigo", "motorista_nome", "origem", "registrado_em",
    ]
    insert_vals = [
        (cod_cliente or "").strip().upper(),
        (codigo_programacao or "").strip().upper(),
        (pedido or "").strip(),
        float(lat_evento) if lat_evento not in (None, "") else None,
        float(lon_evento) if lon_evento not in (None, "") else None,
        str(endereco_evento or "").strip().upper(),
        str(cidade_evento or "").strip().upper(),
        str(bairro_evento or "").strip().upper(),
        str(status_pedido or "").strip().upper(),
        str(motorista_codigo or "").strip().upper(),
        str(motorista_nome or "").strip().upper(),
        str(origem or "APP").strip().upper(),
        _now_iso(),
    ]
    if "company_id" in cols:
        insert_cols.append("company_id")
        insert_vals.append(int(company_id or 1))
    placeholders = ", ".join(["?"] * len(insert_cols))
    cur.execute(
        f"INSERT INTO cliente_localizacao_amostras ({', '.join(insert_cols)}) VALUES ({placeholders})",
        tuple(insert_vals),
    )


def _serialize_substituicao_row(row: sqlite3.Row) -> Dict[str, Any]:
    return {
        "id": row["id"],
        "codigo_programacao": row["codigo_programacao"],
        "status": row["status"],
        "motivo": row["motivo"],
        "km_evento": row["km_evento"],
        "lat_evento": row["lat_evento"],
        "lon_evento": row["lon_evento"],
        "snapshot_json": row["snapshot_json"],
        "origem_motorista_nome": row["origem_motorista_nome"],
        "origem_motorista_codigo": row["origem_motorista_codigo"],
        "origem_motorista_id": row["origem_motorista_id"],
        "origem_veiculo": row["origem_veiculo"],
        "destino_motorista_nome": row["destino_motorista_nome"],
        "destino_motorista_codigo": row["destino_motorista_codigo"],
        "destino_motorista_id": row["destino_motorista_id"],
        "destino_veiculo": row["destino_veiculo"],
        "solicitado_em": row["solicitado_em"],
        "aceito_em": row["aceito_em"],
        "atualizado_em": row["atualizado_em"],
    }


@app.get("/substituicoes/pendentes")
def listar_substituicoes_pendentes(m=Depends(get_current_motorista)):
    codigo_mot = (m.get("codigo") or "").strip().upper()
    nome_mot = (m.get("nome") or "").strip().upper()
    mot_id = int(m.get("id") or 0)
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT *
            FROM rota_substituicoes
            WHERE (
                UPPER(TRIM(COALESCE(destino_motorista_codigo,'')))=?
                OR destino_motorista_id=?
                OR UPPER(TRIM(COALESCE(destino_motorista_nome,'')))=?
            )
              AND UPPER(TRIM(COALESCE(status,''))) IN ('PENDENTE_ACEITE', 'PENDENTE', 'PENDENTE ACEITE')
            ORDER BY solicitado_em DESC
            """,
            (codigo_mot, mot_id, nome_mot),
        )
        return [_serialize_substituicao_row(r) for r in (cur.fetchall() or [])]


@app.get("/cadastros/motoristas")
def listar_cad_motoristas(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {str(r[1]) for r in (cur.fetchall() or [])}
        if "nome" not in cols:
            return []
        sel = ["id", "nome"]
        if "codigo" in cols:
            sel.append("codigo")
        if "status" in cols:
            sel.append("status")
        where = "WHERE TRIM(COALESCE(nome, '')) <> ''"
        if "status" in cols:
            where += " AND UPPER(TRIM(COALESCE(status, 'ATIVO'))) = 'ATIVO'"
        recursos_ocupados = _recursos_ocupados_programacoes_ativas(cur)
        cur.execute(
            f"""
            SELECT {', '.join(sel)}
            FROM motoristas
            {where}
            ORDER BY nome
            """
        )
        out = []
        for r in (cur.fetchall() or []):
            nome = str(r["nome"] or "").strip().upper()
            codigo = str(r["codigo"] or "").strip().upper() if "codigo" in r.keys() else ""
            if codigo in recursos_ocupados["motoristas_codigos"] or nome in recursos_ocupados["motoristas_nomes"]:
                continue
            out.append(
                {
                    "id": int(r["id"]) if r["id"] is not None else None,
                    "nome": nome,
                    "codigo": codigo,
                }
            )
        return out


@app.get("/cadastros/ajudantes")
def listar_cad_ajudantes(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(ajudantes)")
        cols = {r[1] for r in (cur.fetchall() or [])}

        if "nome" not in cols:
            return []

        sel = ["id", "nome"]
        if "sobrenome" in cols:
            sel.append("sobrenome")
        if "telefone" in cols:
            sel.append("telefone")
        if "status" in cols:
            sel.append("status")

        where = "WHERE TRIM(COALESCE(nome, '')) <> ''"
        if "status" in cols:
            where += " AND UPPER(TRIM(COALESCE(status, 'ATIVO'))) = 'ATIVO'"

        ajudantes_ocupados = _recursos_ocupados_programacoes_ativas(cur)["ajudantes"]

        cur.execute(
            f"""
            SELECT {', '.join(sel)}
            FROM ajudantes
            {where}
            ORDER BY nome
            """
        )
        out = []
        for r in (cur.fetchall() or []):
            nome = str(r["nome"] or "").strip().upper()
            sobrenome = (
                str(r["sobrenome"] or "").strip().upper()
                if "sobrenome" in r.keys()
                else ""
            )
            nomes_check = {
                nome,
                f"{nome} {sobrenome}".strip(),
            }
            if nomes_check.intersection(ajudantes_ocupados):
                continue
            item = {
                "id": int(r["id"]) if "id" in r.keys() and r["id"] is not None else None,
                "nome": nome,
            }
            if "sobrenome" in r.keys():
                item["sobrenome"] = sobrenome
            if "telefone" in r.keys():
                item["telefone"] = str(r["telefone"] or "").strip()
            if "status" in r.keys():
                item["status"] = str(r["status"] or "ATIVO").strip().upper()
            out.append(item)
        return out


def _recursos_ocupados_programacoes_ativas(cur: sqlite3.Cursor) -> dict[str, set[str]]:
    ocupados = {
        "motoristas_codigos": set(),
        "motoristas_nomes": set(),
        "veiculos": set(),
        "ajudantes": set(),
    }

    def resolve_ajudantes_ocupados(raw: Any) -> list[str]:
        tokens = _split_people_tokens(_clean_text(raw))
        if not tokens:
            return []
        resolved: list[str] = []
        for token_raw in tokens:
            token = _clean_text(token_raw).upper()
            if not token:
                continue
            found = ""
            try:
                cur.execute(
                    """
                    SELECT nome, sobrenome
                    FROM ajudantes
                    WHERE UPPER(TRIM(CAST(id AS TEXT))) = UPPER(TRIM(?))
                    LIMIT 1
                    """,
                    (token,),
                )
                row_aj = cur.fetchone()
                if row_aj:
                    nome = _clean_text(row_aj["nome"] if hasattr(row_aj, "keys") else row_aj[0]).upper()
                    sobrenome = _clean_text(row_aj["sobrenome"] if hasattr(row_aj, "keys") else row_aj[1]).upper()
                    found = f"{nome} {sobrenome}".strip()
            except Exception:
                found = ""
            if not found:
                try:
                    cur.execute("PRAGMA table_info(equipes)")
                    cols_eq = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    cand_cols = [c for c in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2") if c in cols_eq]
                    if cand_cols and "codigo" in cols_eq:
                        cur.execute(
                            f"""
                            SELECT {', '.join(cand_cols)}
                            FROM equipes
                            WHERE UPPER(TRIM(codigo)) = UPPER(TRIM(?))
                               OR UPPER(TRIM(CAST(id AS TEXT))) = UPPER(TRIM(?))
                            LIMIT 1
                            """,
                            (token, token),
                        )
                        row_eq = cur.fetchone()
                        if row_eq:
                            for col in cand_cols:
                                value = row_eq[col] if col in row_eq.keys() else None
                                for item in resolve_ajudantes_ocupados(value):
                                    if item and item not in resolved:
                                        resolved.append(item)
                            continue
                except Exception:
                    pass
            value = found or token
            if value and value not in resolved:
                resolved.append(value)
        return resolved

    try:
        cur.execute("PRAGMA table_info(programacoes)")
        cols = {str(r[1]) for r in (cur.fetchall() or [])}
        if not cols:
            return ocupados

        select_cols = [
            c
            for c in (
                "motorista_codigo",
                "codigo_motorista",
                "motorista",
                "veiculo",
                "ajudante1",
                "ajudante_1",
                "ajudante2",
                "ajudante_2",
                "equipe",
            )
            if c in cols
        ]
        if not select_cols:
            return ocupados

        where = []
        if "status" in cols:
            where.append(
                "UPPER(TRIM(COALESCE(status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')"
            )
        if "status_operacional" in cols:
            where.append(
                "UPPER(TRIM(COALESCE(status_operacional,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')"
            )
        if "finalizada_no_app" in cols:
            where.append("COALESCE(finalizada_no_app,0)=0")
        if "prestacao_status" in cols:
            where.append("UPPER(TRIM(COALESCE(prestacao_status,'PENDENTE'))) <> 'FECHADA'")
        where_sql = "WHERE " + " AND ".join(where) if where else ""

        cur.execute(
            f"""
            SELECT {', '.join(select_cols)}
            FROM programacoes
            {where_sql}
            """
        )
        for row in (cur.fetchall() or []):
            motorista_codigo = ""
            for col in ("motorista_codigo", "codigo_motorista"):
                if col in row.keys():
                    motorista_codigo = str(row[col] or "").strip().upper()
                    if motorista_codigo:
                        ocupados["motoristas_codigos"].add(motorista_codigo)
            if "motorista" in row.keys():
                motorista_nome = str(row["motorista"] or "").strip().upper()
                if motorista_nome:
                    ocupados["motoristas_nomes"].add(motorista_nome)
                    if not motorista_codigo:
                        ocupados["motoristas_codigos"].add(motorista_nome)
            if "veiculo" in row.keys():
                veiculo = str(row["veiculo"] or "").strip().upper()
                if veiculo:
                    ocupados["veiculos"].add(veiculo)
            for col in ("ajudante1", "ajudante_1", "ajudante2", "ajudante_2", "equipe"):
                if col not in row.keys():
                    continue
                for nome in resolve_ajudantes_ocupados(row[col]):
                    if nome and nome != "-":
                        ocupados["ajudantes"].add(nome)
        return ocupados
    except Exception:
        logging.exception("Falha ao calcular recursos ocupados em programacoes ativas")
        return ocupados


def _ajudantes_ocupados_programacoes_ativas(cur: sqlite3.Cursor) -> set[str]:
    return _recursos_ocupados_programacoes_ativas(cur)["ajudantes"]


@app.get("/cadastros/veiculos")
def listar_cad_veiculos(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(veiculos)")
        cols = {r[1] for r in (cur.fetchall() or [])}

        if "placa" not in cols:
            return []

        sel = ["placa"]
        if "modelo" in cols:
            sel.append("modelo")
        if "capacidade_cx" in cols:
            sel.append("capacidade_cx")
        elif "capacidade" in cols:
            sel.append("capacidade")
        if "status" in cols:
            sel.append("status")
        km_col = next(
            (
                c
                for c in (
                    "ultimo_km",
                    "km_atual",
                    "km_veiculo",
                    "quilometragem_atual",
                    "quilometragem",
                    "odometro_atual",
                    "odometro",
                    "hodometro_atual",
                    "hodometro",
                )
                if c in cols
            ),
            None,
        )
        if km_col:
            sel.append(km_col)

        where = ""
        if "status" in cols:
            where = "WHERE UPPER(TRIM(COALESCE(status, 'ATIVO'))) = 'ATIVO'"
        veiculos_ocupados = _recursos_ocupados_programacoes_ativas(cur)["veiculos"]
        cur.execute(f"SELECT {', '.join(sel)} FROM veiculos {where} ORDER BY placa")
        out = []
        for r in (cur.fetchall() or []):
            placa = str(r["placa"] or "").strip().upper()
            if not placa or placa in veiculos_ocupados:
                continue
            d = {"placa": placa}
            if "modelo" in r.keys():
                d["modelo"] = str(r["modelo"] or "").strip().upper()
            if "capacidade_cx" in r.keys():
                d["capacidade_cx"] = r["capacidade_cx"]
            elif "capacidade" in r.keys():
                d["capacidade_cx"] = r["capacidade"]
            if km_col and km_col in r.keys():
                d["ultimo_km"] = r[km_col]
                d["km_atual"] = r[km_col]
            if "status" in r.keys():
                d["status"] = str(r["status"] or "ATIVO").strip().upper()
            out.append(d)
        return out


@app.post("/rotas/{codigo_programacao}/substituicoes/solicitar")
def solicitar_substituicao_rota(
    codigo_programacao: str,
    payload: SubstituicaoRotaIn,
    m=Depends(get_current_motorista),
):
    codigo = (codigo_programacao or "").strip()
    if not codigo:
        raise HTTPException(status_code=400, detail="Codigo da programacao obrigatorio.")

    destino_cod = (payload.motorista_destino_codigo or "").strip().upper()
    motivo = (payload.motivo or "").strip()
    destino_veic = (payload.veiculo_destino or "").strip().upper()

    if not destino_cod:
        raise HTTPException(status_code=400, detail="Codigo do motorista destino obrigatorio.")
    if not motivo:
        raise HTTPException(status_code=400, detail="Motivo obrigatorio.")

    with get_conn() as conn:
        cur = conn.cursor()
        pr = _fetch_programacao_owned(
            cur,
            codigo,
            m,
            "p.id, p.status, p.motorista, p.veiculo, p.codigo_programacao",
        )
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nao encontrada para este motorista.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS"):
            raise HTTPException(
                status_code=409,
                detail=f"Substituicao permitida apenas com rota em andamento (status atual: {status_atual or 'N/D'}).",
            )

        destino = _get_motorista_by_codigo(cur, destino_cod)
        if not destino:
            raise HTTPException(status_code=404, detail="Motorista destino nao encontrado.")
        if int(destino["id"]) == int(m["id"]):
            raise HTTPException(status_code=400, detail="Motorista destino deve ser diferente do atual.")

        cur.execute(
            """
            SELECT COUNT(*)
            FROM rota_substituicoes
            WHERE codigo_programacao=?
              AND UPPER(TRIM(COALESCE(status,'')))='PENDENTE_ACEITE'
            """,
            (codigo,),
        )
        pend = int((cur.fetchone() or [0])[0] or 0)
        if pend > 0:
            raise HTTPException(status_code=409, detail="Ja existe substituicao pendente para esta rota.")

        snapshot = {
            "status": status_atual,
            "motorista": pr["motorista"],
            "veiculo": pr["veiculo"],
            "solicitado_por": m.get("nome"),
            "solicitado_por_codigo": m.get("codigo"),
        }
        sid = str(uuid4())
        now = datetime.now().isoformat(timespec="seconds")
        cur.execute(
            """
            INSERT INTO rota_substituicoes (
                id, codigo_programacao, status, motivo,
                km_evento, lat_evento, lon_evento, snapshot_json,
                origem_motorista_nome, origem_motorista_codigo, origem_motorista_id, origem_veiculo,
                destino_motorista_nome, destino_motorista_codigo, destino_motorista_id, destino_veiculo,
                solicitado_em, atualizado_em
            ) VALUES (?, ?, 'PENDENTE_ACEITE', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                sid,
                codigo,
                motivo,
                payload.km_evento,
                payload.lat_evento,
                payload.lon_evento,
                json.dumps(snapshot, ensure_ascii=False),
                (m.get("nome") or "").strip().upper(),
                (m.get("codigo") or "").strip().upper(),
                int(m["id"]),
                str(pr["veiculo"] or "").strip().upper(),
                str(destino["nome"] or "").strip().upper(),
                str(destino["codigo"] or "").strip().upper(),
                int(destino["id"]),
                destino_veic or str(pr["veiculo"] or "").strip().upper(),
                now,
                now,
            ),
        )
        conn.commit()

        cur.execute("SELECT * FROM rota_substituicoes WHERE id=? LIMIT 1", (sid,))
        row = cur.fetchone()
        return _serialize_substituicao_row(row)


@app.post("/substituicoes/{substituicao_id}/aceitar")
def aceitar_substituicao_rota(
    substituicao_id: str,
    payload: Optional[SubstituicaoRotaDecisaoIn] = None,
    m=Depends(get_current_motorista),
):
    sid = (substituicao_id or "").strip()
    if not sid:
        raise HTTPException(status_code=400, detail="ID da substituicao invalido.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM rota_substituicoes WHERE id=? LIMIT 1", (sid,))
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Substituicao nao encontrada.")

        item = _serialize_substituicao_row(row)
        st = (item.get("status") or "").strip().upper()
        if st != "PENDENTE_ACEITE":
            raise HTTPException(status_code=409, detail=f"Substituicao nao esta pendente (status={st}).")

        cod_dest = (item.get("destino_motorista_codigo") or "").strip().upper()
        if cod_dest != (m.get("codigo") or "").strip().upper():
            raise HTTPException(status_code=403, detail="Apenas o motorista destino pode aceitar esta substituicao.")

        pr = _fetch_programacao_owned(
            cur,
            item["codigo_programacao"],
            {"id": item["origem_motorista_id"], "codigo": item["origem_motorista_codigo"], "nome": item["origem_motorista_nome"]},
            "p.id, p.status",
        )
        # fallback: se ownership antigo nao bater, tenta por codigo direto
        if not pr:
            cur.execute(
                "SELECT id, status FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                (item["codigo_programacao"],),
            )
            pr = cur.fetchone()
        if not pr:
            raise HTTPException(status_code=404, detail="Programacao da substituicao nao encontrada.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_atual}).")

        cur.execute("PRAGMA table_info(programacoes)")
        cols = {r[1] for r in (cur.fetchall() or [])}
        sets = []
        params: List[Any] = []

        if "motorista" in cols:
            sets.append("motorista=?")
            params.append((m.get("nome") or "").strip().upper())
        if "motorista_id" in cols:
            sets.append("motorista_id=?")
            params.append(int(m["id"]))
        if "motorista_codigo" in cols:
            sets.append("motorista_codigo=?")
            params.append((m.get("codigo") or "").strip().upper())
        if "codigo_motorista" in cols:
            sets.append("codigo_motorista=?")
            params.append((m.get("codigo") or "").strip().upper())
        if "veiculo" in cols:
            sets.append("veiculo=?")
            params.append((item.get("destino_veiculo") or "").strip().upper())
        if "status" in cols:
            sets.append("status=?")
            params.append("EM_ROTA")

        if not sets:
            return {"ok": True, "status": status_atual, "warning": "Nenhuma coluna compatível encontrada para atualizar."}

        params.append(int(pr["id"]))
        cur.execute(f"UPDATE programacoes SET {', '.join(sets)} WHERE id=?", tuple(params))

        now = datetime.now().isoformat(timespec="seconds")
        motivo_aceite = (payload.motivo if payload else None) or ""
        cur.execute(
            """
            UPDATE rota_substituicoes
               SET status='ACEITA',
                   aceito_em=?,
                   atualizado_em=?,
                   motivo=CASE
                     WHEN ? <> '' THEN TRIM(COALESCE(motivo,'') || ' | ACEITE: ' || ?)
                     ELSE motivo
                   END
             WHERE id=?
            """,
            (now, now, motivo_aceite, motivo_aceite, sid),
        )
        conn.commit()

        cur.execute("SELECT * FROM rota_substituicoes WHERE id=? LIMIT 1", (sid,))
        out = cur.fetchone()
        return _serialize_substituicao_row(out)


@app.post("/substituicoes/{substituicao_id}/recusar")
def recusar_substituicao_rota(
    substituicao_id: str,
    payload: Optional[SubstituicaoRotaDecisaoIn] = None,
    m=Depends(get_current_motorista),
):
    sid = (substituicao_id or "").strip()
    if not sid:
        raise HTTPException(status_code=400, detail="ID da substituicao invalido.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM rota_substituicoes WHERE id=? LIMIT 1", (sid,))
        row = cur.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Substituicao nao encontrada.")

        item = _serialize_substituicao_row(row)
        st = (item.get("status") or "").strip().upper()
        if st != "PENDENTE_ACEITE":
            raise HTTPException(status_code=409, detail=f"Substituicao nao esta pendente (status={st}).")

        cod_dest = (item.get("destino_motorista_codigo") or "").strip().upper()
        if cod_dest != (m.get("codigo") or "").strip().upper():
            raise HTTPException(status_code=403, detail="Apenas o motorista destino pode recusar esta substituicao.")

        now = datetime.now().isoformat(timespec="seconds")
        motivo_recusa = (payload.motivo if payload else None) or ""
        cur.execute(
            """
            UPDATE rota_substituicoes
               SET status='RECUSADA',
                   atualizado_em=?,
                   motivo=CASE
                     WHEN ? <> '' THEN TRIM(COALESCE(motivo,'') || ' | RECUSA: ' || ?)
                     ELSE motivo
                   END
             WHERE id=?
            """,
            (now, motivo_recusa, motivo_recusa, sid),
        )
        conn.commit()

        cur.execute("SELECT * FROM rota_substituicoes WHERE id=? LIMIT 1", (sid,))
        out = cur.fetchone()
        return _serialize_substituicao_row(out)


# âœ… CRIAR TRANSFERÊNCIA (origem envia)
@app.post("/rotas/{codigo_programacao}/transferencias")
def criar_transferencia(
    codigo_programacao: str,
    payload: TransferenciaCreateIn,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo_origem = (codigo_programacao or "").strip()
    codigo_destino = (payload.codigo_destino or "").strip()

    if not codigo_origem:
        raise HTTPException(status_code=400, detail="Código de origem inválido.")
    if not codigo_destino:
        raise HTTPException(status_code=400, detail="Código de destino inválido.")
    if codigo_destino == codigo_origem:
        raise HTTPException(status_code=400, detail="Destino não pode ser igual à origem.")

    if not _rota_pertence_ao_motorista(codigo_origem, m):
        raise HTTPException(status_code=403, detail="Rota de origem não pertence ao motorista logado.")

    pedido = (payload.pedido or "").strip()
    cod_cliente = (payload.cod_cliente or "").strip()
    qtd = int(payload.qtd_caixas or 0)

    if qtd <= 0:
        raise HTTPException(status_code=400, detail="Quantidade de caixas inválida (deve ser > 0).")

    snapshot_raw = None
    if payload.snapshot:
        try:
            snapshot_raw = json.dumps(payload.snapshot, ensure_ascii=False)
        except Exception:
            snapshot_raw = None

    obs = (payload.obs or "").strip() or None
    now = _now_iso()
    tid = str(uuid4())

    with get_conn() as conn:
        cur = conn.cursor()
        if _has_pending_substituicao(cur, codigo_origem) or _has_pending_substituicao(cur, codigo_destino):
            raise HTTPException(
                status_code=409,
                detail="Transferencias bloqueadas enquanto houver substituicao de motorista pendente.",
            )
        cur.execute(
            """
            SELECT codigo_programacao, status, status_operacional, tipo_estimativa, operacao_tipo,
                   nf_caixas, total_caixas, caixas_carregadas, caixas_estimado
            FROM programacoes
            WHERE codigo_programacao=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo_origem,),
        )
        origem_row = cur.fetchone()
        if not origem_row:
            raise HTTPException(status_code=404, detail="Rota origem nao encontrada.")
        origem_dict = row_to_dict(origem_row)
        origem_transbordo = _is_transbordo_row(origem_dict)
        carga_raiz_programacao = _resolve_carga_raiz_programacao(cur, codigo_origem)
        if pedido and cod_cliente:
            try:
                cur.execute(
                    """
                    SELECT carga_raiz_programacao
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                      AND UPPER(TRIM(COALESCE(cod_cliente,'')))=UPPER(TRIM(?))
                      AND COALESCE(TRIM(COALESCE(pedido,'')),'')=COALESCE(TRIM(?),'')
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (codigo_origem, cod_cliente, pedido),
                )
                item_raiz = cur.fetchone()
                if item_raiz and str(item_raiz["carga_raiz_programacao"] or "").strip():
                    carga_raiz_programacao = str(item_raiz["carga_raiz_programacao"] or "").strip().upper()
            except Exception:
                pass
        snapshot_obj = payload.snapshot if isinstance(payload.snapshot, dict) else {}
        payload.snapshot = {
            **snapshot_obj,
            "carga_raiz_programacao": carga_raiz_programacao or codigo_origem,
            "carga_origem_imediata": codigo_origem,
            "codigo_origem": codigo_origem,
            "codigo_destino": codigo_destino,
        }
        status_origem = str(origem_dict.get("status_operacional") or origem_dict.get("status") or "").strip().upper().replace(" ", "_")
        transferencia_sem_cliente = origem_transbordo and (not pedido or not cod_cliente)
        if transferencia_sem_cliente:
            if status_origem in ("EM_ENTREGAS", "FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
                raise HTTPException(status_code=409, detail="Transbordo de carga bloqueado para rota em entregas ou encerrada.")
            pedido = "TRANSBORDO"
            cod_cliente = "TRANSBORDO"
            payload.snapshot = {
                **payload.snapshot,
                "transbordo": True,
                "carga_sem_cliente": True,
                "operacao_tipo": "TRANSBORDO",
            }
        else:
            if not pedido:
                raise HTTPException(status_code=400, detail="Pedido e obrigatorio para transferencia de rota de venda.")
            if not cod_cliente:
                raise HTTPException(status_code=400, detail="Codigo do cliente e obrigatorio para transferencia de rota de venda.")
        if payload.snapshot:
            try:
                snapshot_raw = json.dumps(payload.snapshot, ensure_ascii=False)
            except Exception:
                snapshot_raw = None

        # valida rota destino e regra de disponibilidade em rota
        caixas_saldo_expr_dest = _caixas_saldo_subquery(conn, "p")
        cur.execute(
            """
            SELECT
                p.codigo_programacao,
                COALESCE(p.status, '') AS status,
                (
                    SELECT CAST(NULLIF(TRIM(v.capacidade_cx), '') AS INTEGER)
                    FROM veiculos v
                    WHERE UPPER(TRIM(v.placa)) = UPPER(TRIM(p.veiculo))
                       OR UPPER(TRIM(v.modelo)) = UPPER(TRIM(p.veiculo))
                    LIMIT 1
                ) AS capacidade_cx,
                """
            + caixas_saldo_expr_dest
            + """
            FROM programacoes p
            WHERE p.codigo_programacao=?
            LIMIT 1
            """,
            (codigo_destino,),
        )
        destino_row = cur.fetchone()
        if not destino_row:
            raise HTTPException(status_code=404, detail="Rota destino nao encontrada.")

        status_dest = str(destino_row["status"] or "").strip().upper()
        if status_dest in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota destino encerrada (status={status_dest}).",
            )

        # Se destino estiver em rota, só recebe transferência se tiver saldo no veículo.
        em_rota_dest = status_dest in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA")
        if em_rota_dest:
            try:
                saldo_dest = int(float(destino_row["caixas_saldo"] or 0))
            except Exception:
                saldo_dest = 0
            if saldo_dest <= 0:
                raise HTTPException(
                    status_code=409,
                    detail="Motorista destino esta em rota sem saldo de caixas para receber transferencia.",
                )
            if qtd > saldo_dest:
                raise HTTPException(
                    status_code=409,
                    detail=f"Transferencia excede saldo do destino em rota. Saldo destino: {saldo_dest} cx.",
                )
        # valida disponibilidade no pedido de origem
        cur.execute("PRAGMA table_info(programacao_itens)")
        cols_itens = {row[1] for row in cur.fetchall() or []}
        has_pedido_col = "pedido" in cols_itens
        has_caixas_atual_col = "caixas_atual" in cols_itens

        if transferencia_sem_cliente:
            base_qnt = (
                int(origem_dict.get("nf_caixas") or 0)
                or int(origem_dict.get("total_caixas") or 0)
                or int(origem_dict.get("caixas_carregadas") or 0)
                or int(origem_dict.get("caixas_estimado") or 0)
            )
            cur.execute(
                """
                SELECT COALESCE(SUM(qtd_caixas), 0) AS qtd
                FROM transferencias
                WHERE codigo_origem=?
                  AND UPPER(TRIM(COALESCE(status,''))) NOT IN ('CANCELADA','CANCELADO','RECUSADA','RECUSADO')
                """,
                (codigo_origem,),
            )
            row_out = cur.fetchone()
            disponivel_liquido = max(base_qnt - int((row_out["qtd"] if row_out else 0) or 0), 0)
            if base_qnt <= 0:
                raise HTTPException(status_code=409, detail="Informe/salve as caixas carregadas da rota de transbordo antes de transferir.")
            if qtd > disponivel_liquido:
                raise HTTPException(
                    status_code=409,
                    detail=f"Transferencia excede saldo de transbordo. Disponivel: {disponivel_liquido} cx.",
                )
            item_origem = None
        elif has_pedido_col:
            cur.execute(
                """
                SELECT qnt_caixas, caixas_atual
                FROM programacao_itens
                WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido, '')=COALESCE(?, '')
                LIMIT 1
                """,
                (codigo_origem, cod_cliente, pedido),
            )
        else:
            cur.execute(
                """
                SELECT qnt_caixas, caixas_atual
                FROM programacao_itens
                WHERE codigo_programacao=? AND cod_cliente=?
                LIMIT 1
                """,
                (codigo_origem, cod_cliente),
            )
        if not transferencia_sem_cliente:
            item_origem = cur.fetchone()
            if not item_origem:
                raise HTTPException(status_code=404, detail="Pedido de origem não encontrado para transferência.")
            base_qnt = int(item_origem["qnt_caixas"] or 0)
            disponivel = base_qnt
            if has_caixas_atual_col and item_origem["caixas_atual"] is not None:
                try:
                    disponivel = int(item_origem["caixas_atual"])
                except Exception:
                    disponivel = int(item_origem["qnt_caixas"] or 0)

            cur.execute(
                """
                SELECT COALESCE(SUM(qtd_caixas), 0) AS pend
                FROM transferencias
                WHERE codigo_origem=? AND cod_cliente=? AND pedido=? AND UPPER(TRIM(status))='PENDENTE'
                """,
                (codigo_origem, cod_cliente, pedido),
            )
            row_pend = cur.fetchone()
            pendente = int((row_pend["pend"] if row_pend else 0) or 0)
            # Compatibilidade:
            # se caixas_atual ja divergiu da base, assumimos saldo ja aplicado.
            if has_caixas_atual_col and item_origem["caixas_atual"] is not None and int(disponivel) != int(base_qnt):
                disponivel_liquido = max(disponivel, 0)
            else:
                disponivel_liquido = max(disponivel - pendente, 0)
            if qtd > disponivel_liquido:
                raise HTTPException(
                    status_code=409,
                    detail=f"Transferência excede disponível do pedido. Disponível: {disponivel_liquido} cx.",
                )
        cur.execute(
            """
            INSERT INTO transferencias
                (id, codigo_origem, codigo_destino, cod_cliente, pedido, qtd_caixas,
                 status, obs, snapshot, motorista_origem, motorista_destino,
                 qtd_convertida, criado_em, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                tid,
                codigo_origem,
                codigo_destino,
                cod_cliente,
                pedido,
                qtd,
                "PENDENTE",
                obs,
                snapshot_raw,
                nome_motorista,
                None,
                0,
                now,
                now,
            ),
        )

        if transferencia_sem_cliente:
            conn.commit()
            return _fetch_transferencia_by_id(conn, tid)

        # Origem da transferencia: atualiza status/observacao conforme acao
        # parcial => ALTERADO | total => CANCELADO
        novo_caixas_atual = max(int(disponivel_liquido) - int(qtd), 0)
        novo_status_origem = "CANCELADO" if novo_caixas_atual == 0 else "ALTERADO"
        detalhe_origem = (
            f"Transferencia de caixas: -{qtd} cx para {codigo_destino} "
            f"(pedido {pedido} / cliente {cod_cliente})"
        )
        if obs:
            detalhe_origem = f"{detalhe_origem}. Obs: {obs}"

        cur.execute(
            """
            UPDATE programacao_itens_controle
               SET status_pedido=?,
                   alteracao_tipo='QUANTIDADE',
                   alteracao_detalhe=?,
                   caixas_atual=?,
                   alterado_em=?,
                   alterado_por=?,
                   updated_at=datetime('now')
             WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido,'')=COALESCE(?, '')
            """,
            (
                novo_status_origem,
                detalhe_origem,
                novo_caixas_atual,
                now,
                nome_motorista,
                codigo_origem,
                cod_cliente,
                pedido,
            ),
        )
        if cur.rowcount == 0:
            cur.execute(
                """
                INSERT INTO programacao_itens_controle
                    (codigo_programacao, cod_cliente, pedido, status_pedido,
                     alteracao_tipo, alteracao_detalhe, caixas_atual,
                     alterado_em, alterado_por, updated_at)
                VALUES (?, ?, ?, ?, 'QUANTIDADE', ?, ?, ?, ?, datetime('now'))
                """,
                (
                    codigo_origem,
                    cod_cliente,
                    pedido,
                    novo_status_origem,
                    detalhe_origem,
                    novo_caixas_atual,
                    now,
                    nome_motorista,
                ),
            )

        sets = []
        params = []
        if "status_pedido" in cols_itens:
            sets.append("status_pedido=?")
            params.append(novo_status_origem)
        if "alteracao_tipo" in cols_itens:
            sets.append("alteracao_tipo=?")
            params.append("QUANTIDADE")
        if "alteracao_detalhe" in cols_itens:
            sets.append("alteracao_detalhe=?")
            params.append(detalhe_origem)
        if "caixas_atual" in cols_itens:
            sets.append("caixas_atual=?")
            params.append(novo_caixas_atual)
        if "alterado_em" in cols_itens:
            sets.append("alterado_em=?")
            params.append(now)
        if "alterado_por" in cols_itens:
            sets.append("alterado_por=?")
            params.append(nome_motorista)

        if sets:
            if has_pedido_col:
                params.extend([codigo_origem, cod_cliente, pedido])
                cur.execute(
                    f"UPDATE programacao_itens SET {', '.join(sets)} "
                    "WHERE codigo_programacao=? AND cod_cliente=? AND COALESCE(pedido,'')=COALESCE(?, '')",
                    tuple(params),
                )
            else:
                params.extend([codigo_origem, cod_cliente])
                cur.execute(
                    f"UPDATE programacao_itens SET {', '.join(sets)} "
                    "WHERE codigo_programacao=? AND cod_cliente=?",
                    tuple(params),
                )

        conn.commit()
        return _fetch_transferencia_by_id(conn, tid)


@app.get("/desktop/vendas-importadas")
def desktop_listar_vendas_importadas(
    busca: str = Query("", description="Filtro textual"),
    codigo_programacao: str = Query("", description="Filtrar por programacao vinculada"),
    limit: int = Query(5000, ge=1, le=20000),
    _ok: bool = Depends(_require_desktop_secret),
):
    term = (busca or "").strip()
    codigo_filtro = str(codigo_programacao or "").strip().upper()
    like = f"%{term}%"
    where_codigo = (
        "UPPER(COALESCE(codigo_programacao,''))=?"
        if codigo_filtro
        else "TRIM(COALESCE(codigo_programacao,''))=''"
    )
    params_base: Tuple[Any, ...] = ((codigo_filtro,) if codigo_filtro else ())
    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        if term:
            cur.execute(
                f"""
                SELECT id, COALESCE(selecionada,0) AS selecionada, COALESCE(pedido,'') AS pedido,
                       COALESCE(data_venda,'') AS data_venda, COALESCE(cliente,'') AS cliente,
                       COALESCE(nome_cliente,'') AS nome_cliente, COALESCE(produto,'') AS produto,
                       COALESCE(vr_total,0) AS vr_total, COALESCE(qnt,0) AS qnt,
                       COALESCE(cidade,'') AS cidade, COALESCE(vendedor,'') AS vendedor,
                       COALESCE(codigo_programacao,'') AS codigo_programacao
                FROM vendas_importadas
                WHERE IFNULL(usada,0)=0
                  AND {where_codigo}
                  AND (
                    pedido LIKE ? OR cliente LIKE ? OR nome_cliente LIKE ? OR vendedor LIKE ? OR produto LIKE ?
                  )
                ORDER BY id DESC
                LIMIT ?
                """,
                (*params_base, like, like, like, like, like, int(limit)),
            )
        else:
            cur.execute(
                f"""
                SELECT id, COALESCE(selecionada,0) AS selecionada, COALESCE(pedido,'') AS pedido,
                       COALESCE(data_venda,'') AS data_venda, COALESCE(cliente,'') AS cliente,
                       COALESCE(nome_cliente,'') AS nome_cliente, COALESCE(produto,'') AS produto,
                       COALESCE(vr_total,0) AS vr_total, COALESCE(qnt,0) AS qnt,
                       COALESCE(cidade,'') AS cidade, COALESCE(vendedor,'') AS vendedor,
                       COALESCE(codigo_programacao,'') AS codigo_programacao
                FROM vendas_importadas
                WHERE IFNULL(usada,0)=0
                  AND {where_codigo}
                ORDER BY id DESC
                LIMIT ?
                """,
                (*params_base, int(limit)),
            )
        out = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "id": int(r["id"] or 0),
                    "selecionada": int(r["selecionada"] or 0),
                    "pedido": str(r["pedido"] or ""),
                    "data_venda": str(r["data_venda"] or ""),
                    "cliente": str(r["cliente"] or ""),
                    "nome_cliente": str(r["nome_cliente"] or ""),
                    "produto": str(r["produto"] or ""),
                    "vr_total": float(r["vr_total"] or 0.0),
                    "qnt": float(r["qnt"] or 0.0),
                    "cidade": str(r["cidade"] or ""),
                    "vendedor": str(r["vendedor"] or ""),
                    "codigo_programacao": str((r["codigo_programacao"] if "codigo_programacao" in r.keys() else "") or "").strip().upper(),
                }
            )
    return {"ok": True, "rows": out}


@app.post("/desktop/vendas-importadas/importar")
def desktop_importar_vendas_importadas(
    payload: Dict[str, Any],
    _ok: bool = Depends(_require_desktop_secret),
):
    rows = payload.get("rows") if isinstance(payload, dict) else []
    if not isinstance(rows, list):
        rows = []
    if len(rows) > 20000:
        raise HTTPException(status_code=400, detail="Quantidade de linhas excede o limite (20000).")

    total = 0
    ignoradas = 0
    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        for rr in rows:
            if not isinstance(rr, dict):
                ignoradas += 1
                continue
            pedido_u = str(rr.get("pedido") or "").strip().upper()
            cliente_u = str(rr.get("cliente") or "").strip().upper()
            produto_u = str(rr.get("produto") or "").strip().upper()
            nome_u = str(rr.get("nome_cliente") or "").strip().upper()
            data_venda = str(rr.get("data_venda") or "").strip()
            if (not pedido_u) or (not cliente_u) or (not produto_u) or (not nome_u):
                ignoradas += 1
                continue
            cur.execute(
                """
                SELECT 1
                FROM vendas_importadas
                WHERE UPPER(TRIM(COALESCE(pedido,'')))=UPPER(TRIM(?))
                  AND UPPER(TRIM(COALESCE(cliente,'')))=UPPER(TRIM(?))
                  AND UPPER(TRIM(COALESCE(produto,'')))=UPPER(TRIM(?))
                  AND COALESCE(TRIM(data_venda),'')=COALESCE(TRIM(?),'')
                LIMIT 1
                """,
                (pedido_u, cliente_u, produto_u, data_venda),
            )
            if cur.fetchone():
                ignoradas += 1
                continue
            cur.execute(
                """
                INSERT INTO vendas_importadas
                (pedido, data_venda, cliente, nome_cliente, vendedor, produto, vr_total, qnt, cidade, valor_unitario, observacao, selecionada, usada, usada_em, codigo_programacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, '', '')
                """,
                (
                    pedido_u,
                    data_venda,
                    cliente_u,
                    nome_u,
                    str(rr.get("vendedor") or "").strip().upper(),
                    produto_u,
                    float(rr.get("vr_total") or 0.0),
                    float(rr.get("qnt") or 0.0),
                    str(rr.get("cidade") or "").strip().upper(),
                    float(rr.get("valor_unitario") or 0.0),
                    str(rr.get("observacao") or "").strip().upper(),
                ),
            )
            total += 1
    return {"ok": True, "importadas": total, "ignoradas": ignoradas}


@app.post("/desktop/vendas-importadas/{venda_id}/toggle-selecao")
def desktop_toggle_venda_importada_selecao(venda_id: int, _ok: bool = Depends(_require_desktop_secret)):
    rid = int(venda_id or 0)
    if rid <= 0:
        raise HTTPException(status_code=400, detail="venda_id invalido.")
    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE vendas_importadas
               SET selecionada = CASE WHEN selecionada=1 THEN 0 ELSE 1 END
             WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''
            """,
            (rid,),
        )
        updated = int(cur.rowcount or 0)
    return {"ok": True, "id": rid, "updated": updated}


@app.post("/desktop/vendas-importadas/marcar-todas")
def desktop_marcar_todas_vendas_importadas(
    selected: int = Query(1, ge=0, le=1),
    _ok: bool = Depends(_require_desktop_secret),
):
    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        cur.execute(
            "UPDATE vendas_importadas SET selecionada=? WHERE IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
            (int(selected),),
        )
        updated = int(cur.rowcount or 0)
    return {"ok": True, "updated": updated, "selected": int(selected)}


@app.post("/desktop/vendas-importadas/marcar-ids")
def desktop_marcar_ids_vendas_importadas(
    ids: str = Query("", description="CSV de ids"),
    _ok: bool = Depends(_require_desktop_secret),
):
    raw_ids = [x.strip() for x in str(ids or "").split(",") if x.strip()]
    id_list: List[int] = []
    for x in raw_ids:
        try:
            v = int(x)
        except Exception:
            continue
        if v > 0:
            id_list.append(v)
    if not id_list:
        return {"ok": True, "updated": 0}

    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        cur.executemany(
            "UPDATE vendas_importadas SET selecionada=1 WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
            [(rid,) for rid in id_list],
        )
        updated = int(cur.rowcount or 0)
    return {"ok": True, "updated": updated}


@app.post("/desktop/vendas-importadas/consumir")
def desktop_consumir_vendas_importadas(
    payload: Dict[str, Any],
    _ok: bool = Depends(_require_desktop_secret),
):
    ids_raw = payload.get("ids") if isinstance(payload, dict) else []
    codigo_programacao = str((payload or {}).get("codigo_programacao") or "").strip().upper()
    usada_em = str((payload or {}).get("usada_em") or "").strip()
    if not codigo_programacao:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")
    if not isinstance(ids_raw, list):
        ids_raw = []

    ids: List[int] = []
    for x in ids_raw:
        try:
            v = int(x)
        except Exception:
            continue
        if v > 0:
            ids.append(v)
    if not ids:
        return {"ok": True, "updated": 0}

    if not usada_em:
        usada_em = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, codigo_programacao)
        cur.executemany(
            """
            UPDATE vendas_importadas
               SET usada=1,
                   usada_em=?,
                   codigo_programacao=?,
                   selecionada=0
             WHERE id=? AND IFNULL(usada,0)=0
            """,
            [(usada_em, codigo_programacao, rid) for rid in ids],
        )
        updated = int(cur.rowcount or 0)
    return {"ok": True, "updated": updated}


@app.post("/desktop/vendas-importadas/vincular")
def desktop_vincular_vendas_importadas(
    payload: Dict[str, Any],
    _ok: bool = Depends(_require_desktop_secret),
):
    ids_raw = payload.get("ids") if isinstance(payload, dict) else []
    codigo_programacao = str((payload or {}).get("codigo_programacao") or "").strip().upper()
    if not codigo_programacao:
        raise HTTPException(status_code=400, detail="codigo_programacao obrigatorio.")
    if not isinstance(ids_raw, list):
        ids_raw = []

    ids: List[int] = []
    for x in ids_raw:
        try:
            v = int(x)
        except Exception:
            continue
        if v > 0:
            ids.append(v)
    if not ids:
        return {"ok": True, "updated": 0}

    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        _ensure_programacao_mutable(cur, codigo_programacao)
        placeholders = ",".join(["?"] * len(ids))
        cur.execute(
            f"""
            SELECT COUNT(1) AS n
            FROM vendas_importadas
            WHERE id IN ({placeholders})
              AND IFNULL(usada,0)=0
              AND TRIM(COALESCE(codigo_programacao,'')) <> ''
              AND UPPER(TRIM(COALESCE(codigo_programacao,''))) <> UPPER(TRIM(?))
            """,
            (*ids, codigo_programacao),
        )
        row = cur.fetchone()
        conflito = int((row["n"] if row else 0) or 0)
        if conflito > 0:
            raise HTTPException(
                status_code=409,
                detail="existem vendas ja vinculadas a outra programacao; desvincule antes de mover.",
            )
        cur.executemany(
            """
            UPDATE vendas_importadas
               SET codigo_programacao=?,
                   selecionada=1
             WHERE id=? AND IFNULL(usada,0)=0
            """,
            [(codigo_programacao, rid) for rid in ids],
        )
        updated = int(cur.rowcount or 0)
    return {"ok": True, "updated": updated, "codigo_programacao": codigo_programacao}


@app.delete("/desktop/vendas-importadas")
def desktop_apagar_vendas_importadas(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT COUNT(1) AS n
            FROM vendas_importadas
            WHERE IFNULL(usada,0)=1 OR TRIM(COALESCE(codigo_programacao,''))<>''
            """
        )
        row = cur.fetchone()
        protegidas = int((row["n"] if row else 0) or 0)
        cur.execute("DELETE FROM vendas_importadas WHERE IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''")
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "deleted": deleted, "preserved": protegidas}


@app.delete("/desktop/vendas-importadas/ids")
def desktop_apagar_ids_vendas_importadas(
    ids: str = Query("", description="CSV de ids"),
    _ok: bool = Depends(_require_desktop_secret),
):
    raw_ids = [x.strip() for x in str(ids or "").split(",") if x.strip()]
    id_list: List[int] = []
    for x in raw_ids:
        try:
            v = int(x)
        except Exception:
            continue
        if v > 0:
            id_list.append(v)
    if not id_list:
        return {"ok": True, "deleted": 0}

    with get_conn() as conn:
        ensure_core_schema(conn)
        cur = conn.cursor()
        placeholders = ",".join(["?"] * len(id_list))
        cur.execute(
            f"""
            SELECT COUNT(1) AS n
            FROM vendas_importadas
            WHERE id IN ({placeholders})
              AND (IFNULL(usada,0)=1 OR TRIM(COALESCE(codigo_programacao,''))<>'')
            """,
            tuple(id_list),
        )
        row = cur.fetchone()
        protegidas = int((row["n"] if row else 0) or 0)
        if protegidas > 0:
            raise HTTPException(
                status_code=409,
                detail="existem vendas selecionadas ja vinculadas ou consumidas; exclusao bloqueada.",
            )
        cur.executemany(
            "DELETE FROM vendas_importadas WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
            [(rid,) for rid in id_list],
        )
        deleted = int(cur.rowcount or 0)
    return {"ok": True, "deleted": deleted}

# âœ… LISTAR TRANSFERÊNCIAS (destino recebe)
@app.get("/rotas/{codigo_programacao}/transferencias")
def listar_transferencias(
    codigo_programacao: str,
    status: Optional[str] = Query(default=None),
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo = (codigo_programacao or "").strip()

    if not codigo:
        raise HTTPException(status_code=400, detail="Código inválido.")

    if not _rota_pertence_ao_motorista(codigo, m):
        raise HTTPException(status_code=403, detail="Rota de destino não pertence ao motorista logado.")

    with get_conn() as conn:
        return _list_transferencias_por_destino(conn, codigo, status)


@app.get("/rotas/{codigo_programacao}/transferencias-enviadas")
def listar_transferencias_enviadas(
    codigo_programacao: str,
    status: Optional[str] = Query(default=None),
    m=Depends(get_current_motorista),
):
    codigo = (codigo_programacao or "").strip()

    if not codigo:
        raise HTTPException(status_code=400, detail="Código inválido.")

    if not _rota_pertence_ao_motorista(codigo, m):
        raise HTTPException(status_code=403, detail="Rota de origem não pertence ao motorista logado.")

    with get_conn() as conn:
        return _list_transferencias_por_origem(conn, codigo, status)


@app.post("/transferencias/{transferencia_id}/aceitar")
def aceitar_transferencia(
    transferencia_id: str,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    tid = (transferencia_id or "").strip()
    if not tid:
        raise HTTPException(status_code=400, detail="ID invalido.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="Transfer?ncia nao encontrada.")

        codigo_destino = str(item.get("codigo_destino", "")).strip()
        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="Transfer?ncia nao pertence ao motorista logado (destino).")

        st = str(item.get("status", "")).upper().strip()
        if st != "PENDENTE":
            raise HTTPException(status_code=409, detail=f"Transfer?ncia nao est? pendente (status={st}).")

        cur.execute(
            """
            UPDATE transferencias
            SET status=?, motorista_destino=?, atualizado_em=?
            WHERE id=?
            """,
            ("ACEITA", nome_motorista, _now_iso(), tid),
        )
        conn.commit()
        return _fetch_transferencia_by_id(conn, tid)


@app.post("/transferencias/{transferencia_id}/recusar")
def recusar_transferencia(
    transferencia_id: str,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    tid = (transferencia_id or "").strip()
    if not tid:
        raise HTTPException(status_code=400, detail="ID inválido.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="Transferência não encontrada.")

        codigo_destino = str(item.get("codigo_destino", "")).strip()

        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="Transferência não pertence ao motorista logado (destino).")

        st = str(item.get("status", "")).upper().strip()
        if st != "PENDENTE":
            raise HTTPException(status_code=409, detail=f"Transferência não está pendente (status={st}).")

        cur.execute(
            """
            UPDATE transferencias
            SET status=?, motorista_destino=?, atualizado_em=?
            WHERE id=?
            """,
            ("RECUSADA", nome_motorista, _now_iso(), tid),
        )

        # Ao recusar, recalcula imediatamente o pedido da origem para evitar caixas "soltas".
        _recalcular_origem_transferencia(
            cur,
            codigo_origem=str(item.get("codigo_origem") or "").strip(),
            cod_cliente=str(item.get("cod_cliente") or "").strip(),
            pedido_ref=str(item.get("pedido") or "").strip(),
            alterado_por=nome_motorista or "SISTEMA",
            evento="RECUSADA",
        )

        conn.commit()
        return _fetch_transferencia_by_id(conn, tid)


# =====================================================
# ===== CONVERTER TRANSFERÊNCIA (RESERVA -> PEDIDO) ====
# =====================================================

class TransferenciaConverterIn(BaseModel):
    pedido_destino: Optional[str] = None
    cod_cliente_destino: Optional[str] = None
    qtd_caixas: int
    obs: Optional[str] = None
    novo_cliente: Optional[Dict[str, Any]] = None


@app.post("/transferencias/{transferencia_id}/converter")
def converter_transferencia(
    transferencia_id: str,
    payload: TransferenciaConverterIn,
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()

    tid = (transferencia_id or "").strip()
    if not tid:
        raise HTTPException(status_code=400, detail="ID inválido.")

    pedido_dest = (payload.pedido_destino or "").strip()
    cod_cli_dest = (payload.cod_cliente_destino or "").strip()
    qtd = int(payload.qtd_caixas or 0)
    obs = (payload.obs or "").strip()

    novo = payload.novo_cliente or None
    nome_novo = ""
    if novo:
        try:
            nome_novo = (novo.get("nome_cliente") or "").strip()
        except Exception:
            nome_novo = ""
        pedido_novo = (str(novo.get("pedido") or "")).strip() if isinstance(novo, dict) else ""
        cod_novo = (str(novo.get("cod_cliente") or "")).strip() if isinstance(novo, dict) else ""
        if not nome_novo:
            raise HTTPException(status_code=400, detail="nome_cliente é obrigatório para novo cliente.")
        if not pedido_novo:
            raise HTTPException(status_code=400, detail="pedido é obrigatório para novo cliente.")
        if not cod_novo:
            cod_novo = f"MANUAL-{uuid4().hex[:8].upper()}"
        pedido_dest = pedido_novo
        cod_cli_dest = cod_novo

    if not novo:
        if not pedido_dest:
            raise HTTPException(status_code=400, detail="pedido_destino é obrigatório.")
        if not cod_cli_dest:
            raise HTTPException(status_code=400, detail="cod_cliente_destino é obrigatório.")
    if qtd <= 0:
        raise HTTPException(status_code=400, detail="qtd_caixas deve ser > 0.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="Transferência não encontrada.")

        st = str(item.get("status", "")).upper().strip()
        if st != "ACEITA":
            raise HTTPException(status_code=409, detail=f"Transferência não está ACEITA (status={st}).")

        codigo_destino = str(item.get("codigo_destino", "")).strip()
        codigo_origem = str(item.get("codigo_origem", "")).strip()
        if _has_pending_substituicao(cur, codigo_origem) or _has_pending_substituicao(cur, codigo_destino):
            raise HTTPException(
                status_code=409,
                detail="Conversao de transferencia bloqueada enquanto houver substituicao pendente.",
            )
        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="Você não é o motorista destino desta transferência.")

        total = int(item.get("qtd_caixas") or 0)
        convertido = int(item.get("qtd_convertida") or 0)
        saldo = total - convertido
        if saldo < 0:
            saldo = 0

        if qtd > saldo:
            raise HTTPException(status_code=409, detail=f"Quantidade maior que o saldo disponÃvel ({saldo}).")

        cur.execute(
            """
            INSERT INTO transferencias_conversoes
                (transferencia_id, pedido_destino, cod_cliente_destino, qtd, obs, nome_cliente_destino, novo_cliente)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                tid,
                pedido_dest,
                cod_cli_dest,
                qtd,
                obs or None,
                nome_novo or None,
                1 if nome_novo else 0,
            ),
        )
        nome_cliente_destino = nome_novo
        if not nome_cliente_destino and novo and isinstance(novo, dict):
            nome_cliente_destino = str(novo.get("nome_cliente") or novo.get("nome_cliente_destino") or "").strip()
        if not nome_cliente_destino:
            try:
                snapshot = item.get("snapshot") or {}
                if isinstance(snapshot, dict):
                    nome_cliente_destino = str(
                        snapshot.get("nome_cliente_destino")
                        or snapshot.get("nome_cliente")
                        or snapshot.get("cliente")
                        or ""
                    ).strip()
            except Exception:
                nome_cliente_destino = ""
        _upsert_item_destino_transferencia(
            cur,
            codigo_destino=codigo_destino,
            cod_cliente=cod_cli_dest,
            pedido=pedido_dest,
            nome_cliente=nome_cliente_destino or cod_cli_dest,
            qtd_caixas=qtd,
            obs=obs,
            alterado_por=nome_motorista or "SISTEMA",
            carga_raiz_programacao=str(
                (item.get("snapshot") or {}).get("carga_raiz_programacao")
                or item.get("carga_raiz_programacao")
                or _resolve_carga_raiz_programacao(cur, codigo_origem)
                or codigo_origem
                or ""
            ).strip().upper(),
            carga_origem_imediata=str(item.get("codigo_origem") or "").strip().upper(),
            transferencia_origem_id=tid,
        )

        novo_convertido = convertido + qtd
        novo_status = "CONVERTIDA" if novo_convertido >= total else "ACEITA"
        cur.execute(
            """
            UPDATE transferencias
            SET qtd_convertida=qtd_convertida + ?, status=?, atualizado_em=?
            WHERE id=?
            """,
            (qtd, novo_status, _now_iso(), tid),
        )

        conn.commit()
        return _fetch_transferencia_by_id(conn, tid)


