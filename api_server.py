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
from uuid import uuid4
from datetime import datetime
from typing import Optional, List, Dict, Any, Iterator
from contextlib import contextmanager

from fastapi import FastAPI, HTTPException, Depends, Query, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, Field


# =========================================================
# CONFIG
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# âœ… Prioridade:
# 1) variÃ¡vel de ambiente ROTA_DB
# 2) rota_granja.db na pasta do projeto
DB_PATH = os.environ.get("ROTA_DB") or os.path.join(BASE_DIR, "rota_granja.db")

SECRET_KEY = os.environ.get("ROTA_SECRET")
if not SECRET_KEY:
    raise RuntimeError("ROTA_SECRET nao definido. Configure a variavel de ambiente para iniciar a API.")
TOKEN_TTL_SECONDS = 60 * 60 * 24 * 7  # 7 dias

app = FastAPI(title="Rota Granja API", version="1.1.4")

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
    if not os.path.exists(DB_PATH):
        raise RuntimeError(f"Banco n?o encontrado em: {DB_PATH}")
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

def authenticate_motorista(cur: sqlite3.Cursor, codigo: str, senha: str) -> tuple[Optional[sqlite3.Row], Optional[str]]:
    has_acesso = col_exists(cur.connection, "motoristas", "acesso_liberado")
    if has_acesso:
        cur.execute(
            """
            SELECT id, nome, codigo, senha, COALESCE(acesso_liberado, 0) AS acesso_liberado
            FROM motoristas
            WHERE UPPER(TRIM(codigo))=?
            LIMIT 1
            """,
            (codigo,),
        )
    else:
        cur.execute(
            "SELECT id, nome, codigo, senha FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1",
            (codigo,),
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
    return _resolve_ajudante_primeiro_nome(cur, row.get("equipe"))


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
    row["tipo_operacao"] = "CIF" if tipo_estimativa == "CX" else "FOB"
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


def _local_rota_expr(conn: sqlite3.Connection) -> str:
    candidates: List[str] = []
    if col_exists(conn, "programacoes", "local_rota"):
        candidates.append("NULLIF(TRIM(p.local_rota), '')")
    if col_exists(conn, "programacoes", "tipo_rota"):
        candidates.append("NULLIF(TRIM(p.tipo_rota), '')")
    if col_exists(conn, "programacoes", "local_carregamento"):
        candidates.append("NULLIF(TRIM(p.local_carregamento), '')")
    if col_exists(conn, "programacoes", "local_carreg"):
        candidates.append("NULLIF(TRIM(p.local_carreg), '')")

    if not candidates:
        return "'-' AS local_rota"
    return f"COALESCE({', '.join(candidates)}, '-') AS local_rota"


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
    try:
        cur = conn.cursor()
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

    join_on = "pc.codigo_programacao = pi.codigo_programacao AND UPPER(TRIM(pc.cod_cliente)) = UPPER(TRIM(pi.cod_cliente))"
    if has_pi_pedido and has_pc_pedido:
        join_on += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"

    st_pi_expr = "COALESCE(NULLIF(TRIM(pi.status_pedido),''), 'PENDENTE')" if has_pi_status else "'PENDENTE'"
    st_pc_expr = "NULLIF(TRIM(pc.status_pedido),'')" if has_pc_status else "NULL"
    base_expr = "COALESCE(pi.qnt_caixas, 0)" if has_pi_qnt_caixas else "0"

    if has_pc_caixas_atual and has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = "COALESCE(pc.caixas_atual, pi.caixas_atual, pi.qnt_caixas, 0)"
    elif has_pc_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = "COALESCE(pc.caixas_atual, pi.qnt_caixas, 0)"
    elif has_pi_caixas_atual and has_pi_qnt_caixas:
        caixas_raw = "COALESCE(pi.caixas_atual, pi.qnt_caixas, 0)"
    elif has_pc_caixas_atual:
        caixas_raw = "COALESCE(pc.caixas_atual, 0)"
    else:
        caixas_raw = base_expr

    status_eff = f"COALESCE({st_pc_expr}, {st_pi_expr}, 'PENDENTE')"
    saldo_item = f"CASE WHEN UPPER({status_eff}) IN ('ENTREGUE','CANCELADO') THEN 0 ELSE COALESCE({caixas_raw},0) END"
    return f"""(
                    SELECT SUM({saldo_item})
                    FROM programacao_itens pi
                    LEFT JOIN programacao_itens_controle pc
                      ON {join_on}
                    WHERE pi.codigo_programacao = {prog_alias}.codigo_programacao
                ) AS caixas_saldo"""


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

                updated_at TEXT DEFAULT (datetime('now')),

                UNIQUE(codigo_programacao, cod_cliente, pedido)
            )
        """)

        # Migra schema legado da tabela de controle:
        # antes a chave Ãºnica era (codigo_programacao, cod_cliente), o que colide pedidos do mesmo cliente.
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
                         caixas_atual, preco_atual, alterado_em, alterado_por, updated_at)
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
                recorded_at TEXT DEFAULT (datetime('now'))
            )
        """)

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

            add_col("status_pedido", "status_pedido TEXT")
            add_col("alteracao_tipo", "alteracao_tipo TEXT")
            add_col("alteracao_detalhe", "alteracao_detalhe TEXT")
            add_col("caixas_atual", "caixas_atual INTEGER")
            add_col("preco_atual", "preco_atual REAL")
            add_col("alterado_em", "alterado_em TEXT")
            add_col("alterado_por", "alterado_por TEXT")
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
            add_ctrl_col("updated_at", "updated_at TEXT")
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
            # FOB/CIF + auditoria (compatível com desktop)
            add_prog_col("tipo_estimativa", "tipo_estimativa TEXT DEFAULT 'KG'")
            add_prog_col("caixas_estimado", "caixas_estimado INTEGER DEFAULT 0")
            add_prog_col("usuario_criacao", "usuario_criacao TEXT")
            add_prog_col("usuario_ultima_edicao", "usuario_ultima_edicao TEXT")
            add_prog_col("status_operacional", "status_operacional TEXT")
            add_prog_col("status_operacional_obs", "status_operacional_obs TEXT")
            add_prog_col("status_operacional_em", "status_operacional_em TEXT")
            add_prog_col("status_operacional_por", "status_operacional_por TEXT")
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


@app.on_event("startup")
def _startup():
    ensure_tables()
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


def create_token(codigo: str) -> str:
    payload = {
        "codigo": codigo,
        "exp": int(time.time()) + TOKEN_TTL_SECONDS,
    }
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

        return {"codigo": codigo, "exp": exp}
    except Exception:
        return None

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


class MotoristaAcessoIn(BaseModel):
    liberado: bool
    admin: Optional[str] = None
    motivo: Optional[str] = None


class MotoristaSenhaIn(BaseModel):
    nova_senha: str
    admin: Optional[str] = None
    motivo: Optional[str] = None


class RotaAtivaOut(BaseModel):
    codigo_programacao: str
    status: str = ""
    motorista: str = ""
    veiculo: str = ""
    equipe: str = ""
    local_rota: str = ""
    data_criacao: str = ""
    tipo_estimativa: Optional[str] = None
    unidade_estimativa: Optional[str] = None
    tipo_operacao: Optional[str] = None
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


class CarregamentoIn(BaseModel):
    # O app (Carregamento2Page) manda isso:
    nf_numero: str = Field(..., min_length=1)
    nf_kg: float = Field(0.0, ge=0)
    kg_carregado: float = Field(..., ge=0)

    caixas_carregadas: int = Field(..., gt=0)

    # podem vir vazios, entÃ£o nÃ£o force min_length=1
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

    # recebimentos (jÃ¡ fica pronto, mesmo que o app ainda nÃ£o use)
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
    evento_em: Optional[str] = None
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
    status: Optional[str] = "ATIVA"
    local_rota: Optional[str] = None
    local_carregamento: Optional[str] = None
    adiantamento: Optional[float] = 0.0
    total_caixas: Optional[int] = 0
    quilos: Optional[float] = 0.0
    usuario_criacao: Optional[str] = None
    usuario_ultima_edicao: Optional[str] = None
    itens: List[DesktopRotaItemIn] = Field(default_factory=list)


class DesktopMotoristaUpsertIn(BaseModel):
    codigo: str
    nome: str
    telefone: Optional[str] = None
    cpf: Optional[str] = None
    status: Optional[str] = "ATIVO"
    senha: Optional[str] = None
    acesso_liberado: Optional[bool] = True
    acesso_liberado_por: Optional[str] = None
    acesso_obs: Optional[str] = None


class DesktopVeiculoUpsertIn(BaseModel):
    placa: str
    modelo: str
    capacidade_cx: Optional[int] = 0


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

    codigo = data.get("codigo")
    if not codigo:
        raise HTTPException(status_code=401, detail="Token sem cÃ³digo do motorista")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols_m = {r[1] for r in (cur.fetchall() or [])}
        if "acesso_liberado" in cols_m:
            cur.execute(
                """
                SELECT id, nome, codigo, COALESCE(acesso_liberado, 0) AS acesso_liberado
                FROM motoristas
                WHERE codigo=?
                """,
                (codigo,),
            )
        else:
            cur.execute("SELECT id, nome, codigo FROM motoristas WHERE codigo=?", (codigo,))
        m = cur.fetchone()
        if not m:
            raise HTTPException(status_code=401, detail="Motorista nÃ£o encontrado")
        if "acesso_liberado" in cols_m and int(m["acesso_liberado"] or 0) != 1:
            raise HTTPException(status_code=403, detail="Acesso bloqueado. Solicite desbloqueio do administrador.")

    return {"codigo": m["codigo"], "nome": m["nome"], "id": m["id"]}


def _owner_filter_for_programacoes(
    conn: sqlite3.Connection,
    motorista: Dict[str, Any],
    alias: str = "p",
) -> tuple[str, tuple]:
    """
    Resolve filtro de posse por motorista com prioridade em chaves estÃ¡veis.
    Fallback por nome existe apenas para bancos legados sem coluna de vÃ­nculo.
    """
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(programacoes)")
    cols = {r[1] for r in cur.fetchall() or []}

    conds: List[str] = []
    params: List[Any] = []

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
    if not secret or secret != str(SECRET_KEY):
        raise HTTPException(status_code=401, detail="Desktop secret inválido")
    return True


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
# ENDPOINTS BÃSICOS
# =========================================================
@app.get("/ping")
def ping():
    return {"ok": True, "db": DB_PATH}


@app.get("/desktop/cadastros/motoristas")
def desktop_motoristas(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='motoristas'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        status_filter = (
            "WHERE UPPER(COALESCE(status,'ATIVO')) IN ('ATIVO','DESATIVADO')"
            if "status" in cols
            else ""
        )
        cur.execute(
            f"""
            SELECT id, COALESCE(codigo,''), COALESCE(nome,''), COALESCE(status,'ATIVO')
            FROM motoristas
            {status_filter}
            ORDER BY UPPER(COALESCE(nome,'')), id
            """
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "codigo": str(r[1] or "").strip().upper(),
                    "nome": str(r[2] or "").strip().upper(),
                    "status": str(r[3] or "ATIVO").strip().upper(),
                }
            )
        return out


@app.get("/desktop/cadastros/veiculos")
def desktop_veiculos(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='veiculos'")
        if not cur.fetchone():
            return []
        cur.execute(
            """
            SELECT id, COALESCE(placa,''), COALESCE(modelo,''), COALESCE(capacidade_cx, 0)
            FROM veiculos
            ORDER BY UPPER(COALESCE(placa,'')), id
            """
        )
        out = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "id": int(r[0] or 0),
                    "placa": str(r[1] or "").strip().upper(),
                    "modelo": str(r[2] or "").strip().upper(),
                    "capacidade_cx": int(r[3] or 0),
                }
            )
        return out


@app.get("/desktop/cadastros/ajudantes")
def desktop_ajudantes(_ok: bool = Depends(_require_desktop_secret)):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ajudantes'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(ajudantes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        status_filter = "WHERE UPPER(COALESCE(status,'ATIVO'))='ATIVO'" if "status" in cols else ""
        cur.execute(
            f"""
            SELECT id, COALESCE(nome,''), COALESCE(sobrenome,''), COALESCE(status,'ATIVO')
            FROM ajudantes
            {status_filter}
            ORDER BY UPPER(COALESCE(nome,'')), UPPER(COALESCE(sobrenome,'')), id
            """
        )
        out = []
        for r in cur.fetchall() or []:
            nome = f"{str(r[1] or '').strip()} {str(r[2] or '').strip()}".strip().upper()
            out.append(
                {
                    "id": int(r[0] or 0),
                    "nome": nome,
                    "status": str(r[3] or "ATIVO").strip().upper(),
                }
            )
        return out


@app.post("/desktop/cadastros/motoristas/upsert")
def desktop_motoristas_upsert(payload: DesktopMotoristaUpsertIn, _ok: bool = Depends(_require_desktop_secret)):
    codigo = _clean_text(payload.codigo).upper()
    nome = _clean_text(payload.nome).upper()
    if not codigo or not nome:
        raise HTTPException(status_code=400, detail="codigo e nome sao obrigatorios.")

    status = _clean_text(payload.status or "ATIVO").upper()
    if status not in {"ATIVO", "DESATIVADO"}:
        status = "ATIVO"

    telefone = _clean_text(payload.telefone)
    cpf = _clean_text(payload.cpf)

    senha_in = _clean_text(payload.senha)
    senha_hash = ""
    if senha_in:
        senha_hash = senha_in if senha_in.startswith("pbkdf2_sha256$") else hash_password_pbkdf2(senha_in)

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(motoristas)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela motoristas indisponivel.")

        cur.execute("SELECT id, COALESCE(senha,'') AS senha FROM motoristas WHERE UPPER(TRIM(codigo))=? LIMIT 1", (codigo,))
        existing = cur.fetchone()

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
        if "senha" in cols:
            cols_ins.append("senha"); vals_ins.append(senha_hash or hash_password_pbkdf2("1234"))
        if "acesso_liberado" in cols:
            cols_ins.append("acesso_liberado"); vals_ins.append(1 if bool(payload.acesso_liberado) else 0)
        if "acesso_liberado_por" in cols:
            cols_ins.append("acesso_liberado_por"); vals_ins.append(_clean_text(payload.acesso_liberado_por or "DESKTOP_SYNC").upper())
        if "acesso_liberado_em" in cols:
            cols_ins.append("acesso_liberado_em"); vals_ins.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        if "acesso_obs" in cols:
            cols_ins.append("acesso_obs"); vals_ins.append(_clean_text(payload.acesso_obs or "Sincronizado via Desktop"))

        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO motoristas ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "codigo": codigo, "created": 1}


@app.post("/desktop/cadastros/veiculos/upsert")
def desktop_veiculos_upsert(payload: DesktopVeiculoUpsertIn, _ok: bool = Depends(_require_desktop_secret)):
    placa = _clean_text(payload.placa).upper()
    modelo = _clean_text(payload.modelo).upper()
    capacidade_cx = int(payload.capacidade_cx or 0)
    if not placa or not modelo:
        raise HTTPException(status_code=400, detail="placa e modelo sao obrigatorios.")
    if capacidade_cx < 0:
        capacidade_cx = 0

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(veiculos)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela veiculos indisponivel.")

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

        if existing:
            if not set_parts:
                return {"ok": True, "placa": placa, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE veiculos SET {', '.join(set_parts)} WHERE id=?", tuple(params))
            return {"ok": True, "placa": placa, "updated": int(cur.rowcount or 0)}

        cols_ins: List[str] = []
        vals_ins: List[Any] = []
        if "placa" in cols:
            cols_ins.append("placa"); vals_ins.append(placa)
        if "modelo" in cols:
            cols_ins.append("modelo"); vals_ins.append(modelo)
        if "capacidade_cx" in cols:
            cols_ins.append("capacidade_cx"); vals_ins.append(capacidade_cx)
        if not cols_ins:
            raise HTTPException(status_code=500, detail="Colunas de veiculos indisponiveis.")
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO veiculos ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "placa": placa, "created": 1}


@app.post("/desktop/cadastros/ajudantes/upsert")
def desktop_ajudantes_upsert(payload: DesktopAjudanteUpsertIn, _ok: bool = Depends(_require_desktop_secret)):
    nome = _clean_text(payload.nome).upper()
    sobrenome = _clean_text(payload.sobrenome).upper()
    telefone = _clean_text(payload.telefone)
    status = _clean_text(payload.status or "ATIVO").upper()
    if status not in {"ATIVO", "DESATIVADO"}:
        status = "ATIVO"
    if not nome or not sobrenome:
        raise HTTPException(status_code=400, detail="nome e sobrenome sao obrigatorios.")

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(ajudantes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela ajudantes indisponivel.")

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
        if not cols_ins:
            raise HTTPException(status_code=500, detail="Colunas de ajudantes indisponiveis.")
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO ajudantes ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "nome": nome, "sobrenome": sobrenome, "created": 1}


@app.post("/desktop/cadastros/clientes/upsert")
def desktop_clientes_upsert(payload: DesktopClienteUpsertIn, _ok: bool = Depends(_require_desktop_secret)):
    cod_cliente = _clean_text(payload.cod_cliente).upper()
    nome_cliente = _clean_text(payload.nome_cliente).upper()
    if not cod_cliente or not nome_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente e nome_cliente sao obrigatorios.")

    endereco = _clean_text(payload.endereco).upper()
    telefone = _clean_text(payload.telefone).upper()
    vendedor = _clean_text(payload.vendedor).upper()

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(clientes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        if not cols:
            raise HTTPException(status_code=500, detail="Tabela clientes indisponivel.")

        cur.execute("SELECT id FROM clientes WHERE UPPER(TRIM(cod_cliente))=? LIMIT 1", (cod_cliente,))
        existing = cur.fetchone()

        set_parts: List[str] = []
        params: List[Any] = []
        if "cod_cliente" in cols:
            set_parts.append("cod_cliente=?"); params.append(cod_cliente)
        if "nome_cliente" in cols:
            set_parts.append("nome_cliente=?"); params.append(nome_cliente)
        if "endereco" in cols:
            set_parts.append("endereco=?"); params.append(endereco)
        if "telefone" in cols:
            set_parts.append("telefone=?"); params.append(telefone)
        if "vendedor" in cols:
            set_parts.append("vendedor=?"); params.append(vendedor)

        if existing:
            if not set_parts:
                return {"ok": True, "cod_cliente": cod_cliente, "updated": 0}
            params.append(int(existing["id"]))
            cur.execute(f"UPDATE clientes SET {', '.join(set_parts)} WHERE id=?", tuple(params))
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
        if not cols_ins:
            raise HTTPException(status_code=500, detail="Colunas de clientes indisponiveis.")
        ph = ", ".join(["?"] * len(cols_ins))
        cur.execute(f"INSERT INTO clientes ({', '.join(cols_ins)}) VALUES ({ph})", tuple(vals_ins))
        return {"ok": True, "cod_cliente": cod_cliente, "created": 1}


@app.get("/desktop/clientes/base")
def desktop_clientes_base(
    q: str = Query("", description="Busca por codigo/nome/cidade"),
    vendedor: str = Query("", description="Filtro por vendedor"),
    cidade: str = Query("", description="Filtro por cidade"),
    ordem: str = Query("nome", description="nome|codigo"),
    limit: int = Query(300, ge=1, le=1000),
    _ok: bool = Depends(_require_desktop_secret),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
        if not cur.fetchone():
            return []
        cur.execute("PRAGMA table_info(clientes)")
        cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

        col_cod = "cod_cliente" if "cod_cliente" in cols else "''"
        col_nome = "nome_cliente" if "nome_cliente" in cols else "''"
        col_end = "endereco" if "endereco" in cols else "''"
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
            ORDER BY {ordem_sql}
            LIMIT ?
            """,
            (term, like, like, like, vend_f, like_vend, cid_f, like_cid, int(limit)),
        )
        out: List[Dict[str, Any]] = []
        for r in cur.fetchall() or []:
            out.append(
                {
                    "cod_cliente": str(r["cod_cliente"] or "").strip(),
                    "nome_cliente": str(r["nome_cliente"] or "").strip(),
                    "endereco": str(r["endereco"] or "").strip(),
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
        if tipo_estimativa not in ("KG", "CX"):
            tipo_estimativa = "KG"

        if (not motorista_codigo) and motorista_id > 0:
            cur.execute("SELECT COALESCE(codigo,'') AS codigo FROM motoristas WHERE id=? LIMIT 1", (motorista_id,))
            rr = cur.fetchone()
            motorista_codigo = str((rr["codigo"] if rr else "") or "").strip().upper()

        if (not motorista_id) and motorista_codigo:
            cur.execute("SELECT id FROM motoristas WHERE UPPER(TRIM(codigo))=UPPER(TRIM(?)) LIMIT 1", (motorista_codigo,))
            rr = cur.fetchone()
            motorista_id = int(rr["id"] or 0) if rr else 0

        if (not motorista_id) and motorista:
            cur.execute("SELECT id, COALESCE(codigo,'') AS codigo FROM motoristas WHERE UPPER(TRIM(nome))=UPPER(TRIM(?)) LIMIT 1", (motorista,))
            rr = cur.fetchone()
            if rr:
                motorista_id = int(rr["id"] or 0)
                if not motorista_codigo:
                    motorista_codigo = str(rr["codigo"] or "").strip().upper()

        cur.execute(
            """
            SELECT id,
                   COALESCE(status,'') AS status,
                   COALESCE(status_operacional,'') AS status_operacional
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

        status_execucao = {"EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"}
        status_fechado = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}
        status_edicao_desktop = {"", "ATIVA", "ABERTA", "PENDENTE", "PROGRAMADA"}

        # Regra de negocio: desktop edita somente programação ainda não iniciada.
        # Se já está em execução, substituição deve ser via APK.
        if pid > 0:
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
            if "local_rota" in cols_prog:
                sets.append("local_rota=?")
                vals.append(str(payload.local_rota or "").strip().upper())
            if "local_carregamento" in cols_prog:
                sets.append("local_carregamento=?")
                vals.append(str(payload.local_carregamento or "").strip().upper())
            if "adiantamento" in cols_prog:
                sets.append("adiantamento=?")
                vals.append(float(payload.adiantamento or 0.0))
            if "adiantamento_rota" in cols_prog:
                sets.append("adiantamento_rota=?")
                vals.append(float(payload.adiantamento or 0.0))
            if "total_caixas" in cols_prog:
                sets.append("total_caixas=?")
                vals.append(int(payload.total_caixas or 0))
            if "quilos" in cols_prog:
                sets.append("quilos=?")
                vals.append(float(payload.quilos or 0.0))
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
            if "local_rota" in cols_prog:
                col_names.append("local_rota")
                values.append(str(payload.local_rota or "").strip().upper())
            if "local_carregamento" in cols_prog:
                col_names.append("local_carregamento")
                values.append(str(payload.local_carregamento or "").strip().upper())
            if "adiantamento" in cols_prog:
                col_names.append("adiantamento")
                values.append(float(payload.adiantamento or 0.0))
            if "adiantamento_rota" in cols_prog:
                col_names.append("adiantamento_rota")
                values.append(float(payload.adiantamento or 0.0))
            if "total_caixas" in cols_prog:
                col_names.append("total_caixas")
                values.append(int(payload.total_caixas or 0))
            if "quilos" in cols_prog:
                col_names.append("quilos")
                values.append(float(payload.quilos or 0.0))
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
        cur.execute("DELETE FROM programacao_itens WHERE codigo_programacao=?", (codigo,))
        for it in (payload.itens or []):
            cod_cli = str(it.cod_cliente or "").strip().upper()
            nome_cli = str(it.nome_cliente or "").strip().upper()
            if not cod_cli or not nome_cli:
                continue
            cur.execute(
                """
                INSERT INTO programacao_itens
                    (codigo_programacao, cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido, produto)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    codigo,
                    cod_cli,
                    nome_cli,
                    int(it.qnt_caixas or 0),
                    float(it.kg or 0.0),
                    float(it.preco or 0.0),
                    str(it.endereco or "").strip().upper(),
                    str(it.vendedor or "").strip().upper(),
                    str(it.pedido or "").strip().upper(),
                    str(it.produto or "").strip().upper(),
                ),
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
    q: str = Query("", description="Busca por cÃ³digo/nome/cidade"),
    limit: int = Query(200, ge=1, le=1000),
    m=Depends(get_current_motorista),
):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
        if not cur.fetchone():
            return []

        term = (q or "").strip().upper()
        like = f"%{term}%"
        cur.execute(
            """
            SELECT
                TRIM(COALESCE(cod_cliente, '')) AS cod_cliente,
                TRIM(COALESCE(nome_cliente, '')) AS nome_cliente,
                TRIM(COALESCE(cidade, '')) AS cidade,
                TRIM(COALESCE(vendedor, '')) AS vendedor
            FROM clientes
            WHERE
                (? = '')
                OR UPPER(TRIM(COALESCE(cod_cliente, ''))) LIKE ?
                OR UPPER(TRIM(COALESCE(nome_cliente, ''))) LIKE ?
                OR UPPER(TRIM(COALESCE(cidade, ''))) LIKE ?
            ORDER BY UPPER(TRIM(COALESCE(nome_cliente, ''))), UPPER(TRIM(COALESCE(cod_cliente, '')))
            LIMIT ?
            """,
            (term, like, like, like, int(limit)),
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
        raise HTTPException(status_code=400, detail="CÃ³digo da programaÃ§Ã£o Ã© obrigatÃ³rio.")
    if not cod_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente Ã© obrigatÃ³rio.")
    if not nome_cliente:
        raise HTTPException(status_code=400, detail="nome_cliente Ã© obrigatÃ³rio.")
    if qnt_caixas <= 0:
        raise HTTPException(status_code=400, detail="qnt_caixas deve ser maior que zero.")

    pedido = pedido_in or f"RES-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    status_final = status_in if status_in in ("PENDENTE", "ALTERADO", "CANCELADO", "ENTREGUE") else "PENDENTE"

    with get_conn() as conn:
        cur = conn.cursor()
        row_prog = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.codigo_programacao")
        if not row_prog:
            raise HTTPException(status_code=404, detail="ProgramaÃ§Ã£o nÃ£o encontrada para este motorista.")

        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
        if not cur.fetchone():
            raise HTTPException(status_code=500, detail="Tabela programacao_itens nÃ£o encontrada.")

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

    token = create_token(m["codigo"])
    return {"token": token, "nome": m["nome"], "codigo": m["codigo"]}


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
    Lista programaÃ§Ãµes ativas de TODOS os motoristas.
    Isso Ã© necessÃ¡rio para a tela de TRANSFERÃŠNCIA (destino).
    """
    if not ENABLE_ROTAS_ATIVAS_TODAS:
        raise HTTPException(status_code=403, detail="Endpoint desabilitado por configuraÃ§Ã£o.")

    with get_conn() as conn:
        cur = conn.cursor()
        local_expr = _local_rota_expr(conn)
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
                p.motorista,
                p.veiculo,
                p.equipe,
                """ + equipe_id_select + """
                """ + equipe_cols_expr + """,
                """ + local_expr + """,
                """ + media_expr + """,
                """ + kg_carregado_expr + """,
                """ + caixas_carregadas_expr + """,
                """ + caixa_final_expr + """,
                p.data_criacao,
                COALESCE(p.tipo_estimativa, 'KG') AS tipo_estimativa,
                COALESCE(p.caixas_estimado, 0) AS caixas_estimado,
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
            WHERE """ + not_finalized_sql + """
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
            response_rows.append(d)

    return response_rows


@app.get("/rotas/ativas", response_model=List[RotaAtivaOut])
def rotas_ativas(m=Depends(get_current_motorista)):
    with get_conn() as conn:
        cur = conn.cursor()
        local_expr = _local_rota_expr(conn)
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
                p.motorista,
                p.veiculo,
                p.equipe,
                """ + equipe_id_select + """
                """ + equipe_cols_expr + """,
                """ + local_expr + """,
                """ + media_expr + """,
                """ + kg_carregado_expr + """,
                """ + caixas_carregadas_expr + """,
                """ + caixa_final_expr + """,
                p.data_criacao,
                COALESCE(p.tipo_estimativa, 'KG') AS tipo_estimativa,
                COALESCE(p.caixas_estimado, 0) AS caixas_estimado,
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
            raise HTTPException(status_code=404, detail="Rota nÃ£o encontrada para este motorista")

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
        rota = _apply_equipe_nome(rota, equipes_map, cur)
        rota = _decorate_rota_row(rota, cur)
        pend_sub = _has_pending_substituicao(cur, codigo_programacao)
        rota["substituicao_pendente"] = 1 if pend_sub else 0
        rota["status_operacional"] = _status_operacional_especial(rota, pend_substituicao=pend_sub)
        rota["substituicoes"] = _list_substituicoes_por_rota(cur, codigo_programacao, limit=20)

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
        rota = _apply_equipe_nome(rota, equipes_map, cur)
        rota = _decorate_rota_row(rota, cur)
        pend_sub = _has_pending_substituicao(cur, codigo_programacao)
        rota["substituicao_pendente"] = 1 if pend_sub else 0
        rota["status_operacional"] = _status_operacional_especial(rota, pend_substituicao=pend_sub)
        rota["substituicoes"] = _list_substituicoes_por_rota(cur, codigo_programacao, limit=20)

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

            clientes.append(d)

        return {"rota": rota, "clientes": clientes}


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
    codigo_programacao = (codigo_programacao or "").strip()

    cod_cliente = (payload.cod_cliente or "").strip()
    if not cod_cliente:
        raise HTTPException(status_code=400, detail="cod_cliente Ã© obrigatÃ³rio")

    with get_conn() as conn:
        cur = conn.cursor()

        # garante que a rota pertence ao motorista
        pr = _fetch_programacao_owned(cur, codigo_programacao, m, "p.id, p.status")
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nÃ£o encontrada para este motorista")
        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(
                status_code=409,
                detail=f"Rota encerrada. Alteracoes bloqueadas para status {status_atual}.",
            )
        if status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
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
            raise HTTPException(status_code=400, detail="pedido Ã© obrigatÃ³rio para controle do cliente.")
        caixas_atual = payload.caixas_atual
        preco_atual = payload.preco_atual
        alterado_por = (payload.alterado_por or nome_motorista or None)
        alteracao_tipo = (payload.alteracao_tipo or None)
        alteracao_detalhe = (payload.alteracao_detalhe or None)

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
            raise HTTPException(status_code=404, detail="Item de cliente/pedido nÃ£o encontrado na programaÃ§Ã£o.")

        base_caixas = item_base["qnt_caixas"] if item_base else None
        base_preco = item_base["preco"] if item_base else None
        nome_cliente = item_base["nome_cliente"] if item_base else ""
        if pedido is None and item_base:
            pedido = item_base["pedido"]

        allowed_status = {"PENDENTE", "ENTREGUE", "CANCELADO", "ALTERADO"}
        if status_in and status_in not in allowed_status:
            raise HTTPException(status_code=400, detail=f"status_pedido invÃ¡lido: {status_in}.")

        # regra: pedido ENTREGUE nÃ£o pode mais ser alterado
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

        if status_atual == "ENTREGUE":
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
        alterado_em = evento_em if status == "ALTERADO" else None

        # valida faixa de caixas para evitar manipulaÃ§Ã£o indevida
        if caixas_atual is not None:
            try:
                caixas_atual = int(caixas_atual)
            except Exception:
                raise HTTPException(status_code=400, detail="caixas_atual invÃ¡lido.")
            if caixas_atual < 0:
                raise HTTPException(status_code=400, detail="caixas_atual nÃ£o pode ser negativo.")
            if base_caixas is not None:
                try:
                    base_caixas_int = int(base_caixas)
                except Exception:
                    base_caixas_int = None
                if base_caixas_int is not None and caixas_atual > base_caixas_int:
                    raise HTTPException(
                        status_code=400,
                        detail=f"caixas_atual ({caixas_atual}) nÃ£o pode ser maior que caixas do pedido ({base_caixas_int}).",
                    )

        # atualiza controle por cliente (compatÃ­vel com bases sem UNIQUE)
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
                     caixas_atual, preco_atual, alterado_em, alterado_por, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
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
                ),
            )

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
            params.append(caixas_atual if caixas_atual is not None else base_caixas)
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
                    cur.execute(
                        """
                        INSERT INTO recebimentos
                            (codigo_programacao, cod_cliente, pedido, nome_cliente, valor, forma_pagamento, observacao, data_registro)
                        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
                        """,
                        (
                            codigo_programacao,
                            cod_cliente,
                            pedido,
                            nome_cliente,
                            float(valor_recebido),
                            (forma_recebimento or "DINHEIRO"),
                            (obs_recebimento or None),
                        ),
                    )
                else:
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=?",
                        (codigo_programacao, cod_cliente),
                    )
                    cur.execute(
                        """
                        INSERT INTO recebimentos
                            (codigo_programacao, cod_cliente, nome_cliente, valor, forma_pagamento, observacao, data_registro)
                        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
                        """,
                        (
                            codigo_programacao,
                            cod_cliente,
                            nome_cliente,
                            float(valor_recebido),
                            (forma_recebimento or "DINHEIRO"),
                            (obs_recebimento or None),
                        ),
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
            log_payload = payload.dict()
            log_payload.update(
                {
                    "motorista": nome_motorista,
                    "status_pedido": status,
                    "pedido": pedido,
                    "caixas_atual": caixas_atual if caixas_atual is not None else base_caixas,
                    "preco_atual": preco_atual if preco_atual is not None else base_preco,
                    "alterado_por": alterado_por,
                    "alterado_em": alterado_em,
                }
            )
            payload_json = json.dumps(log_payload, ensure_ascii=False)
            cur.execute(
                """
                INSERT INTO programacao_itens_log
                    (codigo_programacao, cod_cliente, evento, payload_json)
                VALUES (?, ?, ?, ?)
                """,
                (codigo_programacao, cod_cliente, "cliente_controle", payload_json),
            )
        except Exception:
            pass

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
        raise HTTPException(status_code=400, detail="CÃ³digo e cliente sÃ£o obrigatÃ³rios")

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
            raise HTTPException(status_code=404, detail="Rota nÃ£o encontrada para este motorista")

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

        cur.execute(
            """
            INSERT INTO rota_gps_pings
                (codigo_programacao, motorista, lat, lon, speed, accuracy, recorded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                codigo_programacao,
                nome_motorista,
                float(payload.lat),
                float(payload.lon),
                (float(payload.speed) if payload.speed is not None else None),
                (float(payload.accuracy) if payload.accuracy is not None else None),
                ts.isoformat(timespec="seconds"),
            ),
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
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "status_operacional", payload.idempotency_key)
        conn.commit()

    return {
        "ok": True,
        "codigo_programacao": codigo_programacao,
        "status_rota": status_rota,
        "status_operacional": status_out,
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
            raise HTTPException(status_code=404, detail="Rota nÃ£o encontrada para este motorista")
        if int(payload.km_inicial or 0) <= 0:
            raise HTTPException(status_code=400, detail="KM inicial deve ser maior que 0.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(status_code=409, detail="Rota jÃ¡ estÃ¡ em andamento.")
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
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "iniciar_rota", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "status": "EM_ROTA"}


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
            raise HTTPException(status_code=404, detail="Rota n?o encontrada para este motorista")
        if int(payload.km_final or 0) < 0:
            raise HTTPException(status_code=400, detail="KM final inv?lido.")

        status_atual = str(pr["status"] or "").strip().upper()
        if status_atual in ("FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"):
            raise HTTPException(status_code=409, detail=f"Rota encerrada (status={status_atual}).")
        if status_atual not in ("EM_ROTA", "EM ROTA", "INICIADA", "EM_ENTREGAS", "EM ENTREGAS", "CARREGADA"):
            raise HTTPException(status_code=409, detail=f"Transi??o inv?lida para finalizar (status={status_atual or 'N/D'}).")
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
        # RECONCILIA??O ANTIFRAUDE
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
            raise HTTPException(
                status_code=409,
                detail="Nao e possivel finalizar: caixas carregadas nao informadas no carregamento.",
            )

        cur.execute(
            """
            SELECT COALESCE(COUNT(*),0)
            FROM transferencias
            WHERE (codigo_origem=? OR codigo_destino=?)
              AND UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
            """,
            (codigo_programacao, codigo_programacao),
        )
        pend_transfer = int((cur.fetchone() or [0])[0] or 0)
        if pend_transfer > 0:
            raise HTTPException(
                status_code=409,
                detail=f"Nao e possivel finalizar: existem {pend_transfer} transferencia(s) pendente(s).",
            )

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
            caixas_expr = "COALESCE(pc.caixas_atual, pi.caixas_atual, pi.qnt_caixas, 0)"
        elif has_pc_caixas_atual and has_pi_qnt_caixas:
            caixas_expr = "COALESCE(pc.caixas_atual, pi.qnt_caixas, 0)"
        elif has_pi_caixas_atual and has_pi_qnt_caixas:
            caixas_expr = "COALESCE(pi.caixas_atual, pi.qnt_caixas, 0)"
        elif has_pc_caixas_atual:
            caixas_expr = "COALESCE(pc.caixas_atual, 0)"
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
        itens = cur.fetchall() or []
        if not itens:
            raise HTTPException(status_code=409, detail="Nao e possivel finalizar: rota sem itens para reconciliacao.")

        total_em_aberto = 0
        itens_em_aberto = []
        pendentes = []
        for it in itens:
            base = _to_int_db(it["base_cx"])
            atual = _to_int_db(it["atual_cx"])
            if base < 0:
                base = 0
            if atual < 0:
                atual = 0

            status_eff = str(it["status_eff"] or "PENDENTE").strip().upper()
            atual_considerado = 0 if status_eff in ("ENTREGUE", "CANCELADO") else atual
            total_em_aberto += atual_considerado
            if atual_considerado > 0:
                itens_em_aberto.append(
                    f"{it['cod_cliente']} / {it['pedido'] or '-'} -> {atual_considerado} cx [{status_eff}]"
                )

            if status_eff == "" or status_eff == "PENDENTE":
                pendentes.append(f"{it['cod_cliente']} / {it['pedido'] or '-'} [{status_eff}]")

        if pendentes:
            raise HTTPException(
                status_code=409,
                detail=f"Nao e possivel finalizar: {len(pendentes)} pedido(s) pendente(s).",
            )

        if total_em_aberto != 0:
            saldo = total_em_aberto
            amostra = "; ".join(itens_em_aberto[:8])
            raise HTTPException(
                status_code=409,
                detail=(
                    "Reconciliacao nao fechou. "
                    f"Carregadas={caixas_carregadas}, saldo em aberto={saldo}. "
                    "Revise cancelamentos/redirecionamentos/entregas. "
                    f"Pedidos com saldo: {amostra}"
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
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "finalizar_rota", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "status": "FINALIZADA"}


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

        # garante que a programaÃ§Ã£o Ã© do motorista logado
        pr = _fetch_programacao_owned(
            cur,
            codigo_programacao,
            m,
            f"p.id, p.status, {carregamento_fechado_sel}, {tipo_estimativa_sel}",
        )
        if not pr:
            raise HTTPException(status_code=404, detail="Rota nÃ£o encontrada para este motorista")
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
        kg_carregado = float(payload.kg_carregado or 0.0)
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

        tipo_estimativa = str(pr["tipo_estimativa"] or "KG").strip().upper()
        if tipo_estimativa not in ("KG", "CX"):
            tipo_estimativa = "KG"
        if tipo_estimativa == "KG" and nf_kg <= 0 and kg_carregado <= 0:
            raise HTTPException(
                status_code=400,
                detail="FOB exige peso informado (nf_kg ou kg_carregado maior que zero).",
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

        # âœ… colunas â€œdesktopâ€/alternativas (se existirem)
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

        # mÃ©dia (sÃ³ se vier)
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
        _idempotency_mark(cur, codigo_motorista, codigo_programacao, "carregamento", payload.idempotency_key)
        conn.commit()

    return {"ok": True, "status": status_result}


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
    return {
        "id": row["id"],
        "status": row["status"],
        "codigo_origem": row["codigo_origem"],
        "codigo_destino": row["codigo_destino"],
        "cod_cliente": row["cod_cliente"],
        "pedido": row["pedido"],
        "qtd_caixas": int(row["qtd_caixas"] or 0),
        "snapshot": _parse_snapshot(row["snapshot"]),
        "obs": row["obs"],
        "motorista_origem": row["motorista_origem"],
        "motorista_destino": row["motorista_destino"],
        "qtd_convertida": int(row["qtd_convertida"] or 0),
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


class TransferenciaCreateIn(BaseModel):
    codigo_destino: str
    pedido: str
    cod_cliente: str
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
        cur.execute(
            """
            SELECT id, nome, codigo
            FROM motoristas
            WHERE TRIM(COALESCE(nome, '')) <> ''
            ORDER BY nome
            """
        )
        out = []
        for r in (cur.fetchall() or []):
            out.append(
                {
                    "id": int(r["id"]) if r["id"] is not None else None,
                    "nome": str(r["nome"] or "").strip().upper(),
                    "codigo": str(r["codigo"] or "").strip().upper(),
                }
            )
        return out


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

        cur.execute(f"SELECT {', '.join(sel)} FROM veiculos ORDER BY placa")
        out = []
        for r in (cur.fetchall() or []):
            placa = str(r["placa"] or "").strip().upper()
            if not placa:
                continue
            d = {"placa": placa}
            if "modelo" in r.keys():
                d["modelo"] = str(r["modelo"] or "").strip().upper()
            if "capacidade_cx" in r.keys():
                d["capacidade_cx"] = r["capacidade_cx"]
            elif "capacidade" in r.keys():
                d["capacidade_cx"] = r["capacidade"]
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


# âœ… CRIAR TRANSFERÃŠNCIA (origem envia)
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
        raise HTTPException(status_code=400, detail="CÃ³digo de origem invÃ¡lido.")
    if not codigo_destino:
        raise HTTPException(status_code=400, detail="CÃ³digo de destino invÃ¡lido.")
    if codigo_destino == codigo_origem:
        raise HTTPException(status_code=400, detail="Destino nÃ£o pode ser igual Ã  origem.")

    if not _rota_pertence_ao_motorista(codigo_origem, m):
        raise HTTPException(status_code=403, detail="Rota de origem nÃ£o pertence ao motorista logado.")

    pedido = (payload.pedido or "").strip()
    cod_cliente = (payload.cod_cliente or "").strip()
    qtd = int(payload.qtd_caixas or 0)

    if not pedido:
        raise HTTPException(status_code=400, detail="Pedido Ã© obrigatÃ³rio.")
    if not cod_cliente:
        raise HTTPException(status_code=400, detail="CÃ³digo do cliente Ã© obrigatÃ³rio.")
    if qtd <= 0:
        raise HTTPException(status_code=400, detail="Quantidade de caixas invÃ¡lida (deve ser > 0).")

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

        if has_pedido_col:
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
        item_origem = cur.fetchone()
        if not item_origem:
            raise HTTPException(status_code=404, detail="Pedido de origem nÃ£o encontrado para transferÃªncia.")
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
                detail=f"TransferÃªncia excede disponÃ­vel do pedido. DisponÃ­vel: {disponivel_liquido} cx.",
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


# âœ… LISTAR TRANSFERÃŠNCIAS (destino recebe)
@app.get("/rotas/{codigo_programacao}/transferencias")
def listar_transferencias(
    codigo_programacao: str,
    status: Optional[str] = Query(default=None),
    m=Depends(get_current_motorista),
):
    nome_motorista = (m["nome"] or "").strip()
    codigo = (codigo_programacao or "").strip()

    if not codigo:
        raise HTTPException(status_code=400, detail="CÃ³digo invÃ¡lido.")

    if not _rota_pertence_ao_motorista(codigo, m):
        raise HTTPException(status_code=403, detail="Rota de destino nÃ£o pertence ao motorista logado.")

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
        raise HTTPException(status_code=400, detail="CÃƒÂ³digo invÃƒÂ¡lido.")

    if not _rota_pertence_ao_motorista(codigo, m):
        raise HTTPException(status_code=403, detail="Rota de origem nÃƒÂ£o pertence ao motorista logado.")

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
        raise HTTPException(status_code=400, detail="ID inv?lido.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="Transfer?ncia n?o encontrada.")

        codigo_destino = str(item.get("codigo_destino", "")).strip()
        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="Transfer?ncia n?o pertence ao motorista logado (destino).")

        st = str(item.get("status", "")).upper().strip()
        if st != "PENDENTE":
            raise HTTPException(status_code=409, detail=f"Transfer?ncia n?o est? pendente (status={st}).")

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
        raise HTTPException(status_code=400, detail="ID invÃ¡lido.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="TransferÃªncia nÃ£o encontrada.")

        codigo_destino = str(item.get("codigo_destino", "")).strip()

        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="TransferÃªncia nÃ£o pertence ao motorista logado (destino).")

        st = str(item.get("status", "")).upper().strip()
        if st != "PENDENTE":
            raise HTTPException(status_code=409, detail=f"TransferÃªncia nÃ£o estÃ¡ pendente (status={st}).")

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
# ===== CONVERTER TRANSFERÃŠNCIA (RESERVA -> PEDIDO) ====
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
        raise HTTPException(status_code=400, detail="ID invÃ¡lido.")

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
            raise HTTPException(status_code=400, detail="nome_cliente Ã© obrigatÃ³rio para novo cliente.")
        if not pedido_novo:
            raise HTTPException(status_code=400, detail="pedido Ã© obrigatÃ³rio para novo cliente.")
        if not cod_novo:
            cod_novo = f"MANUAL-{uuid4().hex[:8].upper()}"
        pedido_dest = pedido_novo
        cod_cli_dest = cod_novo

    if not novo:
        if not pedido_dest:
            raise HTTPException(status_code=400, detail="pedido_destino Ã© obrigatÃ³rio.")
        if not cod_cli_dest:
            raise HTTPException(status_code=400, detail="cod_cliente_destino Ã© obrigatÃ³rio.")
    if qtd <= 0:
        raise HTTPException(status_code=400, detail="qtd_caixas deve ser > 0.")

    with get_conn() as conn:
        cur = conn.cursor()
        item = _fetch_transferencia_by_id(conn, tid)
        if item is None:
            raise HTTPException(status_code=404, detail="TransferÃªncia nÃ£o encontrada.")

        st = str(item.get("status", "")).upper().strip()
        if st != "ACEITA":
            raise HTTPException(status_code=409, detail=f"TransferÃªncia nÃ£o estÃ¡ ACEITA (status={st}).")

        codigo_destino = str(item.get("codigo_destino", "")).strip()
        codigo_origem = str(item.get("codigo_origem", "")).strip()
        if _has_pending_substituicao(cur, codigo_origem) or _has_pending_substituicao(cur, codigo_destino):
            raise HTTPException(
                status_code=409,
                detail="Conversao de transferencia bloqueada enquanto houver substituicao pendente.",
            )
        if not _rota_pertence_ao_motorista(codigo_destino, m):
            raise HTTPException(status_code=403, detail="VocÃª nÃ£o Ã© o motorista destino desta transferÃªncia.")

        total = int(item.get("qtd_caixas") or 0)
        convertido = int(item.get("qtd_convertida") or 0)
        saldo = total - convertido
        if saldo < 0:
            saldo = 0

        if qtd > saldo:
            raise HTTPException(status_code=409, detail=f"Quantidade maior que o saldo disponÃ­vel ({saldo}).")

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

        cur.execute(
            """
            UPDATE transferencias
            SET qtd_convertida=qtd_convertida + ?, atualizado_em=?
            WHERE id=?
            """,
            (qtd, _now_iso(), tid),
        )

        conn.commit()
        return _fetch_transferencia_by_id(conn, tid)


