from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
from typing import Dict


PBKDF2_ITERATIONS = 200_000
DEFAULT_ADMIN_PASSWORD = "123456"
OPERATIONAL_TABLES = [
    "motoristas",
    "vendedores",
    "veiculos",
    "caixas",
    "ajudantes",
    "clientes",
    "fornecedores",
    "equipes",
    "programacoes",
    "programacao_itens",
    "programacao_itens_controle",
    "programacao_itens_log",
    "recebimentos",
    "despesas",
    "rota_gps_override_log",
    "rota_gps_pings",
    "rota_substituicoes",
    "transferencias",
    "transferencias_conversoes",
    "vendas_importadas",
    "programacoes_avulsas",
    "programacoes_avulsas_itens",
    "vendedor_rascunho_itens",
    "vendedor_pre_programacoes",
    "vendedor_pre_programacao_itens",
    "mobile_sync_idempotency",
]
TENANT_SCOPED_TABLES = OPERATIONAL_TABLES + [
    "usuarios",
]

DEFAULT_COMPANY_CODE = "default"
DEFAULT_COMPANY_NAME = "Empresa Inicial"
DEFAULT_PLAN_CODE = "starter"
DEFAULT_PLANS = [
    {
        "code": "starter",
        "name": "Inicial 5 Veiculos",
        "description": "Para quem esta comecando: organiza cadastros, programacoes, recebimentos, despesas e app do motorista em ate 5 veiculos.",
        "monthly_price": 199.0,
        "vehicle_limit": 5,
        "user_limit": 6,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": False,
            "centro_custos": False,
            "relatorios": True,
            "rotas": False,
            "escala": False,
            "app_motorista": True,
            "realtime_tracking": False,
            "financial_reports": False,
            "advanced_reports": False,
            "api_access": False,
        },
    },
    {
        "code": "growth",
        "name": "Crescimento 10 Veiculos",
        "description": "Para frotas que precisam enxergar perdas, rotas e custos: inclui mortalidade, centro de custos e financeiro em ate 10 veiculos.",
        "monthly_price": 399.0,
        "vehicle_limit": 10,
        "user_limit": 15,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": True,
            "centro_custos": True,
            "relatorios": False,
            "rotas": True,
            "escala": False,
            "app_motorista": True,
            "realtime_tracking": True,
            "financial_reports": True,
            "advanced_reports": False,
            "api_access": False,
        },
    },
    {
        "code": "professional",
        "name": "Profissional 15 Veiculos",
        "description": "Para operacoes mais exigentes: adiciona escala, relatorios avancados, rotas e controles completos em ate 15 veiculos.",
        "monthly_price": 699.0,
        "vehicle_limit": 15,
        "user_limit": 30,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": True,
            "centro_custos": True,
            "relatorios": True,
            "rotas": True,
            "escala": True,
            "app_motorista": True,
            "realtime_tracking": True,
            "financial_reports": True,
            "advanced_reports": True,
            "api_access": False,
        },
    },
    {
        "code": "enterprise",
        "name": "Empresarial Mais Veiculos",
        "description": "Para empresas com mais de 15 veiculos, necessidade de API, suporte prioritario e contrato ajustado a operacao.",
        "monthly_price": 999.0,
        "vehicle_limit": None,
        "user_limit": None,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": True,
            "centro_custos": True,
            "relatorios": True,
            "rotas": True,
            "escala": True,
            "app_motorista": True,
            "realtime_tracking": True,
            "financial_reports": True,
            "advanced_reports": True,
            "api_access": True,
            "custom_contract": True,
            "priority_support": True,
        },
    },
    {
        "code": "corporate_private",
        "name": "Corporativo Privado 50",
        "description": "Plano privado para implantacao corporativa com acesso total e ate 50 veiculos.",
        "monthly_price": 0.0,
        "vehicle_limit": 50,
        "user_limit": None,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": True,
            "centro_custos": True,
            "relatorios": True,
            "rotas": True,
            "escala": True,
            "app_motorista": True,
            "realtime_tracking": True,
            "financial_reports": True,
            "advanced_reports": True,
            "api_access": True,
            "custom_contract": True,
            "priority_support": True,
            "private_deployment": True,
        },
    },
    {
        "code": "internal",
        "name": "Internal",
        "description": "Plano interno para desenvolvimento e testes.",
        "monthly_price": 0.0,
        "vehicle_limit": None,
        "user_limit": None,
        "features": {
            "cadastros": True,
            "importar_vendas": True,
            "programacao": True,
            "recebimentos": True,
            "despesas": True,
            "mortalidade": True,
            "centro_custos": True,
            "relatorios": True,
            "rotas": True,
            "escala": True,
            "app_motorista": True,
            "realtime_tracking": True,
            "financial_reports": True,
            "advanced_reports": True,
            "api_access": True,
            "custom_contract": True,
            "private_deployment": True,
        },
    },
]


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


def table_exists(cur: sqlite3.Cursor, table: str) -> bool:
    cur.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1", (table,))
    return cur.fetchone() is not None


def count_rows(cur: sqlite3.Cursor, table: str) -> int:
    cur.execute(f'SELECT COUNT(*) FROM "{table}"')
    row = cur.fetchone()
    return int(row[0] if row else 0)


def _safe_add_column(cur: sqlite3.Cursor, table: str, col: str, coltype: str) -> None:
    cur.execute(f"PRAGMA table_info({table})")
    cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if col.lower() not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")


def _table_columns(cur: sqlite3.Cursor, table: str) -> set[str]:
    cur.execute(f"PRAGMA table_info({table})")
    return {str(r[1]).lower() for r in (cur.fetchall() or [])}


def _safe_add_columns(cur: sqlite3.Cursor, table: str, columns: Dict[str, str]) -> None:
    cols = _table_columns(cur, table)
    for col, coltype in columns.items():
        if col.lower() not in cols:
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
            cols.add(col.lower())


def _normalize_existing_programacoes(cur: sqlite3.Cursor) -> None:
    if not table_exists(cur, "programacoes"):
        return

    _safe_add_columns(
        cur,
        "programacoes",
        {
            "codigo_programacao": "TEXT",
            "codigo": "TEXT",
            "data": "TEXT",
            "data_criacao": "TEXT",
            "motorista_id": "INTEGER",
            "motorista_codigo": "TEXT",
            "codigo_motorista": "TEXT",
            "veiculo": "TEXT",
            "equipe": "TEXT",
            "total_caixas": "INTEGER DEFAULT 0",
            "quilos": "REAL DEFAULT 0",
            "kg_estimado": "REAL DEFAULT 0",
            "tipo_estimativa": "TEXT DEFAULT 'KG'",
            "caixas_estimado": "INTEGER DEFAULT 0",
            "operacao_tipo": "TEXT DEFAULT 'VENDA'",
            "transbordo_modalidade": "TEXT",
            "transbordo_observacao": "TEXT",
            "transbordo_grupo": "TEXT",
            "status": "TEXT DEFAULT 'ATIVA'",
            "status_operacional": "TEXT",
            "status_operacional_obs": "TEXT",
            "status_operacional_em": "TEXT",
            "status_operacional_por": "TEXT",
            "finalizada_no_app": "INTEGER DEFAULT 0",
            "prestacao_status": "TEXT DEFAULT 'PENDENTE'",
            "local_rota": "TEXT",
            "tipo_rota": "TEXT",
            "local_carregamento": "TEXT",
            "granja_carregada": "TEXT",
            "local_carregado": "TEXT",
            "local_carreg": "TEXT",
            "num_nf": "TEXT",
            "nf_numero": "TEXT",
            "nf_caixas": "INTEGER DEFAULT 0",
            "caixas_carregadas": "INTEGER DEFAULT 0",
            "qnt_cx_carregada": "INTEGER DEFAULT 0",
            "kg_carregado": "REAL DEFAULT 0",
            "nf_kg": "REAL DEFAULT 0",
            "nf_kg_carregado": "REAL DEFAULT 0",
            "nf_saldo": "REAL DEFAULT 0",
            "data_saida": "TEXT",
            "hora_saida": "TEXT",
            "saida_dt": "TEXT",
            "data_chegada": "TEXT",
            "hora_chegada": "TEXT",
            "chegada_dt": "TEXT",
            "km_final": "REAL DEFAULT 0",
            "foto_doa_path": "TEXT",
            "doa_foto_path": "TEXT",
            "mortalidade_transbordo_foto_path": "TEXT",
            "foto_doa_ref_json": "TEXT",
            "mortalidade_transbordo_aves": "INTEGER DEFAULT 0",
            "mortalidade_transbordo_kg": "REAL DEFAULT 0",
            "obs_transbordo": "TEXT",
            "ajudantes_alteracao_motivo": "TEXT",
            "ajudantes_alterado_em": "TEXT",
            "historico_ajudantes": "TEXT",
        },
    )

    cur.execute("CREATE INDEX IF NOT EXISTS idx_programacoes_codigo_programacao ON programacoes(codigo_programacao)")
    cur.execute(
        """
        UPDATE programacoes
           SET codigo_programacao=TRIM(COALESCE(codigo, ''))
         WHERE TRIM(COALESCE(codigo_programacao, ''))=''
           AND TRIM(COALESCE(codigo, ''))<>''
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET codigo=TRIM(COALESCE(codigo_programacao, ''))
         WHERE TRIM(COALESCE(codigo, ''))=''
           AND TRIM(COALESCE(codigo_programacao, ''))<>''
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET data_criacao=COALESCE(NULLIF(TRIM(data_criacao), ''), NULLIF(TRIM(data), ''), date('now')),
               data=COALESCE(NULLIF(TRIM(data), ''), NULLIF(TRIM(data_criacao), ''), date('now')),
               status=UPPER(COALESCE(NULLIF(TRIM(status), ''), 'ATIVA')),
               prestacao_status=UPPER(COALESCE(NULLIF(TRIM(prestacao_status), ''), 'PENDENTE')),
               tipo_estimativa=UPPER(COALESCE(NULLIF(TRIM(tipo_estimativa), ''), 'KG'))
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET local_rota=COALESCE(NULLIF(TRIM(local_rota), ''), NULLIF(TRIM(tipo_rota), '')),
               tipo_rota=COALESCE(NULLIF(TRIM(tipo_rota), ''), NULLIF(TRIM(local_rota), '')),
               local_carregamento=COALESCE(NULLIF(TRIM(local_carregamento), ''), NULLIF(TRIM(granja_carregada), ''), NULLIF(TRIM(local_carregado), ''), NULLIF(TRIM(local_carreg), '')),
               granja_carregada=COALESCE(NULLIF(TRIM(granja_carregada), ''), NULLIF(TRIM(local_carregamento), '')),
               local_carregado=COALESCE(NULLIF(TRIM(local_carregado), ''), NULLIF(TRIM(local_carregamento), '')),
               local_carreg=COALESCE(NULLIF(TRIM(local_carreg), ''), NULLIF(TRIM(local_carregamento), '')),
               nf_numero=COALESCE(NULLIF(TRIM(nf_numero), ''), NULLIF(TRIM(num_nf), '')),
               num_nf=COALESCE(NULLIF(TRIM(num_nf), ''), NULLIF(TRIM(nf_numero), '')),
               saida_dt=COALESCE(NULLIF(TRIM(saida_dt), ''), TRIM(COALESCE(data_saida, '') || ' ' || COALESCE(hora_saida, ''))),
               chegada_dt=COALESCE(NULLIF(TRIM(chegada_dt), ''), TRIM(COALESCE(data_chegada, '') || ' ' || COALESCE(hora_chegada, '')))
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET operacao_tipo=CASE
                   WHEN UPPER(TRIM(COALESCE(operacao_tipo, '')))='TRANSBORDO'
                     OR UPPER(TRIM(COALESCE(tipo_estimativa, '')))='CX'
                     OR TRIM(COALESCE(transbordo_grupo, ''))<>''
                   THEN 'TRANSBORDO'
                   ELSE 'VENDA'
               END
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET transbordo_modalidade=CASE
                   WHEN UPPER(TRIM(COALESCE(transbordo_modalidade, '')))='FOB' THEN 'EMPRESA_BUSCA'
                   ELSE COALESCE(NULLIF(TRIM(transbordo_modalidade), ''), 'EMPRESA_BUSCA')
               END,
               transbordo_grupo=COALESCE(NULLIF(TRIM(transbordo_grupo), ''), NULLIF(TRIM(codigo_programacao), ''), NULLIF(TRIM(codigo), ''))
         WHERE UPPER(TRIM(COALESCE(operacao_tipo, '')))='TRANSBORDO'
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET status='FINALIZADA',
               status_operacional='FINALIZADA',
               finalizada_no_app=1
         WHERE UPPER(TRIM(COALESCE(prestacao_status, '')))='FECHADA'
            OR UPPER(TRIM(COALESCE(status, ''))) IN ('FINALIZADA','FINALIZADO')
            OR UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('FINALIZADA','FINALIZADO')
            OR COALESCE(finalizada_no_app, 0)=1
            OR TRIM(COALESCE(data_chegada, ''))<>''
            OR TRIM(COALESCE(hora_chegada, ''))<>''
            OR COALESCE(km_final, 0)>0
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET status='CANCELADA',
               status_operacional='CANCELADA',
               finalizada_no_app=1
         WHERE UPPER(TRIM(COALESCE(status, ''))) IN ('CANCELADA','CANCELADO')
            OR UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('CANCELADA','CANCELADO')
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET status_operacional=NULL
         WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
           AND UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
           AND UPPER(TRIM(COALESCE(prestacao_status, 'PENDENTE'))) <> 'FECHADA'
           AND COALESCE(finalizada_no_app, 0)=0
           AND TRIM(COALESCE(data_chegada, ''))=''
           AND TRIM(COALESCE(hora_chegada, ''))=''
           AND COALESCE(km_final, 0)=0
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET motorista_codigo=COALESCE(NULLIF(TRIM(motorista_codigo), ''), NULLIF(TRIM(codigo_motorista), '')),
               codigo_motorista=COALESCE(NULLIF(TRIM(codigo_motorista), ''), NULLIF(TRIM(motorista_codigo), ''))
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET foto_doa_path=COALESCE(NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(doa_foto_path), ''), NULLIF(TRIM(mortalidade_transbordo_foto_path), '')),
               doa_foto_path=COALESCE(NULLIF(TRIM(doa_foto_path), ''), NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(mortalidade_transbordo_foto_path), '')),
               mortalidade_transbordo_foto_path=COALESCE(NULLIF(TRIM(mortalidade_transbordo_foto_path), ''), NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(doa_foto_path), ''))
        """
    )
    cur.execute(
        """
        UPDATE programacoes
           SET nf_saldo=ROUND(MAX(COALESCE(nf_kg, 0) - COALESCE(NULLIF(nf_kg_carregado, 0), kg_carregado, 0), 0), 2)
         WHERE COALESCE(nf_kg, 0)>0
        """
    )

    if table_exists(cur, "programacao_itens"):
        cur.execute(
            """
            UPDATE programacoes
               SET total_caixas=COALESCE(NULLIF(total_caixas, 0), (
                       SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                         FROM programacao_itens pi
                        WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                   ), 0),
                   nf_caixas=COALESCE(NULLIF(nf_caixas, 0), (
                       SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                         FROM programacao_itens pi
                        WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                   ), 0),
                   caixas_carregadas=COALESCE(NULLIF(caixas_carregadas, 0), (
                       SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                         FROM programacao_itens pi
                        WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                   ), 0),
                   qnt_cx_carregada=COALESCE(NULLIF(qnt_cx_carregada, 0), (
                       SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                         FROM programacao_itens pi
                        WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                   ), 0)
             WHERE TRIM(COALESCE(codigo_programacao, ''))<>''
               AND EXISTS (
                   SELECT 1 FROM programacao_itens pi
                    WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
               )
            """
        )


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


def ensure_tenant_columns(conn: sqlite3.Connection, company_id: int | None = None) -> Dict[str, int]:
    """Garante company_id nas tabelas legadas e faz backfill para a empresa inicial."""
    cur = conn.cursor()
    if company_id is None:
        cur.execute("SELECT id FROM companies ORDER BY id ASC LIMIT 1")
        row = cur.fetchone()
        if row is None:
            return {}
        company_id = int(row[0])

    result: Dict[str, int] = {}
    for table in TENANT_SCOPED_TABLES:
        if not table_exists(cur, table):
            continue
        cols = _table_columns(cur, table)
        if "company_id" not in cols:
            row_count = count_rows(cur, table)
            cur.execute(f'ALTER TABLE "{table}" ADD COLUMN company_id INTEGER DEFAULT {int(company_id)}')
            result[table] = row_count
        else:
            cur.execute(f'SELECT COUNT(*) FROM "{table}" WHERE company_id IS NULL OR company_id=0')
            row = cur.fetchone()
            pending = int(row[0] if row else 0)
            if pending:
                cur.execute(
                    f'UPDATE "{table}" SET company_id=? WHERE company_id IS NULL OR company_id=0',
                    (int(company_id),),
                )
            result[table] = pending
        cur.execute(f'CREATE INDEX IF NOT EXISTS idx_{table}_company_id ON "{table}"(company_id)')
    return result


def ensure_saas_schema(conn: sqlite3.Connection) -> int:
    """Cria a base SaaS multiempresa sem alterar dados operacionais existentes."""
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            legal_name TEXT,
            document TEXT,
            email TEXT,
            phone TEXT,
            status TEXT NOT NULL DEFAULT 'active',
            timezone TEXT NOT NULL DEFAULT 'America/Fortaleza',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )
        """
    )
    _safe_add_columns(
        cur,
        "companies",
        {
            "code": "TEXT",
            "name": "TEXT",
            "legal_name": "TEXT",
            "document": "TEXT",
            "email": "TEXT",
            "phone": "TEXT",
            "status": "TEXT",
            "timezone": "TEXT",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        },
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            description TEXT,
            monthly_price REAL NOT NULL DEFAULT 0,
            vehicle_limit INTEGER,
            user_limit INTEGER,
            features_json TEXT NOT NULL DEFAULT '{}',
            status TEXT NOT NULL DEFAULT 'active',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )
        """
    )
    _safe_add_columns(
        cur,
        "plans",
        {
            "code": "TEXT",
            "name": "TEXT",
            "description": "TEXT",
            "monthly_price": "REAL DEFAULT 0",
            "vehicle_limit": "INTEGER",
            "user_limit": "INTEGER",
            "features_json": "TEXT DEFAULT '{}'",
            "status": "TEXT DEFAULT 'active'",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        },
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS subscriptions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            plan_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'active',
            billing_cycle TEXT NOT NULL DEFAULT 'monthly',
            current_period_start TEXT,
            current_period_end TEXT,
            next_due_date TEXT,
            cancelled_at TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(plan_id) REFERENCES plans(id)
        )
        """
    )
    _safe_add_columns(
        cur,
        "subscriptions",
        {
            "company_id": "INTEGER",
            "plan_id": "INTEGER",
            "status": "TEXT DEFAULT 'active'",
            "billing_cycle": "TEXT DEFAULT 'monthly'",
            "current_period_start": "TEXT",
            "current_period_end": "TEXT",
            "next_due_date": "TEXT",
            "cancelled_at": "TEXT",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        },
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subscription_id INTEGER,
            company_id INTEGER NOT NULL,
            amount REAL NOT NULL DEFAULT 0,
            due_date TEXT,
            paid_at TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            method TEXT,
            reference TEXT,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(subscription_id) REFERENCES subscriptions(id),
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )
        """
    )
    _safe_add_columns(
        cur,
        "payments",
        {
            "subscription_id": "INTEGER",
            "company_id": "INTEGER",
            "amount": "REAL DEFAULT 0",
            "due_date": "TEXT",
            "paid_at": "TEXT",
            "status": "TEXT DEFAULT 'pending'",
            "method": "TEXT",
            "reference": "TEXT",
            "notes": "TEXT",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        },
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            user_id INTEGER,
            actor_type TEXT,
            action TEXT NOT NULL,
            entity_type TEXT,
            entity_id TEXT,
            severity TEXT NOT NULL DEFAULT 'info',
            ip_address TEXT,
            metadata_json TEXT NOT NULL DEFAULT '{}',
            created_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )
        """
    )
    _safe_add_columns(
        cur,
        "audit_logs",
        {
            "company_id": "INTEGER",
            "user_id": "INTEGER",
            "actor_type": "TEXT",
            "action": "TEXT",
            "entity_type": "TEXT",
            "entity_id": "TEXT",
            "severity": "TEXT DEFAULT 'info'",
            "ip_address": "TEXT",
            "metadata_json": "TEXT DEFAULT '{}'",
            "created_at": "TEXT",
        },
    )

    index_statements = [
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_companies_code ON companies(code)",
        "CREATE INDEX IF NOT EXISTS idx_companies_status ON companies(status)",
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_plans_code ON plans(code)",
        "CREATE INDEX IF NOT EXISTS idx_plans_status ON plans(status)",
        "CREATE INDEX IF NOT EXISTS idx_subscriptions_company_status ON subscriptions(company_id, status)",
        "CREATE INDEX IF NOT EXISTS idx_subscriptions_plan ON subscriptions(plan_id)",
        "CREATE INDEX IF NOT EXISTS idx_payments_company_status ON payments(company_id, status)",
        "CREATE INDEX IF NOT EXISTS idx_payments_subscription ON payments(subscription_id)",
        "CREATE INDEX IF NOT EXISTS idx_payments_due_date ON payments(due_date)",
        "CREATE INDEX IF NOT EXISTS idx_audit_logs_company_created ON audit_logs(company_id, created_at)",
        "CREATE INDEX IF NOT EXISTS idx_audit_logs_action ON audit_logs(action)",
    ]
    for sql in index_statements:
        cur.execute(sql)

    for plan in DEFAULT_PLANS:
        features_json = json.dumps(plan["features"], ensure_ascii=True, separators=(",", ":"), sort_keys=True)
        cur.execute(
            """
            INSERT INTO plans (
                code, name, description, monthly_price, vehicle_limit, user_limit,
                features_json, status, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, 'active', datetime('now'), datetime('now'))
            ON CONFLICT(code) DO UPDATE SET
                name=excluded.name,
                description=excluded.description,
                monthly_price=excluded.monthly_price,
                vehicle_limit=excluded.vehicle_limit,
                user_limit=excluded.user_limit,
                features_json=excluded.features_json,
                status='active',
                updated_at=datetime('now')
            """,
            (
                plan["code"],
                plan["name"],
                plan["description"],
                plan["monthly_price"],
                plan["vehicle_limit"],
                plan["user_limit"],
                features_json,
            ),
        )

    company_code = (os.environ.get("ROTA_DEFAULT_COMPANY_CODE") or DEFAULT_COMPANY_CODE).strip() or DEFAULT_COMPANY_CODE
    company_name = (os.environ.get("ROTA_DEFAULT_COMPANY_NAME") or DEFAULT_COMPANY_NAME).strip() or DEFAULT_COMPANY_NAME
    cur.execute("SELECT id FROM companies WHERE code=? LIMIT 1", (company_code,))
    company_row = cur.fetchone()
    if company_row is None:
        cur.execute(
            """
            INSERT INTO companies (code, name, status, timezone, created_at, updated_at)
            VALUES (?, ?, 'active', 'America/Fortaleza', datetime('now'), datetime('now'))
            """,
            (company_code, company_name),
        )
        company_id = int(cur.lastrowid)
    else:
        company_id = int(company_row[0])
        cur.execute(
            """
            UPDATE companies
            SET
                name=COALESCE(NULLIF(name, ''), ?),
                status=COALESCE(NULLIF(status, ''), 'active'),
                timezone=COALESCE(NULLIF(timezone, ''), 'America/Fortaleza'),
                updated_at=datetime('now')
            WHERE id=?
            """,
            (company_name, company_id),
        )

    cur.execute("SELECT id FROM plans WHERE code=? LIMIT 1", (DEFAULT_PLAN_CODE,))
    plan_row = cur.fetchone()
    if plan_row is not None:
        plan_id = int(plan_row[0])
        cur.execute(
            """
            SELECT id
            FROM subscriptions
            WHERE company_id=? AND status IN ('active', 'trialing', 'past_due')
            ORDER BY id ASC
            LIMIT 1
            """,
            (company_id,),
        )
        if cur.fetchone() is None:
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
                (company_id, plan_id),
            )

    cur.execute(
        """
        SELECT id
        FROM audit_logs
        WHERE company_id=? AND action='saas_schema_initialized'
        LIMIT 1
        """,
        (company_id,),
    )
    if cur.fetchone() is None:
        cur.execute(
            """
            INSERT INTO audit_logs (
                company_id, actor_type, action, entity_type, entity_id, severity,
                metadata_json, created_at
            )
            VALUES (?, 'system', 'saas_schema_initialized', 'company', ?, 'info', ?, datetime('now'))
            """,
            (
                company_id,
                str(company_id),
                json.dumps({"default_company_code": company_code}, ensure_ascii=True, sort_keys=True),
            ),
        )
    return company_id


def ensure_core_schema(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            permissoes TEXT,
            cpf TEXT,
            telefone TEXT,
            codigo TEXT,
            senha TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS motoristas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            cpf TEXT,
            telefone TEXT,
            codigo TEXT,
            senha TEXT,
            perfil_app TEXT DEFAULT 'MOTORISTA',
            acesso_liberado INTEGER DEFAULT 0,
            acesso_liberado_por TEXT,
            acesso_liberado_em TEXT,
            acesso_obs TEXT,
            status TEXT DEFAULT 'ATIVO'
        )
        """
    )
    cur.execute("PRAGMA table_info(motoristas)")
    motoristas_cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if "perfil_app" not in motoristas_cols:
        cur.execute("ALTER TABLE motoristas ADD COLUMN perfil_app TEXT DEFAULT 'MOTORISTA'")
    cur.execute(
        """
        UPDATE motoristas
        SET perfil_app='MOTORISTA'
        WHERE perfil_app IS NULL OR TRIM(perfil_app)=''
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS vendedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            nome TEXT,
            telefone TEXT,
            cidade_base TEXT,
            status TEXT DEFAULT 'ATIVO',
            senha TEXT,
            ultimo_login_em TEXT,
            ultimo_login_ip TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS veiculos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            placa TEXT,
            modelo TEXT,
            capacidade_cx INTEGER,
            status TEXT DEFAULT 'ATIVO'
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS caixas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            lote TEXT,
            cor TEXT,
            veiculo_placa TEXT,
            status TEXT DEFAULT 'EM_ESTOQUE',
            data_compra TEXT,
            observacao TEXT,
            company_id INTEGER
        )
        """
    )
    _safe_add_columns(
        cur,
        "caixas",
        {
            "codigo": "TEXT",
            "lote": "TEXT",
            "cor": "TEXT",
            "veiculo_placa": "TEXT",
            "status": "TEXT DEFAULT 'EM_ESTOQUE'",
            "data_compra": "TEXT",
            "observacao": "TEXT",
            "company_id": "INTEGER",
        },
    )
    cur.execute(
        """
        UPDATE caixas
           SET codigo=UPPER(TRIM(COALESCE(NULLIF(codigo, ''), 'CX-' || printf('%05d', id)))),
               lote=UPPER(TRIM(COALESCE(lote, ''))),
               cor=UPPER(TRIM(COALESCE(cor, ''))),
               veiculo_placa=UPPER(TRIM(COALESCE(veiculo_placa, ''))),
               status=UPPER(TRIM(COALESCE(NULLIF(status, ''), 'EM_ESTOQUE')))
        """
    )
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_caixas_codigo ON caixas(codigo)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_caixas_lote ON caixas(lote)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_caixas_veiculo ON caixas(veiculo_placa)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_caixas_status ON caixas(status)")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS caixas_movimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            caixa_id INTEGER,
            codigo TEXT,
            movimento TEXT,
            veiculo_origem TEXT,
            veiculo_destino TEXT,
            status_origem TEXT,
            status_destino TEXT,
            observacao TEXT,
            criado_em TEXT,
            company_id INTEGER
        )
        """
    )
    _safe_add_columns(
        cur,
        "caixas_movimentos",
        {
            "caixa_id": "INTEGER",
            "codigo": "TEXT",
            "movimento": "TEXT",
            "veiculo_origem": "TEXT",
            "veiculo_destino": "TEXT",
            "status_origem": "TEXT",
            "status_destino": "TEXT",
            "observacao": "TEXT",
            "criado_em": "TEXT",
            "company_id": "INTEGER",
        },
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_caixas_movimentos_caixa ON caixas_movimentos(caixa_id, criado_em)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_caixas_movimentos_codigo ON caixas_movimentos(codigo, criado_em)")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS ajudantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            sobrenome TEXT,
            telefone TEXT,
            status TEXT DEFAULT 'ATIVO'
        )
        """
    )
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
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS fornecedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            razao_social TEXT NOT NULL,
            nome_fantasia TEXT,
            documento TEXT,
            tipo_pessoa TEXT DEFAULT 'CNPJ',
            perfil_fornecedor TEXT DEFAULT 'OUTROS',
            telefone TEXT,
            email TEXT,
            cidade TEXT,
            uf TEXT,
            status TEXT DEFAULT 'ATIVO',
            observacao TEXT,
            certificado_nome TEXT,
            certificado_path TEXT,
            certificado_status TEXT DEFAULT 'NAO_INSTALADO',
            certificado_instalado_em TEXT,
            company_id INTEGER
        )
        """
    )
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
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fornecedor_perfis_status ON fornecedor_perfis(status)")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS produtos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            nome TEXT NOT NULL,
            descricao TEXT,
            categoria TEXT DEFAULT 'AVES',
            unidade TEXT DEFAULT 'KG',
            unidade_estoque TEXT DEFAULT 'KG',
            controla_estoque_fisico INTEGER DEFAULT 1,
            controla_estoque_fiscal INTEGER DEFAULT 1,
            estoque_min_kg REAL DEFAULT 0,
            estoque_min_caixas INTEGER DEFAULT 0,
            ncm TEXT,
            cest TEXT,
            cfop_entrada TEXT,
            cfop_saida TEXT,
            ean TEXT,
            custo_padrao REAL DEFAULT 0,
            preco_padrao REAL DEFAULT 0,
            status TEXT DEFAULT 'ATIVO',
            company_id INTEGER
        )
        """
    )
    _safe_add_columns(
        cur,
        "produtos",
        {
            "codigo": "TEXT",
            "nome": "TEXT",
            "descricao": "TEXT",
            "categoria": "TEXT DEFAULT 'AVES'",
            "unidade": "TEXT DEFAULT 'KG'",
            "unidade_estoque": "TEXT DEFAULT 'KG'",
            "controla_estoque_fisico": "INTEGER DEFAULT 1",
            "controla_estoque_fiscal": "INTEGER DEFAULT 1",
            "estoque_min_kg": "REAL DEFAULT 0",
            "estoque_min_caixas": "INTEGER DEFAULT 0",
            "ncm": "TEXT",
            "cest": "TEXT",
            "cfop_entrada": "TEXT",
            "cfop_saida": "TEXT",
            "ean": "TEXT",
            "custo_padrao": "REAL DEFAULT 0",
            "preco_padrao": "REAL DEFAULT 0",
            "status": "TEXT DEFAULT 'ATIVO'",
            "company_id": "INTEGER",
        },
    )
    cur.execute(
        """
        UPDATE produtos
           SET categoria=COALESCE(NULLIF(TRIM(categoria), ''), 'AVES'),
               unidade=COALESCE(NULLIF(TRIM(unidade), ''), 'KG'),
               unidade_estoque=COALESCE(NULLIF(TRIM(unidade_estoque), ''), 'KG'),
               controla_estoque_fisico=COALESCE(controla_estoque_fisico, 1),
               controla_estoque_fiscal=COALESCE(controla_estoque_fiscal, 1),
               estoque_min_kg=COALESCE(estoque_min_kg, 0),
               estoque_min_caixas=COALESCE(estoque_min_caixas, 0),
               status=COALESCE(NULLIF(TRIM(status), ''), 'ATIVO')
         WHERE categoria IS NULL OR TRIM(categoria)=''
            OR unidade IS NULL OR TRIM(unidade)=''
            OR unidade_estoque IS NULL OR TRIM(unidade_estoque)=''
            OR controla_estoque_fisico IS NULL
            OR controla_estoque_fiscal IS NULL
            OR estoque_min_kg IS NULL
            OR estoque_min_caixas IS NULL
            OR status IS NULL OR TRIM(status)=''
        """
    )
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_produtos_codigo ON produtos(codigo)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_produtos_nome ON produtos(nome)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_produtos_status ON produtos(status)")
    _safe_add_columns(
        cur,
        "fornecedores",
        {
            "razao_social": "TEXT",
            "nome_fantasia": "TEXT",
            "documento": "TEXT",
            "tipo_pessoa": "TEXT DEFAULT 'CNPJ'",
            "perfil_fornecedor": "TEXT DEFAULT 'OUTROS'",
            "telefone": "TEXT",
            "email": "TEXT",
            "cidade": "TEXT",
            "uf": "TEXT",
            "status": "TEXT DEFAULT 'ATIVO'",
            "observacao": "TEXT",
            "certificado_nome": "TEXT",
            "certificado_path": "TEXT",
            "certificado_status": "TEXT DEFAULT 'NAO_INSTALADO'",
            "certificado_instalado_em": "TEXT",
            "company_id": "INTEGER",
        },
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fornecedores_documento ON fornecedores(documento)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fornecedores_perfil ON fornecedores(perfil_fornecedor)")
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
    _safe_add_columns(
        cur,
        "compras_nfe",
        {
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
        },
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_numero ON compras_nfe(numero)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_vinculo ON compras_nfe(codigo_programacao)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compras_nfe_fornecedor ON compras_nfe(fornecedor_documento)")
    cur.execute("PRAGMA table_info(clientes)")
    clientes_cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    clientes_expected_cols = {
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
    for col, decl in clientes_expected_cols.items():
        if col not in clientes_cols:
            cur.execute(f"ALTER TABLE clientes ADD COLUMN {col} {decl}")
            clientes_cols.add(col)
    try:
        if "company_id" in clientes_cols:
            _drop_single_column_unique_index(cur, "clientes", "cod_cliente")
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_company_cod ON clientes(company_id, cod_cliente)")
        else:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_cod ON clientes(cod_cliente)")
    except Exception:
        pass
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS equipes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            ajudante1 TEXT,
            ajudante2 TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS vendedor_rascunho_itens (
            id TEXT PRIMARY KEY,
            cod_cliente TEXT NOT NULL,
            nome_cliente TEXT NOT NULL,
            cidade TEXT,
            bairro TEXT,
            endereco TEXT,
            vendedor_cadastro TEXT,
            vendedor_origem TEXT NOT NULL,
            preco REAL DEFAULT 0,
            caixas INTEGER DEFAULT 0,
            status TEXT DEFAULT 'PENDENTE',
            observacao TEXT DEFAULT '',
            alerta_codigo_programacao TEXT,
            alerta_status_rota TEXT,
            criado_em TEXT DEFAULT (datetime('now')),
            atualizado_em TEXT DEFAULT (datetime('now')),
            criado_por_codigo TEXT,
            atualizado_por_codigo TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS vendedor_pre_programacoes (
            id TEXT PRIMARY KEY,
            titulo TEXT NOT NULL,
            observacao TEXT DEFAULT '',
            status TEXT DEFAULT 'ABERTA',
            criado_em TEXT DEFAULT (datetime('now')),
            atualizado_em TEXT DEFAULT (datetime('now')),
            criado_por_codigo TEXT,
            atualizado_por_codigo TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS vendedor_pre_programacao_itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pre_programacao_id TEXT NOT NULL,
            rascunho_item_id TEXT NOT NULL,
            ordem INTEGER DEFAULT 0,
            criado_em TEXT DEFAULT (datetime('now')),
            atualizado_em TEXT DEFAULT (datetime('now')),
            UNIQUE(pre_programacao_id, rascunho_item_id)
        )
        """
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_pre_programacao_itens_pp ON vendedor_pre_programacao_itens(pre_programacao_id, ordem, id)"
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_pre_programacao_itens_rascunho ON vendedor_pre_programacao_itens(rascunho_item_id)"
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS programacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_programacao TEXT,
            data_criacao TEXT,
            motorista TEXT,
            veiculo TEXT,
            equipe TEXT,
            kg_estimado REAL,
            status TEXT DEFAULT 'ATIVA',
            prestacao_status TEXT DEFAULT 'PENDENTE',
            local_rota TEXT,
            tipo_rota TEXT,
            local_carregamento TEXT,
            granja_carregada TEXT,
            local_carregado TEXT,
            local_carreg TEXT,
            data_saida TEXT,
            hora_saida TEXT,
            data_chegada TEXT,
            hora_chegada TEXT,
            adiantamento REAL DEFAULT 0,
            adiantamento_origem TEXT,
            num_nf TEXT,
            kg_carregado REAL DEFAULT 0,
            media REAL DEFAULT 0,
            adiantamento_rota REAL DEFAULT 0,
            nf_numero TEXT,
            nf_kg REAL DEFAULT 0,
            nf_caixas INTEGER DEFAULT 0,
            nf_kg_carregado REAL DEFAULT 0,
            nf_kg_vendido REAL DEFAULT 0,
            nf_saldo REAL DEFAULT 0,
            km_inicial REAL DEFAULT 0,
            km_final REAL DEFAULT 0,
            litros REAL DEFAULT 0,
            km_rodado REAL DEFAULT 0,
            media_km_l REAL DEFAULT 0,
            custo_km REAL DEFAULT 0,
            ced_200_qtd INTEGER DEFAULT 0,
            ced_100_qtd INTEGER DEFAULT 0,
            ced_50_qtd INTEGER DEFAULT 0,
            ced_20_qtd INTEGER DEFAULT 0,
            ced_10_qtd INTEGER DEFAULT 0,
            ced_5_qtd INTEGER DEFAULT 0,
            ced_2_qtd INTEGER DEFAULT 0,
            valor_dinheiro REAL DEFAULT 0,
            pix_motorista REAL DEFAULT 0,
            diaria_motorista_valor REAL DEFAULT 0,
            rota_observacao TEXT,
            motorista_id INTEGER,
            codigo TEXT,
            data TEXT,
            total_caixas INTEGER DEFAULT 0,
            quilos REAL DEFAULT 0,
            saida_dt TEXT,
            chegada_dt TEXT,
            aves_caixa_final INTEGER,
            qnt_aves_caixa_final INTEGER,
            media_1 REAL,
            media_2 REAL,
            media_3 REAL,
            carregamento_fechado INTEGER DEFAULT 0,
            carregamento_salvo_em TEXT,
            nf_preco REAL DEFAULT 0,
            motorista_codigo TEXT,
            codigo_motorista TEXT,
            tipo_estimativa TEXT DEFAULT 'KG',
            caixas_estimado INTEGER DEFAULT 0,
            usuario_criacao TEXT,
            usuario_ultima_edicao TEXT,
            status_operacional TEXT,
            status_operacional_obs TEXT,
            status_operacional_em TEXT,
            status_operacional_por TEXT,
            finalizada_no_app INTEGER DEFAULT 0
        )
        """
    )
    cur.execute("PRAGMA table_info(programacoes)")
    programacoes_cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
    if "adiantamento_origem" not in programacoes_cols:
        cur.execute("ALTER TABLE programacoes ADD COLUMN adiantamento_origem TEXT")
    _safe_add_column(cur, "programacoes", "local_rota", "TEXT")
    _safe_add_column(cur, "programacoes", "tipo_rota", "TEXT")
    _safe_add_column(cur, "programacoes", "local_carregamento", "TEXT")
    _safe_add_column(cur, "programacoes", "granja_carregada", "TEXT")
    _safe_add_column(cur, "programacoes", "local_carregado", "TEXT")
    _safe_add_column(cur, "programacoes", "local_carreg", "TEXT")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS programacao_itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_programacao TEXT,
            cod_cliente TEXT,
            nome_cliente TEXT,
            qnt_caixas INTEGER,
            kg REAL,
            preco REAL,
            endereco TEXT,
            vendedor TEXT,
            pedido TEXT,
            produto_id INTEGER,
            produto TEXT,
            status_pedido TEXT,
            caixas_atual INTEGER,
            preco_atual REAL,
            alterado_em TEXT,
            alterado_por TEXT,
            alteracao_tipo TEXT,
            alteracao_detalhe TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS recebimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_programacao TEXT,
            cod_cliente TEXT,
            nome_cliente TEXT,
            valor REAL,
            forma_pagamento TEXT,
            observacao TEXT,
            num_nf TEXT,
            data_registro TEXT,
            pedido TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS despesas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_programacao TEXT,
            descricao TEXT,
            valor REAL,
            data_registro TEXT,
            tipo_despesa TEXT DEFAULT 'ROTA',
            categoria TEXT,
            motorista TEXT,
            veiculo TEXT,
            observacao TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS vendas_importadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido TEXT,
            data_venda TEXT,
            cliente TEXT,
            nome_cliente TEXT,
            vendedor TEXT,
            produto_id INTEGER,
            produto TEXT,
            vr_total REAL,
            qnt REAL,
            cidade TEXT,
            valor_unitario REAL,
            observacao TEXT,
            selecionada INTEGER DEFAULT 0,
            usada INTEGER DEFAULT 0,
            usada_em TEXT,
            codigo_programacao TEXT
        )
        """
    )
    _safe_add_column(cur, "programacao_itens", "produto_id", "INTEGER")
    _safe_add_column(cur, "vendas_importadas", "pedido", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "selecionada", "INTEGER DEFAULT 0")
    _safe_add_column(cur, "vendas_importadas", "data_venda", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "cliente", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "nome_cliente", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "vendedor", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "produto", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "produto_id", "INTEGER")
    _safe_add_column(cur, "vendas_importadas", "vr_total", "REAL")
    _safe_add_column(cur, "vendas_importadas", "qnt", "REAL")
    _safe_add_column(cur, "vendas_importadas", "cidade", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "valor_unitario", "REAL")
    _safe_add_column(cur, "vendas_importadas", "observacao", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "usada", "INTEGER DEFAULT 0")
    _safe_add_column(cur, "vendas_importadas", "usada_em", "TEXT")
    _safe_add_column(cur, "vendas_importadas", "codigo_programacao", "TEXT")
    _safe_add_column(cur, "programacoes", "pix_motorista", "REAL DEFAULT 0")
    _normalize_existing_programacoes(cur)
    company_id = ensure_saas_schema(conn)
    ensure_tenant_columns(conn, company_id)
    conn.commit()


def ensure_admin_user(conn: sqlite3.Connection, password: str | None = None) -> str:
    cur = conn.cursor()
    cur.execute("SELECT id, COALESCE(senha, '') FROM usuarios WHERE UPPER(COALESCE(nome,''))='ADMIN' LIMIT 1")
    row = cur.fetchone()
    admin_password = (password or os.environ.get("ROTA_ADMIN_PASS") or os.environ.get("ROTA_ADMIN_PASSWORD") or DEFAULT_ADMIN_PASSWORD).strip()
    if not admin_password:
        admin_password = DEFAULT_ADMIN_PASSWORD
    if row is None:
        cur.execute(
            "INSERT INTO usuarios (nome, permissoes, codigo, senha) VALUES (?, ?, ?, ?)",
            ("ADMIN", "ADMIN", "ADMIN", hash_password_pbkdf2(admin_password)),
        )
    elif not str(row[1] or "").startswith("pbkdf2_sha256$"):
        cur.execute(
            "UPDATE usuarios SET senha=?, permissoes='ADMIN', codigo=COALESCE(codigo, 'ADMIN') WHERE id=?",
            (hash_password_pbkdf2(str(row[1] or admin_password)), int(row[0])),
        )
    else:
        cur.execute(
            "UPDATE usuarios SET permissoes='ADMIN', codigo=COALESCE(codigo, 'ADMIN') WHERE id=?",
            (int(row[0]),),
        )
    try:
        company_id = ensure_saas_schema(conn)
        ensure_tenant_columns(conn, company_id)
    except Exception:
        pass
    conn.commit()
    return admin_password


def reset_operational_data(conn: sqlite3.Connection) -> Dict[str, int]:
    cur = conn.cursor()
    result: Dict[str, int] = {}
    for table in OPERATIONAL_TABLES:
        if not table_exists(cur, table):
            continue
        before = count_rows(cur, table)
        cur.execute(f'DELETE FROM "{table}"')
        result[table] = before
    cur.execute("DELETE FROM usuarios WHERE UPPER(COALESCE(nome,'')) <> 'ADMIN'")
    conn.commit()
    return result


def ensure_permission_system(conn: sqlite3.Connection) -> None:
    """
    Inicializa o sistema de permissões com as permissões padrão do sistema.
    Garante que todas as permissões existem e o ADMIN tem acesso total.
    """
    cur = conn.cursor()
    
    # Criar tabelas se não existirem
    cur.execute("""
        CREATE TABLE IF NOT EXISTS permissoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            modulo TEXT NOT NULL,
            nome_permissao TEXT NOT NULL,
            descricao TEXT,
            ativo INTEGER DEFAULT 1
        )
    """)
    
    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuario_permissoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER NOT NULL,
            permissao_id INTEGER NOT NULL,
            concedida_em TEXT DEFAULT (datetime('now')),
            concedida_por TEXT,
            UNIQUE(usuario_id, permissao_id),
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id),
            FOREIGN KEY(permissao_id) REFERENCES permissoes(id)
        )
    """)
    
    # Permissões padrão por módulo
    permissoes_padrao = [
        # Programações
        ("programacoes", "visualizar_programacoes", "Visualizar programações"),
        ("programacoes", "criar_programacoes", "Criar novas programações"),
        ("programacoes", "editar_programacoes", "Editar programações"),
        ("programacoes", "deletar_programacoes", "Deletar programações"),
        ("programacoes", "finalizar_programacoes", "Finalizar programações"),
        
        # Prestação de Contas
        ("prestacao", "gerar_prestacao", "Gerar prestação de contas"),
        ("prestacao", "editar_prestacao", "Editar prestação"),
        ("prestacao", "fechar_prestacao", "Fechar prestação"),
        
        # Cadastros
        ("cadastros", "gerenciar_clientes", "Gerenciar clientes"),
        ("cadastros", "gerenciar_motoristas", "Gerenciar motoristas"),
        ("cadastros", "gerenciar_vendedores", "Gerenciar vendedores"),
        
        # Relatórios
        ("relatorios", "gerar_relatorios", "Gerar relatórios"),
        ("relatorios", "exportar_dados", "Exportar dados para Excel/PDF"),
        
        # Sistema
        ("sistema", "gerenciar_usuarios", "Gerenciar usuários e permissões"),
        ("sistema", "acessar_ferramentas", "Acessar ferramentas do sistema"),
        ("sistema", "fazer_backup", "Fazer backup do banco de dados"),
        ("sistema", "restaurar_backup", "Restaurar backup"),
        ("sistema", "limpar_logs", "Limpar logs do sistema"),
        ("sistema", "ver_configuracoes", "Visualizar configurações"),
        ("sistema", "editar_configuracoes", "Editar configurações"),
    ]
    
    # Inserir permissões padrão se não existirem
    for modulo, nome_perm, descricao in permissoes_padrao:
        cur.execute(
            "SELECT id FROM permissoes WHERE modulo=? AND nome_permissao=? LIMIT 1",
            (modulo, nome_perm)
        )
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO permissoes (modulo, nome_permissao, descricao, ativo) VALUES (?, ?, ?, 1)",
                (modulo, nome_perm, descricao)
            )
    
    # Garantir que o ADMIN tem acesso a todas as permissões
    cur.execute("SELECT id FROM usuarios WHERE UPPER(COALESCE(nome,''))='ADMIN' LIMIT 1")
    admin_row = cur.fetchone()
    if admin_row:
        admin_id = admin_row[0]
        cur.execute("SELECT id FROM permissoes WHERE ativo=1")
        for (perm_id,) in cur.fetchall():
            cur.execute(
                "INSERT OR IGNORE INTO usuario_permissoes (usuario_id, permissao_id, concedida_por) VALUES (?, ?, ?)",
                (admin_id, perm_id, "SISTEMA")
            )
    
    conn.commit()
