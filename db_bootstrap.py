from __future__ import annotations

import base64
import hashlib
import hmac
import os
import secrets
import sqlite3
from typing import Dict


PBKDF2_ITERATIONS = 200_000
OPERATIONAL_TABLES = [
    "motoristas",
    "veiculos",
    "ajudantes",
    "clientes",
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
    "mobile_sync_idempotency",
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
            acesso_liberado INTEGER DEFAULT 0,
            acesso_liberado_por TEXT,
            acesso_liberado_em TEXT,
            acesso_obs TEXT,
            status TEXT DEFAULT 'ATIVO'
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS veiculos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            placa TEXT,
            modelo TEXT,
            capacidade_cx INTEGER
        )
        """
    )
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
            cod_cliente TEXT UNIQUE,
            nome_cliente TEXT,
            endereco TEXT,
            bairro TEXT,
            cidade TEXT,
            uf TEXT,
            telefone TEXT,
            rota TEXT,
            vendedor TEXT
        )
        """
    )
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
            tipo_rota TEXT,
            granja_carregada TEXT,
            data_saida TEXT,
            hora_saida TEXT,
            data_chegada TEXT,
            hora_chegada TEXT,
            adiantamento REAL DEFAULT 0,
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
    conn.commit()


def ensure_admin_user(conn: sqlite3.Connection, password: str | None = None) -> str:
    cur = conn.cursor()
    cur.execute("SELECT id, COALESCE(senha, '') FROM usuarios WHERE UPPER(COALESCE(nome,''))='ADMIN' LIMIT 1")
    row = cur.fetchone()
    admin_password = (password or os.environ.get("ROTA_ADMIN_PASS") or os.environ.get("ROTA_ADMIN_PASSWORD") or "").strip()
    if not admin_password:
        admin_password = secrets.token_urlsafe(8)
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
