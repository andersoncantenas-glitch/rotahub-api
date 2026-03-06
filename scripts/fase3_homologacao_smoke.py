from __future__ import annotations

import importlib
import os
import shutil
import sqlite3
import sys
import tempfile
import urllib.request
from pathlib import Path


def ok(msg: str) -> None:
    print(f"[OK] {msg}")


def fail(msg: str) -> None:
    print(f"[FALHA] {msg}")


def _scalar(cur: sqlite3.Cursor, sql: str, params: tuple = ()) -> int:
    cur.execute(sql, params)
    row = cur.fetchone()
    return int((row[0] if row else 0) or 0)


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    src_db = Path(r"C:\rotahub\rota_granja.db")
    temp_dir = Path(tempfile.mkdtemp(prefix="rotahub_fase3_"))
    test_db = temp_dir / "rota_granja_test.db"

    try:
        if src_db.exists():
            shutil.copy2(src_db, test_db)
            ok(f"Banco base copiado: {src_db}")
        else:
            test_db.touch()
            ok("Banco base nao encontrado; usando banco vazio para smoke.")

        os.environ["ROTA_DB"] = str(test_db)
        os.environ.setdefault("ROTA_SERVER_URL", "https://rotahub-api.onrender.com")
        os.environ["ROTA_SQL_MIRROR_API"] = "0"
        if str(project_root) not in sys.path:
            sys.path.insert(0, str(project_root))

        main_mod = importlib.import_module("main")
        importlib.reload(main_mod)
        ok("main.py importado com sucesso")

        # 1) Migração/estrutura de banco
        main_mod.db_init()
        with sqlite3.connect(test_db) as conn:
            cur = conn.cursor()
            for t in ("motoristas", "veiculos", "ajudantes", "clientes", "programacoes", "programacao_itens", "recebimentos", "despesas"):
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (t,))
                if not cur.fetchone():
                    fail(f"Tabela obrigatoria ausente: {t}")
                    return 1
        ok("Estrutura de tabelas obrigatorias validada")

        # 2) Cadastros básicos
        with sqlite3.connect(test_db) as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO motoristas (nome, codigo, senha, telefone, status) VALUES (?, ?, ?, ?, ?)", ("TESTE MOTORISTA", "MOT-99", "123456", "88999999999", "ATIVO"))
            cur.execute("INSERT INTO veiculos (placa, modelo, capacidade_cx) VALUES (?, ?, ?)", ("ABC1D23", "TRUCK", 500))
            cur.execute("INSERT INTO ajudantes (nome, sobrenome, telefone, status) VALUES (?, ?, ?, ?)", ("AJUDA", "UM", "88988888888", "ATIVO"))
            cur.execute("INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor) VALUES (?, ?, ?, ?, ?)", ("C001", "CLIENTE TESTE", "RUA 1", "88977777777", "VENDEDOR"))
            conn.commit()
        ok("Cadastros basicos inseridos")

        # 3) Programação + itens + fetch helpers
        codigo = main_mod.generate_program_code()
        with sqlite3.connect(test_db) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO programacoes (codigo_programacao, data_criacao, motorista, veiculo, equipe, kg_estimado, status) VALUES (?, datetime('now'), ?, ?, ?, ?, ?)",
                (codigo, "TESTE MOTORISTA", "ABC1D23", "AJUDA UM", 1000.0, "ATIVA"),
            )
            cur.execute(
                "INSERT INTO programacao_itens (codigo_programacao, cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido, produto) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (codigo, "C001", "CLIENTE TESTE", 10, 120.0, 8.5, "RUA 1", "VENDEDOR", "PED-1", "FRANGO"),
            )
            conn.commit()

        ativas = main_mod.fetch_programacoes_ativas()
        itens = main_mod.fetch_programacao_itens(codigo)
        if not ativas:
            fail("fetch_programacoes_ativas retornou vazio apos inserir programacao")
            return 1
        if not itens:
            fail("fetch_programacao_itens retornou vazio apos inserir item")
            return 1
        ok("Fluxo programacao/itens validado")

        # 4) Recebimentos + despesas (prestação parcial)
        with sqlite3.connect(test_db) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO recebimentos (codigo_programacao, cod_cliente, nome_cliente, valor, forma_pagamento, observacao, num_nf, data_registro) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))",
                (codigo, "C001", "CLIENTE TESTE", 900.0, "DINHEIRO", "", "",),
            )
            cur.execute(
                "INSERT INTO despesas (codigo_programacao, descricao, valor, data_registro, categoria, observacao) VALUES (?, ?, ?, datetime('now'), ?, ?)",
                (codigo, "COMBUSTIVEL", 150.0, "COMBUSTIVEL", ""),
            )
            conn.commit()
            rec_total = _scalar(cur, "SELECT COALESCE(SUM(valor),0) FROM recebimentos WHERE codigo_programacao=?", (codigo,))
            desp_total = _scalar(cur, "SELECT COALESCE(SUM(valor),0) FROM despesas WHERE codigo_programacao=?", (codigo,))
        if rec_total <= 0 or desp_total <= 0:
            fail("Totais de recebimentos/despesas invalidos no smoke")
            return 1
        ok("Fluxo recebimentos/despesas validado")

        # 5) API online (saúde básica)
        api_url = os.environ.get("ROTA_SERVER_URL", "").rstrip("/") + "/openapi.json"
        try:
            req = urllib.request.Request(api_url, method="GET")
            with urllib.request.urlopen(req, timeout=10) as resp:
                status = int(getattr(resp, "status", 0) or 0)
            if status == 200:
                ok(f"API respondeu 200 em {api_url}")
            else:
                fail(f"API respondeu status {status} em {api_url}")
                return 1
        except Exception as e:
            fail(f"Falha ao consultar API ({api_url}): {e}")
            return 1

        print("\n=== RESULTADO FASE 3 (SMOKE) ===")
        print("Aprovado: banco, helpers de fluxo e conectividade API basica.")
        print("Pendente: homologacao manual de interface Desktop/APK motorista.")
        return 0

    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
