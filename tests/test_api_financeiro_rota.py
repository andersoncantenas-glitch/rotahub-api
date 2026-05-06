import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402


class ApiFinanceiroRotaTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        api_server.DB_PATH = self.db_path
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE programacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    data_criacao TEXT,
                    prestacao_status TEXT,
                    status TEXT,
                    status_operacional TEXT,
                    motorista TEXT,
                    motorista_id INTEGER,
                    motorista_codigo TEXT,
                    codigo_motorista TEXT,
                    veiculo TEXT,
                    equipe TEXT,
                    kg_estimado REAL,
                    tipo_estimativa TEXT,
                    caixas_estimado INTEGER,
                    local_rota TEXT,
                    tipo_rota TEXT,
                    local_carregamento TEXT,
                    granja_carregada TEXT,
                    local_carregado TEXT,
                    local_carreg TEXT,
                    total_caixas INTEGER,
                    quilos REAL,
                    nf_kg REAL,
                    nf_kg_carregado REAL,
                    nf_kg_vendido REAL,
                    nf_saldo REAL,
                    km_inicial REAL,
                    km_final REAL,
                    litros REAL,
                    km_rodado REAL,
                    media_km_l REAL,
                    valor_dinheiro REAL,
                    pix_motorista REAL,
                    ced_100_qtd INTEGER,
                    adiantamento REAL,
                    adiantamento_origem TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE motoristas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    nome TEXT
                )
                """
            )
            conn.execute("INSERT INTO motoristas (codigo, nome) VALUES ('M1', 'MOTORISTA UM')")
            conn.execute(
                """
                CREATE TABLE programacao_itens (
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
                    produto TEXT
                )
                """
            )
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, prestacao_status, status, status_operacional,
                    nf_kg, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                    km_inicial, km_final, litros, km_rodado, media_km_l,
                    valor_dinheiro, pix_motorista, ced_100_qtd, adiantamento, adiantamento_origem
                )
                VALUES ('PG1', 'PENDENTE', 'EM_ENTREGAS', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '')
                """
            )

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def test_financeiro_accepts_partial_km_without_final_odometer(self):
        payload = api_server.DesktopRotaFinanceiroIn(
            km_inicial=1000,
            km_final=0,
            litros=20,
            km_rodado=0,
            media_km_l=0,
            valor_dinheiro=2960,
            pix_motorista=540,
            ced_100_qtd=20,
        )

        result = api_server.desktop_atualizar_financeiro_rota("PG1", payload, _ok=True)

        self.assertTrue(result["ok"])
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT km_inicial, km_final, valor_dinheiro, pix_motorista, ced_100_qtd FROM programacoes WHERE codigo_programacao='PG1'"
            ).fetchone()
        self.assertEqual(row, (1000.0, 0.0, 2960.0, 540.0, 20))

    def test_financeiro_accepts_derived_nf_variance_without_blocking_cash_save(self):
        payload = api_server.DesktopRotaFinanceiroIn(
            nf_kg=0,
            nf_kg_carregado=500,
            nf_kg_vendido=620,
            nf_saldo=-120,
            valor_dinheiro=2960,
            adiantamento=800,
            adiantamento_origem="Caixa Matriz",
        )

        result = api_server.desktop_atualizar_financeiro_rota("PG1", payload, _ok=True)

        self.assertTrue(result["ok"])
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT nf_kg_carregado, nf_kg_vendido, nf_saldo, valor_dinheiro, adiantamento, adiantamento_origem FROM programacoes WHERE codigo_programacao='PG1'"
            ).fetchone()
        self.assertEqual(row, (500.0, 620.0, -120.0, 2960.0, 800.0, "CAIXA MATRIZ"))

    def test_upsert_persists_local_rota_aliases_for_desktop_and_reports(self):
        payload = api_server.DesktopRotaUpsertIn(
            codigo_programacao="PG2",
            motorista="MOTORISTA UM",
            motorista_codigo="M1",
            veiculo="ABC1234",
            equipe="EQUIPE A",
            kg_estimado=1000,
            tipo_estimativa="KG",
            local_rota="SERTAO",
            local_carregamento="GRANJA MATRIZ",
        )

        result = api_server.desktop_rotas_upsert(payload, _ok=True)

        self.assertTrue(result["ok"])
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                """
                SELECT local_rota, tipo_rota, local_carregamento, granja_carregada,
                       local_carregado, local_carreg
                FROM programacoes
                WHERE codigo_programacao='PG2'
                """
            ).fetchone()
        self.assertEqual(row, ("SERTAO", "SERTAO", "GRANJA MATRIZ", "GRANJA MATRIZ", "GRANJA MATRIZ", "GRANJA MATRIZ"))


if __name__ == "__main__":
    unittest.main()
