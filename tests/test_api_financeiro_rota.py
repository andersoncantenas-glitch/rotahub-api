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
                    prestacao_status TEXT,
                    status TEXT,
                    status_operacional TEXT,
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
                    ced_100_qtd INTEGER,
                    adiantamento REAL,
                    adiantamento_origem TEXT
                )
                """
            )
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, prestacao_status, status, status_operacional,
                    nf_kg, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                    km_inicial, km_final, litros, km_rodado, media_km_l,
                    valor_dinheiro, ced_100_qtd, adiantamento, adiantamento_origem
                )
                VALUES ('PG1', 'PENDENTE', 'EM_ENTREGAS', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '')
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
            ced_100_qtd=20,
        )

        result = api_server.desktop_atualizar_financeiro_rota("PG1", payload, _ok=True)

        self.assertTrue(result["ok"])
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT km_inicial, km_final, valor_dinheiro, ced_100_qtd FROM programacoes WHERE codigo_programacao='PG1'"
            ).fetchone()
        self.assertEqual(row, (1000.0, 0.0, 2960.0, 20))

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


if __name__ == "__main__":
    unittest.main()
