import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402


class ApiClientesLegacySchemaTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        api_server.DB_PATH = self.db_path
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("CREATE TABLE clientes (cod_cliente TEXT, nome_cliente TEXT)")
            conn.execute(
                "INSERT INTO clientes (cod_cliente, nome_cliente) VALUES (?, ?)",
                ("001", "CLIENTE ANTIGO"),
            )

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def test_upsert_clientes_accepts_legacy_table_without_id_or_vendedor(self):
        payload = api_server.DesktopClienteUpsertIn(
            cod_cliente="001",
            nome_cliente="Cliente Novo",
            endereco="Rua A",
            telefone="9999",
            vendedor="Vend 1",
        )

        result = api_server.desktop_clientes_upsert(payload, _ok=True)

        self.assertTrue(result["ok"])
        self.assertEqual(result["updated"], 1)
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT cod_cliente, nome_cliente, endereco, telefone, vendedor FROM clientes WHERE cod_cliente='001'"
            ).fetchone()
        self.assertEqual(row, ("001", "CLIENTE NOVO", "RUA A", "9999", "VEND 1"))

    def test_bulk_upsert_clientes_accepts_legacy_table(self):
        payload = api_server.DesktopClientesBulkUpsertIn(
            clientes=[
                api_server.DesktopClienteUpsertIn(
                    cod_cliente="001",
                    nome_cliente="Cliente Novo",
                    telefone="9999",
                ),
                api_server.DesktopClienteUpsertIn(
                    cod_cliente="002",
                    nome_cliente="Cliente Dois",
                    endereco="Rua Dois",
                    vendedor="Vend 2",
                ),
            ]
        )

        result = api_server.desktop_clientes_bulk_upsert(payload, _ok=True)

        self.assertTrue(result["ok"])
        self.assertEqual(result["total"], 2)
        self.assertEqual(result["updated"], 1)
        self.assertEqual(result["created"], 1)
        with sqlite3.connect(self.db_path) as conn:
            codes = [
                row[0]
                for row in conn.execute(
                    "SELECT cod_cliente FROM clientes ORDER BY cod_cliente"
                ).fetchall()
            ]
            row_001 = conn.execute(
                "SELECT nome_cliente, telefone FROM clientes WHERE cod_cliente='001'"
            ).fetchone()
            row_002 = conn.execute(
                "SELECT nome_cliente, endereco, vendedor FROM clientes WHERE cod_cliente='002'"
            ).fetchone()
        self.assertIn("001", codes)
        self.assertIn("002", codes)
        self.assertEqual(row_001, ("CLIENTE NOVO", "9999"))
        self.assertEqual(row_002, ("CLIENTE DOIS", "RUA DOIS", "VEND 2"))

    def test_clientes_base_returns_telefone_after_schema_repair(self):
        api_server.desktop_clientes_upsert(
            api_server.DesktopClienteUpsertIn(
                cod_cliente="002",
                nome_cliente="Cliente Dois",
                telefone="8888",
            ),
            _ok=True,
        )

        rows = api_server.desktop_clientes_base(q="", vendedor="", cidade="", ordem="codigo", limit=10, _ok=True)

        by_code = {row["cod_cliente"]: row for row in rows}
        self.assertEqual(by_code["002"]["telefone"], "8888")


if __name__ == "__main__":
    unittest.main()
