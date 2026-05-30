import os
import sqlite3
import tempfile
import unittest

from db_bootstrap import OPERATIONAL_TABLES, TENANT_SCOPED_TABLES, ensure_core_schema, ensure_tenant_columns


class SaasTenantBackfillTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)

    def tearDown(self):
        try:
            os.remove(self.db_path)
        except OSError:
            pass

    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def test_company_id_is_added_and_backfilled_on_legacy_tables(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE motoristas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT)")
            cur.execute("INSERT INTO motoristas (nome) VALUES ('MOTORISTA LEGADO')")
            cur.execute("CREATE TABLE usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT, senha TEXT)")
            cur.execute("INSERT INTO usuarios (nome, senha) VALUES ('ADMIN LEGADO', '123456')")
            cur.execute("CREATE TABLE rota_gps_pings (id INTEGER PRIMARY KEY AUTOINCREMENT, lat REAL)")
            cur.execute("INSERT INTO rota_gps_pings (lat) VALUES (-3.7)")

            ensure_core_schema(conn)

            cur.execute("SELECT id FROM companies WHERE code='default' LIMIT 1")
            company_id = int(cur.fetchone()[0])

            for table in ("motoristas", "rota_gps_pings", "usuarios"):
                cur.execute(f"PRAGMA table_info({table})")
                cols = {str(row[1]).lower() for row in cur.fetchall()}
                self.assertIn("company_id", cols)
                cur.execute(f"SELECT COUNT(*) FROM {table} WHERE company_id=?", (company_id,))
                self.assertGreaterEqual(int(cur.fetchone()[0]), 1)
                cur.execute("SELECT 1 FROM sqlite_master WHERE type='index' AND name=?", (f"idx_{table}_company_id",))
                self.assertIsNotNone(cur.fetchone())

    def test_backfill_is_idempotent_and_reports_only_null_rows(self):
        with self._connect() as conn:
            ensure_core_schema(conn)
            cur = conn.cursor()
            cur.execute("SELECT id FROM companies WHERE code='default' LIMIT 1")
            company_id = int(cur.fetchone()[0])

            cur.execute("INSERT INTO veiculos (placa, modelo, company_id) VALUES ('ABC1D23', 'TRUCK', NULL)")
            summary = ensure_tenant_columns(conn, company_id)
            self.assertEqual(summary.get("veiculos"), 1)

            second_summary = ensure_tenant_columns(conn, company_id)
            self.assertEqual(second_summary.get("veiculos"), 0)
            cur.execute("SELECT company_id FROM veiculos WHERE placa='ABC1D23' LIMIT 1")
            self.assertEqual(int(cur.fetchone()[0]), company_id)

    def test_tenant_scoped_tables_include_operational_tables_and_users(self):
        for table in OPERATIONAL_TABLES:
            self.assertIn(table, TENANT_SCOPED_TABLES)
        self.assertIn("usuarios", TENANT_SCOPED_TABLES)


if __name__ == "__main__":
    unittest.main()
