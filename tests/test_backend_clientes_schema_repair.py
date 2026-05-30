import os
import tempfile
import unittest
from pathlib import Path

from sqlalchemy import create_engine, text

from backend.config.database import _ensure_backend_columns


class BackendClientesSchemaRepairTests(unittest.TestCase):
    def _temp_db_path(self) -> str:
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        return path

    def test_rebuilds_clientes_table_without_id(self):
        db_path = self._temp_db_path()
        engine = create_engine(f"sqlite:///{Path(db_path).as_posix()}")
        try:
            with engine.begin() as conn:
                conn.execute(text("CREATE TABLE clientes (cod_cliente TEXT, nome_cliente TEXT, endereco TEXT)"))
                conn.execute(
                    text("INSERT INTO clientes (cod_cliente, nome_cliente, endereco) VALUES ('001', 'Cliente Um', 'Rua A')")
                )
                conn.execute(
                    text("INSERT INTO clientes (cod_cliente, nome_cliente, endereco) VALUES ('001', 'Duplicado', 'Rua B')")
                )

                _ensure_backend_columns(conn)

                columns = {row[1] for row in conn.execute(text("PRAGMA table_info(clientes)")).fetchall()}
                self.assertIn("id", columns)
                self.assertIn("telefone", columns)
                self.assertIn("vendedor", columns)

                rows = conn.execute(
                    text("SELECT cod_cliente, nome_cliente, endereco FROM clientes ORDER BY id")
                ).fetchall()
                self.assertEqual(rows, [("001", "CLIENTE UM", "RUA A")])
                backup_count = conn.execute(text("SELECT COUNT(*) FROM clientes_legacy_backup")).scalar()
                self.assertEqual(backup_count, 2)
        finally:
            engine.dispose()
            try:
                os.unlink(db_path)
            except OSError:
                pass

    def test_adds_missing_clientes_columns_and_copies_legacy_nome(self):
        db_path = self._temp_db_path()
        engine = create_engine(f"sqlite:///{Path(db_path).as_posix()}")
        try:
            with engine.begin() as conn:
                conn.execute(text("CREATE TABLE clientes (id INTEGER PRIMARY KEY, cod_cliente TEXT, nome TEXT)"))
                conn.execute(text("INSERT INTO clientes (cod_cliente, nome) VALUES ('002', 'Cliente Dois')"))

                _ensure_backend_columns(conn)

                columns = {row[1] for row in conn.execute(text("PRAGMA table_info(clientes)")).fetchall()}
                self.assertIn("nome_cliente", columns)
                self.assertIn("telefone", columns)
                self.assertIn("vendedor", columns)
                row = conn.execute(
                    text("SELECT cod_cliente, nome_cliente, telefone, vendedor FROM clientes WHERE cod_cliente='002'")
                ).fetchone()
                self.assertEqual(row, ("002", "Cliente Dois", None, None))
        finally:
            engine.dispose()
            try:
                os.unlink(db_path)
            except OSError:
                pass


if __name__ == "__main__":
    unittest.main()
