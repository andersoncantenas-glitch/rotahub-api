import contextlib
import os
import sqlite3
import tempfile
import unittest

import main


class RetornoTransbordoEventsTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.addCleanup(self._cleanup_db_file)
        conn = sqlite3.connect(self.db_path)
        try:
            conn.execute(
                """
                CREATE TABLE programacao_itens_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    cod_cliente TEXT,
                    pedido TEXT,
                    evento TEXT,
                    payload_json TEXT,
                    registrado_em TEXT,
                    created_at TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE roteiro_operacional (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo_evento TEXT,
                    codigo_programacao TEXT,
                    pedido TEXT,
                    cod_cliente TEXT,
                    cliente_nome TEXT,
                    caixas INTEGER DEFAULT 0,
                    kg REAL DEFAULT 0,
                    observacao TEXT,
                    payload_json TEXT,
                    data_hora TEXT,
                    created_at TEXT
                )
                """
            )
            conn.execute(
                """
                INSERT INTO programacao_itens_log
                    (codigo_programacao, cod_cliente, pedido, evento, payload_json, registrado_em, created_at)
                VALUES
                    ('PG-TB', '__TRANSBORDO__', '', 'transbordo',
                     '{"aves_mortas_transbordo":119,"mortalidade_transbordo_kg":360.09,"foto_doa_path":"foto.jpg"}',
                     '2026-05-22T13:44:17', '2026-05-22T13:44:17')
                """
            )
            for i in range(100):
                conn.execute(
                    """
                    INSERT INTO programacao_itens_log
                        (codigo_programacao, cod_cliente, pedido, evento, payload_json, registrado_em, created_at)
                    VALUES (?, ?, ?, 'cliente_controle', '{"status_pedido":"ENTREGUE"}', ?, ?)
                    """,
                    ("PG-TB", f"C{i:03d}", str(i), "2026-05-23T10:00:00", "2026-05-23T10:00:00"),
                )
            conn.execute(
                """
                INSERT INTO roteiro_operacional
                    (tipo_evento, codigo_programacao, kg, payload_json, data_hora, created_at)
                VALUES
                    ('TRANSBORDO', 'PG-TB', 360.09,
                     '{"aves_mortas_transbordo":119,"mortalidade_transbordo_kg":360.09,"foto_doa_path":"foto.jpg"}',
                     '2026-05-22T13:44:17', '2026-05-22T13:44:17')
                """
            )
            conn.execute(
                """
                INSERT INTO roteiro_operacional
                    (tipo_evento, codigo_programacao, observacao, data_hora, created_at)
                VALUES ('NF_WEB', 'PG-TB', 'nao deve aparecer', '2026-05-24T08:00:00', '2026-05-24T08:00:00')
                """
            )
            conn.commit()
        finally:
            conn.close()

    def _cleanup_db_file(self):
        try:
            if os.path.exists(self.db_path):
                os.unlink(self.db_path)
        except PermissionError:
            pass

    @contextlib.contextmanager
    def _db(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def test_transbordo_log_is_kept_outside_recent_limit_and_deduped(self):
        original_get_db = main.get_db
        main.get_db = self._db
        try:
            logs = main._fetch_logs_retorno_operacional("PG-TB", limit=80)
            roteiro = main._fetch_roteiro_retorno_operacional("PG-TB", limit=80)
            eventos = main._dedupe_retorno_eventos(logs + roteiro)
        finally:
            main.get_db = original_get_db

        transbordos = [evt for evt in eventos if evt.get("cliente") == "TRANSBORDO"]
        self.assertEqual(len(transbordos), 1)
        self.assertEqual(transbordos[0]["st"], "TRANSBORDO")
        self.assertEqual(transbordos[0]["kg"], "360,09")
        self.assertEqual(transbordos[0]["mort"], "119")
        self.assertIn("MORT 119 / 360,09KG", transbordos[0]["detalhe"])
        self.assertIn("FOTO", transbordos[0]["detalhe"])
        self.assertFalse(any(evt.get("st") == "NF_WEB" for evt in eventos))


if __name__ == "__main__":
    unittest.main()
