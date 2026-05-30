import os
import sqlite3
import tempfile
import unittest
from contextlib import contextmanager

import main


class EscalaFolgasTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".sqlite")
        os.close(fd)
        self._orig_get_db = main.get_db

        @contextmanager
        def _tmp_db():
            conn = sqlite3.connect(self.db_path)
            try:
                yield conn
            finally:
                conn.close()

        main.get_db = _tmp_db

    def tearDown(self):
        main.get_db = self._orig_get_db
        try:
            os.remove(self.db_path)
        except OSError:
            pass

    def test_registra_folga_motorista_ativa_por_codigo_e_nome(self):
        folga_id = main.registrar_escala_folga(
            "MOTORISTA",
            "Joao Silva",
            pessoa_codigo="M01",
            data_inicio="2026-05-27",
            data_fim="2026-05-28",
            motivo="Descanso",
        )

        folgas = main.fetch_escala_folgas_ativas("2026-05-27", "2026-05-27")

        self.assertGreater(folga_id, 0)
        self.assertIn("M01", folgas["motoristas_codigos"])
        self.assertIn("JOAO SILVA", folgas["motoristas_nomes"])
        self.assertTrue(main.pessoa_em_folga("MOTORISTA", nome="Joao Silva", codigo="M01", data_inicio="2026-05-27"))

    def test_folga_fora_do_periodo_nao_bloqueia_e_encerrar_remove(self):
        folga_id = main.registrar_escala_folga(
            "AJUDANTE",
            "Maria Souza",
            pessoa_id="7",
            data_inicio="2026-06-01",
            data_fim="2026-06-02",
        )

        fora_periodo = main.fetch_escala_folgas_ativas("2026-05-27", "2026-05-27")
        self.assertNotIn("7", fora_periodo["ajudantes_ids"])

        self.assertTrue(main.encerrar_escala_folga(folga_id))
        no_periodo_original = main.fetch_escala_folgas_ativas("2026-06-01", "2026-06-01")
        self.assertNotIn("7", no_periodo_original["ajudantes_ids"])

    def test_nao_permite_folga_duplicada_sobreposta(self):
        main.registrar_escala_folga(
            "MOTORISTA",
            "Joao Silva",
            pessoa_codigo="M01",
            data_inicio="2026-05-27",
            data_fim="2026-05-28",
        )

        with self.assertRaises(ValueError):
            main.registrar_escala_folga(
                "MOTORISTA",
                "Joao Silva",
                pessoa_codigo="M01",
                data_inicio="2026-05-28",
                data_fim="2026-05-29",
            )


if __name__ == "__main__":
    unittest.main()
