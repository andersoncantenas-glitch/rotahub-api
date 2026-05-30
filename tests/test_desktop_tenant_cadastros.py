import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402
from db_bootstrap import ensure_core_schema  # noqa: E402


class DesktopTenantCadastrosTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        api_server.DB_PATH = self.db_path
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            ensure_core_schema(conn)
            cur = conn.cursor()
            cur.execute("SELECT id FROM companies WHERE code='default' LIMIT 1")
            self.company_a = int(cur.fetchone()["id"])
            self.company_b = self._create_company(cur, "cliente-b")

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def _create_company(self, cur, code: str) -> int:
        cur.execute(
            "INSERT INTO companies (code, name, status) VALUES (?, ?, 'active')",
            (code, code.upper()),
        )
        company_id = int(cur.lastrowid)
        cur.execute("SELECT id FROM plans WHERE code='professional' LIMIT 1")
        plan_id = int(cur.fetchone()[0])
        cur.execute(
            """
            INSERT INTO subscriptions (
                company_id, plan_id, status, billing_cycle,
                current_period_start, current_period_end, next_due_date
            )
            VALUES (?, ?, 'active', 'monthly', date('now'), date('now', '+30 day'), date('now', '+30 day'))
            """,
            (company_id, plan_id),
        )
        return company_id

    def _insert_dynamic(self, cur, table: str, values: dict) -> None:
        cur.execute(f"PRAGMA table_info({table})")
        cols = {str(row[1]) for row in cur.fetchall()}
        data = {key: value for key, value in values.items() if key in cols}
        placeholders = ", ".join(["?"] * len(data))
        cur.execute(
            f"INSERT INTO {table} ({', '.join(data.keys())}) VALUES ({placeholders})",
            tuple(data.values()),
        )

    def test_desktop_clientes_upsert_list_and_delete_are_company_scoped(self):
        api_server.desktop_clientes_upsert(
            api_server.DesktopClienteUpsertIn(cod_cliente="CLI-01", nome_cliente="Cliente A"),
            _ok=True,
            x_company_id=str(self.company_a),
        )
        api_server.desktop_clientes_upsert(
            api_server.DesktopClienteUpsertIn(cod_cliente="CLI-01", nome_cliente="Cliente B"),
            _ok=True,
            x_company_id=str(self.company_b),
        )

        rows_a = api_server.desktop_clientes_base(q="", vendedor="", cidade="", ordem="nome", limit=300, _ok=True, x_company_id=str(self.company_a))
        rows_b = api_server.desktop_clientes_base(q="", vendedor="", cidade="", ordem="nome", limit=300, _ok=True, x_company_id=str(self.company_b))

        self.assertEqual([row["nome_cliente"] for row in rows_a], ["CLIENTE A"])
        self.assertEqual([row["nome_cliente"] for row in rows_b], ["CLIENTE B"])

        deleted = api_server.desktop_clientes_delete("CLI-01", _ok=True, x_company_id=str(self.company_a))
        self.assertEqual(deleted["deleted"], 1)

        rows_a = api_server.desktop_clientes_base(q="", vendedor="", cidade="", ordem="nome", limit=300, _ok=True, x_company_id=str(self.company_a))
        rows_b = api_server.desktop_clientes_base(q="", vendedor="", cidade="", ordem="nome", limit=300, _ok=True, x_company_id=str(self.company_b))

        self.assertEqual(rows_a, [])
        self.assertEqual([row["nome_cliente"] for row in rows_b], ["CLIENTE B"])

    def test_desktop_clientes_replaces_legacy_global_unique_index(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("DROP INDEX IF EXISTS idx_clientes_company_cod")
            conn.execute("CREATE UNIQUE INDEX idx_clientes_cod_legacy ON clientes(cod_cliente)")

        api_server.desktop_clientes_upsert(
            api_server.DesktopClienteUpsertIn(cod_cliente="CLI-02", nome_cliente="Cliente A"),
            _ok=True,
            x_company_id=str(self.company_a),
        )
        api_server.desktop_clientes_upsert(
            api_server.DesktopClienteUpsertIn(cod_cliente="CLI-02", nome_cliente="Cliente B"),
            _ok=True,
            x_company_id=str(self.company_b),
        )

        with sqlite3.connect(self.db_path) as conn:
            indexes = conn.execute("PRAGMA index_list(clientes)").fetchall()
            names = {row[1] for row in indexes}
            total = conn.execute("SELECT COUNT(*) FROM clientes WHERE cod_cliente='CLI-02'").fetchone()[0]

        self.assertIn("idx_clientes_company_cod", names)
        self.assertNotIn("idx_clientes_cod_legacy", names)
        self.assertEqual(total, 2)

    def test_desktop_motorista_upsert_rejects_invalid_registration_data(self):
        cases = [
            api_server.DesktopMotoristaUpsertIn(
                codigo="AA",
                nome="Motorista Valido",
                telefone="88999999999",
                senha="1234",
            ),
            api_server.DesktopMotoristaUpsertIn(
                codigo="MOT-10",
                nome="AB",
                telefone="88999999999",
                senha="1234",
            ),
            api_server.DesktopMotoristaUpsertIn(
                codigo="MOT-10",
                nome="Motorista Valido",
                telefone="123",
                senha="1234",
            ),
            api_server.DesktopMotoristaUpsertIn(
                codigo="MOT-10",
                nome="Motorista Valido",
                telefone="88999999999",
                cpf="11111111111",
                senha="1234",
            ),
            api_server.DesktopMotoristaUpsertIn(
                codigo="MOT-10",
                nome="Motorista Valido",
                telefone="88999999999",
                status="BLOQUEADO",
                senha="1234",
            ),
            api_server.DesktopMotoristaUpsertIn(
                codigo="MOT-10",
                nome="Motorista Valido",
                telefone="88999999999",
            ),
        ]

        for payload in cases:
            with self.subTest(payload=payload.model_dump()):
                with self.assertRaises(api_server.HTTPException) as ctx:
                    api_server.desktop_motoristas_upsert(payload, _ok=True, x_company_id=str(self.company_a))
                self.assertEqual(ctx.exception.status_code, 400)

    def test_desktop_vendedor_and_ajudante_upsert_reject_invalid_registration_data(self):
        vendedor_payload = api_server.DesktopVendedorUpsertIn(
            codigo="VEN-01",
            nome="Vendedor Valido",
            telefone="123",
            senha="1234",
        )
        with self.assertRaises(api_server.HTTPException) as vendedor_ctx:
            api_server.desktop_vendedores_upsert(vendedor_payload, _ok=True, x_company_id=str(self.company_a))
        self.assertEqual(vendedor_ctx.exception.status_code, 400)

        vendedor_sem_senha = api_server.DesktopVendedorUpsertIn(
            codigo="VEN-01",
            nome="Vendedor Valido",
        )
        with self.assertRaises(api_server.HTTPException) as senha_ctx:
            api_server.desktop_vendedores_upsert(vendedor_sem_senha, _ok=True, x_company_id=str(self.company_a))
        self.assertEqual(senha_ctx.exception.status_code, 400)

        ajudante_payload = api_server.DesktopAjudanteUpsertIn(
            nome="Ajudante",
            sobrenome="Valido",
            telefone="123",
        )
        with self.assertRaises(api_server.HTTPException) as ajudante_ctx:
            api_server.desktop_ajudantes_upsert(ajudante_payload, _ok=True, x_company_id=str(self.company_a))
        self.assertEqual(ajudante_ctx.exception.status_code, 400)

    def test_desktop_veiculo_upsert_rejects_invalid_registration_data(self):
        cases = [
            api_server.DesktopVeiculoUpsertIn(placa="@@@", modelo="Truck", capacidade_cx=100),
            api_server.DesktopVeiculoUpsertIn(placa="ABC1D23", modelo="Truck", capacidade_cx=-1),
        ]

        for payload in cases:
            with self.subTest(payload=payload.model_dump()):
                with self.assertRaises(api_server.HTTPException) as ctx:
                    api_server.desktop_veiculos_upsert(payload, _ok=True, x_company_id=str(self.company_a))
                self.assertEqual(ctx.exception.status_code, 400)

    def test_desktop_rota_upsert_creates_transbordo_metadata_for_cx(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            self._insert_dynamic(
                cur,
                "motoristas",
                {
                    "codigo": "MOTCX",
                    "nome": "MOTORISTA CX",
                    "status": "ATIVO",
                    "telefone": "88999991111",
                    "senha": "123456",
                    "company_id": self.company_a,
                },
            )

        api_server.desktop_rotas_upsert(
            api_server.DesktopRotaUpsertIn(
                codigo_programacao="CX-001",
                data_criacao="2026-05-21 08:00:00",
                motorista="MOTORISTA CX",
                motorista_codigo="MOTCX",
                veiculo="ABC1D23",
                equipe="AJUDANTE A|AJUDANTE B",
                tipo_estimativa="CX",
                caixas_estimado=25,
                total_caixas=25,
                local_rota="SERRA",
                local_carregamento="GRANJA TESTE",
            ),
            _ok=True,
        )

        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                """
                SELECT codigo_programacao, codigo, data, operacao_tipo,
                       transbordo_modalidade, transbordo_grupo,
                       tipo_estimativa, caixas_estimado, total_caixas,
                       local_rota, local_carregamento
                  FROM programacoes
                 WHERE codigo_programacao='CX-001'
                """
            ).fetchone()

        self.assertEqual(row["codigo"], "CX-001")
        self.assertEqual(row["data"], "2026-05-21 08:00:00")
        self.assertEqual(row["operacao_tipo"], "TRANSBORDO")
        self.assertEqual(row["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(row["transbordo_grupo"], "CX-001")
        self.assertEqual(row["tipo_estimativa"], "CX")
        self.assertEqual(int(row["caixas_estimado"]), 25)
        self.assertEqual(int(row["total_caixas"]), 25)
        self.assertEqual(row["local_rota"], "SERRA")
        self.assertEqual(row["local_carregamento"], "GRANJA TESTE")

        listed = api_server.desktop_programacoes_listar(modo="todas", limit=10, _ok=True)
        by_codigo = {item["codigo_programacao"]: item for item in listed["programacoes"]}
        self.assertEqual(by_codigo["CX-001"]["tipo_estimativa"], "CX")
        self.assertEqual(by_codigo["CX-001"]["operacao_tipo"], "TRANSBORDO")
        self.assertEqual(by_codigo["CX-001"]["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(by_codigo["CX-001"]["transbordo_grupo"], "CX-001")

        centro = api_server.desktop_centro_custos_rows(limit=10, _ok=True)
        centro_by_codigo = {item["codigo_programacao"]: item for item in centro["rows"]}
        self.assertEqual(centro_by_codigo["CX-001"]["tipo_estimativa"], "CX")
        self.assertEqual(centro_by_codigo["CX-001"]["operacao_tipo"], "TRANSBORDO")
        self.assertEqual(centro_by_codigo["CX-001"]["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(centro_by_codigo["CX-001"]["transbordo_grupo"], "CX-001")

    def test_legacy_mobile_cadastros_only_return_active_available_resources(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            self._insert_dynamic(
                cur,
                "motoristas",
                {"nome": "MOTORISTA OCUPADO", "codigo": "MOBUSY", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "motoristas",
                {"nome": "MOTORISTA LIVRE", "codigo": "MOFREE", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "motoristas",
                {"nome": "MOTORISTA INATIVO", "codigo": "MOINA", "status": "INATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "motoristas",
                {"nome": "MOTORISTA LIBERADO", "codigo": "MOCLOSED", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "veiculos",
                {"placa": "BUS1A01", "modelo": "TRUCK", "capacidade_cx": 20, "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "veiculos",
                {"placa": "FRE1A01", "modelo": "TRUCK", "capacidade_cx": 20, "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "veiculos",
                {"placa": "INA1A01", "modelo": "TRUCK", "capacidade_cx": 20, "status": "DESATIVADO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "veiculos",
                {"placa": "CLO1A01", "modelo": "TRUCK", "capacidade_cx": 20, "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "ajudantes",
                {"nome": "AJUDANTE", "sobrenome": "OCUPADO", "telefone": "88999990001", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "ajudantes",
                {"nome": "AJUDANTE", "sobrenome": "LIVRE", "telefone": "88999990002", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "ajudantes",
                {"nome": "AJUDANTE", "sobrenome": "INATIVO", "telefone": "88999990003", "status": "DESATIVADO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "ajudantes",
                {"nome": "AJUDANTE", "sobrenome": "LIBERADO", "telefone": "88999990004", "status": "ATIVO", "company_id": self.company_a},
            )
            self._insert_dynamic(
                cur,
                "programacoes",
                {
                    "codigo_programacao": "ROTA-ABERTA",
                    "motorista": "MOTORISTA OCUPADO",
                    "motorista_codigo": "MOBUSY",
                    "veiculo": "BUS1A01",
                    "equipe": "AJUDANTE OCUPADO",
                    "status": "ATIVA",
                    "status_operacional": "EM_ROTA",
                    "prestacao_status": "PENDENTE",
                    "company_id": self.company_a,
                },
            )
            self._insert_dynamic(
                cur,
                "programacoes",
                {
                    "codigo_programacao": "ROTA-FECHADA",
                    "motorista": "MOTORISTA LIBERADO",
                    "motorista_codigo": "MOCLOSED",
                    "veiculo": "CLO1A01",
                    "equipe": "AJUDANTE LIBERADO",
                    "status": "FINALIZADA",
                    "status_operacional": "FINALIZADA",
                    "prestacao_status": "FECHADA",
                    "finalizada_no_app": 1,
                    "company_id": self.company_a,
                },
            )

        motoristas = api_server.listar_cad_motoristas(m={"codigo": "MOB"})
        veiculos = api_server.listar_cad_veiculos(m={"codigo": "MOB"})
        ajudantes = api_server.listar_cad_ajudantes(m={"codigo": "MOB"})

        self.assertEqual({row["codigo"] for row in motoristas}, {"MOFREE", "MOCLOSED"})
        self.assertEqual({row["placa"] for row in veiculos}, {"FRE1A01", "CLO1A01"})
        self.assertEqual({f'{row["nome"]} {row.get("sobrenome", "")}'.strip() for row in ajudantes}, {"AJUDANTE LIVRE", "AJUDANTE LIBERADO"})

    def test_core_schema_adapts_existing_legacy_programacoes(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("DROP TABLE IF EXISTS programacao_itens")
            cur.execute("DROP TABLE IF EXISTS programacoes")
            cur.execute(
                """
                CREATE TABLE programacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    data TEXT,
                    motorista TEXT,
                    veiculo TEXT,
                    equipe TEXT,
                    tipo_estimativa TEXT,
                    status TEXT,
                    status_operacional TEXT,
                    data_chegada TEXT,
                    km_final REAL,
                    num_nf TEXT,
                    local_rota TEXT,
                    granja_carregada TEXT
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE programacao_itens (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    qnt_caixas INTEGER
                )
                """
            )
            cur.execute(
                """
                INSERT INTO programacoes (
                    codigo, data, motorista, veiculo, equipe, tipo_estimativa,
                    status, status_operacional, data_chegada, km_final,
                    num_nf, local_rota, granja_carregada
                )
                VALUES ('LEG-001', '2026-05-20', 'MOTORISTA LEGADO', 'AAA1A11', '1|2', 'CX',
                        '', 'EM_ROTA', '2026-05-21', 120, 'NF123', 'SERRA', 'GRANJA A')
                """
            )
            cur.executemany(
                "INSERT INTO programacao_itens (codigo_programacao, qnt_caixas) VALUES (?, ?)",
                [("LEG-001", 10), ("LEG-001", 5)],
            )
            ensure_core_schema(conn)
            row = conn.execute(
                """
                SELECT codigo_programacao, codigo, data_criacao, status, status_operacional,
                       finalizada_no_app, prestacao_status, operacao_tipo, transbordo_modalidade,
                       transbordo_grupo, nf_numero, local_carregamento, total_caixas, nf_caixas,
                       caixas_carregadas, qnt_cx_carregada
                  FROM programacoes
                 WHERE codigo_programacao='LEG-001'
                """
            ).fetchone()

        self.assertEqual(row["codigo"], "LEG-001")
        self.assertEqual(row["status"], "FINALIZADA")
        self.assertEqual(row["status_operacional"], "FINALIZADA")
        self.assertEqual(int(row["finalizada_no_app"]), 1)
        self.assertEqual(row["prestacao_status"], "PENDENTE")
        self.assertEqual(row["operacao_tipo"], "TRANSBORDO")
        self.assertEqual(row["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(row["transbordo_grupo"], "LEG-001")
        self.assertEqual(row["nf_numero"], "NF123")
        self.assertEqual(row["local_carregamento"], "GRANJA A")
        self.assertEqual(int(row["total_caixas"]), 15)
        self.assertEqual(int(row["nf_caixas"]), 15)
        self.assertEqual(int(row["caixas_carregadas"]), 15)
        self.assertEqual(int(row["qnt_cx_carregada"]), 15)


if __name__ == "__main__":
    unittest.main()
