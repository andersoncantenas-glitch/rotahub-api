import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402


class ApiTransbordoFlowTests(unittest.TestCase):
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
                    status TEXT,
                    status_operacional TEXT,
                    motorista TEXT,
                    motorista_id INTEGER,
                    motorista_codigo TEXT,
                    codigo_motorista TEXT,
                    veiculo TEXT,
                    equipe TEXT,
                    data_criacao TEXT,
                    local_rota TEXT,
                    local_carregamento TEXT,
                    tipo_estimativa TEXT,
                    operacao_tipo TEXT,
                    transbordo_modalidade TEXT,
                    transbordo_observacao TEXT,
                    transbordo_grupo TEXT,
                    usuario_criacao TEXT,
                    usuario_ultima_edicao TEXT,
                    nf_caixas INTEGER,
                    total_caixas INTEGER,
                    caixas_carregadas INTEGER,
                    caixas_estimado INTEGER,
                    km_inicial REAL,
                    km_final REAL,
                    data_chegada TEXT,
                    hora_chegada TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE veiculos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    placa TEXT,
                    modelo TEXT,
                    capacidade_cx INTEGER
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE equipes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    nome TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE programacao_itens (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    cod_cliente TEXT,
                    nome_cliente TEXT,
                    qnt_caixas INTEGER,
                    preco REAL,
                    caixas_atual INTEGER,
                    pedido TEXT,
                    produto TEXT,
                    status_pedido TEXT,
                    alteracao_tipo TEXT,
                    alteracao_detalhe TEXT,
                    alterado_em TEXT,
                    alterado_por TEXT,
                    carga_raiz_programacao TEXT,
                    carga_origem_imediata TEXT,
                    transferencia_origem_id TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE programacao_itens_controle (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    cod_cliente TEXT,
                    pedido TEXT,
                    status_pedido TEXT,
                    alteracao_tipo TEXT,
                    alteracao_detalhe TEXT,
                    mortalidade_aves INTEGER,
                    media_aplicada REAL,
                    peso_previsto REAL,
                    valor_recebido REAL,
                    forma_recebimento TEXT,
                    obs_recebimento TEXT,
                    preco_atual REAL,
                    lat_evento REAL,
                    lon_evento REAL,
                    endereco_evento TEXT,
                    cidade_evento TEXT,
                    bairro_evento TEXT,
                    ordem_sugerida INTEGER,
                    eta TEXT,
                    distancia REAL,
                    confianca_localizacao REAL,
                    caixas_atual INTEGER,
                    alterado_em TEXT,
                    alterado_por TEXT,
                    updated_at TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE transferencias (
                    id TEXT PRIMARY KEY,
                    codigo_origem TEXT,
                    codigo_destino TEXT,
                    cod_cliente TEXT,
                    pedido TEXT,
                    qtd_caixas INTEGER,
                    status TEXT,
                    obs TEXT,
                    snapshot TEXT,
                    motorista_origem TEXT,
                    motorista_destino TEXT,
                    qtd_convertida INTEGER DEFAULT 0,
                    criado_em TEXT,
                    atualizado_em TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE transferencias_conversoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    transferencia_id TEXT NOT NULL,
                    pedido_destino TEXT,
                    cod_cliente_destino TEXT,
                    qtd INTEGER,
                    obs TEXT,
                    nome_cliente_destino TEXT,
                    novo_cliente INTEGER DEFAULT 0,
                    criado_em TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE recebimentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    cod_cliente TEXT,
                    pedido TEXT,
                    nome_cliente TEXT,
                    valor REAL,
                    forma_pagamento TEXT,
                    observacao TEXT,
                    num_nf TEXT,
                    data_registro TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE despesas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    descricao TEXT,
                    valor REAL,
                    categoria TEXT,
                    observacao TEXT,
                    data_registro TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE rota_substituicoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_programacao TEXT,
                    status TEXT
                )
                """
            )
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, status, status_operacional, motorista, motorista_codigo,
                    codigo_motorista, veiculo, data_criacao, local_rota, local_carregamento,
                    tipo_estimativa, operacao_tipo, transbordo_modalidade, transbordo_observacao, transbordo_grupo,
                    nf_caixas, total_caixas, caixas_carregadas, caixas_estimado, km_inicial
                )
                VALUES ('TR-ORIG', 'EM_ENTREGAS', '', 'MOTORISTA ORIGEM', 'M1', 'M1',
                        'TRK1A11', '2026-05-20 08:00:00', 'SERRA', 'GRANJA ORIGEM',
                        'CX', 'TRANSBORDO', 'EMPRESA_BUSCA', 'ROTA DE TRANSBORDO', 'TR-ORIG',
                        10, 10, 10, 10, 100)
                """
            )
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, status, status_operacional, motorista, motorista_codigo,
                    codigo_motorista, veiculo, data_criacao, local_rota, local_carregamento,
                    tipo_estimativa, operacao_tipo, nf_caixas, total_caixas,
                    caixas_carregadas, caixas_estimado, km_inicial
                )
                VALUES ('TR-DEST', 'EM_ENTREGAS', '', 'MOTORISTA DESTINO', 'M2', 'M2',
                        'TRK2A22', '2026-05-20 09:00:00', 'SERRA', 'GRANJA DESTINO',
                        'KG', 'VENDA', 0, 0, 0, 10, 50)
                """
            )
            conn.execute("INSERT INTO veiculos (placa, modelo, capacidade_cx) VALUES ('TRK1A11', 'TRUCK', 20)")
            conn.execute("INSERT INTO veiculos (placa, modelo, capacidade_cx) VALUES ('TRK2A22', 'TRUCK', 20)")

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def _motorista(self):
        return {"id": 1, "codigo": "M1", "nome": "MOTORISTA ORIGEM", "is_admin": True}

    def test_app_rota_outputs_include_transbordo_metadata(self):
        rotas = api_server.rotas_ativas(self._motorista())
        by_codigo = {row["codigo_programacao"]: row for row in rotas}
        origem = by_codigo["TR-ORIG"]
        self.assertEqual(origem["tipo_estimativa"], "CX")
        self.assertEqual(origem["operacao_tipo"], "TRANSBORDO")
        self.assertTrue(origem["transbordo"])
        self.assertEqual(origem["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(origem["transbordo_observacao"], "ROTA DE TRANSBORDO")
        self.assertEqual(origem["transbordo_grupo"], "TR-ORIG")

        detalhe = api_server.rota_detalhe("TR-ORIG", self._motorista())
        rota = detalhe["rota"]
        self.assertEqual(rota["operacao_tipo"], "TRANSBORDO")
        self.assertTrue(rota["transbordo"])
        self.assertEqual(rota["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(rota["transbordo_grupo"], "TR-ORIG")

        bundle = api_server.desktop_rota_bundle("TR-ORIG", _ok=True)
        self.assertEqual(bundle["rota"]["operacao_tipo"], "TRANSBORDO")
        self.assertTrue(bundle["rota"]["transbordo"])
        self.assertEqual(bundle["rota"]["transbordo_modalidade"], "EMPRESA_BUSCA")
        self.assertEqual(bundle["rota"]["transbordo_grupo"], "TR-ORIG")

    def test_transbordo_planejado_nao_ocupa_capacidade_antes_do_aceite(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, status, status_operacional, motorista, motorista_codigo,
                    codigo_motorista, veiculo, data_criacao, local_rota, local_carregamento,
                    tipo_estimativa, operacao_tipo, transbordo_modalidade, transbordo_grupo,
                    nf_caixas, total_caixas, caixas_carregadas, caixas_estimado, km_inicial
                )
                VALUES ('TR-PLAN', 'ATIVA', '', 'MOTORISTA DESTINO', 'M2', 'M2',
                        'TRK3A33', '2026-05-20 09:30:00', 'SERRA', 'TRANSBORDO',
                        'CX', 'TRANSBORDO', 'EMPRESA_BUSCA', 'TR-PLAN',
                        0, 290, 0, 290, 60)
                """
            )
            conn.execute("INSERT INTO veiculos (placa, modelo, capacidade_cx) VALUES ('TRK3A33', 'TRUCK', 420)")
            conn.execute(
                """
                INSERT INTO programacao_itens (
                    codigo_programacao, cod_cliente, nome_cliente, qnt_caixas,
                    caixas_atual, pedido, status_pedido
                )
                VALUES ('TR-PLAN', 'CLI-PLAN', 'CLIENTE PLANEJADO', 290, 290, 'PED-PLAN', 'PENDENTE')
                """
            )

        rota_vazia = api_server.rota_detalhe("TR-PLAN", self._motorista())["rota"]
        self.assertEqual(rota_vazia["capacidade_cx"], 420)
        self.assertEqual(rota_vazia["caixas_saldo"], 0)

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO transferencias (
                    id, codigo_origem, codigo_destino, cod_cliente, pedido, qtd_caixas,
                    status, snapshot, qtd_convertida, criado_em, atualizado_em
                )
                VALUES (
                    'TRF-PLAN', 'TR-ORIG', 'TR-PLAN', 'TRANSBORDO', 'TRANSBORDO',
                    290, 'ACEITA',
                    '{"carga_raiz_programacao":"TR-ORIG","carga_origem_imediata":"TR-ORIG","transbordo":true}',
                    0, '2026-05-20 10:00:00', '2026-05-20 10:05:00'
                )
                """
            )

        rota_carregada = api_server.rota_detalhe("TR-PLAN", self._motorista())["rota"]
        self.assertEqual(rota_carregada["caixas_saldo"], 290)
        self.assertEqual(rota_carregada["capacidade_cx"] - rota_carregada["caixas_saldo"], 130)

    def test_planejamento_de_venda_nao_ocupa_capacidade_antes_do_carregamento(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO programacoes (
                    codigo_programacao, status, status_operacional, motorista, motorista_codigo,
                    codigo_motorista, veiculo, data_criacao, local_rota, local_carregamento,
                    tipo_estimativa, operacao_tipo, nf_caixas, total_caixas,
                    caixas_carregadas, caixas_estimado, km_inicial
                )
                VALUES ('SALE-PLAN', 'ATIVA', '', 'MOTORISTA DESTINO', 'M2', 'M2',
                        'TRK4A44', '2026-05-20 09:40:00', 'SERRA', 'GRANJA DESTINO',
                        'KG', 'VENDA', 290, 290, 0, 0, 70)
                """
            )
            conn.execute("INSERT INTO veiculos (placa, modelo, capacidade_cx) VALUES ('TRK4A44', 'TRUCK', 420)")
            conn.execute(
                """
                INSERT INTO programacao_itens (
                    codigo_programacao, cod_cliente, nome_cliente, qnt_caixas,
                    caixas_atual, pedido, status_pedido
                )
                VALUES ('SALE-PLAN', 'CLI-SALE', 'CLIENTE VENDA', 290, 290, 'PED-SALE', 'PENDENTE')
                """
            )

        rota_planejada = api_server.rota_detalhe("SALE-PLAN", self._motorista())["rota"]
        self.assertEqual(rota_planejada["capacidade_cx"], 420)
        self.assertEqual(rota_planejada["caixas_saldo"], 0)

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                UPDATE programacoes
                   SET caixas_carregadas=290
                 WHERE codigo_programacao='SALE-PLAN'
                """
            )

        rota_carregada = api_server.rota_detalhe("SALE-PLAN", self._motorista())["rota"]
        self.assertEqual(rota_carregada["caixas_saldo"], 290)
        self.assertEqual(rota_carregada["capacidade_cx"] - rota_carregada["caixas_saldo"], 130)

    def test_finalizacao_bloqueia_transferencia_aceita_nao_convertida(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO transferencias (
                    id, codigo_origem, codigo_destino, cod_cliente, pedido, qtd_caixas,
                    status, qtd_convertida, criado_em, atualizado_em
                )
                VALUES ('TRF-PARCIAL', 'TR-ORIG', 'TR-DEST', 'TRANSBORDO', 'TRANSBORDO',
                        4, 'ACEITA', 2, '2026-05-20 10:00:00', '2026-05-20 10:05:00')
                """
            )

        with self.assertRaises(api_server.HTTPException) as ctx:
            api_server.finalizar_rota(
                "TR-ORIG",
                api_server.FinalizarRotaIn(data_chegada="2026-05-20", hora_chegada="11:00", km_final=150),
                self._motorista(),
            )

        self.assertEqual(ctx.exception.status_code, 409)
        self.assertIn("nao convertida", ctx.exception.detail)

    def test_conversao_total_marca_transferencia_como_convertida(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO transferencias (
                    id, codigo_origem, codigo_destino, cod_cliente, pedido, qtd_caixas,
                    status, snapshot, qtd_convertida, criado_em, atualizado_em
                )
                VALUES (
                    'TRF-TOTAL', 'TR-ORIG', 'TR-DEST', 'TRANSBORDO', 'TRANSBORDO',
                    4, 'ACEITA',
                    '{"carga_raiz_programacao":"TR-ORIG","carga_origem_imediata":"TR-ORIG","transbordo":true}',
                    0, '2026-05-20 10:00:00', '2026-05-20 10:05:00'
                )
                """
            )

        result = api_server.converter_transferencia(
            "TRF-TOTAL",
            api_server.TransferenciaConverterIn(
                pedido_destino="PED-DEST",
                cod_cliente_destino="CLI-DEST",
                qtd_caixas=4,
                novo_cliente={"pedido": "PED-DEST", "cod_cliente": "CLI-DEST", "nome_cliente": "CLIENTE DESTINO"},
            ),
            self._motorista(),
        )

        self.assertEqual(result["status"], "CONVERTIDA")
        self.assertEqual(result["qtd_convertida"], 4)
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT status, qtd_convertida FROM transferencias WHERE id='TRF-TOTAL'"
            ).fetchone()
            item = conn.execute(
                """
                SELECT qnt_caixas, carga_raiz_programacao, carga_origem_imediata, transferencia_origem_id
                FROM programacao_itens
                WHERE codigo_programacao='TR-DEST' AND cod_cliente='CLI-DEST' AND pedido='PED-DEST'
                """
            ).fetchone()

        self.assertEqual(row, ("CONVERTIDA", 4))
        self.assertEqual(item, (4, "TR-ORIG", "TR-ORIG", "TRF-TOTAL"))

    def test_fluxo_destino_entrega_recebimento_e_finalizacao(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO transferencias (
                    id, codigo_origem, codigo_destino, cod_cliente, pedido, qtd_caixas,
                    status, snapshot, qtd_convertida, criado_em, atualizado_em
                )
                VALUES (
                    'TRF-FECHA', 'TR-ORIG', 'TR-DEST', 'TRANSBORDO', 'TRANSBORDO',
                    4, 'ACEITA',
                    '{"carga_raiz_programacao":"TR-ORIG","carga_origem_imediata":"TR-ORIG","transbordo":true}',
                    0, '2026-05-20 10:00:00', '2026-05-20 10:05:00'
                )
                """
            )

        api_server.converter_transferencia(
            "TRF-FECHA",
            api_server.TransferenciaConverterIn(
                pedido_destino="PED-FECHA",
                cod_cliente_destino="CLI-FECHA",
                qtd_caixas=4,
                novo_cliente={"pedido": "PED-FECHA", "cod_cliente": "CLI-FECHA", "nome_cliente": "CLIENTE FECHA"},
            ),
            self._motorista(),
        )

        entrega = api_server.salvar_controle_cliente(
            "TR-DEST",
            api_server.ClienteControleIn(
                cod_cliente="CLI-FECHA",
                pedido="PED-FECHA",
                status_pedido="ENTREGUE",
                valor_recebido=200,
                forma_recebimento="PIX",
                obs_recebimento="RECEBIDO NO TRANSBORDO",
                evento_em="2026-05-20 12:00:00",
            ),
            self._motorista(),
        )
        self.assertTrue(entrega["ok"])

        finalizada = api_server.finalizar_rota(
            "TR-DEST",
            api_server.FinalizarRotaIn(data_chegada="2026-05-20", hora_chegada="13:00", km_final=90),
            self._motorista(),
        )
        self.assertEqual(finalizada["status"], "FINALIZADA")

        with sqlite3.connect(self.db_path) as conn:
            receb = conn.execute(
                """
                SELECT valor, forma_pagamento, observacao
                FROM recebimentos
                WHERE codigo_programacao='TR-DEST' AND cod_cliente='CLI-FECHA' AND pedido='PED-FECHA'
                """
            ).fetchone()
            controle = conn.execute(
                """
                SELECT status_pedido, caixas_atual, valor_recebido, forma_recebimento
                FROM programacao_itens_controle
                WHERE codigo_programacao='TR-DEST' AND cod_cliente='CLI-FECHA' AND pedido='PED-FECHA'
                """
            ).fetchone()
            prog = conn.execute(
                "SELECT status, data_chegada, hora_chegada, km_final FROM programacoes WHERE codigo_programacao='TR-DEST'"
            ).fetchone()

        self.assertEqual(receb, (200.0, "PIX", "RECEBIDO NO TRANSBORDO"))
        self.assertEqual(controle, ("ENTREGUE", 0, 200.0, "PIX"))
        self.assertEqual(prog, ("FINALIZADA", "2026-05-20", "13:00", 90.0))


if __name__ == "__main__":
    unittest.main()
