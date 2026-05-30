from __future__ import annotations

import asyncio
import os
import tempfile
import unittest
from pathlib import Path


fd, DB_PATH = tempfile.mkstemp(suffix=".db")
os.close(fd)

os.environ["DATABASE_URL"] = f"sqlite+aiosqlite:///{Path(DB_PATH).as_posix()}"
os.environ["DEBUG"] = "false"
os.environ["JWT_SECRET_KEY"] = "test-secret"
os.environ["ALLOWED_HOSTS"] = "localhost,127.0.0.1,testserver"
os.environ["RATE_LIMIT_REQUESTS"] = "10000"

from fastapi.testclient import TestClient  # noqa: E402

from backend.config.database import async_session, create_tables, engine  # noqa: E402
from backend.main import app  # noqa: E402
from backend.models.user import UserDB  # noqa: E402
from backend.services.auth import get_password_hash  # noqa: E402


class LogisticaConfigTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        asyncio.run(create_tables())
        asyncio.run(cls._seed_admin())

    @classmethod
    def tearDownClass(cls):
        asyncio.run(engine.dispose())
        try:
            os.unlink(DB_PATH)
        except OSError:
            pass

    @classmethod
    async def _seed_admin(cls):
        async with async_session() as session:
            session.add(
                UserDB(
                    username="admin",
                    nome="ADMIN",
                    senha=get_password_hash("Admin@123456"),
                    permissoes="ADMIN",
                )
            )
            await session.commit()

    def _headers(self, client: TestClient):
        login = client.post("/api/v1/auth/login", data={"username": "admin", "password": "Admin@123456"})
        self.assertEqual(login.status_code, 200)
        return {"Authorization": f"Bearer {login.json()['access_token']}"}

    def test_default_profile_preserves_frango_vivo_and_allows_generic_update(self):
        with TestClient(app, base_url="http://testserver") as client:
            headers = self._headers(client)

            current = client.get("/api/v1/logistica/config", headers=headers)
            self.assertEqual(current.status_code, 200, current.text)
            self.assertEqual(current.json()["config"]["perfil_codigo"], "FRANGO_VIVO")
            self.assertEqual(current.json()["config"]["produto_padrao"], "AVES VIVAS")

            updated = client.put(
                "/api/v1/logistica/config",
                headers=headers,
                json={
                    "company_id": current.json()["company_id"],
                    "perfil_codigo": "DISTRIBUICAO_GERAL",
                    "produto_padrao": "POLPA",
                    "unidade_padrao": "KG",
                    "embalagem_label": "Caixas",
                    "perda_label": "Avaria",
                    "quantidade_embalagem_label": "Qtd. por caixa",
                    "usa_mortalidade": 0,
                    "usa_aves_por_caixa": 0,
                    "usa_nota_fiscal_motorista": 1,
                    "usa_estoque_fisico": 1,
                    "usa_estoque_fiscal": 1,
                },
            )
            self.assertEqual(updated.status_code, 200, updated.text)
            self.assertEqual(updated.json()["config"]["perfil_codigo"], "DISTRIBUICAO_GERAL")
            self.assertEqual(updated.json()["config"]["perda_label"], "Avaria")

            ajuste = client.post(
                "/api/v1/compras/estoque/ajuste",
                headers=headers,
                json={
                    "tipo_estoque": "FISICO",
                    "tipo_movimento": "ENTRADA",
                    "quantidade_kg": 10,
                    "quantidade_caixas": 2,
                    "observacao": "perfil generico",
                },
            )
            self.assertEqual(ajuste.status_code, 200, ajuste.text)

            resumo = client.get("/api/v1/compras/estoque/resumo", headers=headers)
            self.assertEqual(resumo.status_code, 200, resumo.text)
            produtos = resumo.json()["produtos"]
            self.assertTrue(any(produto["nome"] == "POLPA" for produto in produtos))

    def test_can_create_operational_profile_unit_and_occurrence(self):
        with TestClient(app, base_url="http://testserver") as client:
            headers = self._headers(client)

            unidade = client.post(
                "/api/v1/logistica/unidades",
                headers=headers,
                json={"codigo": "BD", "nome": "Balde", "tipo": "EMBALAGEM"},
            )
            self.assertEqual(unidade.status_code, 200, unidade.text)

            perfil = client.post(
                "/api/v1/logistica/perfis",
                headers=headers,
                json={
                    "codigo": "POLPAS_CONGELADAS",
                    "nome": "Polpas congeladas",
                    "descricao": "Distribuicao de polpas com controle por balde.",
                    "produto_padrao": "POLPA",
                    "unidade_padrao": "BD",
                    "embalagem_label": "Baldes",
                    "perda_label": "Avaria",
                    "quantidade_embalagem_label": "Unidades por balde",
                    "usa_mortalidade": 0,
                    "usa_aves_por_caixa": 0,
                    "usa_nota_fiscal_motorista": 1,
                    "usa_estoque_fisico": 1,
                    "usa_estoque_fiscal": 1,
                },
            )
            self.assertEqual(perfil.status_code, 200, perfil.text)

            ocorrencia = client.post(
                "/api/v1/logistica/ocorrencias",
                headers=headers,
                json={
                    "codigo": "DESCONGELAMENTO",
                    "nome": "Descongelamento",
                    "categoria": "PERDA",
                    "perfil_codigo": "POLPAS_CONGELADAS",
                },
            )
            self.assertEqual(ocorrencia.status_code, 200, ocorrencia.text)

            current = client.get("/api/v1/logistica/config", headers=headers)
            self.assertEqual(current.status_code, 200, current.text)
            data = current.json()
            self.assertTrue(any(item["codigo"] == "BD" for item in data["unidades"]))
            self.assertTrue(any(item["codigo"] == "POLPAS_CONGELADAS" for item in data["perfis"]))
            self.assertTrue(any(item["codigo"] == "DESCONGELAMENTO" for item in data["ocorrencias"]))


if __name__ == "__main__":
    unittest.main()
