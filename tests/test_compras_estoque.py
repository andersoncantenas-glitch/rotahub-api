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


class ComprasEstoqueTests(unittest.TestCase):
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

    def test_ajuste_e_zeragem_fisica_preservam_movimentos(self):
        with TestClient(app, base_url="http://testserver") as client:
            headers = self._headers(client)

            ajuste = client.post(
                "/api/v1/compras/estoque/ajuste",
                headers=headers,
                json={
                    "tipo_estoque": "FISICO",
                    "tipo_movimento": "ENTRADA",
                    "produto": "FRANGO VIVO",
                    "quantidade_kg": 125.5,
                    "quantidade_caixas": 10,
                    "observacao": "teste ajuste",
                },
            )
            self.assertEqual(ajuste.status_code, 200, ajuste.text)
            self.assertEqual(ajuste.json()["movimentos_criados"], 1)

            resumo = client.get("/api/v1/compras/estoque/resumo", headers=headers)
            self.assertEqual(resumo.status_code, 200, resumo.text)
            self.assertEqual(resumo.json()["fisico_saldo_kg"], 125.5)

            zerar = client.post("/api/v1/compras/estoque/zerar-fisico", headers=headers)
            self.assertEqual(zerar.status_code, 200, zerar.text)
            self.assertEqual(zerar.json()["movimentos_criados"], 1)

            resumo_zerado = client.get("/api/v1/compras/estoque/resumo", headers=headers)
            self.assertEqual(resumo_zerado.status_code, 200, resumo_zerado.text)
            self.assertEqual(resumo_zerado.json()["fisico_saldo_kg"], 0)


if __name__ == "__main__":
    unittest.main()
