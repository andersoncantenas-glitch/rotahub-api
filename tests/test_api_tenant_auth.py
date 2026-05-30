import os
import sqlite3
import tempfile
import unittest
from types import SimpleNamespace

from fastapi.security import HTTPAuthorizationCredentials


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402
from db_bootstrap import ensure_core_schema  # noqa: E402


class ApiTenantAuthTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        api_server.DB_PATH = self.db_path
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            ensure_core_schema(conn)
            cur = conn.cursor()
            cur.execute("SELECT id FROM companies WHERE code='default' LIMIT 1")
            self.company_id = int(cur.fetchone()["id"])
            cur.execute(
                """
                INSERT INTO motoristas (
                    nome, codigo, senha, perfil_app, acesso_liberado, company_id
                )
                VALUES ('MOTORISTA UM', 'MOT-01', '123456', 'MOTORISTA', 1, ?)
                """,
                (self.company_id,),
            )
            cur.execute(
                """
                INSERT INTO vendedores (codigo, nome, senha, status, company_id)
                VALUES ('VEND-01', 'VENDEDOR UM', '123456', 'ATIVO', ?)
                """,
                (self.company_id,),
            )

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def test_token_carries_company_and_user_context(self):
        token = api_server.create_token(
            "MOT-01",
            perfil="motorista",
            company_id=self.company_id,
            user_id=123,
            username="MOTORISTA UM",
            role="MOTORISTA",
        )

        payload = api_server.verify_token(token)

        self.assertEqual(payload["codigo"], "MOT-01")
        self.assertEqual(payload["company_id"], self.company_id)
        self.assertEqual(payload["user_id"], 123)
        self.assertEqual(payload["username"], "MOTORISTA UM")
        self.assertEqual(payload["role"], "MOTORISTA")

    def test_motorista_login_and_dependency_validate_company_id(self):
        result = api_server.autenticar_motorista(api_server.LoginIn(codigo="MOT-01", senha="123456"))

        self.assertEqual(result["company_id"], self.company_id)
        payload = api_server.verify_token(result["token"])
        self.assertEqual(payload["company_id"], self.company_id)

        credentials = HTTPAuthorizationCredentials(scheme="Bearer", credentials=result["token"])
        current = api_server.get_current_motorista(credentials)
        self.assertEqual(current["company_id"], self.company_id)
        self.assertEqual(current["codigo"], "MOT-01")

        wrong_token = api_server.create_token(
            "MOT-01",
            perfil="motorista",
            company_id=self.company_id + 1,
            user_id=1,
            username="MOTORISTA UM",
        )
        wrong_credentials = HTTPAuthorizationCredentials(scheme="Bearer", credentials=wrong_token)
        with self.assertRaises(api_server.HTTPException) as ctx:
            api_server.get_current_motorista(wrong_credentials)
        self.assertEqual(ctx.exception.status_code, 403)

    def test_vendedor_login_includes_company_id(self):
        request = SimpleNamespace(client=SimpleNamespace(host="127.0.0.1"))
        result = api_server.autenticar_vendedor(api_server.LoginIn(codigo="VEND-01", senha="123456"), request)

        self.assertEqual(result["company_id"], self.company_id)
        payload = api_server.verify_token(result["token"])
        self.assertEqual(payload["company_id"], self.company_id)

        credentials = HTTPAuthorizationCredentials(scheme="Bearer", credentials=result["token"])
        current = api_server.get_current_vendedor(credentials)
        self.assertEqual(current["company_id"], self.company_id)
        self.assertEqual(current["codigo"], "VEND-01")


if __name__ == "__main__":
    unittest.main()
