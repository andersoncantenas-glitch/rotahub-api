from unittest.mock import patch

from app.services import auth_service

from tests._contract_test_helpers import ContractTestCase, _FakeConn, _FakeCtx, _FakeCursor


class AuthServiceContractTests(ContractTestCase):
    def test_auth_invalid_credentials_returns_contract(self):
        result = auth_service.autenticar_usuario("", "")
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "local")
        self.assertEqual(result["error"], "Credenciais inválidas.")

    def test_auth_success_returns_contract_with_user_data(self):
        rows = [
            (0, "id", "INTEGER", 0, None, 1),
            (1, "nome", "TEXT", 0, None, 0),
            (2, "senha", "TEXT", 0, None, 0),
            (3, "permissoes", "TEXT", 0, None, 0),
            (4, "cpf", "TEXT", 0, None, 0),
            (5, "telefone", "TEXT", 0, None, 0),
        ]
        user_row = (1, "ADMIN", "1234", "ADMIN", "12345678901", "11999999999")
        fake_ctx = _FakeCtx(_FakeConn(_FakeCursor(table_info_rows=rows, fetchone_rows=[user_row])))
        with patch.object(auth_service, "get_db", return_value=fake_ctx):
            result = auth_service.autenticar_usuario("ADMIN", "1234")

        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "local")
        self.assertIsNone(result["error"])
        self.assertIsInstance(result["data"], dict)
        self.assertEqual(result["data"].get("nome"), "ADMIN")
        self.assertTrue(bool(result["data"].get("is_admin")))

    def test_auth_ensure_admin_without_password_column_returns_contract_error(self):
        rows = [
            (0, "id", "INTEGER", 0, None, 1),
            (1, "nome", "TEXT", 0, None, 0),
        ]
        fake_ctx = _FakeCtx(_FakeConn(_FakeCursor(table_info_rows=rows)))
        with patch.object(auth_service, "get_db", return_value=fake_ctx):
            result = auth_service.ensure_admin_user()
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "local")
        self.assertEqual(result["error"], "Base sem coluna de senha.")
