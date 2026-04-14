import os
from unittest.mock import patch

from app.services import cliente_service

from tests._contract_test_helpers import ContractTestCase


class ClienteServiceContractTests(ContractTestCase):
    def test_cliente_sync_error_uses_contextual_default_message(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(cliente_service, "_call_api", side_effect=Exception()):
                result = cliente_service.sync_cliente_upsert_api(
                    "C1",
                    "Cliente",
                    "Endereco",
                    "11999999999",
                    "Vend",
                    is_desktop_api_sync_enabled=lambda: True,
                )
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertEqual(result["error"], "Falha ao sincronizar cliente na API.")

    def test_cliente_sync_success_returns_contract_ok(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(cliente_service, "_call_api", return_value={"status": "ok"}):
                result = cliente_service.sync_cliente_upsert_api(
                    "C1",
                    "Cliente",
                    "Endereco",
                    "11999999999",
                    "Vend",
                    is_desktop_api_sync_enabled=lambda: True,
                )
        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertIsNone(result["error"])
