import os
from unittest.mock import patch

from app.services import vendedor_service

from tests._contract_test_helpers import ContractTestCase


class VendedorServiceContractTests(ContractTestCase):
    def test_vendedor_sync_error_uses_contextual_default_message(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(vendedor_service, "_call_api", side_effect=Exception()):
                result = vendedor_service.sync_vendedor_upsert_api(
                    {"codigo": "V1", "nome": "Vendedor"},
                    is_desktop_api_sync_enabled=lambda: True,
                    norm=lambda v: str(v or "").strip().upper(),
                    normalize_phone=lambda v: str(v or "").strip(),
                )
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertEqual(result["error"], "Falha ao sincronizar vendedor na API.")

    def test_vendedor_sync_success_returns_contract_ok(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(vendedor_service, "_call_api", return_value={"status": "ok"}):
                result = vendedor_service.sync_vendedor_upsert_api(
                    {"codigo": "V1", "nome": "Vendedor"},
                    is_desktop_api_sync_enabled=lambda: True,
                    norm=lambda v: str(v or "").strip().upper(),
                    normalize_phone=lambda v: str(v or "").strip(),
                )
        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertIsNone(result["error"])
