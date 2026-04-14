import os
from unittest.mock import patch

from app.services import motorista_service

from tests._contract_test_helpers import ContractTestCase


class MotoristaServiceContractTests(ContractTestCase):
    def test_motorista_fetch_error_uses_contextual_default_message(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(motorista_service, "_call_api", side_effect=Exception()):
                result = motorista_service.fetch_motoristas_rows(
                    fields=[("codigo", "Codigo")],
                    can_read_from_api=lambda: True,
                )
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertEqual(result["error"], "Falha ao carregar motoristas na API.")

    def test_motorista_fetch_success_returns_contract_ok(self):
        api_rows = [{"id": 1, "codigo": "MOT-01", "nome": "MOTORISTA 1", "status": "ATIVO"}]
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(motorista_service, "_call_api", return_value=api_rows):
                with patch.object(motorista_service, "fetch_motoristas_cache_local_by_codigo", return_value={}):
                    result = motorista_service.fetch_motoristas_rows(
                        fields=[("codigo", "Codigo"), ("nome", "Nome")],
                        can_read_from_api=lambda: True,
                    )
        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "api")
        self.assertIsNone(result["error"])
        self.assertIsInstance(result["data"], list)
        self.assertEqual(len(result["data"]), 1)
