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

    def test_salvar_clientes_linhas_uses_bulk_api_and_cleans_nan_values(self):
        with patch.dict(os.environ, {"ROTA_SECRET": "secret"}, clear=False):
            with patch.object(cliente_service, "_call_api", return_value={"total": 2}) as call_api:
                result = cliente_service.salvar_clientes_linhas(
                    [
                        (1.0, "Cliente Um", "Rua Um", float("nan"), "nan"),
                        (2.0, "Cliente Dois", None, "8888", "Vend"),
                    ],
                    is_desktop_api_sync_enabled=lambda: True,
                )

        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "api")
        call_api.assert_called_once()
        args, kwargs = call_api.call_args
        self.assertEqual(args[:2], ("POST", "desktop/cadastros/clientes/bulk-upsert"))
        self.assertEqual(
            kwargs["payload"],
            {
                "clientes": [
                    {
                        "cod_cliente": "1",
                        "nome_cliente": "CLIENTE UM",
                        "endereco": "RUA UM",
                        "telefone": "",
                        "vendedor": "",
                    },
                    {
                        "cod_cliente": "2",
                        "nome_cliente": "CLIENTE DOIS",
                        "endereco": "",
                        "telefone": "8888",
                        "vendedor": "VEND",
                    },
                ]
            },
        )
