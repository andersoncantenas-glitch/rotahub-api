from unittest.mock import patch

from app.services import programacao_service

from tests._contract_test_helpers import ContractTestCase


class ProgramacaoServiceContractTests(ContractTestCase):
    def test_programacao_fetch_success_local_fallback_returns_contract_ok(self):
        with patch.object(programacao_service, "fetch_programacao_itens_local", return_value=[{"cod_cliente": "C1"}]):
            result = programacao_service.fetch_programacao_itens(
                codigo_programacao="prog-001",
                limit=100,
                upper=lambda s: str(s or "").strip().upper(),
                safe_int=lambda v, d=0: int(v) if str(v or "").isdigit() else d,
                safe_float=lambda v, d=0.0: float(v) if str(v or "").strip() else d,
                get_db=lambda: None,
                call_api=lambda *args, **kwargs: {},
                quote=lambda s: str(s),
                is_desktop_api_sync_enabled=lambda: False,
                desktop_secret="",
            )
        self.assert_contract(result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["source"], "local")
        self.assertIsNone(result["error"])
        self.assertEqual(result["data"], [{"cod_cliente": "C1"}])

    def test_programacao_fetch_local_failure_returns_contract_error(self):
        with patch.object(programacao_service, "fetch_programacao_itens_local", side_effect=Exception("db fail")):
            result = programacao_service.fetch_programacao_itens(
                codigo_programacao="prog-001",
                limit=100,
                upper=lambda s: str(s or "").strip().upper(),
                safe_int=lambda v, d=0: int(v) if str(v or "").isdigit() else d,
                safe_float=lambda v, d=0.0: float(v) if str(v or "").strip() else d,
                get_db=lambda: None,
                call_api=lambda *args, **kwargs: {},
                quote=lambda s: str(s),
                is_desktop_api_sync_enabled=lambda: False,
                desktop_secret="",
            )
        self.assert_contract(result)
        self.assertFalse(result["ok"])
        self.assertEqual(result["source"], "local")
        self.assertEqual(result["error"], "Falha ao obter itens da programação.")
        self.assertEqual(result["data"], [])
