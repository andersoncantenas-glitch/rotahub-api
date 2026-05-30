import unittest

from fastapi import FastAPI
from fastapi.testclient import TestClient

from app.middleware.feature_middleware import FeatureGateMiddleware
from app.middleware.tenant_middleware import TenantContextMiddleware


class FeatureFlagsTests(unittest.TestCase):
    def _client(self, *, allowed: bool):
        app = FastAPI()

        def verify_token(_token):
            return {"company_id": 10, "user_id": 1, "username": "tester", "role": "admin"}

        def can_use_feature(_company_id, _feature):
            return allowed

        app.add_middleware(
            FeatureGateMiddleware,
            endpoint_features={"/premium": "advanced_reports"},
            can_use_feature=can_use_feature,
        )
        app.add_middleware(TenantContextMiddleware, verify_token=verify_token)

        @app.get("/premium/report")
        def premium_report():
            return {"ok": True}

        return TestClient(app)

    def test_feature_gate_blocks_disabled_feature(self):
        response = self._client(allowed=False).get("/premium/report", headers={"Authorization": "Bearer token"})
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.json()["feature"], "advanced_reports")

    def test_feature_gate_uses_company_header_without_bearer_token(self):
        response = self._client(allowed=False).get("/premium/report", headers={"X-Company-ID": "10"})
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.json()["company_id"], 10)

    def test_feature_gate_allows_enabled_feature(self):
        response = self._client(allowed=True).get("/premium/report", headers={"Authorization": "Bearer token"})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"ok": True})


if __name__ == "__main__":
    unittest.main()
