import asyncio
import json
import unittest

from starlette.requests import Request
from starlette.responses import JSONResponse

from app.middleware.billing_middleware import BillingProtectionMiddleware
from app.middleware.tenant_middleware import TenantContextMiddleware


async def _noop_app(_scope, _receive, _send):
    return None


def _request(method: str, path: str) -> Request:
    return Request(
        {
            "type": "http",
            "asgi": {"version": "3.0"},
            "http_version": "1.1",
            "method": method,
            "scheme": "http",
            "path": path,
            "raw_path": path.encode("ascii"),
            "query_string": b"",
            "headers": [(b"authorization", b"Bearer token")],
            "client": ("127.0.0.1", 12345),
            "server": ("testserver", 80),
        }
    )


def _json(response):
    return json.loads(response.body.decode("utf-8"))


class BillingMiddlewareTests(unittest.TestCase):
    def _dispatch(self, method: str, path: str, *, company_status="active", subscription_status="active", audit_events=None):
        def verify_token(_token):
            return {"company_id": 10, "user_id": 99, "username": "tester", "role": "motorista"}

        def billing_context(_company_id):
            return {
                "company_status": company_status,
                "subscription_status": subscription_status,
            }

        def audit_block(company_id, request, payload):
            if audit_events is not None:
                audit_events.append((company_id, request.url.path, payload["billing_status"]))

        tenant = TenantContextMiddleware(_noop_app, verify_token=verify_token)
        billing = BillingProtectionMiddleware(_noop_app, get_billing_context=billing_context, audit_block=audit_block)
        request = _request(method, path)

        async def terminal(_request):
            return JSONResponse({"ok": True})

        async def call_billing(req):
            return await billing.dispatch(req, terminal)

        return asyncio.run(tenant.dispatch(request, call_billing))

    def test_get_requests_are_allowed_when_suspended(self):
        response = self._dispatch("GET", "/programacoes", subscription_status="suspended")
        self.assertEqual(response.status_code, 200)

    def test_auth_paths_are_allowed_when_suspended(self):
        response = self._dispatch("POST", "/auth/motorista/login", subscription_status="suspended")
        self.assertEqual(response.status_code, 200)

    def test_mutation_is_blocked_and_audited_when_subscription_suspended(self):
        audit_events = []
        response = self._dispatch("POST", "/programacoes", subscription_status="suspended", audit_events=audit_events)

        self.assertEqual(response.status_code, 402)
        self.assertEqual(_json(response)["billing_status"], "suspended")
        self.assertEqual(audit_events, [(10, "/programacoes", "suspended")])


if __name__ == "__main__":
    unittest.main()
