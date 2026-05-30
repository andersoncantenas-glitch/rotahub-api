from __future__ import annotations

from typing import Any, Callable

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request


class TenantContextMiddleware(BaseHTTPMiddleware):
    def __init__(self, app, verify_token: Callable[[str], dict | None]):
        super().__init__(app)
        self._verify_token = verify_token

    async def dispatch(self, request: Request, call_next):
        request.state.company_id = None
        request.state.tenant = None

        auth_header = str(request.headers.get("authorization") or "").strip()
        if auth_header.lower().startswith("bearer "):
            token = auth_header.split(" ", 1)[1].strip()
            try:
                payload = self._verify_token(token)
            except Exception:
                payload = None
            if isinstance(payload, dict):
                company_id = _safe_int(payload.get("company_id"))
                request.state.company_id = company_id
                request.state.tenant = {
                    "company_id": company_id,
                    "user_id": _safe_int(payload.get("user_id")),
                    "username": str(payload.get("username") or payload.get("codigo") or "").strip(),
                    "role": str(payload.get("role") or payload.get("perfil") or "").strip(),
                }

        if request.state.company_id is None:
            company_id = _safe_int(request.headers.get("x-company-id"))
            if company_id:
                request.state.company_id = company_id
                request.state.tenant = {
                    "company_id": company_id,
                    "user_id": None,
                    "username": "",
                    "role": "desktop",
                }

        return await call_next(request)


def _safe_int(value: Any) -> int | None:
    try:
        if value is None or value == "":
            return None
        return int(value)
    except Exception:
        return None
