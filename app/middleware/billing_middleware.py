from __future__ import annotations

from typing import Any, Callable

from jose import JWTError, jwt
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse

from backend.config.settings import settings


SAFE_METHODS = {"GET", "HEAD", "OPTIONS"}
BLOCKED_STATUSES = {"suspended", "cancelled", "inactive", "blocked", "trial_expired"}
ALLOWED_PATH_PREFIXES = (
    "/api/v1/auth",
    "/api/v1/billing",
    "/api/v1/public",
    "/api/v1/saas-admin",
    "/docs",
    "/openapi.json",
    "/redoc",
    "/health",
    "/ready",
)


class BillingProtectionMiddleware(BaseHTTPMiddleware):
    def __init__(
        self,
        app,
        get_billing_context: Callable[[int], dict | None],
        audit_block: Callable[[int, Request, dict], None] | None = None,
    ):
        super().__init__(app)
        self._get_billing_context = get_billing_context
        self._audit_block = audit_block

    async def dispatch(self, request: Request, call_next):
        if _is_allowed_without_billing_check(request):
            return await call_next(request)

        company_id = _safe_int(getattr(request.state, "company_id", None)) or _company_id_from_token(request)
        if not company_id:
            return await call_next(request)

        try:
            context = self._get_billing_context(company_id) or {}
        except Exception:
            context = {}

        company_status = _norm_status(context.get("company_status") or context.get("status"))
        subscription_status = _norm_status(context.get("subscription_status"))
        blocking_status = company_status if company_status in BLOCKED_STATUSES else ""
        if not blocking_status and subscription_status in BLOCKED_STATUSES:
            blocking_status = subscription_status

        if not blocking_status:
            return await call_next(request)

        payload = {
            "detail": "Operacao bloqueada por status de cobranca.",
            "company_id": company_id,
            "billing_status": blocking_status,
        }
        if callable(self._audit_block):
            try:
                self._audit_block(company_id, request, payload)
            except Exception:
                pass
        return JSONResponse(status_code=402, content=payload)


def _is_allowed_without_billing_check(request: Request) -> bool:
    if request.method.upper() in SAFE_METHODS:
        return True
    path = str(request.url.path or "")
    return any(path.startswith(prefix) for prefix in ALLOWED_PATH_PREFIXES)


def _norm_status(value: Any) -> str:
    return str(value or "").strip().lower()


def _safe_int(value: Any) -> int | None:
    try:
        if value is None or value == "":
            return None
        return int(value)
    except Exception:
        return None


def _company_id_from_token(request: Request) -> int | None:
    auth_header = str(request.headers.get("Authorization") or "")
    if not auth_header.lower().startswith("bearer "):
        return None
    token = auth_header.split(" ", 1)[1].strip()
    if not token:
        return None
    try:
        payload = jwt.decode(token, settings.JWT_SECRET_KEY, algorithms=[settings.JWT_ALGORITHM])
        return _safe_int(payload.get("company_id"))
    except (JWTError, ValueError, TypeError):
        return None
