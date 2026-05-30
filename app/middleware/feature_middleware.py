from __future__ import annotations

from typing import Callable

from jose import JWTError, jwt
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse

from backend.config.settings import settings


class FeatureGateMiddleware(BaseHTTPMiddleware):
    def __init__(self, app, endpoint_features: dict[str, str], can_use_feature: Callable[[int, str], bool]):
        super().__init__(app)
        self._endpoint_features = dict(endpoint_features or {})
        self._can_use_feature = can_use_feature

    async def dispatch(self, request: Request, call_next):
        feature = _match_feature(str(request.url.path or ""), self._endpoint_features)
        if not feature:
            return await call_next(request)

        company_id = _safe_int(getattr(request.state, "company_id", None)) or _company_id_from_token(request)
        if not company_id:
            return await call_next(request)

        try:
            allowed = bool(self._can_use_feature(company_id, feature))
        except Exception:
            allowed = False
        if allowed:
            return await call_next(request)

        return JSONResponse(
            status_code=403,
            content={
                "detail": "Recurso nao disponivel no plano atual.",
                "feature": feature,
                "company_id": company_id,
            },
        )


def _match_feature(path: str, endpoint_features: dict[str, str]) -> str:
    for prefix, feature in endpoint_features.items():
        if path.startswith(prefix):
            return str(feature or "").strip()
    return ""


def _safe_int(value):
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
