# backend/middleware/tenant.py
"""
Tenant middleware for multi-tenant SaaS support
"""
import logging
from typing import Optional
from fastapi import Request, HTTPException
from jose import JWTError, jwt
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse

from backend.config.settings import settings

logger = logging.getLogger(__name__)


class TenantContext:
    """Context for current tenant"""
    def __init__(self):
        self.tenant_id: Optional[str] = None
        self.company_id: Optional[int] = None
        self.plan_code: Optional[str] = None

    def set_tenant(self, tenant_id: str, company_id: int, plan_code: str):
        self.tenant_id = tenant_id
        self.company_id = company_id
        self.plan_code = plan_code

    def clear(self):
        self.tenant_id = None
        self.company_id = None
        self.plan_code = None


# Global tenant context
tenant_context = TenantContext()


class TenantMiddleware(BaseHTTPMiddleware):
    """Middleware to handle tenant context from request headers"""

    async def dispatch(self, request: Request, call_next):
        # Extract tenant from header (e.g., X-Tenant-ID)
        tenant_header = request.headers.get("X-Tenant-ID")
        company_id = 1
        auth_header = request.headers.get("Authorization", "")
        if auth_header.lower().startswith("bearer "):
            token = auth_header.split(" ", 1)[1].strip()
            try:
                payload = jwt.decode(token, settings.JWT_SECRET_KEY, algorithms=[settings.JWT_ALGORITHM])
                company_id = int(payload.get("company_id") or 1)
            except (JWTError, ValueError, TypeError):
                company_id = 1

        if tenant_header:
            # In production, validate tenant exists in database
            # For now, accept any tenant header
            tenant_context.set_tenant(
                tenant_id=tenant_header,
                company_id=company_id,
                plan_code="PROFESSIONAL"  # Placeholder
            )
        else:
            # Default tenant for development
            tenant_context.set_tenant(
                tenant_id=f"company-{company_id}",
                company_id=company_id,
                plan_code="PROFESSIONAL"
            )

        # Add tenant info to request state
        request.state.tenant_id = tenant_context.tenant_id
        request.state.company_id = tenant_context.company_id
        request.state.plan_code = tenant_context.plan_code

        response = await call_next(request)

        # Clear context after request
        tenant_context.clear()

        return response
