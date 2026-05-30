# backend/middleware/logging.py
"""
Logging middleware for request/response logging
"""
import time
import logging
from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware

logger = logging.getLogger(__name__)


class LoggingMiddleware(BaseHTTPMiddleware):
    """Middleware to log HTTP requests and responses"""

    async def dispatch(self, request: Request, call_next):
        start_time = time.time()

        # Log request
        logger.info(
            f"Request: {request.method} {request.url} "
            f"Client: {request.client.host if request.client else 'unknown'} "
            f"Tenant: {getattr(request.state, 'tenant_id', 'unknown')}"
        )

        response = await call_next(request)

        # Calculate processing time
        process_time = time.time() - start_time

        # Log response
        logger.info(
            f"Response: {response.status_code} "
            f"Time: {process_time:.3f}s "
            f"URL: {request.url}"
        )

        return response