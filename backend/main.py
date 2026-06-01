# backend/main.py
"""
RotaHub SaaS API - Backend Principal
FastAPI application with multi-tenant support
"""
import os
import sys
import uuid
from datetime import date
from pathlib import Path
from contextlib import asynccontextmanager
from urllib.parse import unquote
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.trustedhost import TrustedHostMiddleware
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import logging
from dotenv import load_dotenv
from sqlalchemy import bindparam, text

# Ensure the project root is on sys.path so backend package imports work when running this script directly
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
load_dotenv(PROJECT_ROOT / ".env")
WEB_DIR = PROJECT_ROOT / "backend" / "web"
OWNER_WEB_DIR = PROJECT_ROOT / "backend" / "web_owner"
PUBLIC_WEB_DIR = PROJECT_ROOT / "backend" / "web_public"
ASSETS_DIR = PROJECT_ROOT / "assets"

from backend.config.database import create_tables
from backend.config.database import async_session
from backend.config.settings import settings
from backend.middleware.tenant import TenantMiddleware
from backend.middleware.rate_limit import RateLimitMiddleware
from backend.middleware.logging import LoggingMiddleware
from backend.api.v1.api import api_router
from backend.models.user import UserDB
from backend.services.auth import get_password_hash
from app.db.connection import configure_connection
from app.middleware.feature_middleware import FeatureGateMiddleware
from app.middleware.billing_middleware import BillingProtectionMiddleware
from app.repositories.base_repository import ensure_saas_ready, get_db as get_saas_db
from app.repositories import company_repository, subscription_repository
from app.services import feature_service

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


class NoCacheStaticFiles(StaticFiles):
    def file_response(self, full_path, stat_result, scope, status_code=200):
        response = super().file_response(full_path, stat_result, scope, status_code)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response


def _env_truthy(name: str, default: str = "0") -> bool:
    return str(os.getenv(name, default) or "").strip().lower() in {"1", "true", "yes", "y", "sim", "on"}


def _sqlite_path_from_url(database_url: str) -> str:
    value = (database_url or "").strip()
    if value.startswith("sqlite+aiosqlite:///"):
        return unquote(value.removeprefix("sqlite+aiosqlite:///"))
    if value.startswith("sqlite:///"):
        return unquote(value.removeprefix("sqlite:///"))
    return ""


def _sync_legacy_mobile_db_env() -> None:
    sqlite_path = _sqlite_path_from_url(settings.DATABASE_URL)
    if sqlite_path:
        os.environ["ROTA_DB"] = sqlite_path
        logger.info("Legacy mobile API using ROTA_DB from DATABASE_URL: %s", sqlite_path)


PLAN_FEATURE_ENDPOINTS = {
    "/api/v1/cadastros": "cadastros",
    "/api/v1/users": "cadastros",
    "/api/v1/permissoes": "cadastros",
    "/api/v1/logistica": "cadastros",
    "/api/v1/importar-vendas": "importar_vendas",
    "/api/v1/programacao": "programacao",
    "/api/v1/recebimentos": "recebimentos",
    "/api/v1/despesas/mortalidade": "mortalidade",
    "/api/v1/despesas": "despesas",
    "/api/v1/rotas": "rotas",
    "/api/v1/escala": "escala",
    "/api/v1/centro-custos": "centro_custos",
    "/api/v1/compras": "despesas",
    "/api/v1/relatorios": "relatorios",
    "/api/v1/audit-logs": "relatorios",
    "/api/v1/system-tools": "private_deployment",
}


def _configure_feature_db() -> None:
    sqlite_path = _sqlite_path_from_url(settings.DATABASE_URL)
    if sqlite_path:
        configure_connection(sqlite_path)


def _temporary_admin_password() -> tuple[str, str]:
    configured_password = (
        os.getenv("ROTA_OWNER_ADMIN_PASSWORD")
        or os.getenv("OWNER_ADMIN_PASSWORD")
        or os.getenv("ROTA_ADMIN_PASS")
        or os.getenv("ROTA_ADMIN_PASSWORD")
        or ""
    ).strip()
    if configured_password:
        return configured_password, "ambiente"
    return "123456", "padrao"


async def _ensure_saas_baseline() -> None:
    _configure_feature_db()
    with get_saas_db() as conn:
        ensure_saas_ready(conn)


async def _ensure_owner_admin_user() -> None:
    owner_names = [item.strip().upper() for item in settings.OWNER_ADMIN_USERS if str(item or "").strip()]
    username = owner_names[0] if owner_names else "ADMIN"
    password, password_source = _temporary_admin_password()

    async with async_session() as db:
        result = await db.execute(
            text(
                """
                SELECT id
                  FROM usuarios
                 WHERE UPPER(COALESCE(username, '')) IN :owner_names
                    OR UPPER(COALESCE(nome, '')) IN :owner_names
                    OR UPPER(COALESCE(permissoes, '')) IN ('DONO', 'OWNER', 'SUPERADMIN', 'SUPER_ADMIN')
                 ORDER BY id ASC
                 LIMIT 1
                """
            ).bindparams(bindparam("owner_names", expanding=True)),
            {"owner_names": owner_names or ["ADMIN"]},
        )
        owner_id = result.scalar_one_or_none()
        if owner_id:
            await db.execute(
                text(
                    """
                    UPDATE usuarios
                       SET permissoes='ADMIN',
                           is_active=1,
                           company_id=COALESCE(company_id, 1),
                           username=COALESCE(NULLIF(TRIM(username), ''), :username),
                           nome=COALESCE(NULLIF(TRIM(nome), ''), :username)
                     WHERE id=:user_id
                    """
                ),
                {"username": username, "user_id": int(owner_id)},
            )
            await db.commit()
            logger.info("Owner admin user verified: %s", username)
            return

        user = UserDB(
            username=username,
            nome=username,
            senha=get_password_hash(password),
            permissoes="ADMIN",
            is_active=True,
            company_id=1,
        )
        db.add(user)
        await db.commit()
        logger.warning(
            "Owner admin user created: %s | temporary password source: %s",
            username,
            password_source,
        )


def _can_use_plan_feature(company_id: int, feature_name: str) -> bool:
    _configure_feature_db()
    result = feature_service.can_use_feature(company_id, feature_name)
    data = result.get("data") or {}
    return bool(result.get("ok") and data.get("allowed"))


def _billing_context(company_id: int) -> dict:
    _configure_feature_db()
    company = company_repository.get_company(company_id) or {}
    subscription = subscription_repository.get_active_subscription(company_id) or {}
    subscription_status = subscription.get("status")
    next_due_date = subscription.get("next_due_date")
    if subscription_status == "trialing" and next_due_date:
        try:
            if date.fromisoformat(str(next_due_date)) < date.today():
                subscription_status = "trial_expired"
        except ValueError:
            pass
    return {
        "company_status": company.get("status"),
        "subscription_status": subscription_status,
    }


LEGACY_MOBILE_APP = None
LEGACY_MOBILE_ENSURE_TABLES = None
_legacy_default = "1" if os.getenv("ROTA_SECRET") and not os.getenv("DATABASE_URL") else "0"
if _env_truthy("ROTA_ENABLE_LEGACY_MOBILE_API", _legacy_default):
    try:
        _sync_legacy_mobile_db_env()
        from api_server import app as LEGACY_MOBILE_APP  # type: ignore
        from api_server import ensure_tables as LEGACY_MOBILE_ENSURE_TABLES  # type: ignore

        logger.info("Legacy mobile API compatibility enabled")
    except Exception:
        logger.exception("Legacy mobile API compatibility could not be enabled")
        LEGACY_MOBILE_APP = None
        LEGACY_MOBILE_ENSURE_TABLES = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan manager"""
    # Startup
    logger.info("Starting RotaHub SaaS API...")
    await create_tables()
    await _ensure_saas_baseline()
    await _ensure_owner_admin_user()
    logger.info("Database tables created/verified")
    if LEGACY_MOBILE_ENSURE_TABLES is not None:
        LEGACY_MOBILE_ENSURE_TABLES()
        logger.info("Legacy mobile API tables created/verified")

    yield

    # Shutdown
    logger.info("Shutting down RotaHub SaaS API...")

# Create FastAPI app
app = FastAPI(
    title="RotaHub SaaS API",
    description="API para o sistema de gestão logística RotaHub SaaS",
    version="1.0.0",
    lifespan=lifespan,
    docs_url="/docs" if settings.DEBUG else None,
    redoc_url="/redoc" if settings.DEBUG else None,
)

# Security middlewares (order matters)
app.add_middleware(TrustedHostMiddleware, allowed_hosts=settings.ALLOWED_HOSTS)
if _env_truthy("ROTA_ENFORCE_PLAN_FEATURES", "1"):
    app.add_middleware(
        FeatureGateMiddleware,
        endpoint_features=PLAN_FEATURE_ENDPOINTS,
        can_use_feature=_can_use_plan_feature,
    )
if _env_truthy("ROTA_ENFORCE_BILLING", "1"):
    app.add_middleware(
        BillingProtectionMiddleware,
        get_billing_context=_billing_context,
    )
app.add_middleware(TenantMiddleware)
app.add_middleware(RateLimitMiddleware)
app.add_middleware(LoggingMiddleware)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    return response

# Global exception handler
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled exception: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": "Internal server error"},
    )

# Health check endpoint
@app.get("/health")
async def health_check():
    return {"status": "healthy", "service": "RotaHub SaaS API"}


@app.get("/ready")
async def readiness_check():
    checks: dict[str, object] = {"database": False, "photos_dir": False}
    errors: dict[str, str] = {}

    try:
        async with async_session() as db:
            await db.execute(text("SELECT 1"))
        checks["database"] = True
    except Exception as exc:
        errors["database"] = str(exc)

    photos_dir = Path(os.getenv("ROTA_MOBILE_PHOTOS_DIR", PROJECT_ROOT / ".rotahub_runtime" / "fotos_rotas")).expanduser()
    try:
        photos_dir.mkdir(parents=True, exist_ok=True)
        test_path = photos_dir / f".ready_{uuid.uuid4().hex}"
        test_path.write_text("ok", encoding="utf-8")
        test_path.unlink(missing_ok=True)
        checks["photos_dir"] = True
    except Exception as exc:
        errors["photos_dir"] = str(exc)

    ready = all(bool(value) for value in checks.values())
    payload = {"status": "ready" if ready else "not_ready", "checks": checks}
    if errors:
        payload["errors"] = errors
    return JSONResponse(status_code=200 if ready else 503, content=payload)


@app.get("/", include_in_schema=False)
async def web_app():
    return RedirectResponse(url="/public/index.html")


# Include API routers
app.include_router(api_router, prefix="/api/v1")

if ASSETS_DIR.exists():
    app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")

if PUBLIC_WEB_DIR.exists():
    app.mount("/public", StaticFiles(directory=PUBLIC_WEB_DIR), name="public-web")

app.mount("/app", NoCacheStaticFiles(directory=WEB_DIR), name="web")

if OWNER_WEB_DIR.exists():
    app.mount("/owner", StaticFiles(directory=OWNER_WEB_DIR), name="owner-web")

if LEGACY_MOBILE_APP is not None:
    app.mount("/", LEGACY_MOBILE_APP, name="legacy-mobile-api")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info",
    )
