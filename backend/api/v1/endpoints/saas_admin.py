# backend/api/v1/endpoints/saas_admin.py
"""
Admin SaaS endpoints mirroring the desktop SaaSAdminPage.
"""
from __future__ import annotations

import json
import re
import secrets
import unicodedata
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy.engine import make_url

from app.db.connection import configure_connection
from app.repositories import audit_repository, company_repository, payment_repository, plan_repository, subscription_repository
from app.repositories.base_repository import ensure_saas_ready, get_db as get_saas_db, row_to_dict, rows_to_dicts
from app.services import feature_service, saas_admin_service
from backend.api.v1.endpoints.users import require_owner_user
from backend.config.settings import settings
from backend.models.user import User
from backend.services.auth import get_password_hash
from backend.services.plan_change_requests import (
    close_plan_change_request,
    ensure_plan_change_requests_table,
    get_plan_change_request,
    list_plan_change_requests,
)

router = APIRouter()

PROJECT_ROOT = Path(__file__).resolve().parents[4]
VALID_COMPANY_STATUSES = {"active", "suspended", "cancelled", "inactive"}
TRIAL_PLAN_PRIORITY = ("enterprise", "professional", "growth", "starter")


class CompanyStatusPayload(BaseModel):
    status: str = Field(min_length=1, max_length=40)
    reason: str | None = Field(default=None, max_length=240)

    @field_validator("status", mode="before")
    @classmethod
    def strip_status(cls, value):
        return str(value or "").strip().lower()


class CompanyPlanPayload(BaseModel):
    plan_code: str = Field(min_length=1, max_length=80)
    reason: str | None = Field(default=None, max_length=240)

    @field_validator("plan_code", mode="before")
    @classmethod
    def strip_plan_code(cls, value):
        return str(value or "").strip().lower()


class PaymentCreatePayload(BaseModel):
    company_id: int = Field(gt=0)
    amount: float = Field(default=0, ge=0)
    due_date: str | None = Field(default=None, max_length=20)
    notes: str | None = Field(default=None, max_length=500)

    @field_validator("due_date", "notes", mode="before")
    @classmethod
    def strip_optional_text(cls, value):
        if value is None:
            return None
        text = str(value).strip()
        return text or None


class PaymentRegisterPayload(BaseModel):
    method: str | None = Field(default="manual", max_length=80)
    reference: str | None = Field(default=None, max_length=160)
    notes: str | None = Field(default=None, max_length=500)

    @field_validator("method", "reference", "notes", mode="before")
    @classmethod
    def strip_register_text(cls, value):
        if value is None:
            return None
        text = str(value).strip()
        return text or None


class BillingAutomationPayload(BaseModel):
    grace_days: int = Field(default=0, ge=0, le=365)


class LeadApprovalPayload(BaseModel):
    plan_code: str = Field(min_length=1, max_length=80)
    trial_days: int = Field(default=30, ge=1, le=90)

    @field_validator("plan_code", mode="before")
    @classmethod
    def strip_plan_code(cls, value):
        return str(value or "").strip().lower()


class LeadRejectPayload(BaseModel):
    reason: str | None = Field(default=None, max_length=240)

    @field_validator("reason", mode="before")
    @classmethod
    def strip_reason(cls, value):
        return str(value or "").strip() or None


class PlanRequestReviewPayload(BaseModel):
    notes: str | None = Field(default=None, max_length=500)

    @field_validator("notes", mode="before")
    @classmethod
    def strip_notes(cls, value):
        return str(value or "").strip() or None


def sqlite_db_path() -> Path:
    url = make_url(settings.DATABASE_URL)
    if not url.drivername.startswith("sqlite"):
        raise HTTPException(status_code=409, detail="Admin SaaS local disponivel somente para banco SQLite.")
    if not url.database or url.database == ":memory:":
        raise HTTPException(status_code=409, detail="Admin SaaS indisponivel para banco SQLite em memoria.")
    path = Path(url.database)
    if not path.is_absolute():
        path = (PROJECT_ROOT / path).resolve()
    return path


def configure_saas_sqlite() -> Path:
    db_path = sqlite_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    configure_connection(str(db_path))
    return db_path


def ensure_signup_leads_table(conn) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS public_signup_leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            document TEXT NOT NULL,
            email TEXT NOT NULL,
            phone TEXT NOT NULL,
            company TEXT NOT NULL,
            plan_code TEXT NOT NULL,
            message TEXT,
            status TEXT DEFAULT 'novo',
            user_agent TEXT,
            ip_address TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute("PRAGMA table_info(public_signup_leads)")
    columns = {str(row[1]).lower() for row in cur.fetchall()}
    for column, definition in {
        "company_id": "INTEGER",
        "reviewed_at": "TEXT",
        "reviewed_by": "TEXT",
        "review_notes": "TEXT",
        "trial_days": "INTEGER",
    }.items():
        if column not in columns:
            cur.execute(f"ALTER TABLE public_signup_leads ADD COLUMN {column} {definition}")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_public_signup_leads_status ON public_signup_leads(status)")


def list_signup_leads(limit: int = 500) -> list[dict]:
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        ensure_signup_leads_table(conn)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT *
              FROM public_signup_leads
             ORDER BY CASE status WHEN 'novo' THEN 0 ELSE 1 END, id DESC
             LIMIT ?
            """,
            (max(1, min(int(limit or 500), 5000)),),
        )
        return rows_to_dicts(cur.fetchall())


def list_companies_with_subscriptions() -> list[dict]:
    companies = company_repository.list_companies(limit=500)
    for company in companies:
        subscription = subscription_repository.get_active_subscription(int(company.get("id") or 0)) or {}
        company["plan_code"] = subscription.get("plan_code")
        company["subscription_status"] = subscription.get("status")
        company["next_due_date"] = subscription.get("next_due_date")
    return companies


def company_code(value: str, lead_id: int) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or "empresa")).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-z0-9]+", "-", normalized.lower()).strip("-")[:42] or "empresa"
    return f"{slug}-{int(lead_id)}"


def best_trial_plan(cur) -> dict:
    placeholders = ",".join("?" for _ in TRIAL_PLAN_PRIORITY)
    cur.execute(
        f"""
        SELECT *
          FROM plans
         WHERE status='active'
           AND code IN ({placeholders})
        """,
        TRIAL_PLAN_PRIORITY,
    )
    plans = {str(row["code"] or "").strip().lower(): row_to_dict(row) for row in cur.fetchall()}
    for code in TRIAL_PLAN_PRIORITY:
        if code in plans:
            return plans[code]
    raise HTTPException(status_code=404, detail="Plano de demonstracao nao encontrado.")


def raise_service_error(result: dict, *, default_status: int = status.HTTP_400_BAD_REQUEST) -> None:
    error = str((result or {}).get("error") or "Falha na operacao.")
    status_code = default_status
    lowered = error.lower()
    if "nao encontrada" in lowered or "nao encontrado" in lowered:
        status_code = status.HTTP_404_NOT_FOUND
    if "downgrade bloqueado" in lowered:
        status_code = status.HTTP_409_CONFLICT
    raise HTTPException(status_code=status_code, detail=error)


def decode_json_field(raw: Any) -> dict:
    if isinstance(raw, dict):
        return dict(raw)
    try:
        parsed = json.loads(str(raw or "{}"))
    except Exception:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def normalize_plan(plan: dict | None) -> dict:
    item = dict(plan or {})
    item["features"] = decode_json_field(item.get("features_json"))
    return item


def normalize_audit(log: dict | None) -> dict:
    item = dict(log or {})
    item["metadata"] = decode_json_field(item.get("metadata_json"))
    return item


def dashboard_payload(company_id: int | None = None) -> dict[str, Any]:
    result = saas_admin_service.get_dashboard(company_id)
    if not result.get("ok"):
        raise_service_error(result)
    data = dict(result.get("data") or {})
    company = data.get("company") or {}
    cid = int(company.get("id") or company_id or 0)
    features_result = feature_service.list_company_features(cid) if cid else {"ok": False, "data": None}
    data["companies"] = list_companies_with_subscriptions()
    data["signup_leads"] = list_signup_leads()
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        data["plan_change_requests"] = list_plan_change_requests(conn)
    data["plans"] = [normalize_plan(plan) for plan in (data.get("plans") or [])]
    data["audit_logs"] = [normalize_audit(log) for log in (data.get("audit_logs") or [])]
    data["features"] = features_result.get("data") if features_result.get("ok") else None
    return data


def actor_name(current_user: User) -> str:
    return str(current_user.nome or current_user.username or "ADMIN").strip() or "ADMIN"


@router.get("/dashboard")
async def get_saas_dashboard(
    company_id: int | None = Query(default=None, ge=1),
    current_user: User = Depends(require_owner_user),
):
    """Load the same dashboard data consumed by the desktop Admin SaaS page."""
    configure_saas_sqlite()
    return dashboard_payload(company_id)


@router.get("/companies")
async def list_companies(
    status_filter: str = Query(default="", alias="status"),
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return company_repository.list_companies(status=status_filter.strip() or None, limit=limit)


@router.get("/signup-leads")
async def get_signup_leads(
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return list_signup_leads(limit)


@router.get("/plan-change-requests")
async def get_plan_change_requests(
    status_filter: str = Query(default="", alias="status"),
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        return list_plan_change_requests(conn, status=status_filter.strip().lower() or None, limit=limit)


@router.post("/plan-change-requests/{request_id}/approve")
async def approve_plan_change_request(
    request_id: int,
    payload: PlanRequestReviewPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    actor = actor_name(current_user)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        ensure_plan_change_requests_table(conn)
        request = get_plan_change_request(conn, request_id)
    if not request:
        raise HTTPException(status_code=404, detail="Solicitacao de plano nao encontrada.")
    if str(request.get("status") or "").strip().lower() != "pending":
        raise HTTPException(status_code=409, detail="Esta solicitacao ja foi analisada.")
    result = saas_admin_service.change_company_plan(
        int(request.get("company_id") or 0),
        str(request.get("requested_plan_code") or ""),
        actor=actor,
        reason=payload.notes or f"Solicitacao #{request_id} aprovada pelo Owner.",
    )
    if not result.get("ok"):
        raise_service_error(result)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        reviewed = close_plan_change_request(
            conn,
            request_id,
            status="approved",
            reviewed_by=actor,
            review_notes=payload.notes or "",
        )
    audit_repository.create_audit_log(
        {
            "company_id": int(request.get("company_id") or 0),
            "actor_type": "admin",
            "action": "solicitacao_plano_aprovada",
            "entity_type": "plan_change_request",
            "entity_id": str(request_id),
            "severity": "info",
            "metadata": {
                "actor": actor,
                "current_plan_code": request.get("current_plan_code"),
                "requested_plan_code": request.get("requested_plan_code"),
                "notes": payload.notes or "",
            },
        }
    )
    return {"ok": True, "request": reviewed, "subscription": result.get("data")}


@router.post("/plan-change-requests/{request_id}/reject")
async def reject_plan_change_request(
    request_id: int,
    payload: PlanRequestReviewPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    actor = actor_name(current_user)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        request = get_plan_change_request(conn, request_id)
        if not request:
            raise HTTPException(status_code=404, detail="Solicitacao de plano nao encontrada.")
        if str(request.get("status") or "").strip().lower() != "pending":
            raise HTTPException(status_code=409, detail="Esta solicitacao ja foi analisada.")
        reviewed = close_plan_change_request(
            conn,
            request_id,
            status="rejected",
            reviewed_by=actor,
            review_notes=payload.notes or "",
        )
    audit_repository.create_audit_log(
        {
            "company_id": int(request.get("company_id") or 0),
            "actor_type": "admin",
            "action": "solicitacao_plano_recusada",
            "entity_type": "plan_change_request",
            "entity_id": str(request_id),
            "severity": "warning",
            "metadata": {
                "actor": actor,
                "current_plan_code": request.get("current_plan_code"),
                "requested_plan_code": request.get("requested_plan_code"),
                "notes": payload.notes or "",
            },
        }
    )
    return {"ok": True, "request": reviewed}


@router.post("/signup-leads/{lead_id}/approve-trial")
async def approve_signup_lead_trial(
    lead_id: int,
    payload: LeadApprovalPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    actor = actor_name(current_user)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        ensure_signup_leads_table(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM public_signup_leads WHERE id=? LIMIT 1", (int(lead_id),))
        lead = row_to_dict(cur.fetchone())
        if not lead:
            raise HTTPException(status_code=404, detail="Solicitacao nao encontrada.")
        if str(lead.get("status") or "").strip().lower() != "novo":
            raise HTTPException(status_code=409, detail="Esta solicitacao ja foi analisada.")
        plan = best_trial_plan(cur)
        cur.execute("SELECT id FROM companies WHERE document=? LIMIT 1", (lead.get("document"),))
        if cur.fetchone():
            raise HTTPException(status_code=409, detail="Ja existe uma empresa cadastrada com este CPF/CNPJ.")
        code = company_code(str(lead.get("company") or ""), lead_id)
        cur.execute(
            """
            INSERT INTO companies (
                code, name, legal_name, document, email, phone, status, timezone, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'active', 'America/Fortaleza', datetime('now'), datetime('now'))
            """,
            (
                code,
                lead.get("company"),
                lead.get("company"),
                lead.get("document"),
                lead.get("email"),
                lead.get("phone"),
            ),
        )
        company_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO subscriptions (
                company_id, plan_id, status, billing_cycle,
                current_period_start, current_period_end, next_due_date,
                created_at, updated_at
            ) VALUES (
                ?, ?, 'trialing', 'monthly',
                date('now'), date('now', ?), date('now', ?),
                datetime('now'), datetime('now')
            )
            """,
            (company_id, int(plan["id"]), f"+{payload.trial_days} day", f"+{payload.trial_days} day"),
        )
        subscription_id = int(cur.lastrowid)
        temporary_username = f"CLIENTE{int(lead_id):04d}"
        temporary_password = secrets.token_urlsafe(8)
        cur.execute(
            """
            INSERT INTO usuarios (
                username, nome, senha, permissoes, cpf, telefone, is_active, company_id
            ) VALUES (?, ?, ?, 'ADMIN', ?, ?, 1, ?)
            """,
            (
                temporary_username,
                lead.get("name"),
                get_password_hash(temporary_password),
                lead.get("document"),
                lead.get("phone"),
                company_id,
            ),
        )
        cur.execute(
            """
            UPDATE public_signup_leads
               SET status='aprovado', company_id=?, trial_days=?,
                   reviewed_at=datetime('now'), reviewed_by=?
             WHERE id=?
            """,
            (company_id, payload.trial_days, actor, int(lead_id)),
        )
    audit_repository.create_audit_log(
        {
            "company_id": company_id,
            "actor_type": "admin",
            "action": "demonstracao_aprovada",
            "entity_type": "subscription",
            "entity_id": str(subscription_id),
            "severity": "info",
            "metadata": {
                "actor": actor,
                "lead_id": int(lead_id),
                "plan_code": plan.get("code"),
                "requested_plan_code": payload.plan_code,
                "trial_days": payload.trial_days,
            },
        }
    )
    trial_end = ""
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT next_due_date FROM subscriptions WHERE id=? LIMIT 1", (subscription_id,))
        row = cur.fetchone()
        trial_end = str(row["next_due_date"] or "") if row else ""
    return {
        "ok": True,
        "company_id": company_id,
        "subscription_id": subscription_id,
        "trial_days": payload.trial_days,
        "trial_end": trial_end,
        "temporary_username": temporary_username,
        "temporary_password": temporary_password,
    }


@router.post("/signup-leads/{lead_id}/reject")
async def reject_signup_lead(
    lead_id: int,
    payload: LeadRejectPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    actor = actor_name(current_user)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        ensure_signup_leads_table(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE public_signup_leads
               SET status='recusado', reviewed_at=datetime('now'),
                   reviewed_by=?, review_notes=?
             WHERE id=? AND status='novo'
            """,
            (actor, payload.reason or "", int(lead_id)),
        )
        if not cur.rowcount:
            raise HTTPException(status_code=409, detail="Solicitacao inexistente ou ja analisada.")
    return {"ok": True}


@router.get("/companies/{company_id}")
async def get_company(
    company_id: int,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    company = company_repository.get_company(company_id)
    if not company:
        raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
    return company


@router.put("/companies/{company_id}/status")
async def update_company_status(
    company_id: int,
    payload: CompanyStatusPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    if payload.status not in VALID_COMPANY_STATUSES:
        raise HTTPException(status_code=422, detail="Status invalido.")
    result = saas_admin_service.set_company_status(
        company_id,
        payload.status,
        actor=actor_name(current_user),
        reason=payload.reason or "",
    )
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "company": result.get("data")}


@router.post("/companies/{company_id}/admin-access")
async def reset_company_admin_access(
    company_id: int,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    actor = actor_name(current_user)
    company = company_repository.get_company(company_id)
    if not company:
        raise HTTPException(status_code=404, detail="Empresa nao encontrada.")
    username = f"CLIENTE{int(company_id):04d}"
    temporary_password = secrets.token_urlsafe(8)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT id FROM usuarios WHERE company_id=? AND username=? LIMIT 1", (int(company_id), username))
        existing = cur.fetchone()
        if existing:
            cur.execute(
                """
                UPDATE usuarios
                   SET senha=?, permissoes='ADMIN', is_active=1
                 WHERE id=?
                """,
                (get_password_hash(temporary_password), int(existing["id"])),
            )
        else:
            cur.execute(
                """
                INSERT INTO usuarios (
                    username, nome, senha, permissoes, cpf, telefone, is_active, company_id
                ) VALUES (?, ?, ?, 'ADMIN', ?, ?, 1, ?)
                """,
                (
                    username,
                    company.get("name") or username,
                    get_password_hash(temporary_password),
                    company.get("document"),
                    company.get("phone"),
                    int(company_id),
                ),
            )
    audit_repository.create_audit_log(
        {
            "company_id": int(company_id),
            "actor_type": "admin",
            "action": "acesso_admin_regerado",
            "entity_type": "user",
            "entity_id": username,
            "severity": "warning",
            "metadata": {"actor": actor, "username": username},
        }
    )
    return {"ok": True, "username": username, "temporary_password": temporary_password}


@router.put("/companies/{company_id}/plan")
async def change_company_plan(
    company_id: int,
    payload: CompanyPlanPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.change_company_plan(
        company_id,
        payload.plan_code,
        actor=actor_name(current_user),
        reason=payload.reason or "",
    )
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "subscription": result.get("data")}


@router.get("/companies/{company_id}/usage")
async def company_usage(
    company_id: int,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.get_dashboard(company_id)
    if not result.get("ok"):
        raise_service_error(result)
    return (result.get("data") or {}).get("usage") or {}


@router.get("/companies/{company_id}/features")
async def company_features(
    company_id: int,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = feature_service.list_company_features(company_id)
    if not result.get("ok"):
        raise_service_error(result)
    return result.get("data")


@router.get("/plans")
async def list_plans(
    include_inactive: bool = False,
    limit: int = Query(default=100, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return [normalize_plan(plan) for plan in plan_repository.list_plans(include_inactive=include_inactive, limit=limit)]


@router.get("/subscriptions")
async def list_subscriptions(
    company_id: int | None = Query(default=None, ge=1),
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return subscription_repository.list_subscriptions(company_id=company_id, limit=limit)


@router.get("/payments")
async def list_payments(
    company_id: int | None = Query(default=None, ge=1),
    payment_status: str = Query(default="", alias="status"),
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return payment_repository.list_payments(
        company_id=company_id,
        status=payment_status.strip() or None,
        limit=limit,
    )


@router.post("/payments")
async def create_payment(
    payload: PaymentCreatePayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.create_payment(
        payload.company_id,
        payload.amount,
        payload.due_date or "",
        notes=payload.notes or "",
    )
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "payment": result.get("data")}


@router.post("/payments/{payment_id}/registrar-pagamento")
async def register_payment(
    payment_id: int,
    payload: PaymentRegisterPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.register_payment(
        payment_id,
        method=payload.method or "manual",
        reference=payload.reference or "",
        notes=payload.notes or "",
        actor=actor_name(current_user),
    )
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "payment": result.get("data")}


@router.post("/payments/{payment_id}/gerar-boleto")
async def generate_payment_boleto(
    payment_id: int,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.generate_boleto(payment_id, actor=actor_name(current_user))
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "payment": result.get("data")}


@router.get("/audit-logs")
async def list_saas_audit_logs(
    company_id: int | None = Query(default=None, ge=1),
    action: str = "",
    limit: int = Query(default=500, ge=1, le=5000),
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    return [
        normalize_audit(log)
        for log in audit_repository.list_audit_logs(
            company_id=company_id,
            action=action.strip() or None,
            limit=limit,
        )
    ]


@router.post("/billing/run-overdue-check")
async def run_overdue_check(
    payload: BillingAutomationPayload,
    current_user: User = Depends(require_owner_user),
):
    configure_saas_sqlite()
    result = saas_admin_service.run_overdue_check(grace_days=payload.grace_days)
    if not result.get("ok"):
        raise_service_error(result)
    return {"ok": True, "summary": result.get("data") or {}}
