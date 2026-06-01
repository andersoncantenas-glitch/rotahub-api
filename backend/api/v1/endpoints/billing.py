from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy.engine import make_url

from app.db.connection import configure_connection
from app.repositories.base_repository import ensure_saas_ready, get_db as get_saas_db, row_to_dict, rows_to_dicts
from app.services.vehicle_limit_service import vehicle_usage_snapshot
from backend.config.settings import settings
from backend.models.user import User
from backend.services.auth import get_current_user
from backend.services.plan_change_requests import create_plan_change_request, list_plan_change_requests

router = APIRouter()

PROJECT_ROOT = Path(__file__).resolve().parents[4]
COMMERCIAL_PLAN_CODES = ("starter", "growth", "professional", "enterprise")


class PlanChangeRequestPayload(BaseModel):
    requested_plan_code: str = Field(min_length=1, max_length=80)
    message: str | None = Field(default=None, max_length=500)

    @field_validator("requested_plan_code", mode="before")
    @classmethod
    def normalize_plan(cls, value):
        return str(value or "").strip().lower()

    @field_validator("message", mode="before")
    @classmethod
    def normalize_message(cls, value):
        return str(value or "").strip() or None


def configure_saas_sqlite() -> None:
    url = make_url(settings.DATABASE_URL)
    if not url.drivername.startswith("sqlite") or not url.database or url.database == ":memory:":
        raise HTTPException(status_code=409, detail="Assinatura indisponivel para este banco.")
    path = Path(url.database)
    if not path.is_absolute():
        path = (PROJECT_ROOT / path).resolve()
    configure_connection(str(path))


def decode_json(raw: Any) -> dict:
    if isinstance(raw, dict):
        return dict(raw)
    try:
        parsed = json.loads(str(raw or "{}"))
    except Exception:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def commercial_plans(conn) -> list[dict]:
    cur = conn.cursor()
    placeholders = ",".join("?" for _ in COMMERCIAL_PLAN_CODES)
    cur.execute(
        f"""
        SELECT id, code, name, description, monthly_price, vehicle_limit, user_limit, features_json
          FROM plans
         WHERE status='active' AND code IN ({placeholders})
         ORDER BY CASE code
            WHEN 'starter' THEN 1
            WHEN 'growth' THEN 2
            WHEN 'professional' THEN 3
            WHEN 'enterprise' THEN 4
            ELSE 99
         END
        """,
        COMMERCIAL_PLAN_CODES,
    )
    plans = rows_to_dicts(cur.fetchall())
    for plan in plans:
        plan["features"] = decode_json(plan.get("features_json"))
    return plans


def active_subscription(conn, company_id: int) -> dict:
    cur = conn.cursor()
    cur.execute(
        """
        SELECT s.*, p.code AS plan_code, p.name AS plan_name, p.monthly_price AS plan_monthly_price,
               p.vehicle_limit AS plan_vehicle_limit, p.user_limit AS plan_user_limit,
               p.features_json AS plan_features_json
          FROM subscriptions s
          JOIN plans p ON p.id=s.plan_id
         WHERE s.company_id=? AND s.status IN ('active', 'trialing', 'past_due')
         ORDER BY s.id DESC
         LIMIT 1
        """,
        (int(company_id),),
    )
    row = row_to_dict(cur.fetchone()) or {}
    row["features"] = decode_json(row.get("plan_features_json"))
    return row


def company_info(conn, company_id: int) -> dict:
    cur = conn.cursor()
    cur.execute("SELECT id, name, document, email, phone, status FROM companies WHERE id=? LIMIT 1", (int(company_id),))
    return row_to_dict(cur.fetchone()) or {}


def plan_by_code(conn, code: str) -> dict | None:
    cur = conn.cursor()
    cur.execute("SELECT * FROM plans WHERE code=? AND status='active' LIMIT 1", (str(code or "").strip().lower(),))
    return row_to_dict(cur.fetchone())


def count_company_table(conn, table: str, company_id: int) -> int:
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
    if not cur.fetchone():
        return 0
    cur.execute(f"PRAGMA table_info({table})")
    columns = {str(row[1]).lower() for row in cur.fetchall()}
    if "company_id" in columns:
        cur.execute(f'SELECT COUNT(*) FROM "{table}" WHERE company_id=?', (int(company_id),))
    else:
        cur.execute(f'SELECT COUNT(*) FROM "{table}"')
    row = cur.fetchone()
    return int(row[0] if row else 0)


def company_usage(conn, company_id: int) -> dict:
    return {
        "company_id": int(company_id),
        "vehicles": vehicle_usage_snapshot(conn, int(company_id)),
        "users": count_company_table(conn, "usuarios", int(company_id)),
        "motoristas": count_company_table(conn, "motoristas", int(company_id)),
        "vendedores": count_company_table(conn, "vendedores", int(company_id)),
        "programacoes": count_company_table(conn, "programacoes", int(company_id)),
        "clientes": count_company_table(conn, "clientes", int(company_id)),
    }


def request_type(current: dict, requested: dict) -> str:
    current_price = float(current.get("plan_monthly_price") or 0)
    requested_price = float(requested.get("monthly_price") or 0)
    if requested_price > current_price:
        return "upgrade"
    if requested_price < current_price:
        return "downgrade"
    return "change"


@router.get("/my-plan")
async def my_plan(current_user: User = Depends(get_current_user)):
    configure_saas_sqlite()
    company_id = int(current_user.company_id or 1)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        subscription = active_subscription(conn, company_id)
        pending_requests = list_plan_change_requests(conn, status="pending", limit=100)
        pending_requests = [item for item in pending_requests if int(item.get("company_id") or 0) == company_id]
        return {
            "company": company_info(conn, company_id),
            "subscription": subscription,
            "usage": company_usage(conn, company_id),
            "plans": commercial_plans(conn),
            "pending_requests": pending_requests,
        }


@router.post("/plan-change-requests", status_code=status.HTTP_201_CREATED)
async def request_plan_change(
    payload: PlanChangeRequestPayload,
    current_user: User = Depends(get_current_user),
):
    configure_saas_sqlite()
    company_id = int(current_user.company_id or 1)
    with get_saas_db() as conn:
        ensure_saas_ready(conn)
        subscription = active_subscription(conn, company_id)
        if not subscription:
            raise HTTPException(status_code=409, detail="Empresa sem assinatura ativa.")
        requested_plan = plan_by_code(conn, payload.requested_plan_code)
        if not requested_plan or payload.requested_plan_code not in COMMERCIAL_PLAN_CODES:
            raise HTTPException(status_code=404, detail="Plano solicitado nao encontrado.")
        if (
            str(subscription.get("plan_code") or "").lower() == payload.requested_plan_code
            and str(subscription.get("status") or "").strip().lower() != "trialing"
        ):
            raise HTTPException(status_code=409, detail="Este ja e o plano atual da empresa.")
        existing = [
            item for item in list_plan_change_requests(conn, status="pending", limit=500)
            if int(item.get("company_id") or 0) == company_id
        ]
        if existing:
            raise HTTPException(status_code=409, detail="Ja existe uma solicitacao de plano pendente.")
        usage = company_usage(conn, company_id)
        vehicles = usage.get("vehicles") or {}
        request = create_plan_change_request(
            conn,
            {
                "company_id": company_id,
                "current_plan_code": subscription.get("plan_code"),
                "requested_plan_code": requested_plan.get("code"),
                "request_type": request_type(subscription, requested_plan),
                "vehicle_count": vehicles.get("vehicle_count") or 0,
                "vehicle_limit_current": subscription.get("plan_vehicle_limit"),
                "vehicle_limit_requested": requested_plan.get("vehicle_limit"),
                "user_count": usage.get("users") or 0,
                "user_limit_current": subscription.get("plan_user_limit"),
                "user_limit_requested": requested_plan.get("user_limit"),
                "message": payload.message or "",
                "requested_by_user_id": current_user.id,
                "requested_by_name": current_user.nome or current_user.username,
            },
        )
        return {"ok": True, "request": request}
