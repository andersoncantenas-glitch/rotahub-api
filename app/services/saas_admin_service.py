from __future__ import annotations

import json

from app.repositories import audit_repository, company_repository, payment_repository, plan_repository, subscription_repository
from app.repositories.base_repository import ensure_saas_ready, get_db
from app.services.billing_automation_service import suspend_overdue_subscriptions
from app.services.saas_result import error_message, service_result
from app.services.usage_service import get_company_usage
from app.services.vehicle_limit_service import vehicle_usage_snapshot


def get_dashboard(company_id: int | None = None) -> dict:
    try:
        company = company_repository.get_company(company_id) if company_id else company_repository.get_default_company()
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        cid = int(company["id"])
        subscription = subscription_repository.get_active_subscription(cid)
        usage = get_company_usage(cid)
        payments = payment_repository.list_payments(company_id=cid, limit=20)
        audits = audit_repository.list_audit_logs(company_id=cid, limit=50)
        return service_result(
            ok=True,
            data={
                "company": company,
                "subscription": subscription,
                "usage": usage.get("data") if isinstance(usage, dict) else {},
                "payments": payments,
                "audit_logs": audits,
                "plans": plan_repository.list_plans(include_inactive=False, limit=100),
            },
        )
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao carregar dashboard SaaS."))


def change_company_plan(company_id: int, plan_code: str, *, actor: str = "ADMIN", reason: str = "") -> dict:
    try:
        company = company_repository.get_company(company_id)
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        plan = plan_repository.get_plan_by_code(plan_code)
        if not plan:
            return service_result(ok=False, data=None, error="Plano nao encontrado.")
        with get_db() as conn:
            ensure_saas_ready(conn)
            usage = vehicle_usage_snapshot(conn, int(company_id))
        vehicle_count = int((usage or {}).get("vehicle_count") or 0)
        vehicle_limit = plan.get("vehicle_limit")
        usage_result = get_company_usage(company_id)
        usage_data = (usage_result.get("data") if isinstance(usage_result, dict) else {}) or {}
        user_count = int(usage_data.get("users") or 0)
        user_limit = plan.get("user_limit")
        if vehicle_limit is not None and vehicle_count > int(vehicle_limit):
            return service_result(
                ok=False,
                data={"usage": usage, "plan": plan},
                error=(
                    f"Downgrade bloqueado: empresa possui {vehicle_count} veiculos, "
                    f"mas o plano {plan.get('name')} permite {int(vehicle_limit)}."
                ),
            )
        if user_limit is not None and user_count > int(user_limit):
            return service_result(
                ok=False,
                data={"usage": usage_data, "plan": plan},
                error=(
                    f"Downgrade bloqueado: empresa possui {user_count} usuarios, "
                    f"mas o plano {plan.get('name')} permite {int(user_limit)}."
                ),
            )
        subscription = subscription_repository.change_company_plan(company_id, int(plan["id"]))
        audit_repository.create_audit_log(
            {
                "company_id": int(company_id),
                "actor_type": "admin",
                "action": "plano_alterado",
                "entity_type": "subscription",
                "entity_id": str(subscription.get("id") or ""),
                "severity": "info",
                "metadata": {
                    "actor": actor,
                    "new_plan_code": plan.get("code"),
                    "reason": reason,
                    "vehicle_count": vehicle_count,
                    "user_count": user_count,
                },
            }
        )
        return service_result(ok=True, data=subscription)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao alterar plano."))


def set_company_status(company_id: int, status: str, *, actor: str = "ADMIN", reason: str = "") -> dict:
    try:
        company = company_repository.update_company(company_id, {"status": str(status or "").strip().lower()})
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        audit_repository.create_audit_log(
            {
                "company_id": int(company_id),
                "actor_type": "admin",
                "action": "empresa_status_alterado",
                "entity_type": "company",
                "entity_id": str(company_id),
                "severity": "warning",
                "metadata": {"actor": actor, "status": status, "reason": reason},
            }
        )
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao alterar status."))


def create_payment(company_id: int, amount: float, due_date: str = "", *, notes: str = "") -> dict:
    try:
        subscription = subscription_repository.get_active_subscription(company_id)
        payment = payment_repository.create_payment(
            {
                "company_id": int(company_id),
                "subscription_id": (subscription or {}).get("id"),
                "amount": float(amount or 0),
                "due_date": str(due_date or "").strip() or None,
                "notes": notes,
            }
        )
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao criar pagamento."))


def register_payment(payment_id: int, *, method: str = "manual", reference: str = "", notes: str = "", actor: str = "ADMIN") -> dict:
    try:
        payment = payment_repository.register_payment(payment_id, method=method, reference=reference, notes=notes)
        if not payment:
            return service_result(ok=False, data=None, error="Pagamento nao encontrado.")
        audit_repository.create_audit_log(
            {
                "company_id": int(payment.get("company_id") or 0),
                "actor_type": "admin",
                "action": "pagamento_registrado",
                "entity_type": "payment",
                "entity_id": str(payment.get("id") or payment_id),
                "severity": "info",
                "metadata": {"actor": actor, "method": method, "reference": reference},
            }
        )
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao registrar pagamento."))


def run_overdue_check(grace_days: int = 0) -> dict:
    return suspend_overdue_subscriptions(grace_days=int(grace_days or 0))


def format_features(features_json: str | None) -> str:
    try:
        data = json.loads(str(features_json or "{}"))
    except Exception:
        data = {}
    if not isinstance(data, dict):
        return ""
    enabled = [key for key, value in sorted(data.items()) if bool(value)]
    return ", ".join(enabled)
