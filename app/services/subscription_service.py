from __future__ import annotations

from app.repositories import company_repository, plan_repository, subscription_repository
from app.services.saas_result import error_message, service_result


def get_company_subscription(company_id: int) -> dict:
    try:
        subscription = subscription_repository.get_active_subscription(company_id)
        if not subscription:
            return service_result(ok=False, data=None, error="Assinatura ativa nao encontrada.")
        return service_result(ok=True, data=subscription)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao buscar assinatura."))


def list_subscriptions(company_id: int | None = None, limit: int | None = 500) -> dict:
    try:
        return service_result(ok=True, data=subscription_repository.list_subscriptions(company_id=company_id, limit=limit))
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao listar assinaturas."))


def change_company_plan(company_id: int, plan_code: str) -> dict:
    try:
        company = company_repository.get_company(company_id)
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        plan = plan_repository.get_plan_by_code(plan_code)
        if not plan:
            return service_result(ok=False, data=None, error="Plano nao encontrado.")
        subscription = subscription_repository.change_company_plan(company_id, int(plan["id"]))
        return service_result(ok=True, data=subscription)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao alterar plano da empresa."))
