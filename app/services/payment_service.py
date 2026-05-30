from __future__ import annotations

from app.repositories import company_repository, payment_repository, subscription_repository
from app.services.saas_result import error_message, service_result


def list_payments(company_id: int | None = None, status: str | None = None, limit: int | None = 500) -> dict:
    try:
        payments = payment_repository.list_payments(company_id=company_id, status=status, limit=limit)
        return service_result(ok=True, data=payments)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao listar pagamentos."))


def create_payment(data: dict) -> dict:
    try:
        payload = dict(data or {})
        company_id = int(payload.get("company_id") or 0)
        if not company_repository.get_company(company_id):
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        if not payload.get("subscription_id"):
            subscription = subscription_repository.get_active_subscription(company_id)
            if subscription:
                payload["subscription_id"] = subscription.get("id")
        payment = payment_repository.create_payment(payload)
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao criar pagamento."))


def register_payment(payment_id: int, *, method: str | None = None, reference: str | None = None, notes: str | None = None) -> dict:
    try:
        payment = payment_repository.register_payment(payment_id, method=method, reference=reference, notes=notes)
        if not payment:
            return service_result(ok=False, data=None, error="Pagamento nao encontrado.")
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao registrar pagamento."))
