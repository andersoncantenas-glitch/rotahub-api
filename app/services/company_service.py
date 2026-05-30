from __future__ import annotations

from app.repositories import company_repository
from app.services.saas_result import error_message, service_result


def list_companies(status: str | None = None, limit: int | None = 500) -> dict:
    try:
        return service_result(ok=True, data=company_repository.list_companies(status=status, limit=limit))
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao listar empresas."))


def get_company(company_id: int) -> dict:
    try:
        company = company_repository.get_company(company_id)
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao buscar empresa."))


def get_default_company() -> dict:
    try:
        company = company_repository.get_default_company()
        if not company:
            return service_result(ok=False, data=None, error="Empresa inicial nao encontrada.")
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao buscar empresa inicial."))


def create_company(data: dict) -> dict:
    try:
        company = company_repository.create_company(data or {})
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao criar empresa."))


def update_company(company_id: int, data: dict) -> dict:
    try:
        company = company_repository.update_company(company_id, data or {})
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao atualizar empresa."))


def set_company_status(company_id: int, status: str) -> dict:
    return update_company(company_id, {"status": str(status or "").strip()})
