from __future__ import annotations

from app.repositories import audit_repository
from app.services.saas_result import error_message, service_result


def record_audit_log(
    *,
    action: str,
    company_id: int | None = None,
    user_id: int | None = None,
    actor_type: str = "system",
    entity_type: str | None = None,
    entity_id: str | None = None,
    severity: str = "info",
    ip_address: str | None = None,
    metadata: dict | str | None = None,
) -> dict:
    try:
        audit_log = audit_repository.create_audit_log(
            {
                "company_id": company_id,
                "user_id": user_id,
                "actor_type": actor_type,
                "action": action,
                "entity_type": entity_type,
                "entity_id": entity_id,
                "severity": severity,
                "ip_address": ip_address,
                "metadata": metadata,
            }
        )
        return service_result(ok=True, data=audit_log)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao registrar audit log."))


def list_audit_logs(company_id: int | None = None, action: str | None = None, limit: int | None = 500) -> dict:
    try:
        logs = audit_repository.list_audit_logs(company_id=company_id, action=action, limit=limit)
        return service_result(ok=True, data=logs)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao listar audit logs."))
