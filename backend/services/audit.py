# backend/services/audit.py
"""
Audit log service
"""
import json
from typing import Any

from fastapi import Request
from sqlalchemy.ext.asyncio import AsyncSession

from backend.models.audit import AuditLogDB
from backend.models.user import User


def client_ip_from_request(request: Request | None) -> str | None:
    if request is None or request.client is None:
        return None
    return str(request.client.host or "").strip() or None


def record_audit_log(
    db: AsyncSession,
    *,
    action: str,
    actor_user: User | None = None,
    actor_type: str = "user",
    entity_type: str | None = None,
    entity_id: int | str | None = None,
    severity: str = "info",
    ip_address: str | None = None,
    metadata: dict[str, Any] | None = None,
) -> AuditLogDB:
    audit_log = AuditLogDB(
        user_id=actor_user.id if actor_user else None,
        actor_type=actor_type if actor_user else "system",
        action=action,
        entity_type=entity_type,
        entity_id=str(entity_id) if entity_id is not None else None,
        severity=severity,
        ip_address=ip_address,
        metadata_json=json.dumps(metadata or {}, ensure_ascii=True, sort_keys=True),
    )
    db.add(audit_log)
    return audit_log
