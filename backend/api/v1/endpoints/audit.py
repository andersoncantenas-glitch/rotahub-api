# backend/api/v1/endpoints/audit.py
"""
Audit log endpoints
"""
import json
from typing import Any, List

from fastapi import APIRouter, Depends, Query
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.audit import AuditLogDB
from backend.models.user import User

router = APIRouter()


class AuditLogResponse(BaseModel):
    id: int
    company_id: int | None
    user_id: int | None
    actor_type: str | None
    action: str
    entity_type: str | None
    entity_id: str | None
    severity: str
    ip_address: str | None
    metadata: dict[str, Any]
    created_at: str | None


def _metadata_dict(metadata_json: str | None) -> dict[str, Any]:
    try:
        parsed = json.loads(metadata_json or "{}")
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _audit_response(log: AuditLogDB) -> AuditLogResponse:
    return AuditLogResponse(
        id=log.id,
        company_id=log.company_id,
        user_id=log.user_id,
        actor_type=log.actor_type,
        action=log.action,
        entity_type=log.entity_type,
        entity_id=log.entity_id,
        severity=log.severity,
        ip_address=log.ip_address,
        metadata=_metadata_dict(log.metadata_json),
        created_at=log.created_at,
    )


@router.get("/", response_model=List[AuditLogResponse])
async def get_audit_logs(
    skip: int = Query(default=0, ge=0),
    limit: int = Query(default=100, ge=1, le=500),
    action: str | None = None,
    entity_type: str | None = None,
    entity_id: str | None = None,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """Get audit logs (admin only)"""
    stmt = select(AuditLogDB)
    if action:
        stmt = stmt.where(AuditLogDB.action == action)
    if entity_type:
        stmt = stmt.where(AuditLogDB.entity_type == entity_type)
    if entity_id:
        stmt = stmt.where(AuditLogDB.entity_id == entity_id)
    stmt = stmt.order_by(AuditLogDB.id.desc()).offset(skip).limit(limit)
    result = await db.execute(stmt)
    return [_audit_response(log) for log in result.scalars().all()]
