# backend/models/audit.py
"""
Audit log model
"""
from datetime import datetime, timezone

from sqlalchemy import Column, Integer, String, Text
from pydantic import BaseModel as PydanticBaseModel, ConfigDict

from backend.config.database import Base


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


class AuditLogDB(Base):
    """SQLAlchemy audit log model"""
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True, index=True)
    company_id = Column(Integer, nullable=True)
    user_id = Column(Integer, nullable=True)
    actor_type = Column(String, default="system")
    action = Column(String, nullable=False, index=True)
    entity_type = Column(String, nullable=True)
    entity_id = Column(String, nullable=True)
    severity = Column(String, nullable=False, default="info")
    ip_address = Column(String, nullable=True)
    metadata_json = Column(Text, nullable=False, default="{}")
    created_at = Column(String, default=utc_now_iso)


class AuditLog(PydanticBaseModel):
    """Pydantic audit log model"""
    model_config = ConfigDict(from_attributes=True)

    id: int
    company_id: int | None = None
    user_id: int | None = None
    actor_type: str | None = "system"
    action: str
    entity_type: str | None = None
    entity_id: str | None = None
    severity: str = "info"
    ip_address: str | None = None
    metadata_json: str = "{}"
    created_at: str | None = None
