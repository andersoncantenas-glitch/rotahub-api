# backend/api/v1/endpoints/users.py
"""
User management endpoints
"""
from typing import List
from fastapi import APIRouter, Depends, HTTPException, Request, status
from sqlalchemy import select, text
from sqlalchemy.ext.asyncio import AsyncSession
from pydantic import BaseModel, ConfigDict, Field, field_validator

from backend.config.database import get_db
from backend.config.settings import settings
from backend.models.user import User, UserDB
from backend.services.audit import client_ip_from_request, record_audit_log
from backend.services.auth import get_current_user, get_password_hash
from backend.services.permissions import assign_permissions_by_profile, normalize_profile

router = APIRouter()
AUDITED_USER_FIELDS = ("username", "nome", "permissoes", "is_active", "cpf", "idade", "telefone", "company_id")


class UserResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    username: str
    nome: str
    permissoes: str
    is_active: bool
    cpf: str | None
    idade: int | None
    telefone: str | None
    company_id: int | None = None


class UserCreate(BaseModel):
    username: str = Field(min_length=1, max_length=80)
    password: str = Field(min_length=6, max_length=128)
    nome: str = Field(min_length=1, max_length=120)
    permissoes: str = Field(default="OPERADOR", min_length=1, max_length=40)
    is_active: bool = True
    cpf: str | None = Field(default=None, max_length=20)
    idade: int | None = Field(default=None, ge=0, le=130)
    telefone: str | None = Field(default=None, max_length=30)
    company_id: int | None = Field(default=None, ge=1)

    @field_validator("username", "nome", "permissoes", "cpf", "telefone", mode="before")
    @classmethod
    def strip_text(cls, value):
        if value is None:
            return None
        return str(value).strip()


class UserUpdate(BaseModel):
    username: str | None = Field(default=None, min_length=1, max_length=80)
    password: str | None = Field(default=None, min_length=6, max_length=128)
    nome: str | None = Field(default=None, min_length=1, max_length=120)
    permissoes: str | None = Field(default=None, min_length=1, max_length=40)
    is_active: bool | None = None
    cpf: str | None = Field(default=None, max_length=20)
    idade: int | None = Field(default=None, ge=0, le=130)
    telefone: str | None = Field(default=None, max_length=30)
    company_id: int | None = Field(default=None, ge=1)

    @field_validator("username", "nome", "permissoes", "cpf", "telefone", mode="before")
    @classmethod
    def strip_text(cls, value):
        if value is None:
            return None
        return str(value).strip()


async def require_admin_user(current_user: User = Depends(get_current_user)) -> User:
    if str(current_user.permissoes or "").strip().upper() != "ADMIN":
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Admin permission required",
        )
    return current_user


async def require_owner_user(current_user: User = Depends(get_current_user)) -> User:
    profile = str(current_user.permissoes or "").strip().upper()
    username = str(current_user.username or "").strip().upper()
    owner_users = {item.strip().upper() for item in settings.OWNER_ADMIN_USERS if item.strip()}
    if profile in {"DONO", "OWNER", "SUPERADMIN", "SUPER_ADMIN"} or username in owner_users:
        return current_user
    raise HTTPException(
        status_code=status.HTTP_403_FORBIDDEN,
        detail="Acesso permitido somente ao administrador dono do sistema.",
    )


async def get_user_or_404(db: AsyncSession, user_id: int) -> UserDB:
    user = await db.get(UserDB, user_id)
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    return user


async def ensure_username_available(db: AsyncSession, username: str, *, exclude_user_id: int | None = None) -> None:
    stmt = select(UserDB).where(UserDB.username == username)
    result = await db.execute(stmt)
    existing = result.scalar_one_or_none()
    if existing and existing.id != exclude_user_id:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Username already exists",
        )


async def enforce_user_limit(db: AsyncSession, current_user: User) -> None:
    company_id = int(current_user.company_id or 1)
    result = await db.execute(
        text(
            """
            SELECT p.user_limit
              FROM subscriptions s
              JOIN plans p ON p.id=s.plan_id
             WHERE s.company_id=:company_id
               AND s.status IN ('active', 'trialing', 'past_due')
             ORDER BY s.id DESC
             LIMIT 1
            """
        ),
        {"company_id": company_id},
    )
    limit = result.scalar_one_or_none()
    if limit is None:
        return
    count = (
        await db.execute(
            text("SELECT COUNT(*) FROM usuarios WHERE company_id=:company_id AND COALESCE(is_active, 1)=1"),
            {"company_id": company_id},
        )
    ).scalar()
    if int(count or 0) >= int(limit):
        raise HTTPException(
            status_code=403,
            detail=f"Limite de {int(limit)} usuarios atingido no plano atual. Faça upgrade para cadastrar mais usuários.",
        )


def user_audit_snapshot(user: UserDB) -> dict:
    return {field: getattr(user, field) for field in AUDITED_USER_FIELDS}


def changed_user_fields(before: dict, after: dict) -> list[str]:
    return [field for field in AUDITED_USER_FIELDS if before.get(field) != after.get(field)]


def user_update_audit_action(before: dict, after: dict) -> str:
    if before.get("is_active") is True and after.get("is_active") is False:
        return "usuario_desativado"
    if before.get("is_active") is False and after.get("is_active") is True:
        return "usuario_reativado"
    return "usuario_alterado"


async def assign_known_profile_permissions(
    db: AsyncSession,
    *,
    user_id: int,
    profile: str | None,
    granted_by: str,
) -> dict | None:
    try:
        normalized_profile = normalize_profile(profile)
    except ValueError:
        return None
    return await assign_permissions_by_profile(
        db,
        user_id=user_id,
        profile=normalized_profile,
        granted_by=granted_by,
        commit=False,
    )


@router.get("/me", response_model=UserResponse)
async def get_current_user_info(current_user: User = Depends(get_current_user)):
    """Get current user information"""
    return current_user


@router.get("/", response_model=List[UserResponse])
async def get_users(
    skip: int = 0,
    limit: int = 100,
    include_inactive: bool = False,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user)
):
    """Get list of users (admin only)"""
    stmt = select(UserDB)
    if not include_inactive:
        stmt = stmt.where(UserDB.is_active.is_(True))
    stmt = stmt.offset(skip).limit(limit)
    result = await db.execute(stmt)
    return result.scalars().all()


@router.post("/", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
async def create_user(
    payload: UserCreate,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user)
):
    """Create a user (admin only)"""
    if payload.is_active:
        await enforce_user_limit(db, current_user)
    await ensure_username_available(db, payload.username)
    user = UserDB(
        username=payload.username,
        nome=payload.nome,
        senha=get_password_hash(payload.password),
        permissoes=payload.permissoes,
        is_active=payload.is_active,
        cpf=payload.cpf,
        idade=payload.idade,
        telefone=payload.telefone,
        company_id=current_user.company_id or 1,
    )
    db.add(user)
    await db.flush()
    profile_assignment = await assign_known_profile_permissions(
        db,
        user_id=user.id,
        profile=user.permissoes,
        granted_by=current_user.username,
    )
    record_audit_log(
        db,
        action="usuario_criado",
        actor_user=current_user,
        entity_type="user",
        entity_id=user.id,
        ip_address=client_ip_from_request(request),
        metadata={
            "target_user_id": user.id,
            "target_username": user.username,
            "is_active": bool(user.is_active),
            "created_fields": changed_user_fields({}, user_audit_snapshot(user)),
            "profile_assignment": profile_assignment,
        },
    )
    await db.commit()
    await db.refresh(user)
    return user


@router.get("/{user_id}", response_model=UserResponse)
async def get_user(
    user_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user)
):
    """Get user by ID (admin only)"""
    return await get_user_or_404(db, user_id)


@router.patch("/{user_id}", response_model=UserResponse)
async def update_user(
    user_id: int,
    payload: UserUpdate,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user)
):
    """Update user by ID (admin only)"""
    user = await get_user_or_404(db, user_id)
    before = user_audit_snapshot(user)
    data = payload.model_dump(exclude_unset=True)

    if user.id == current_user.id and data.get("is_active") is False:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Cannot deactivate your own user",
        )

    username = data.pop("username", None)
    if username is not None:
        await ensure_username_available(db, username, exclude_user_id=user.id)
        user.username = username

    password = data.pop("password", None)
    password_changed = password is not None
    if password is not None:
        user.senha = get_password_hash(password)

    profile_changed_to = data.get("permissoes")
    for field, value in data.items():
        setattr(user, field, value)

    profile_assignment = None
    if profile_changed_to is not None:
        profile_assignment = await assign_known_profile_permissions(
            db,
            user_id=user.id,
            profile=profile_changed_to,
            granted_by=current_user.username,
        )

    after = user_audit_snapshot(user)
    changed_fields = changed_user_fields(before, after)
    if password_changed:
        changed_fields.append("password")

    if changed_fields:
        record_audit_log(
            db,
            action=user_update_audit_action(before, after),
            actor_user=current_user,
            entity_type="user",
            entity_id=user.id,
            ip_address=client_ip_from_request(request),
            metadata={
                "target_user_id": user.id,
                "target_username": user.username,
                "changed_fields": changed_fields,
                "previous_is_active": bool(before.get("is_active")),
                "new_is_active": bool(after.get("is_active")),
                "password_changed": password_changed,
                "profile_assignment": profile_assignment,
            },
        )

    await db.commit()
    await db.refresh(user)
    return user


@router.delete("/{user_id}", response_model=UserResponse)
async def deactivate_user(
    user_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user)
):
    """Deactivate user by ID (admin only)"""
    user = await get_user_or_404(db, user_id)
    if user.id == current_user.id:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Cannot deactivate your own user",
        )

    before = user_audit_snapshot(user)
    user.is_active = False
    after = user_audit_snapshot(user)
    if before.get("is_active") != after.get("is_active"):
        record_audit_log(
            db,
            action="usuario_desativado",
            actor_user=current_user,
            entity_type="user",
            entity_id=user.id,
            severity="warning",
            ip_address=client_ip_from_request(request),
            metadata={
                "target_user_id": user.id,
                "target_username": user.username,
                "changed_fields": ["is_active"],
                "previous_is_active": bool(before.get("is_active")),
                "new_is_active": bool(after.get("is_active")),
            },
        )
    await db.commit()
    await db.refresh(user)
    return user
