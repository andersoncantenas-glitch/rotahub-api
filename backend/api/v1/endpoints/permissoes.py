# backend/api/v1/endpoints/permissoes.py
"""
Fine-grained permissions endpoints mirroring the desktop PermissionsPage.
"""
from __future__ import annotations

from typing import List

from fastapi import APIRouter, Depends, HTTPException, Request, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import delete, func, select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.permissions import PermissaoDB, UsuarioPermissaoDB
from backend.models.user import User, UserDB
from backend.services.audit import client_ip_from_request, record_audit_log
from backend.services.permissions import assign_permissions_by_profile, ensure_default_permissions, normalize_profile

router = APIRouter()


class PermissionResponse(BaseModel):
    id: int
    modulo: str
    nome: str
    descricao: str | None
    ativo: bool


class GrantedPermissionResponse(PermissionResponse):
    concessao_id: int
    concedida_em: str | None
    concedida_por: str | None


class PermissionUserResponse(BaseModel):
    id: int
    username: str
    nome: str
    permissoes: str
    is_active: bool
    is_admin: bool
    granted_count: int


class PermissionsOverviewResponse(BaseModel):
    usuarios: list[PermissionUserResponse]
    permissoes: list[PermissionResponse]
    modulos: list[str]


class GrantPermissionPayload(BaseModel):
    permissao_id: int = Field(gt=0)


class AssignProfilePayload(BaseModel):
    perfil: str = Field(min_length=1, max_length=40)

    @field_validator("perfil", mode="before")
    @classmethod
    def strip_profile(cls, value):
        return str(value or "").strip().upper()


class MutatingPermissionResponse(BaseModel):
    ok: bool
    mensagem: str
    usuario_id: int
    permissao_id: int
    permissao: GrantedPermissionResponse | None = None
    ja_existia: bool = False


class ProfileAssignResponse(BaseModel):
    ok: bool
    usuario_id: int
    perfil: str
    permissoes_atribuidas: int


class ModuleUserResponse(BaseModel):
    id: int
    username: str
    nome: str
    qtd_permissoes: int


def _is_admin_user(user: UserDB) -> bool:
    return (
        str(user.permissoes or "").strip().upper() == "ADMIN"
        or str(user.nome or "").strip().upper() == "ADMIN"
        or str(user.username or "").strip().upper() == "ADMIN"
    )


def _permission_response(permission: PermissaoDB) -> PermissionResponse:
    return PermissionResponse(
        id=permission.id,
        modulo=permission.modulo or "",
        nome=permission.nome_permissao or "",
        descricao=permission.descricao,
        ativo=bool(permission.ativo),
    )


def _granted_response(permission: PermissaoDB, grant: UsuarioPermissaoDB) -> GrantedPermissionResponse:
    return GrantedPermissionResponse(
        **_permission_response(permission).model_dump(),
        concessao_id=grant.id,
        concedida_em=grant.concedida_em,
        concedida_por=grant.concedida_por,
    )


async def _get_user_or_404(db: AsyncSession, user_id: int) -> UserDB:
    user = await db.get(UserDB, user_id)
    if user is None:
        raise HTTPException(status_code=404, detail="Usuario nao encontrado")
    return user


async def _get_permission_or_404(db: AsyncSession, permission_id: int) -> PermissaoDB:
    permission = await db.get(PermissaoDB, permission_id)
    if permission is None:
        raise HTTPException(status_code=404, detail="Permissao nao encontrada")
    return permission


async def _list_user_permissions(db: AsyncSession, user_id: int) -> list[GrantedPermissionResponse]:
    result = await db.execute(
        select(PermissaoDB, UsuarioPermissaoDB)
        .join(UsuarioPermissaoDB, UsuarioPermissaoDB.permissao_id == PermissaoDB.id)
        .where(UsuarioPermissaoDB.usuario_id == user_id)
        .order_by(PermissaoDB.modulo, PermissaoDB.nome_permissao)
    )
    return [_granted_response(permission, grant) for permission, grant in result.all()]


async def _permission_counts_by_user(db: AsyncSession) -> dict[int, int]:
    result = await db.execute(
        select(UsuarioPermissaoDB.usuario_id, func.count(UsuarioPermissaoDB.id)).group_by(
            UsuarioPermissaoDB.usuario_id
        )
    )
    return {int(user_id): int(count or 0) for user_id, count in result.all()}


@router.get("/overview", response_model=PermissionsOverviewResponse)
async def permissions_overview(
    include_inactive: bool = True,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """Load users and available permissions for the web manager."""
    await ensure_default_permissions(db)
    users_stmt = select(UserDB)
    if not include_inactive:
        users_stmt = users_stmt.where(UserDB.is_active.is_(True))
    users_stmt = users_stmt.order_by(UserDB.nome, UserDB.username, UserDB.id)
    users = (await db.execute(users_stmt)).scalars().all()
    counts = await _permission_counts_by_user(db)

    permissions = (
        (
            await db.execute(
                select(PermissaoDB).order_by(PermissaoDB.modulo, PermissaoDB.nome_permissao, PermissaoDB.id)
            )
        )
        .scalars()
        .all()
    )
    permission_responses = [_permission_response(permission) for permission in permissions]
    return PermissionsOverviewResponse(
        usuarios=[
            PermissionUserResponse(
                id=user.id,
                username=user.username,
                nome=user.nome,
                permissoes=user.permissoes or "",
                is_active=bool(user.is_active),
                is_admin=_is_admin_user(user),
                granted_count=counts.get(user.id, 0),
            )
            for user in users
        ],
        permissoes=permission_responses,
        modulos=sorted({permission.modulo for permission in permission_responses if permission.modulo}),
    )


@router.get("/disponiveis", response_model=List[PermissionResponse])
async def available_permissions(
    modulo: str | None = None,
    apenas_ativas: bool = True,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """List permissions in the catalog, optionally filtered by module."""
    await ensure_default_permissions(db)
    stmt = select(PermissaoDB)
    if modulo:
        stmt = stmt.where(func.lower(PermissaoDB.modulo) == modulo.strip().lower())
    if apenas_ativas:
        stmt = stmt.where(PermissaoDB.ativo == 1)
    stmt = stmt.order_by(PermissaoDB.modulo, PermissaoDB.nome_permissao)
    result = await db.execute(stmt)
    return [_permission_response(permission) for permission in result.scalars().all()]


@router.get("/usuarios/{user_id}", response_model=List[GrantedPermissionResponse])
async def user_permissions(
    user_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """List permissions granted to a user."""
    await ensure_default_permissions(db)
    await _get_user_or_404(db, user_id)
    return await _list_user_permissions(db, user_id)


@router.post("/usuarios/{user_id}/conceder", response_model=MutatingPermissionResponse)
async def grant_permission(
    user_id: int,
    payload: GrantPermissionPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """Grant the selected permission to the selected user."""
    await ensure_default_permissions(db)
    user = await _get_user_or_404(db, user_id)
    permission = await _get_permission_or_404(db, payload.permissao_id)

    existing = (
        await db.execute(
            select(UsuarioPermissaoDB).where(
                UsuarioPermissaoDB.usuario_id == user_id,
                UsuarioPermissaoDB.permissao_id == payload.permissao_id,
            )
        )
    ).scalar_one_or_none()
    if existing is not None:
        return MutatingPermissionResponse(
            ok=True,
            mensagem="Permissao ja concedida.",
            usuario_id=user_id,
            permissao_id=payload.permissao_id,
            permissao=_granted_response(permission, existing),
            ja_existia=True,
        )

    grant = UsuarioPermissaoDB(
        usuario_id=user_id,
        permissao_id=payload.permissao_id,
        concedida_por=current_user.username,
    )
    db.add(grant)
    await db.flush()
    record_audit_log(
        db,
        action="permissao_concedida",
        actor_user=current_user,
        entity_type="user_permission",
        entity_id=user_id,
        ip_address=client_ip_from_request(request),
        metadata={
            "target_user_id": user.id,
            "target_username": user.username,
            "permission_id": permission.id,
            "modulo": permission.modulo,
            "nome_permissao": permission.nome_permissao,
        },
    )
    await db.commit()
    await db.refresh(grant)
    return MutatingPermissionResponse(
        ok=True,
        mensagem="Permissao concedida.",
        usuario_id=user_id,
        permissao_id=payload.permissao_id,
        permissao=_granted_response(permission, grant),
    )


@router.delete("/usuarios/{user_id}/{permission_id}", response_model=MutatingPermissionResponse)
async def revoke_permission(
    user_id: int,
    permission_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """Revoke a permission from the selected user."""
    await ensure_default_permissions(db)
    user = await _get_user_or_404(db, user_id)
    permission = await _get_permission_or_404(db, permission_id)
    grant = (
        await db.execute(
            select(UsuarioPermissaoDB).where(
                UsuarioPermissaoDB.usuario_id == user_id,
                UsuarioPermissaoDB.permissao_id == permission_id,
            )
        )
    ).scalar_one_or_none()
    if grant is None:
        raise HTTPException(status_code=404, detail="Permissao nao encontrada para este usuario")

    record_audit_log(
        db,
        action="permissao_revogada",
        actor_user=current_user,
        entity_type="user_permission",
        entity_id=user_id,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={
            "target_user_id": user.id,
            "target_username": user.username,
            "permission_id": permission.id,
            "modulo": permission.modulo,
            "nome_permissao": permission.nome_permissao,
        },
    )
    await db.execute(delete(UsuarioPermissaoDB).where(UsuarioPermissaoDB.id == grant.id))
    await db.commit()
    return MutatingPermissionResponse(
        ok=True,
        mensagem="Permissao revogada.",
        usuario_id=user_id,
        permissao_id=permission_id,
    )


@router.post("/usuarios/{user_id}/perfil", response_model=ProfileAssignResponse)
async def assign_profile(
    user_id: int,
    payload: AssignProfilePayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """Apply one of the desktop profiles and replace the user's permissions."""
    await ensure_default_permissions(db)
    user = await _get_user_or_404(db, user_id)
    try:
        profile = normalize_profile(payload.perfil)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc)) from exc

    before = [permission.model_dump() for permission in await _list_user_permissions(db, user_id)]
    result = await assign_permissions_by_profile(
        db,
        user_id=user_id,
        profile=profile,
        granted_by=current_user.username,
        commit=False,
    )
    user.permissoes = profile
    record_audit_log(
        db,
        action="permissoes_perfil_aplicado",
        actor_user=current_user,
        entity_type="user",
        entity_id=user_id,
        ip_address=client_ip_from_request(request),
        metadata={
            "target_user_id": user.id,
            "target_username": user.username,
            "perfil": profile,
            "permissoes_antes": before,
            "permissoes_atribuidas": result["permissoes_atribuidas"],
        },
    )
    await db.commit()
    return ProfileAssignResponse(ok=True, **result)


@router.get("/modulos/{modulo}/usuarios", response_model=List[ModuleUserResponse])
async def users_with_module(
    modulo: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    """List users with at least one permission in a module."""
    await ensure_default_permissions(db)
    result = await db.execute(
        select(UserDB.id, UserDB.username, UserDB.nome, func.count(UsuarioPermissaoDB.id).label("qtd"))
        .join(UsuarioPermissaoDB, UsuarioPermissaoDB.usuario_id == UserDB.id)
        .join(PermissaoDB, UsuarioPermissaoDB.permissao_id == PermissaoDB.id)
        .where(func.lower(PermissaoDB.modulo) == modulo.strip().lower())
        .group_by(UserDB.id, UserDB.username, UserDB.nome)
        .order_by(UserDB.nome, UserDB.username)
    )
    return [
        ModuleUserResponse(id=user_id, username=username, nome=nome, qtd_permissoes=int(qtd or 0))
        for user_id, username, nome, qtd in result.all()
    ]
