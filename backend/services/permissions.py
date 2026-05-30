# backend/services/permissions.py
"""
Permission catalog and profile assignment helpers for the browser API.
"""
from __future__ import annotations

from sqlalchemy import delete, func, or_, select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.models.permissions import PermissaoDB, UsuarioPermissaoDB
from backend.models.user import UserDB


DEFAULT_PERMISSIONS: tuple[tuple[str, str, str], ...] = (
    ("programacoes", "visualizar_programacoes", "Visualizar programacoes"),
    ("programacoes", "criar_programacoes", "Criar novas programacoes"),
    ("programacoes", "editar_programacoes", "Editar programacoes"),
    ("programacoes", "deletar_programacoes", "Deletar programacoes"),
    ("programacoes", "finalizar_programacoes", "Finalizar programacoes"),
    ("prestacao", "gerar_prestacao", "Gerar prestacao de contas"),
    ("prestacao", "editar_prestacao", "Editar prestacao"),
    ("prestacao", "fechar_prestacao", "Fechar prestacao"),
    ("cadastros", "gerenciar_clientes", "Gerenciar clientes"),
    ("cadastros", "gerenciar_motoristas", "Gerenciar motoristas"),
    ("cadastros", "gerenciar_vendedores", "Gerenciar vendedores"),
    ("relatorios", "gerar_relatorios", "Gerar relatorios"),
    ("relatorios", "exportar_dados", "Exportar dados para Excel/PDF"),
    ("sistema", "gerenciar_usuarios", "Gerenciar usuarios e permissoes"),
    ("sistema", "acessar_ferramentas", "Acessar ferramentas do sistema"),
    ("sistema", "fazer_backup", "Fazer backup do banco de dados"),
    ("sistema", "restaurar_backup", "Restaurar backup"),
    ("sistema", "limpar_logs", "Limpar logs do sistema"),
    ("sistema", "ver_configuracoes", "Visualizar configuracoes"),
    ("sistema", "editar_configuracoes", "Editar configuracoes"),
)


PROFILE_PERMISSION_NAMES: dict[str, dict[str, tuple[str, ...]]] = {
    "GERENTE": {
        "programacoes": (
            "visualizar_programacoes",
            "criar_programacoes",
            "editar_programacoes",
            "finalizar_programacoes",
        ),
        "prestacao": ("gerar_prestacao", "editar_prestacao"),
        "cadastros": ("gerenciar_clientes", "gerenciar_motoristas"),
        "relatorios": ("gerar_relatorios", "exportar_dados"),
        "sistema": ("ver_configuracoes",),
    },
    "OPERADOR": {
        "programacoes": ("visualizar_programacoes", "criar_programacoes"),
        "prestacao": ("gerar_prestacao",),
        "cadastros": ("gerenciar_clientes",),
        "relatorios": ("gerar_relatorios",),
    },
    "VISUALIZADOR": {
        "programacoes": ("visualizar_programacoes",),
        "relatorios": ("gerar_relatorios",),
    },
}

VALID_PROFILES = ("ADMIN", "GERENTE", "OPERADOR", "VISUALIZADOR")


def normalize_profile(profile: str | None) -> str:
    value = str(profile or "").strip().upper()
    if value not in VALID_PROFILES:
        raise ValueError(f"Perfil '{profile}' nao reconhecido")
    return value


async def ensure_default_permissions(db: AsyncSession, *, commit: bool = True) -> None:
    """
    Keep the permission catalog present and guarantee ADMIN users have full access.
    """
    dirty = False
    existing_result = await db.execute(select(PermissaoDB))
    existing = {
        (str(permission.modulo or "").lower(), str(permission.nome_permissao or "").lower()): permission
        for permission in existing_result.scalars().all()
    }

    for modulo, nome_permissao, descricao in DEFAULT_PERMISSIONS:
        key = (modulo.lower(), nome_permissao.lower())
        if key not in existing:
            permission = PermissaoDB(
                modulo=modulo,
                nome_permissao=nome_permissao,
                descricao=descricao,
                ativo=1,
            )
            db.add(permission)
            existing[key] = permission
            dirty = True

    if dirty:
        await db.flush()

    permissions_result = await db.execute(select(PermissaoDB).where(PermissaoDB.ativo == 1))
    active_permissions = permissions_result.scalars().all()

    admins_result = await db.execute(
        select(UserDB).where(
            or_(
                func.upper(func.coalesce(UserDB.permissoes, "")) == "ADMIN",
                func.upper(func.coalesce(UserDB.nome, "")) == "ADMIN",
                func.upper(func.coalesce(UserDB.username, "")) == "ADMIN",
            )
        )
    )
    admins = admins_result.scalars().all()
    if admins and active_permissions:
        admin_ids = [admin.id for admin in admins]
        permission_ids = [permission.id for permission in active_permissions]
        grants_result = await db.execute(
            select(UsuarioPermissaoDB.usuario_id, UsuarioPermissaoDB.permissao_id).where(
                UsuarioPermissaoDB.usuario_id.in_(admin_ids),
                UsuarioPermissaoDB.permissao_id.in_(permission_ids),
            )
        )
        existing_grants = {(usuario_id, permissao_id) for usuario_id, permissao_id in grants_result.all()}
        for admin_id in admin_ids:
            for permission_id in permission_ids:
                if (admin_id, permission_id) not in existing_grants:
                    db.add(
                        UsuarioPermissaoDB(
                            usuario_id=admin_id,
                            permissao_id=permission_id,
                            concedida_por="SISTEMA",
                        )
                    )
                    dirty = True

    if dirty:
        if commit:
            await db.commit()
        else:
            await db.flush()


async def assign_permissions_by_profile(
    db: AsyncSession,
    *,
    user_id: int,
    profile: str,
    granted_by: str,
    commit: bool = False,
) -> dict[str, int | str]:
    """Replace a user's fine-grained permissions with the selected profile."""
    normalized_profile = normalize_profile(profile)
    await ensure_default_permissions(db, commit=False)

    user = await db.get(UserDB, user_id)
    if not user:
        raise ValueError("Usuario nao encontrado")

    permissions_result = await db.execute(select(PermissaoDB).where(PermissaoDB.ativo == 1))
    permissions = permissions_result.scalars().all()
    if normalized_profile == "ADMIN":
        target_ids = {permission.id for permission in permissions}
    else:
        wanted = {
            (module.lower(), permission_name.lower())
            for module, names in PROFILE_PERMISSION_NAMES[normalized_profile].items()
            for permission_name in names
        }
        target_ids = {
            permission.id
            for permission in permissions
            if (str(permission.modulo or "").lower(), str(permission.nome_permissao or "").lower()) in wanted
        }

    await db.execute(delete(UsuarioPermissaoDB).where(UsuarioPermissaoDB.usuario_id == user_id))
    for permission_id in sorted(target_ids):
        db.add(
            UsuarioPermissaoDB(
                usuario_id=user_id,
                permissao_id=permission_id,
                concedida_por=granted_by,
            )
        )

    if commit:
        await db.commit()
    else:
        await db.flush()

    return {"usuario_id": user_id, "perfil": normalized_profile, "permissoes_atribuidas": len(target_ids)}
