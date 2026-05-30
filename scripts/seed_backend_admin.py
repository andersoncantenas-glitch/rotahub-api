from __future__ import annotations

import argparse
import asyncio
import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Cria ou atualiza o usuario admin da API FastAPI.")
    parser.add_argument("--database-url", help="DATABASE_URL alvo. Se omitido, usa o ambiente/.env.")
    parser.add_argument(
        "--username",
        default=os.environ.get("ROTA_BACKEND_ADMIN_USERNAME", "admin"),
        help="Login do usuario admin.",
    )
    parser.add_argument(
        "--password",
        default=os.environ.get("ROTA_BACKEND_ADMIN_PASSWORD"),
        help="Senha do usuario admin. Tambem pode usar ROTA_BACKEND_ADMIN_PASSWORD.",
    )
    parser.add_argument(
        "--nome",
        default=os.environ.get("ROTA_BACKEND_ADMIN_NOME", "ADMIN"),
        help="Nome exibido para o usuario admin.",
    )
    parser.add_argument(
        "--permissoes",
        default=os.environ.get("ROTA_BACKEND_ADMIN_PERMISSOES", "ADMIN"),
        help="Perfil/permissoes do usuario admin.",
    )
    return parser


async def _seed_admin(args: argparse.Namespace) -> tuple[str, int]:
    if args.database_url:
        os.environ["DATABASE_URL"] = args.database_url

    environment = os.environ.get("ENVIRONMENT", "development").strip().lower()
    password = args.password
    used_default_password = False
    if not password:
        if environment in {"prod", "production"}:
            raise RuntimeError("Informe --password ou ROTA_BACKEND_ADMIN_PASSWORD em producao.")
        password = "Admin@123456"
        used_default_password = True

    from sqlalchemy import select

    from backend.config.database import async_session, create_tables
    from backend.models.user import UserDB
    from backend.services.auth import get_password_hash

    await create_tables()

    async with async_session() as session:
        result = await session.execute(select(UserDB).where(UserDB.username == args.username))
        user = result.scalar_one_or_none()
        password_hash = get_password_hash(password)

        if user is None:
            user = UserDB(
                username=args.username,
                nome=args.nome,
                senha=password_hash,
                permissoes=args.permissoes,
            )
            session.add(user)
            action = "created"
        else:
            user.nome = args.nome
            user.senha = password_hash
            user.permissoes = args.permissoes
            action = "updated"

        await session.commit()
        await session.refresh(user)

    if used_default_password:
        print("Senha padrao aplicada para desenvolvimento: Admin@123456")

    return action, int(user.id)


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    action, user_id = asyncio.run(_seed_admin(args))
    print(f"Backend admin {action}: id={user_id} username={args.username} permissoes={args.permissoes}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
