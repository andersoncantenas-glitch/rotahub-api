from __future__ import annotations

import argparse
import os
from pathlib import Path

from runtime_config import ensure_runtime_files, load_app_config


def add_runtime_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--target",
        choices=("desktop", "server"),
        default=os.environ.get("ROTA_SCRIPT_TARGET", "desktop"),
        help="Runtime alvo. Use desktop para banco local/legado e server para banco da API.",
    )
    parser.add_argument(
        "--db",
        default="",
        help="Caminho explicito do SQLite. Se omitido, usa ROTA_DB ou o runtime do target.",
    )


def resolve_script_db_path(args: argparse.Namespace) -> tuple[object, str]:
    db_override = str(args.db or os.environ.get("ROTA_DB") or "").strip()
    original_rota_db = os.environ.get("ROTA_DB")

    if args.db:
        os.environ["ROTA_DB"] = str(args.db).strip()
    elif args.target == "desktop":
        # load_app_config("desktop") ignora ROTA_DB em modo local; preservamos
        # esse valor abaixo para permitir scripts sobre bancos desktop legados.
        os.environ.pop("ROTA_DB", None)

    try:
        app_config = load_app_config(args.target)
    except RuntimeError as exc:
        msg = str(exc)
        if "banco desktop legado" in msg:
            raise RuntimeError(
                "O target 'server' recusou um banco desktop legado. "
                "Para atualizar esse banco local, rode com '--target desktop' "
                "ou informe um '--db' de servidor."
            ) from exc
        raise
    finally:
        if original_rota_db is None:
            os.environ.pop("ROTA_DB", None)
        else:
            os.environ["ROTA_DB"] = original_rota_db

    db_path = os.path.abspath(db_override or app_config.db_path)
    if args.target == "server" and Path(db_path).name.lower() == "rota_granja.db":
        raise RuntimeError(
            f"Target server recusou banco desktop legado: {db_path}. "
            "Use '--target desktop' para esse banco."
        )

    ensure_runtime_files(app_config)
    return app_config, db_path
