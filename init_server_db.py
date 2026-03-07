from __future__ import annotations

import argparse
import json
import sqlite3

from db_bootstrap import ensure_admin_user, ensure_core_schema, reset_operational_data
from database_runtime import log_startup_diagnostics
from runtime_config import apply_process_environment, ensure_runtime_files, load_app_config


def main() -> int:
    parser = argparse.ArgumentParser(description="Inicializa o banco publicado com schema limpo e usuario ADMIN.")
    parser.add_argument("--reset", action="store_true", help="Apaga os dados operacionais existentes antes de manter apenas o ADMIN.")
    parser.add_argument("--admin-pass", default="", help="Senha inicial do ADMIN.")
    args = parser.parse_args()

    config = load_app_config("server")
    apply_process_environment(config)
    ensure_runtime_files(config)

    with sqlite3.connect(config.db_path) as conn:
        ensure_core_schema(conn)
        if args.reset:
            cleared = reset_operational_data(conn)
        else:
            cleared = {}
        ensure_admin_user(conn, args.admin_pass or None)

    counts = log_startup_diagnostics(config.db_path, config)
    print(
        json.dumps(
            {
                "app_env": config.app_env,
                "db_path": config.db_path,
                "api_base_url": config.api_base_url,
                "allow_dev_data_upload": config.allow_dev_data_upload,
                "sync_enabled": config.sync_enabled,
                "source_of_truth": config.source_of_truth,
                "reset": args.reset,
                "cleared_tables": cleared,
                "counts": counts,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
