from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path
import argparse


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from db_bootstrap import ensure_saas_schema, ensure_tenant_columns
from saas_script_common import add_runtime_args, resolve_script_db_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Aplica backfill de company_id nas tabelas legadas.")
    add_runtime_args(parser)
    args = parser.parse_args()

    app_config, db_path = resolve_script_db_path(args)
    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        company_id = ensure_saas_schema(conn)
        summary = ensure_tenant_columns(conn, company_id)
        conn.commit()

    print(f"Backfill company_id aplicado em: {db_path}")
    print(f"target: {app_config.app_kind} | env: {app_config.app_env}")
    if not summary:
        print("Nenhuma tabela operacional encontrada para backfill.")
        return 0
    for table, rows in sorted(summary.items()):
        print(f"{table}: {rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
