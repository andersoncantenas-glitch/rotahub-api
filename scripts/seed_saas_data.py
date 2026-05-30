from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path
import argparse


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from db_bootstrap import ensure_admin_user, ensure_core_schema, ensure_permission_system, ensure_saas_schema
from saas_script_common import add_runtime_args, resolve_script_db_path


def _count(cur: sqlite3.Cursor, table: str) -> int:
    cur.execute(f"SELECT COUNT(*) FROM {table}")
    row = cur.fetchone()
    return int(row[0] if row else 0)


def main() -> int:
    parser = argparse.ArgumentParser(description="Aplica seed das tabelas SaaS.")
    add_runtime_args(parser)
    args = parser.parse_args()

    app_config, db_path = resolve_script_db_path(args)
    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        ensure_core_schema(conn)
        ensure_saas_schema(conn)
        ensure_admin_user(conn, os.environ.get("ROTA_ADMIN_PASS") or os.environ.get("ROTA_ADMIN_PASSWORD"))
        ensure_permission_system(conn)
        cur = conn.cursor()
        summary = {
            "companies": _count(cur, "companies"),
            "plans": _count(cur, "plans"),
            "subscriptions": _count(cur, "subscriptions"),
            "payments": _count(cur, "payments"),
            "audit_logs": _count(cur, "audit_logs"),
        }

    print(f"SaaS seed aplicado em: {db_path}")
    print(f"target: {app_config.app_kind} | env: {app_config.app_env}")
    for key, value in summary.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
