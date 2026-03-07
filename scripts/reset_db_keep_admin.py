from __future__ import annotations

import argparse
import json
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path


OPERATIONAL_TABLES = [
    "motoristas",
    "veiculos",
    "ajudantes",
    "clientes",
    "equipes",
    "programacoes",
    "programacao_itens",
    "programacao_itens_controle",
    "programacao_itens_log",
    "recebimentos",
    "despesas",
    "rota_gps_override_log",
    "rota_gps_pings",
    "rota_substituicoes",
    "transferencias",
    "transferencias_conversoes",
    "vendas_importadas",
    "programacoes_avulsas",
    "programacoes_avulsas_itens",
    "mobile_sync_idempotency",
]


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,))
    return cur.fetchone() is not None


def count_rows(conn: sqlite3.Connection, table: str) -> int:
    cur = conn.cursor()
    cur.execute(f'SELECT COUNT(*) FROM "{table}"')
    return int(cur.fetchone()[0] or 0)


def main() -> int:
    parser = argparse.ArgumentParser(description="Zera dados operacionais preservando somente o usuario ADMIN.")
    parser.add_argument("--db", required=True, help="Caminho do banco SQLite a ser resetado")
    parser.add_argument("--backup-dir", default="diagnostics/reset_backups", help="Diretorio para backup")
    args = parser.parse_args()

    db_path = Path(args.db).expanduser().resolve()
    if not db_path.exists():
        raise SystemExit(f"Banco nao encontrado: {db_path}")

    backup_dir = Path(args.backup_dir).expanduser().resolve()
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir / f"{db_path.stem}_before_reset_{datetime.now().strftime('%Y%m%d_%H%M%S')}{db_path.suffix}"
    shutil.copy2(db_path, backup_path)

    conn = sqlite3.connect(db_path)
    report: dict[str, object] = {
        "db_path": str(db_path),
        "backup_path": str(backup_path),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "tables": {},
        "usuarios": {},
    }
    try:
        conn.execute("PRAGMA foreign_keys=OFF")
        conn.execute("BEGIN")
        for table in OPERATIONAL_TABLES:
            if not table_exists(conn, table):
                report["tables"][table] = {"status": "missing"}
                continue
            before = count_rows(conn, table)
            conn.execute(f'DELETE FROM "{table}"')
            report["tables"][table] = {"status": "cleared", "before": before, "after": 0}

        if table_exists(conn, "usuarios"):
            before_users = count_rows(conn, "usuarios")
            conn.execute("DELETE FROM usuarios WHERE UPPER(COALESCE(nome,'')) <> 'ADMIN'")
            after_users = count_rows(conn, "usuarios")
            report["usuarios"] = {"before": before_users, "after": after_users}
        else:
            report["usuarios"] = {"status": "missing"}

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
