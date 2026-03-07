from __future__ import annotations

import csv
import json
import os
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path


SOURCE_DB = Path(r"C:\rotahub\rota_granja.db")
OUTPUT_ROOT = Path("diagnostics")
SYSTEM_TABLES = {"sqlite_sequence"}


def _connect_readonly(db_path: Path) -> sqlite3.Connection:
    return sqlite3.connect(f"file:{db_path}?mode=ro&immutable=1", uri=True)


def _fetch_table_names(conn: sqlite3.Connection) -> list[str]:
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    return [str(row[0]) for row in cur.fetchall() if str(row[0]) not in SYSTEM_TABLES]


def _fetch_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    return [str(row[1]) for row in cur.fetchall()]


def _fetch_rows(conn: sqlite3.Connection, table: str) -> list[tuple]:
    cur = conn.cursor()
    cur.execute(f"SELECT * FROM {table}")
    return cur.fetchall()


def _write_csv(path: Path, headers: list[str], rows: list[tuple]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        writer.writerows(rows)


def main() -> int:
    if not SOURCE_DB.exists():
        raise SystemExit(f"Banco nao encontrado: {SOURCE_DB}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = OUTPUT_ROOT / f"installed_db_snapshot_{timestamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    snapshot_db = out_dir / SOURCE_DB.name
    shutil.copy2(SOURCE_DB, snapshot_db)

    conn = _connect_readonly(SOURCE_DB)
    try:
        table_names = _fetch_table_names(conn)
        summary: dict[str, object] = {
            "source_db": str(SOURCE_DB),
            "snapshot_db": str(snapshot_db.resolve()),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "tables": {},
        }

        for table in table_names:
            headers = _fetch_columns(conn, table)
            rows = _fetch_rows(conn, table)
            csv_path = out_dir / f"{table}.csv"
            _write_csv(csv_path, headers, rows)
            summary["tables"][table] = {
                "row_count": len(rows),
                "columns": headers,
                "csv": csv_path.name,
            }

        summary_path = out_dir / "summary.json"
        with summary_path.open("w", encoding="utf-8") as fh:
            json.dump(summary, fh, ensure_ascii=False, indent=2)

        print(f"Snapshot gerado em: {out_dir.resolve()}")
        print(f"Banco copiado para: {snapshot_db.resolve()}")
        for table, metadata in summary["tables"].items():
            print(f"{table}: {metadata['row_count']} linhas")
    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
