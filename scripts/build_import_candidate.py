from __future__ import annotations

import json
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path


SOURCE_DB = Path(r"C:\rotahub\rota_granja.db")
TARGET_DB = Path(r"C:\pdc_rota\.rotahub_runtime\desktop\development\dev-local\rota_granja.db")
OUTPUT_ROOT = Path("diagnostics")
TABLES_TO_REPLACE = [
    "usuarios",
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
    "rota_substituicoes",
    "transferencias",
    "transferencias_conversoes",
    "vendas_importadas",
]


def connect_source(db_path: Path) -> sqlite3.Connection:
    return sqlite3.connect(f"file:{db_path}?mode=ro&immutable=1", uri=True)


def connect_rw(db_path: Path) -> sqlite3.Connection:
    return sqlite3.connect(db_path)


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    cur = conn.cursor()
    cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table,),
    )
    return cur.fetchone() is not None


def get_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    return [str(row[1]) for row in cur.fetchall()]


def get_count(conn: sqlite3.Connection, table: str) -> int:
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {table}")
    return int(cur.fetchone()[0] or 0)


def reset_identity(conn: sqlite3.Connection, table: str) -> None:
    cur = conn.cursor()
    cur.execute("DELETE FROM sqlite_sequence WHERE name=?", (table,))


def copy_table(source: sqlite3.Connection, target: sqlite3.Connection, table: str) -> dict[str, object]:
    if not table_exists(source, table) or not table_exists(target, table):
        return {"status": "skipped_missing_table"}

    source_columns = get_columns(source, table)
    target_columns = get_columns(target, table)
    shared_columns = [column for column in source_columns if column in target_columns]

    if not shared_columns:
        return {"status": "skipped_no_shared_columns"}

    src_cur = source.cursor()
    dst_cur = target.cursor()
    before_count = get_count(target, table)
    source_count = get_count(source, table)

    quoted_columns = ", ".join(f'"{column}"' for column in shared_columns)
    placeholders = ", ".join("?" for _ in shared_columns)

    src_cur.execute(f"SELECT {quoted_columns} FROM {table}")
    rows = src_cur.fetchall()

    dst_cur.execute(f'DELETE FROM "{table}"')
    try:
        reset_identity(target, table)
    except Exception:
        pass
    if rows:
        dst_cur.executemany(
            f'INSERT INTO "{table}" ({quoted_columns}) VALUES ({placeholders})',
            rows,
        )

    after_count = get_count(target, table)
    return {
        "status": "replaced",
        "source_count": source_count,
        "target_before": before_count,
        "candidate_after": after_count,
        "shared_columns": shared_columns,
    }


def main() -> int:
    if not SOURCE_DB.exists():
        raise SystemExit(f"Banco de origem nao encontrado: {SOURCE_DB}")
    if not TARGET_DB.exists():
        raise SystemExit(f"Banco de destino nao encontrado: {TARGET_DB}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = OUTPUT_ROOT / f"merge_candidate_{timestamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    candidate_db = out_dir / "rota_granja_merged_candidate.db"
    shutil.copy2(TARGET_DB, candidate_db)

    source = connect_source(SOURCE_DB)
    candidate = connect_rw(candidate_db)
    report: dict[str, object] = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_db": str(SOURCE_DB),
        "target_db": str(TARGET_DB),
        "candidate_db": str(candidate_db.resolve()),
        "tables": {},
    }

    try:
        candidate.execute("PRAGMA foreign_keys=OFF")
        candidate.execute("BEGIN")
        for table in TABLES_TO_REPLACE:
            report["tables"][table] = copy_table(source, candidate, table)
        candidate.commit()
    except Exception:
        candidate.rollback()
        raise
    finally:
        source.close()
        candidate.close()

    report_path = out_dir / "merge_report.json"
    with report_path.open("w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)

    print(f"Base candidata gerada em: {candidate_db.resolve()}")
    print(f"Relatorio gerado em: {report_path.resolve()}")
    for table, metadata in report["tables"].items():
        if metadata.get("status") == "replaced":
            print(
                f"{table}: origem={metadata['source_count']} "
                f"destino_atual={metadata['target_before']} candidato={metadata['candidate_after']}"
            )
        else:
            print(f"{table}: {metadata['status']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
