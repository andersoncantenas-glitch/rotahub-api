import argparse
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).with_name("rota_granja.db")

TABLES_TO_CLEAR = [
    "motoristas",
    "usuarios",
    "veiculos",
    "equipes",
    "clientes",
    "vendas_importadas",
    "programacoes",
    "programacao_itens",
    "programacao_itens_controle",
    "programacao_itens_log",
    "recebimentos",
    "despesas",
]


def confirm() -> bool:
    resp = input(
        "Este script vai apagar todos os dados principais do banco (programações, entregas, equipes etc.).\n"
        "Certifique-se de ter um backup antes de continuar.\n"
        "Deseja continuar? [s/N]: "
    )
    return resp.strip().lower() in {"s", "sim", "y", "yes"}


def clear_tables(conn: sqlite3.Connection):
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys=OFF")
    try:
        for table in TABLES_TO_CLEAR:
            cur.execute(f"DELETE FROM {table}")
            cur.execute("DELETE FROM sqlite_sequence WHERE name=?", (table,))
    finally:
        cur.execute("PRAGMA foreign_keys=ON")


def backup_database(path: Path) -> Path:
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Banco nao encontrado em {DB_PATH}")

    backup_path = path or Path(f"rota_granja_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    shutil.copy2(DB_PATH, backup_path)
    return backup_path


def parse_args():
    parser = argparse.ArgumentParser(description="Limpa dados de produção do banco rota_granja.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Mostra o que seria limpo sem executar alterações.",
    )
    parser.add_argument(
        "--backup",
        type=Path,
        help="Caminho para salvar o backup antes da limpeza. Se omitido, gera rota_granja_backup_YYYYmmdd_HHMMSS.db no diretório atual.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.dry_run:
        print("Dry run: as seguintes tabelas seriam limpas:")
        for table in TABLES_TO_CLEAR:
            print(f"  - {table}")
        print("Nenhuma alteração foi feita.")
        raise SystemExit(0)

    if not confirm():
        print("Operacao cancelada.")
        raise SystemExit(0)

    backup_path = args.backup or Path(f"rota_granja_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")

    with sqlite3.connect(DB_PATH) as conn:
        try:
            clear_tables(conn)
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            conn.execute("VACUUM")
        except Exception as exc:
            raise RuntimeError("Falha ao limpar o banco") from exc

    print("Banco limpo. Reinicie o sistema para recarregar a aplicacao.")
