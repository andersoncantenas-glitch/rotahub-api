from __future__ import annotations

import argparse
import json
import os
import shutil
import sqlite3
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _env_path(name: str, default: Path) -> Path:
    value = str(os.getenv(name) or "").strip()
    return Path(value).expanduser().resolve() if value else default.resolve()


def _sqlite_backup(src: Path, dst: Path) -> None:
    if not src.exists():
        raise FileNotFoundError(f"Banco SQLite nao encontrado: {src}")
    src_conn = sqlite3.connect(str(src))
    dst_conn = sqlite3.connect(str(dst))
    try:
        with dst_conn:
            src_conn.backup(dst_conn)
        result = dst_conn.execute("PRAGMA integrity_check").fetchone()
        if not result or str(result[0]).lower() != "ok":
            raise RuntimeError(f"Falha na integridade do backup SQLite: {result}")
    finally:
        dst_conn.close()
        src_conn.close()


def _add_directory(archive: zipfile.ZipFile, source: Path, prefix: str) -> int:
    if not source.exists():
        return 0
    count = 0
    for path in sorted(source.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(source)
        archive.write(path, Path(prefix) / relative)
        count += 1
    return count


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Exporta banco SQLite e fotos do RotaHub em um pacote portatil.",
    )
    parser.add_argument("--db", help="Caminho do SQLite. Padrao: ROTA_DB.")
    parser.add_argument("--photos-dir", help="Pasta das fotos. Padrao: ROTA_MOBILE_PHOTOS_DIR.")
    parser.add_argument("--output-dir", help="Pasta de destino. Padrao: ROTA_EXPORT_DIR ou ./exports.")
    args = parser.parse_args()

    db_path = Path(args.db).expanduser().resolve() if args.db else _env_path("ROTA_DB", PROJECT_ROOT / "rotadb.db")
    photos_dir = (
        Path(args.photos_dir).expanduser().resolve()
        if args.photos_dir
        else _env_path("ROTA_MOBILE_PHOTOS_DIR", PROJECT_ROOT / ".rotahub_runtime" / "fotos_rotas")
    )
    output_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else _env_path("ROTA_EXPORT_DIR", PROJECT_ROOT / "exports")
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_path = output_dir / f"rotahub_migration_{timestamp}.zip"

    with tempfile.TemporaryDirectory(prefix="rotahub_export_") as temp_dir:
        temp_root = Path(temp_dir)
        backup_db = temp_root / "rotadb.db"
        _sqlite_backup(db_path, backup_db)

        manifest = {
            "format": "rotahub-migration-v1",
            "created_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "database": {
                "type": "sqlite",
                "archive_path": "database/rotadb.db",
                "source_name": db_path.name,
                "size_bytes": backup_db.stat().st_size,
            },
            "photos": {
                "archive_path": "fotos_rotas/",
                "included": photos_dir.exists(),
                "files": 0,
            },
            "restore": {
                "rota_db": "/var/rotahub/data/rotadb.db",
                "photos_dir": "/var/rotahub/data/fotos_rotas",
            },
        }

        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(backup_db, "database/rotadb.db")
            manifest["photos"]["files"] = _add_directory(archive, photos_dir, "fotos_rotas")
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))

    print(f"Pacote criado: {archive_path}")
    print(f"Banco: {db_path}")
    print(f"Fotos incluidas: {manifest['photos']['files']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
