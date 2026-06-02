from __future__ import annotations

import os
import sys
from pathlib import Path

import uvicorn


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    host = str(os.getenv("HOST") or "0.0.0.0").strip()
    port = int(os.getenv("PORT") or "8000")
    data_root = str(os.getenv("ROTA_DATA_ROOT") or "").strip()
    rota_db = str(os.getenv("ROTA_DB") or os.getenv("DATABASE_URL") or "").strip()
    photos_dir = str(os.getenv("ROTA_MOBILE_PHOTOS_DIR") or "").strip()
    backup_dir = str(os.getenv("BACKUP_DIR") or "").strip()

    print(f"RotaHub server start | host={host} | port={port}", flush=True)
    print(f"RotaHub server start | python={sys.version.split()[0]}", flush=True)
    print(f"RotaHub server start | data_root={data_root}", flush=True)
    print(f"RotaHub server start | rota_db={rota_db}", flush=True)
    print(f"RotaHub server start | photos_dir={photos_dir}", flush=True)
    print(f"RotaHub server start | backup_dir={backup_dir}", flush=True)

    uvicorn.run(
        "backend.main:app",
        host=host,
        port=port,
        log_level=str(os.getenv("UVICORN_LOG_LEVEL") or "info").strip().lower(),
        proxy_headers=True,
        forwarded_allow_ips=str(os.getenv("FORWARDED_ALLOW_IPS") or "*").strip(),
    )


if __name__ == "__main__":
    main()
