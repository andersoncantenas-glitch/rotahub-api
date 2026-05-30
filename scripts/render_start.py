from __future__ import annotations

import os
import sys
from pathlib import Path

import uvicorn


def main() -> None:
    port = int(os.getenv("PORT") or "10000")
    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    rota_db = os.getenv("ROTA_DB") or os.getenv("DATABASE_URL") or ""
    photos_dir = os.getenv("ROTA_MOBILE_PHOTOS_DIR") or ""
    print(f"RotaHub Render start | port={port}", flush=True)
    print(f"RotaHub Render start | python={sys.version.split()[0]}", flush=True)
    print(f"RotaHub Render start | rota_db={rota_db}", flush=True)
    print(f"RotaHub Render start | photos_dir={photos_dir}", flush=True)

    uvicorn.run(
        "backend.main:app",
        host="0.0.0.0",
        port=port,
        log_level="info",
        proxy_headers=True,
        forwarded_allow_ips="*",
    )


if __name__ == "__main__":
    main()
