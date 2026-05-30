import os
import sys
from pathlib import Path

import uvicorn

PROJECT_ROOT = Path(__file__).resolve().parent.parent
os.chdir(PROJECT_ROOT)
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

use_external_env = os.getenv("ROTA_USE_EXTERNAL_ENV") == "1"

if not os.getenv("ROTA_DB") or not use_external_env:
    os.environ["ROTA_DB"] = str(PROJECT_ROOT / "rotadb.db")
if not os.getenv("DATABASE_URL") or not use_external_env:
    db_path = Path(os.environ["ROTA_DB"]).expanduser().resolve().as_posix()
    os.environ["DATABASE_URL"] = f"sqlite+aiosqlite:///{db_path}"
if not os.getenv("ROTA_MOBILE_PHOTOS_DIR") or not use_external_env:
    os.environ["ROTA_MOBILE_PHOTOS_DIR"] = str(PROJECT_ROOT / ".rotahub_runtime" / "fotos_rotas")
os.environ.setdefault("ROTA_ENABLE_LEGACY_MOBILE_API", "1")

if __name__ == "__main__":
    uvicorn.run(
        "backend.main:app",
        host=os.getenv("HOST", "0.0.0.0"),
        port=int(os.getenv("PORT", "8000")),
        reload=False,
        log_level=os.getenv("LOG_LEVEL", "info"),
    )
