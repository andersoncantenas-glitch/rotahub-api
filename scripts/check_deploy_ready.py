from __future__ import annotations

import importlib.util
import os
import sqlite3
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SECRETS = {
    "",
    "your-secret-key-change-in-production",
    "your-super-secret-key-change-in-production",
    "your-jwt-secret-key-change-in-production",
    "change-this-mobile-sync-secret",
    "troque-por-um-segredo-longo",
    "troque-por-outro-segredo-longo",
    "troque-por-segredo-do-app-motorista",
}


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def status(ok: bool, message: str) -> bool:
    print(f"{'OK' if ok else 'ERRO'} - {message}")
    return ok


def warn(message: str) -> None:
    print(f"AVISO - {message}")


def writable_directory(path: Path, label: str) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        test_file = path / ".write_test"
        test_file.write_text("ok", encoding="utf-8")
        test_file.unlink(missing_ok=True)
        return status(True, f"{label} gravavel: {path}")
    except Exception as exc:
        return status(False, f"{label} sem gravacao: {exc}")


def sqlite_path_from_env() -> Path | None:
    rota_db = os.getenv("ROTA_DB", "").strip()
    if rota_db:
        return Path(rota_db).expanduser()
    database_url = os.getenv("DATABASE_URL", "").strip()
    if database_url.startswith("sqlite"):
        raw = database_url.split("///", 1)[-1]
        return Path(raw).expanduser()
    return None


def main() -> int:
    load_env_file(ROOT / ".env")
    env = os.getenv("ENVIRONMENT", "development").strip().lower()
    production = env in {"prod", "production"}
    use_external_env = os.getenv("ROTA_USE_EXTERNAL_ENV") == "1"
    if not production and not use_external_env:
        os.environ["ROTA_DB"] = str(ROOT / "rotadb.db")
        os.environ["DATABASE_URL"] = f"sqlite+aiosqlite:///{(ROOT / 'rotadb.db').as_posix()}"
        os.environ.setdefault("ROTA_MOBILE_PHOTOS_DIR", str(ROOT / ".rotahub_runtime" / "fotos_rotas"))
        os.environ.setdefault("ROTA_ENABLE_LEGACY_MOBILE_API", "1")
    ok = True

    ok &= status((ROOT / "backend" / "main.py").exists(), "backend/main.py encontrado")
    ok &= status((ROOT / "backend" / "web" / "index.html").exists(), "frontend web encontrado")
    ok &= status((ROOT / "backend" / "web_public" / "index.html").exists(), "pagina publica encontrada")
    ok &= status((ROOT / "backend" / "web_owner" / "index.html").exists(), "painel Owner SaaS encontrado")
    ok &= status((ROOT / "backend" / "requirements.txt").exists(), "backend/requirements.txt encontrado")
    ok &= status((ROOT / "scripts" / "server_start.py").exists(), "startup portatil encontrado")
    ok &= status((ROOT / "scripts" / "export_runtime_backup.py").exists(), "exportador de migracao encontrado")

    for relative in ("backend/web/app.js", "backend/web_public/public.js", "backend/web_owner/owner.js"):
        js_path = ROOT / relative
        if js_path.exists():
            try:
                subprocess.run(["node", "--check", str(js_path)], check=True, capture_output=True, text=True)
                ok &= status(True, f"JavaScript valido: {relative}")
            except FileNotFoundError:
                warn("Node.js nao encontrado; validacao de JavaScript ignorada.")
                break
            except subprocess.CalledProcessError as exc:
                detail = (exc.stderr or exc.stdout or "").strip().splitlines()
                ok &= status(False, f"JavaScript invalido em {relative}: {detail[-1] if detail else exc}")

    for module_name in ("fastapi", "uvicorn", "sqlalchemy", "pydantic_settings", "jose"):
        ok &= status(importlib.util.find_spec(module_name) is not None, f"dependencia Python disponivel: {module_name}")

    secret_key = os.getenv("SECRET_KEY", "")
    jwt_secret = os.getenv("JWT_SECRET_KEY", "")
    rota_secret = os.getenv("ROTA_SECRET", "")
    if production:
        ok &= status(secret_key not in DEFAULT_SECRETS and len(secret_key) >= 32, "SECRET_KEY de producao configurado")
        ok &= status(jwt_secret not in DEFAULT_SECRETS and len(jwt_secret) >= 32, "JWT_SECRET_KEY de producao configurado")
        ok &= status(rota_secret not in DEFAULT_SECRETS and len(rota_secret) >= 24, "ROTA_SECRET do app motorista configurado")
    else:
        warn("ENVIRONMENT nao esta como production; validacao de segredos ficou em modo aviso.")

    allowed_hosts = [item.strip() for item in os.getenv("ALLOWED_HOSTS", "").split(",") if item.strip()]
    cors_origins = [item.strip() for item in os.getenv("CORS_ORIGINS", "").split(",") if item.strip()]
    if production:
        ok &= status(bool(allowed_hosts), "ALLOWED_HOSTS configurado")
        ok &= status(bool(cors_origins), "CORS_ORIGINS configurado")
        ok &= status("*" not in allowed_hosts, "ALLOWED_HOSTS sem wildcard em producao")
        ok &= status(all(origin.startswith("https://") for origin in cors_origins), "CORS_ORIGINS usa HTTPS em producao")
    else:
        if not allowed_hosts:
            warn("ALLOWED_HOSTS nao definido no ambiente atual.")
        if not cors_origins:
            warn("CORS_ORIGINS nao definido no ambiente atual.")

    db_path = sqlite_path_from_env()
    if db_path:
        if not db_path.is_absolute():
            db_path = (ROOT / db_path).resolve()
        ok &= status(db_path.parent.exists(), f"diretorio do SQLite existe: {db_path.parent}")
        try:
            db_path.parent.mkdir(parents=True, exist_ok=True)
            with sqlite3.connect(db_path) as conn:
                conn.execute("PRAGMA schema_version")
            ok &= status(True, f"SQLite abre corretamente: {db_path}")
        except Exception as exc:
            ok &= status(False, f"SQLite nao abriu: {exc}")
        if production and str(db_path).startswith(str(ROOT)):
            warn("ROTA_DB aponta para dentro do projeto; em VPS prefira /var/rotahub/data ou volume persistente.")
    elif os.getenv("DATABASE_URL", "").startswith("postgresql"):
        ok &= status(True, "DATABASE_URL PostgreSQL configurado")
    else:
        ok &= status(False, "ROTA_DB ou DATABASE_URL precisa estar configurado")

    photos_dir = Path(os.getenv("ROTA_MOBILE_PHOTOS_DIR", ROOT / ".rotahub_runtime" / "fotos_rotas")).expanduser()
    if not photos_dir.is_absolute():
        photos_dir = (ROOT / photos_dir).resolve()
    ok &= writable_directory(photos_dir, "diretorio de fotos")

    backup_dir = Path(os.getenv("BACKUP_DIR", ROOT / "backup")).expanduser()
    if not backup_dir.is_absolute():
        backup_dir = (ROOT / backup_dir).resolve()
    ok &= writable_directory(backup_dir, "diretorio de backups")

    export_dir = Path(os.getenv("ROTA_EXPORT_DIR", ROOT / "exports")).expanduser()
    if not export_dir.is_absolute():
        export_dir = (ROOT / export_dir).resolve()
    ok &= writable_directory(export_dir, "diretorio de exportacao")

    legacy = os.getenv("ROTA_ENABLE_LEGACY_MOBILE_API", "0").strip().lower() in {"1", "true", "yes", "on"}
    if production:
        ok &= status(legacy, "compatibilidade com app motorista habilitada")
    elif not legacy:
        warn("ROTA_ENABLE_LEGACY_MOBILE_API nao esta habilitado no ambiente atual.")

    print("\nResultado:", "PRONTO PARA DEPLOY" if ok else "AJUSTES NECESSARIOS ANTES DO DEPLOY")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
