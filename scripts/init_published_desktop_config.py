from __future__ import annotations

import argparse
import json
import os
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Gera o config.json correto do desktop publicado, isolado do development.")
    parser.add_argument("--api-url", default="https://rotahub-api.onrender.com")
    parser.add_argument("--tenant-id", default="default-company")
    parser.add_argument("--company-id", default="default-company")
    parser.add_argument("--config-path", default=str(Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "RotaHubDesktop" / "config.json"))
    args = parser.parse_args()

    config_path = Path(args.config_path).expanduser().resolve()
    data_root = config_path.parent
    db_path = data_root / "desktop" / "production" / args.tenant_id / "rotahub_desktop.db"
    db_path.parent.mkdir(parents=True, exist_ok=True)

    config = {
        "app_env": "production",
        "tenant": {
            "tenant_id": args.tenant_id,
            "company_id": args.company_id,
        },
        "runtime": {
            "data_root": str(data_root),
            "db_path": str(db_path),
            "sync_enabled": True,
            "sql_mirror_api": True,
            "require_server_binding": True,
            "allow_remote_write": True,
            "allow_remote_read": True,
            "allow_dev_data_upload": False,
            "allow_seed_db": False,
            "allow_version_update": True,
            "tenant_mode": "database-per-tenant",
            "source_of_truth": "api-central",
            "schema_version": 1,
            "desktop_secret": "",
        },
        "api": {
            "base_url": args.api_url.rstrip("/"),
            "timeout": 60.0,
        },
        "update": {
            "channel": "stable",
            "manifest_url": "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/version.json",
            "setup_url": "https://github.com/andersoncantenas-glitch/rotahub-api/releases/latest",
            "changelog_url": "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/changelog.txt",
        },
        "logging": {
            "level": "INFO",
        },
        "support": {
            "whatsapp": "",
            "email": "",
        },
    }
    config_path.parent.mkdir(parents=True, exist_ok=True)
    with config_path.open("w", encoding="utf-8") as fh:
        json.dump(config, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    print(config_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
