import json
import os
import sqlite3
import urllib.request
import urllib.error


def check_db(path: str):
    out = {"path": path, "exists": os.path.exists(path), "counts": {}}
    if not out["exists"]:
        return out
    tables = [
        "motoristas",
        "usuarios",
        "veiculos",
        "ajudantes",
        "clientes",
        "programacoes",
        "programacao_itens",
        "recebimentos",
        "despesas",
        "vendas_importadas",
    ]
    try:
        con = sqlite3.connect(path)
        cur = con.cursor()
        for t in tables:
            try:
                cur.execute(f"SELECT COUNT(*) FROM {t}")
                out["counts"][t] = int(cur.fetchone()[0] or 0)
            except Exception:
                pass
        con.close()
    except Exception as exc:
        out["error"] = str(exc)
    return out


def check_api(base: str, secret: str):
    out = {"base": base, "ok": False}
    if not base:
        out["error"] = "ROTA_SERVER_URL vazio"
        return out
    try:
        req = urllib.request.Request(
            f"{base.rstrip('/')}/admin/motoristas/acesso",
            headers={"X-Desktop-Secret": secret or ""},
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            payload = json.loads(resp.read().decode("utf-8", errors="replace"))
        out["ok"] = True
        out["motoristas_total_api"] = len((payload or {}).get("motoristas") or [])
        return out
    except urllib.error.HTTPError as exc:
        out["error"] = f"{exc.code} {exc.reason}"
        return out
    except Exception as exc:
        out["error"] = str(exc)
        return out


if __name__ == "__main__":
    db_path = os.environ.get("ROTA_DB", "").strip()
    api_base = os.environ.get("ROTA_SERVER_URL", "").strip()
    secret = os.environ.get("ROTA_SECRET", "").strip()
    result = {
        "env": {
            "ROTA_DB": db_path,
            "ROTA_SERVER_URL": api_base,
            "ROTA_SECRET_SET": bool(secret),
            "ROTA_DESKTOP_SYNC_API": os.environ.get("ROTA_DESKTOP_SYNC_API", ""),
        },
        "db": check_db(db_path) if db_path else {"error": "ROTA_DB nao definido"},
        "api": check_api(api_base, secret),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
