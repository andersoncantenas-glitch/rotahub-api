from __future__ import annotations

from app.repositories.base_repository import ensure_saas_ready, get_db
from app.services.saas_result import error_message, service_result
from app.services.vehicle_limit_service import vehicle_usage_snapshot


def get_company_usage(company_id: int) -> dict:
    try:
        with get_db() as conn:
            ensure_saas_ready(conn)
            vehicle_snapshot = vehicle_usage_snapshot(conn, int(company_id))
            cur = conn.cursor()
            users = _count_table(cur, "usuarios", int(company_id))
            motoristas = _count_table(cur, "motoristas", int(company_id))
            vendedores = _count_table(cur, "vendedores", int(company_id))
            programacoes = _count_table(cur, "programacoes", int(company_id))
            clientes = _count_table(cur, "clientes", int(company_id))
        return service_result(
            ok=True,
            data={
                "company_id": int(company_id),
                "vehicles": vehicle_snapshot,
                "users": users,
                "motoristas": motoristas,
                "vendedores": vendedores,
                "programacoes": programacoes,
                "clientes": clientes,
            },
        )
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao consultar uso da empresa."))


def _count_table(cur, table: str, company_id: int) -> int:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
    if not cur.fetchone():
        return 0
    cur.execute(f"PRAGMA table_info({table})")
    cols = {str(row[1]).lower() for row in (cur.fetchall() or [])}
    if "company_id" in cols:
        cur.execute(f'SELECT COUNT(*) FROM "{table}" WHERE company_id=?', (int(company_id),))
    else:
        cur.execute(f'SELECT COUNT(*) FROM "{table}"')
    row = cur.fetchone()
    return int(row[0] if row else 0)
