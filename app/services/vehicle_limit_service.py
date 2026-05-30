from __future__ import annotations

import json
import sqlite3
from typing import Any

from app.services.saas_result import service_result


def check_vehicle_limit(
    conn: sqlite3.Connection,
    company_id: int,
    *,
    exclude_placa: str | None = None,
) -> dict:
    try:
        snapshot = vehicle_usage_snapshot(conn, company_id, exclude_placa=exclude_placa)
        limit = snapshot.get("vehicle_limit")
        if limit is None:
            return service_result(ok=True, data=snapshot)
        if int(snapshot["vehicle_count"]) >= int(limit):
            plan_name = str(snapshot.get("plan_name") or "Plano atual")
            error = f"Limite de {int(limit)} veiculos atingido. Plano {plan_name}. Faca upgrade."
            _audit_vehicle_limit_attempt(conn, company_id, snapshot, exclude_placa, error)
            return service_result(ok=False, data=snapshot, error=error)
        return service_result(ok=True, data=snapshot)
    except Exception as exc:
        return service_result(ok=False, data=None, error=str(exc) or "Falha ao validar limite de veiculos.")


def vehicle_usage_snapshot(conn: sqlite3.Connection, company_id: int, *, exclude_placa: str | None = None) -> dict:
    cur = conn.cursor()
    plan = _active_plan(cur, company_id)
    vehicle_count = _vehicle_count(cur, company_id, exclude_placa=exclude_placa)
    return {
        "company_id": int(company_id),
        "plan_code": plan.get("code"),
        "plan_name": plan.get("name"),
        "vehicle_limit": plan.get("vehicle_limit"),
        "vehicle_count": vehicle_count,
    }


def _active_plan(cur: sqlite3.Cursor, company_id: int) -> dict:
    cur.execute(
        """
        SELECT p.code, p.name, p.vehicle_limit
        FROM subscriptions s
        JOIN plans p ON p.id = s.plan_id
        WHERE s.company_id=? AND s.status IN ('active', 'trialing', 'past_due')
        ORDER BY s.id DESC
        LIMIT 1
        """,
        (int(company_id),),
    )
    row = cur.fetchone()
    if not row:
        return {"code": None, "name": "Sem plano", "vehicle_limit": 0}
    return {
        "code": _row_get(row, "code", 0),
        "name": _row_get(row, "name", 1),
        "vehicle_limit": _row_get(row, "vehicle_limit", 2),
    }


def _vehicle_count(cur: sqlite3.Cursor, company_id: int, *, exclude_placa: str | None = None) -> int:
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='veiculos'")
    if not cur.fetchone():
        return 0
    cur.execute("PRAGMA table_info(veiculos)")
    cols = {str(row[1]).lower() for row in (cur.fetchall() or [])}
    clauses = []
    params: list[Any] = []
    if "company_id" in cols:
        clauses.append("company_id=?")
        params.append(int(company_id))
    placa = str(exclude_placa or "").strip().upper()
    if placa and "placa" in cols:
        clauses.append("UPPER(TRIM(COALESCE(placa,'')))<>?")
        params.append(placa)
    where = "WHERE " + " AND ".join(clauses) if clauses else ""
    cur.execute(f"SELECT COUNT(*) FROM veiculos {where}", tuple(params))
    row = cur.fetchone()
    return int(row[0] if row else 0)


def _audit_vehicle_limit_attempt(
    conn: sqlite3.Connection,
    company_id: int,
    snapshot: dict,
    placa: str | None,
    error: str,
) -> None:
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='audit_logs'")
        if not cur.fetchone():
            return
        cur.execute(
            """
            INSERT INTO audit_logs (
                company_id, actor_type, action, entity_type, entity_id,
                severity, metadata_json, created_at
            )
            VALUES (?, 'system', 'tentativa_exceder_limite_veiculo', 'vehicle', ?, 'warning', ?, datetime('now'))
            """,
            (
                int(company_id),
                str(placa or "").strip().upper(),
                json.dumps(
                    {
                        "error": error,
                        "plan_code": snapshot.get("plan_code"),
                        "plan_name": snapshot.get("plan_name"),
                        "vehicle_count": snapshot.get("vehicle_count"),
                        "vehicle_limit": snapshot.get("vehicle_limit"),
                    },
                    ensure_ascii=True,
                    sort_keys=True,
                ),
            ),
        )
    except Exception:
        pass


def _row_get(row: Any, key: str, index: int):
    try:
        return row[key]
    except Exception:
        return row[index]
