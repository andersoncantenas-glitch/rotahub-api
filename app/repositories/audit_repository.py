from __future__ import annotations

import json

from app.repositories.base_repository import ensure_saas_ready, get_db, normalize_limit, row_to_dict, rows_to_dicts


def create_audit_log(data: dict) -> dict:
    metadata = data.get("metadata")
    if metadata is None:
        metadata_json = data.get("metadata_json") or "{}"
    elif isinstance(metadata, str):
        metadata_json = metadata
    else:
        metadata_json = json.dumps(metadata, ensure_ascii=True, sort_keys=True)
    payload = {
        "company_id": data.get("company_id"),
        "user_id": data.get("user_id"),
        "actor_type": data.get("actor_type") or "system",
        "action": data.get("action"),
        "entity_type": data.get("entity_type"),
        "entity_id": data.get("entity_id"),
        "severity": data.get("severity") or "info",
        "ip_address": data.get("ip_address"),
        "metadata_json": metadata_json,
    }
    if not str(payload.get("action") or "").strip():
        raise ValueError("Acao do audit log e obrigatoria.")
    fields = list(payload.keys())
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            INSERT INTO audit_logs ({", ".join(fields)}, created_at)
            VALUES ({", ".join(["?"] * len(fields))}, datetime('now'))
            """,
            tuple(payload[field] for field in fields),
        )
        audit_id = int(cur.lastrowid)
        cur.execute("SELECT * FROM audit_logs WHERE id=? LIMIT 1", (audit_id,))
        return row_to_dict(cur.fetchone()) or {}


def list_audit_logs(company_id: int | None = None, action: str | None = None, limit: int | None = 500) -> list[dict]:
    clauses: list[str] = []
    params: list[object] = []
    if company_id:
        clauses.append("company_id=?")
        params.append(int(company_id))
    if action:
        clauses.append("action=?")
        params.append(str(action).strip())
    where = "WHERE " + " AND ".join(clauses) if clauses else ""
    params.append(normalize_limit(limit))
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT *
            FROM audit_logs
            {where}
            ORDER BY id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return rows_to_dicts(cur.fetchall())
