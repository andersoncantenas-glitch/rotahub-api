from __future__ import annotations

import sqlite3
from typing import Any

from app.repositories.base_repository import ensure_saas_ready, row_to_dict, rows_to_dicts


def ensure_plan_change_requests_table(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS plan_change_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            current_plan_code TEXT,
            requested_plan_code TEXT NOT NULL,
            request_type TEXT DEFAULT 'change',
            vehicle_count INTEGER DEFAULT 0,
            vehicle_limit_current INTEGER,
            vehicle_limit_requested INTEGER,
            user_count INTEGER DEFAULT 0,
            user_limit_current INTEGER,
            user_limit_requested INTEGER,
            status TEXT DEFAULT 'pending',
            message TEXT,
            requested_by_user_id INTEGER,
            requested_by_name TEXT,
            reviewed_at TEXT,
            reviewed_by TEXT,
            review_notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute("PRAGMA table_info(plan_change_requests)")
    columns = {str(row[1]).lower() for row in cur.fetchall()}
    for column, definition in {
        "vehicle_count": "INTEGER DEFAULT 0",
        "vehicle_limit_current": "INTEGER",
        "vehicle_limit_requested": "INTEGER",
        "user_count": "INTEGER DEFAULT 0",
        "user_limit_current": "INTEGER",
        "user_limit_requested": "INTEGER",
        "reviewed_at": "TEXT",
        "reviewed_by": "TEXT",
        "review_notes": "TEXT",
        "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
    }.items():
        if column not in columns:
            cur.execute(f"ALTER TABLE plan_change_requests ADD COLUMN {column} {definition}")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_plan_change_requests_status ON plan_change_requests(status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_plan_change_requests_company ON plan_change_requests(company_id)")


def plan_change_request_select() -> str:
    return """
        SELECT r.*,
               c.name AS company_name,
               c.document AS company_document,
               cp.name AS current_plan_name,
               rp.name AS requested_plan_name,
               rp.monthly_price AS requested_plan_price
          FROM plan_change_requests r
          LEFT JOIN companies c ON c.id = r.company_id
          LEFT JOIN plans cp ON cp.code = r.current_plan_code
          LEFT JOIN plans rp ON rp.code = r.requested_plan_code
    """


def list_plan_change_requests(conn: sqlite3.Connection, *, status: str | None = None, limit: int = 500) -> list[dict]:
    ensure_saas_ready(conn)
    ensure_plan_change_requests_table(conn)
    params: list[Any] = []
    where = ""
    if status:
        where = "WHERE r.status=?"
        params.append(str(status).strip().lower())
    params.append(max(1, min(int(limit or 500), 5000)))
    cur = conn.cursor()
    cur.execute(
        f"""
        {plan_change_request_select()}
        {where}
         ORDER BY CASE r.status WHEN 'pending' THEN 0 ELSE 1 END, r.id DESC
         LIMIT ?
        """,
        tuple(params),
    )
    return rows_to_dicts(cur.fetchall())


def get_plan_change_request(conn: sqlite3.Connection, request_id: int) -> dict | None:
    ensure_saas_ready(conn)
    ensure_plan_change_requests_table(conn)
    cur = conn.cursor()
    cur.execute(f"{plan_change_request_select()} WHERE r.id=? LIMIT 1", (int(request_id),))
    return row_to_dict(cur.fetchone())


def create_plan_change_request(conn: sqlite3.Connection, payload: dict) -> dict:
    ensure_saas_ready(conn)
    ensure_plan_change_requests_table(conn)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO plan_change_requests (
            company_id, current_plan_code, requested_plan_code, request_type,
            vehicle_count, vehicle_limit_current, vehicle_limit_requested,
            user_count, user_limit_current, user_limit_requested,
            status, message, requested_by_user_id, requested_by_name,
            created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?, datetime('now'), datetime('now'))
        """,
        (
            int(payload.get("company_id") or 0),
            payload.get("current_plan_code"),
            payload.get("requested_plan_code"),
            payload.get("request_type") or "change",
            int(payload.get("vehicle_count") or 0),
            payload.get("vehicle_limit_current"),
            payload.get("vehicle_limit_requested"),
            int(payload.get("user_count") or 0),
            payload.get("user_limit_current"),
            payload.get("user_limit_requested"),
            payload.get("message") or "",
            payload.get("requested_by_user_id"),
            payload.get("requested_by_name") or "",
        ),
    )
    return get_plan_change_request(conn, int(cur.lastrowid)) or {}


def close_plan_change_request(
    conn: sqlite3.Connection,
    request_id: int,
    *,
    status: str,
    reviewed_by: str,
    review_notes: str = "",
) -> dict | None:
    ensure_saas_ready(conn)
    ensure_plan_change_requests_table(conn)
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE plan_change_requests
           SET status=?, reviewed_at=datetime('now'), reviewed_by=?,
               review_notes=?, updated_at=datetime('now')
         WHERE id=?
        """,
        (str(status).strip().lower(), reviewed_by, review_notes or "", int(request_id)),
    )
    return get_plan_change_request(conn, request_id)
