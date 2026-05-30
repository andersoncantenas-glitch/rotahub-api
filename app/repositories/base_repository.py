from __future__ import annotations

from typing import Any, Iterable

from app.db.connection import get_db
from db_bootstrap import ensure_saas_schema, ensure_tenant_columns


def ensure_saas_ready(conn) -> None:
    if hasattr(conn, "_suspend_sql_mirror"):
        previous = bool(getattr(conn, "_suspend_sql_mirror", False))
        conn._suspend_sql_mirror = True
        try:
            company_id = ensure_saas_schema(conn)
            ensure_tenant_columns(conn, company_id)
        finally:
            conn._suspend_sql_mirror = previous
        return
    company_id = ensure_saas_schema(conn)
    ensure_tenant_columns(conn, company_id)


def row_to_dict(row: Any) -> dict | None:
    if row is None:
        return None
    if isinstance(row, dict):
        return dict(row)
    try:
        return {key: row[key] for key in row.keys()}
    except Exception:
        return dict(row)


def rows_to_dicts(rows: Iterable[Any]) -> list[dict]:
    return [row_to_dict(row) for row in (rows or []) if row is not None]


def normalize_limit(limit: int | None, *, default: int = 500, maximum: int = 5000) -> int:
    try:
        value = int(limit or default)
    except Exception:
        value = default
    return max(1, min(value, maximum))


__all__ = [
    "ensure_saas_ready",
    "get_db",
    "normalize_limit",
    "row_to_dict",
    "rows_to_dicts",
]
