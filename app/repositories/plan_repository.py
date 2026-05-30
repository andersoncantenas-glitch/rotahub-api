from __future__ import annotations

from app.repositories.base_repository import ensure_saas_ready, get_db, normalize_limit, row_to_dict, rows_to_dicts


def get_plan(plan_id: int) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM plans WHERE id=? LIMIT 1", (int(plan_id or 0),))
        return row_to_dict(cur.fetchone())


def get_plan_by_code(code: str) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM plans WHERE code=? LIMIT 1", (str(code or "").strip(),))
        return row_to_dict(cur.fetchone())


def list_plans(*, include_inactive: bool = False, limit: int | None = 100) -> list[dict]:
    where = "" if include_inactive else "WHERE status='active'"
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT *
            FROM plans
            {where}
            ORDER BY monthly_price ASC, id ASC
            LIMIT ?
            """,
            (normalize_limit(limit, default=100),),
        )
        return rows_to_dicts(cur.fetchall())
