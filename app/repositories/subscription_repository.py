from __future__ import annotations

from app.repositories.base_repository import ensure_saas_ready, get_db, normalize_limit, row_to_dict, rows_to_dicts


SUBSCRIPTION_SELECT = """
SELECT
    s.*,
    p.code AS plan_code,
    p.name AS plan_name,
    p.monthly_price AS plan_monthly_price,
    p.vehicle_limit AS plan_vehicle_limit,
    p.user_limit AS plan_user_limit,
    p.features_json AS plan_features_json
FROM subscriptions s
JOIN plans p ON p.id = s.plan_id
"""


def get_subscription(subscription_id: int) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"{SUBSCRIPTION_SELECT} WHERE s.id=? LIMIT 1",
            (int(subscription_id or 0),),
        )
        return row_to_dict(cur.fetchone())


def get_active_subscription(company_id: int) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            {SUBSCRIPTION_SELECT}
            WHERE s.company_id=? AND s.status IN ('active', 'trialing', 'past_due')
            ORDER BY s.id DESC
            LIMIT 1
            """,
            (int(company_id or 0),),
        )
        return row_to_dict(cur.fetchone())


def list_subscriptions(company_id: int | None = None, limit: int | None = 500) -> list[dict]:
    params: list[object] = []
    where = ""
    if company_id:
        where = "WHERE s.company_id=?"
        params.append(int(company_id))
    params.append(normalize_limit(limit))
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            {SUBSCRIPTION_SELECT}
            {where}
            ORDER BY s.id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return rows_to_dicts(cur.fetchall())


def create_subscription(
    *,
    company_id: int,
    plan_id: int,
    status: str = "active",
    billing_cycle: str = "monthly",
    next_due_date: str | None = None,
) -> dict:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO subscriptions (
                company_id, plan_id, status, billing_cycle,
                current_period_start, current_period_end, next_due_date,
                created_at, updated_at
            )
            VALUES (
                ?, ?, ?, ?,
                date('now'), date('now', '+30 day'), COALESCE(?, date('now', '+30 day')),
                datetime('now'), datetime('now')
            )
            """,
            (int(company_id), int(plan_id), str(status or "active"), str(billing_cycle or "monthly"), next_due_date),
        )
        subscription_id = int(cur.lastrowid)
    return get_subscription(subscription_id) or {}


def change_company_plan(company_id: int, plan_id: int) -> dict:
    current = get_active_subscription(company_id)
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        if current:
            cur.execute(
                """
                UPDATE subscriptions
                SET plan_id=?, status='active', updated_at=datetime('now')
                WHERE id=?
                """,
                (int(plan_id), int(current["id"])),
            )
            subscription_id = int(current["id"])
        else:
            cur.execute(
                """
                INSERT INTO subscriptions (
                    company_id, plan_id, status, billing_cycle,
                    current_period_start, current_period_end, next_due_date,
                    created_at, updated_at
                )
                VALUES (
                    ?, ?, 'active', 'monthly',
                    date('now'), date('now', '+30 day'), date('now', '+30 day'),
                    datetime('now'), datetime('now')
                )
                """,
                (int(company_id), int(plan_id)),
            )
            subscription_id = int(cur.lastrowid)
    return get_subscription(subscription_id) or {}
