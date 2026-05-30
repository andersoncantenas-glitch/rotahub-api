from __future__ import annotations

from app.repositories.base_repository import ensure_saas_ready, get_db
from app.services.saas_result import error_message, service_result


def suspend_overdue_subscriptions(*, grace_days: int = 0) -> dict:
    try:
        with get_db() as conn:
            ensure_saas_ready(conn)
            summary = suspend_overdue_subscriptions_conn(conn, grace_days=grace_days)
        return service_result(ok=True, data=summary)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao suspender assinaturas vencidas."))


def suspend_overdue_subscriptions_conn(conn, *, grace_days: int = 0) -> dict:
    cur = conn.cursor()
    cur.execute(
        """
        SELECT s.id, s.company_id, s.next_due_date
        FROM subscriptions s
        WHERE s.status IN ('active', 'trialing', 'past_due')
          AND s.next_due_date IS NOT NULL
          AND date(s.next_due_date, ?) < date('now')
        ORDER BY s.id ASC
        """,
        (f"+{max(0, int(grace_days or 0))} day",),
    )
    rows = cur.fetchall() or []
    suspended = 0
    for row in rows:
        subscription_id = int(row["id"] if hasattr(row, "keys") else row[0])
        company_id = int(row["company_id"] if hasattr(row, "keys") else row[1])
        cur.execute(
            "UPDATE subscriptions SET status='suspended', updated_at=datetime('now') WHERE id=?",
            (subscription_id,),
        )
        cur.execute(
            "UPDATE companies SET status='suspended', updated_at=datetime('now') WHERE id=?",
            (company_id,),
        )
        _audit_suspension(cur, company_id, subscription_id)
        suspended += 1
    return {"checked": len(rows), "suspended": suspended, "grace_days": max(0, int(grace_days or 0))}


def _audit_suspension(cur, company_id: int, subscription_id: int) -> None:
    try:
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='audit_logs'")
        if not cur.fetchone():
            return
        cur.execute(
            """
            INSERT INTO audit_logs (
                company_id, actor_type, action, entity_type, entity_id,
                severity, metadata_json, created_at
            )
            VALUES (?, 'system', 'empresa_suspensa_por_atraso', 'subscription', ?, 'warning', '{}', datetime('now'))
            """,
            (int(company_id), str(subscription_id)),
        )
    except Exception:
        pass
