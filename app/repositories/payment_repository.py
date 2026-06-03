from __future__ import annotations

from app.repositories.base_repository import ensure_saas_ready, get_db, normalize_limit, row_to_dict, rows_to_dicts


def get_payment(payment_id: int) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (int(payment_id or 0),))
        return row_to_dict(cur.fetchone())


def list_payments(company_id: int | None = None, status: str | None = None, limit: int | None = 500) -> list[dict]:
    clauses: list[str] = []
    params: list[object] = []
    if company_id:
        clauses.append("company_id=?")
        params.append(int(company_id))
    if status:
        clauses.append("status=?")
        params.append(str(status).strip())
    where = "WHERE " + " AND ".join(clauses) if clauses else ""
    params.append(normalize_limit(limit))
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT *
            FROM payments
            {where}
            ORDER BY COALESCE(due_date, created_at) DESC, id DESC
            LIMIT ?
            """,
            tuple(params),
        )
        return rows_to_dicts(cur.fetchall())


def create_payment(data: dict) -> dict:
    payload = {
        "subscription_id": data.get("subscription_id"),
        "company_id": data.get("company_id"),
        "amount": data.get("amount", 0),
        "due_date": data.get("due_date"),
        "paid_at": data.get("paid_at"),
        "status": data.get("status") or "pending",
        "method": data.get("method"),
        "reference": data.get("reference"),
        "boleto_status": data.get("boleto_status"),
        "boleto_our_number": data.get("boleto_our_number"),
        "boleto_digitable_line": data.get("boleto_digitable_line"),
        "boleto_pdf_url": data.get("boleto_pdf_url"),
        "boleto_pdf_path": data.get("boleto_pdf_path"),
        "boleto_generated_at": data.get("boleto_generated_at"),
        "notes": data.get("notes"),
    }
    if not payload.get("company_id"):
        raise ValueError("Empresa do pagamento e obrigatoria.")
    fields = list(payload.keys())
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            INSERT INTO payments ({", ".join(fields)}, created_at, updated_at)
            VALUES ({", ".join(["?"] * len(fields))}, datetime('now'), datetime('now'))
            """,
            tuple(payload[field] for field in fields),
        )
        payment_id = int(cur.lastrowid)
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (payment_id,))
        return row_to_dict(cur.fetchone()) or {}


def update_boleto(payment_id: int, data: dict) -> dict | None:
    payload = {
        "method": data.get("method") or "boleto",
        "reference": data.get("reference"),
        "boleto_status": data.get("boleto_status") or "generated",
        "boleto_our_number": data.get("boleto_our_number"),
        "boleto_digitable_line": data.get("boleto_digitable_line"),
        "boleto_pdf_url": data.get("boleto_pdf_url"),
        "boleto_pdf_path": data.get("boleto_pdf_path"),
        "boleto_generated_at": data.get("boleto_generated_at"),
    }
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE payments
            SET
                method=?,
                reference=COALESCE(?, reference),
                boleto_status=?,
                boleto_our_number=?,
                boleto_digitable_line=?,
                boleto_pdf_url=?,
                boleto_pdf_path=?,
                boleto_generated_at=COALESCE(?, datetime('now')),
                updated_at=datetime('now')
            WHERE id=?
            """,
            (
                payload["method"],
                payload["reference"],
                payload["boleto_status"],
                payload["boleto_our_number"],
                payload["boleto_digitable_line"],
                payload["boleto_pdf_url"],
                payload["boleto_pdf_path"],
                payload["boleto_generated_at"],
                int(payment_id or 0),
            ),
        )
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (int(payment_id or 0),))
        return row_to_dict(cur.fetchone())


def register_payment(payment_id: int, *, method: str | None = None, reference: str | None = None, notes: str | None = None) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE payments
            SET
                status='paid',
                paid_at=datetime('now'),
                method=COALESCE(?, method),
                reference=COALESCE(?, reference),
                notes=COALESCE(?, notes),
                updated_at=datetime('now')
            WHERE id=?
            """,
            (method, reference, notes, int(payment_id or 0)),
        )
        cur.execute("SELECT subscription_id FROM payments WHERE id=? LIMIT 1", (int(payment_id or 0),))
        payment_ref = cur.fetchone()
        subscription_id = None
        if payment_ref:
            try:
                subscription_id = payment_ref["subscription_id"]
            except Exception:
                subscription_id = payment_ref[0]
        if subscription_id:
            cur.execute(
                """
                UPDATE subscriptions
                SET
                    status='active',
                    current_period_start=date('now'),
                    current_period_end=date('now', '+30 day'),
                    next_due_date=date('now', '+30 day'),
                    updated_at=datetime('now')
                WHERE id=?
                """,
                (int(subscription_id),),
            )
            cur.execute(
                """
                UPDATE companies
                SET status='active', updated_at=datetime('now')
                WHERE id=(SELECT company_id FROM subscriptions WHERE id=?)
                """,
                (int(subscription_id),),
            )
        cur.execute("SELECT * FROM payments WHERE id=? LIMIT 1", (int(payment_id or 0),))
        return row_to_dict(cur.fetchone())
