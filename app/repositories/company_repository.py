from __future__ import annotations

from app.repositories.base_repository import (
    ensure_saas_ready,
    get_db,
    normalize_limit,
    row_to_dict,
    rows_to_dicts,
)


COMPANY_FIELDS = {
    "code",
    "name",
    "legal_name",
    "document",
    "email",
    "phone",
    "status",
    "timezone",
}


def get_company(company_id: int) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (int(company_id or 0),))
        return row_to_dict(cur.fetchone())


def get_company_by_code(code: str) -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM companies WHERE code=? LIMIT 1", (str(code or "").strip(),))
        return row_to_dict(cur.fetchone())


def get_default_company() -> dict | None:
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute("SELECT * FROM companies ORDER BY id ASC LIMIT 1")
        return row_to_dict(cur.fetchone())


def list_companies(status: str | None = None, limit: int | None = 500) -> list[dict]:
    params: list[object] = []
    where = ""
    if status:
        where = "WHERE status=?"
        params.append(str(status).strip())
    params.append(normalize_limit(limit))
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT *
            FROM companies
            {where}
            ORDER BY id ASC
            LIMIT ?
            """,
            tuple(params),
        )
        return rows_to_dicts(cur.fetchall())


def create_company(data: dict) -> dict:
    payload = {k: v for k, v in dict(data or {}).items() if k in COMPANY_FIELDS}
    if not str(payload.get("code") or "").strip():
        raise ValueError("Codigo da empresa e obrigatorio.")
    if not str(payload.get("name") or "").strip():
        raise ValueError("Nome da empresa e obrigatorio.")
    payload.setdefault("status", "active")
    payload.setdefault("timezone", "America/Fortaleza")

    fields = list(payload.keys())
    placeholders = ", ".join(["?"] * len(fields))
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"""
            INSERT INTO companies ({", ".join(fields)}, created_at, updated_at)
            VALUES ({placeholders}, datetime('now'), datetime('now'))
            """,
            tuple(payload[field] for field in fields),
        )
        company_id = int(cur.lastrowid)
        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (company_id,))
        return row_to_dict(cur.fetchone()) or {}


def update_company(company_id: int, data: dict) -> dict | None:
    payload = {k: v for k, v in dict(data or {}).items() if k in COMPANY_FIELDS and k != "code"}
    if not payload:
        return get_company(company_id)
    set_clause = ", ".join(f"{field}=?" for field in payload)
    with get_db() as conn:
        ensure_saas_ready(conn)
        cur = conn.cursor()
        cur.execute(
            f"UPDATE companies SET {set_clause}, updated_at=datetime('now') WHERE id=?",
            tuple(payload.values()) + (int(company_id or 0),),
        )
        cur.execute("SELECT * FROM companies WHERE id=? LIMIT 1", (int(company_id or 0),))
        return row_to_dict(cur.fetchone())
