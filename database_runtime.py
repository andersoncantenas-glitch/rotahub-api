import hashlib
import json
import logging
import os
import sqlite3
import uuid
from datetime import datetime, timedelta
from typing import Callable, Dict, Iterable, List, Optional

from runtime_config import AppConfig


CRITICAL_TABLES = [
    "usuarios",
    "motoristas",
    "veiculos",
    "ajudantes",
    "programacoes",
    "programacao_itens",
    "recebimentos",
    "despesas",
    "centro_custos",
    "rotas",
]


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _db_fingerprint(db_path: str) -> str:
    payload = f"{db_path}|{os.path.getsize(db_path) if os.path.exists(db_path) else 0}"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def _table_exists(cur: sqlite3.Cursor, name: str) -> bool:
    cur.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1", (name,))
    return bool(cur.fetchone())


def ensure_runtime_schema(db_path: str, config: AppConfig) -> None:
    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS app_metadata (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT,
                updated_at TEXT
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS schema_migrations (
                version INTEGER PRIMARY KEY,
                applied_at TEXT NOT NULL,
                notes TEXT
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sync_queue (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tenant_id TEXT NOT NULL,
                company_id TEXT NOT NULL,
                endpoint TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                attempt_count INTEGER NOT NULL DEFAULT 0,
                next_attempt_at TEXT,
                last_error TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sync_queue_status_next ON sync_queue(status, next_attempt_at)")

        metadata = {
            "app_env": config.app_env,
            "tenant_id": config.tenant_id,
            "company_id": config.company_id,
            "source_type": config.app_kind,
            "source_of_truth": config.source_of_truth,
            "tenant_mode": config.tenant_mode,
            "schema_version": str(config.schema_version),
            "installation_id": str(uuid.uuid5(uuid.NAMESPACE_URL, f"{config.app_kind}:{config.runtime_dir}:{config.tenant_id}")),
            "db_fingerprint": _db_fingerprint(db_path),
            "created_at": _now_str(),
        }
        for key, value in metadata.items():
            cur.execute(
                """
                INSERT INTO app_metadata(meta_key, meta_value, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value, updated_at=excluded.updated_at
                """,
                (key, str(value), _now_str()),
            )
        cur.execute(
            "INSERT OR IGNORE INTO schema_migrations(version, applied_at, notes) VALUES (?, ?, ?)",
            (int(config.schema_version), _now_str(), "bootstrap"),
        )
        conn.commit()
    finally:
        conn.close()


def read_metadata(db_path: str) -> Dict[str, str]:
    if not os.path.exists(db_path):
        return {}
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        if not _table_exists(cur, "app_metadata"):
            return {}
        cur.execute("SELECT meta_key, meta_value FROM app_metadata")
        return {str(k): str(v) for k, v in cur.fetchall() or []}
    finally:
        conn.close()


def validate_database_identity(db_path: str, config: AppConfig) -> None:
    meta = read_metadata(db_path)
    if not meta:
        return
    db_env = str(meta.get("app_env") or "").strip()
    db_tenant = str(meta.get("tenant_id") or "").strip()
    db_company = str(meta.get("company_id") or "").strip()
    if db_env and db_env != config.app_env:
        raise RuntimeError(f"Banco incompatível com APP_ENV atual: {db_env} != {config.app_env}")
    if db_tenant and db_tenant != config.tenant_id:
        raise RuntimeError(f"Banco incompatível com tenant atual: {db_tenant} != {config.tenant_id}")
    if db_company and db_company != config.company_id:
        raise RuntimeError(f"Banco incompatível com company atual: {db_company} != {config.company_id}")


def count_core_tables(db_path: str) -> Dict[str, int]:
    if not os.path.exists(db_path):
        return {name: 0 for name in CRITICAL_TABLES}
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        counts: Dict[str, int] = {}
        for table in CRITICAL_TABLES:
            if not _table_exists(cur, table):
                counts[table] = 0
                continue
            try:
                cur.execute(f"SELECT COUNT(*) FROM {table}")
                row = cur.fetchone()
                counts[table] = int(row[0] if row else 0)
            except Exception:
                counts[table] = 0
        return counts
    finally:
        conn.close()


def log_startup_diagnostics(db_path: str, config: AppConfig) -> Dict[str, int]:
    ensure_runtime_schema(db_path, config)
    validate_database_identity(db_path, config)
    counts = count_core_tables(db_path)
    logging.log(
        getattr(logging, str(config.log_level or "INFO").upper(), logging.INFO),
        "Startup diagnostics | env=%s | db=%s | api=%s | sync=%s | sql_mirror=%s | channel=%s | tenant=%s | company=%s | persistence=%s | source=%s | version=%s | schema=%s | secret=%s | counts=%s",
        config.app_env,
        os.path.abspath(db_path),
        config.api_base_url,
        config.sync_enabled,
        config.sql_mirror_api,
        config.update_channel,
        config.tenant_id,
        config.company_id,
        config.tenant_mode,
        config.source_of_truth,
        config.app_version,
        config.schema_version,
        "present" if config.desktop_secret else "missing",
        json.dumps(counts, ensure_ascii=False),
    )
    if sum(counts.values()) == 0 and os.path.exists(db_path) and os.path.getsize(db_path) > 0:
        logging.warning("Banco aberto sem registros nas tabelas críticas. Verifique se o DB_PATH está correto: %s", db_path)
    return counts


def enqueue_sql_statements(db_path: str, config: AppConfig, statements: List[Dict[str, object]], endpoint: str = "desktop/sql/mutate") -> None:
    if not statements:
        return
    ensure_runtime_schema(db_path, config)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        now = _now_str()
        cur.execute(
            """
            INSERT INTO sync_queue(tenant_id, company_id, endpoint, payload_json, status, attempt_count, next_attempt_at, last_error, created_at, updated_at)
            VALUES (?, ?, ?, ?, 'pending', 0, ?, '', ?, ?)
            """,
            (
                config.tenant_id,
                config.company_id,
                endpoint,
                json.dumps({"statements": statements}, ensure_ascii=False),
                now,
                now,
                now,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def process_sync_queue(
    db_path: str,
    config: AppConfig,
    api_caller: Callable[[str, str, object, Optional[str], Optional[Dict[str, str]]], object],
    *,
    batch_size: int = 20,
    max_attempts: int = 5,
) -> Dict[str, int]:
    result = {"sent": 0, "failed": 0, "dead_letter": 0}
    if not config.sync_enabled or not config.sql_mirror_api or not config.allow_remote_write:
        return result
    if not config.desktop_secret:
        logging.warning("Sync queue não processada: ROTA_SECRET ausente.")
        return result
    ensure_runtime_schema(db_path, config)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        now = _now_str()
        cur.execute(
            """
            SELECT id, endpoint, payload_json, attempt_count
            FROM sync_queue
            WHERE status IN ('pending', 'failed')
              AND (next_attempt_at IS NULL OR next_attempt_at <= ?)
            ORDER BY id
            LIMIT ?
            """,
            (now, int(batch_size)),
        )
        rows = cur.fetchall() or []
        for row in rows:
            queue_id, endpoint, payload_json, attempt_count = int(row[0]), str(row[1]), str(row[2]), int(row[3] or 0)
            try:
                payload = json.loads(payload_json)
                api_caller(
                    "POST",
                    endpoint,
                    payload=payload,
                    token=None,
                    extra_headers={"X-Desktop-Secret": config.desktop_secret},
                )
                cur.execute(
                    "UPDATE sync_queue SET status='sent', last_error='', updated_at=? WHERE id=?",
                    (_now_str(), queue_id),
                )
                result["sent"] += 1
            except Exception as exc:
                next_attempt = datetime.now() + timedelta(minutes=min(30, 2 ** min(attempt_count, 4)))
                new_attempts = attempt_count + 1
                new_status = "dead_letter" if new_attempts >= max_attempts else "failed"
                cur.execute(
                    """
                    UPDATE sync_queue
                    SET status=?, attempt_count=?, next_attempt_at=?, last_error=?, updated_at=?
                    WHERE id=?
                    """,
                    (
                        new_status,
                        new_attempts,
                        next_attempt.strftime("%Y-%m-%d %H:%M:%S"),
                        str(exc)[:1000],
                        _now_str(),
                        queue_id,
                    ),
                )
                if new_status == "dead_letter":
                    result["dead_letter"] += 1
                else:
                    result["failed"] += 1
        conn.commit()
        return result
    finally:
        conn.close()
