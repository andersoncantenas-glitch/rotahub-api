import logging
import os
import sqlite3
from contextlib import contextmanager

_DB_PATH = None
_CALL_API = None
_IS_SYNC_ENABLED = None


def configure_connection(db_path, call_api=None, is_sync_enabled=None):
    global _DB_PATH, _CALL_API, _IS_SYNC_ENABLED
    _DB_PATH = db_path
    _CALL_API = call_api
    _IS_SYNC_ENABLED = is_sync_enabled


def _configured_db_path():
    if not _DB_PATH:
        raise RuntimeError("SQLite connection is not configured.")
    return _DB_PATH


def _configure_sqlite(conn: sqlite3.Connection) -> sqlite3.Connection:
    """Aplica pragmas de desempenho/concorrencia."""
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA busy_timeout = 5000")
    except Exception:
        logging.debug("Falha ignorada")
    return conn


def _is_mutating_sql(sql: str) -> bool:
    s = str(sql or "").lstrip().upper()
    return s.startswith("INSERT ") or s.startswith("UPDATE ") or s.startswith("DELETE ") or s.startswith("REPLACE ")


def _normalize_sql_scalar(v):
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    if isinstance(v, (bytes, bytearray)):
        try:
            return bytes(v).decode("utf-8", errors="replace")
        except Exception:
            return str(v)
    return str(v)


def _normalize_sql_params(params):
    if params is None:
        return []
    if isinstance(params, dict):
        out = {}
        for k, v in params.items():
            out[str(k)] = _normalize_sql_scalar(v)
        return out
    if not isinstance(params, (list, tuple)):
        return [_normalize_sql_scalar(params)]
    return [_normalize_sql_scalar(v) for v in params]


def _is_sql_mirror_enabled() -> bool:
    if not str(os.environ.get("ROTA_SECRET", "") or "").strip():
        return False
    raw = str(os.environ.get("ROTA_SQL_MIRROR_API", "1") or "").strip().lower()
    if not callable(_IS_SYNC_ENABLED):
        return False
    return _IS_SYNC_ENABLED() and raw in {"1", "true", "yes", "y", "sim", "on"}


def _push_sql_mutations_to_api(statements):
    if not statements:
        return
    if not _is_sql_mirror_enabled():
        return
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        logging.warning("Espelhamento SQL desativado: ROTA_SECRET ausente.")
        return
    if not callable(_CALL_API):
        raise RuntimeError("SQL mirror API caller is not configured.")
    _CALL_API(
        "POST",
        "desktop/sql/mutate",
        payload={"statements": statements},
        extra_headers={"X-Desktop-Secret": desktop_secret},
    )


class SyncedCursor(sqlite3.Cursor):
    def execute(self, sql, parameters=()):
        result = super().execute(sql, parameters)
        try:
            conn = getattr(self, "connection", None)
            if conn is not None and hasattr(conn, "_track_sql_mutation"):
                conn._track_sql_mutation(sql, parameters)
        except Exception:
            logging.debug("Falha ao rastrear mutacao SQL", exc_info=True)
        return result

    def executemany(self, sql, seq_of_parameters):
        seq = list(seq_of_parameters or [])
        result = super().executemany(sql, seq)
        try:
            conn = getattr(self, "connection", None)
            if conn is not None and hasattr(conn, "_track_sql_mutation"):
                for p in seq:
                    conn._track_sql_mutation(sql, p)
        except Exception:
            logging.debug("Falha ao rastrear mutacoes SQL (executemany)", exc_info=True)
        return result


class SyncedConnection(sqlite3.Connection):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pending_sql_mutations = []
        self._suspend_sql_mirror = False

    def cursor(self, *args, **kwargs):
        kwargs.setdefault("factory", SyncedCursor)
        return super().cursor(*args, **kwargs)

    def _track_sql_mutation(self, sql, params):
        if self._suspend_sql_mirror:
            return
        if not _is_mutating_sql(sql):
            return
        self._pending_sql_mutations.append(
            {
                "sql": str(sql),
                "params": _normalize_sql_params(params),
            }
        )

    def commit(self):
        if self._pending_sql_mutations and not self._suspend_sql_mirror:
            _push_sql_mutations_to_api(self._pending_sql_mutations)
            self._pending_sql_mutations = []
        return super().commit()

    def rollback(self):
        self._pending_sql_mutations = []
        return super().rollback()


@contextmanager
def get_db():
    """Gerenciador de contexto para conexoes com o banco."""
    conn = sqlite3.connect(_configured_db_path(), factory=SyncedConnection)
    _configure_sqlite(conn)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_connect():
    """Funcao compativel para codigo existente."""
    conn = sqlite3.connect(_configured_db_path(), factory=SyncedConnection)
    _configure_sqlite(conn)
    return conn
