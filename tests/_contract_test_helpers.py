import unittest


class ContractTestCase(unittest.TestCase):
    def assert_contract(self, result):
        self.assertIsInstance(result, dict)
        self.assertEqual(set(result.keys()), {"ok", "data", "error", "source"})
        self.assertIsInstance(result["ok"], bool)
        self.assertIn(result["source"], {"local", "api", "both"})


class _FakeCursor:
    def __init__(self, table_info_rows=None, fetchone_rows=None):
        self._table_info_rows = table_info_rows or []
        self._fetchone_rows = list(fetchone_rows or [])

    def execute(self, *_args, **_kwargs):
        return None

    def fetchall(self):
        return list(self._table_info_rows)

    def fetchone(self):
        if not self._fetchone_rows:
            return None
        return self._fetchone_rows.pop(0)


class _FakeConn:
    def __init__(self, cursor):
        self._cursor = cursor

    def cursor(self):
        return self._cursor


class _FakeCtx:
    def __init__(self, conn):
        self._conn = conn

    def __enter__(self):
        return self._conn

    def __exit__(self, exc_type, exc, tb):
        return False
