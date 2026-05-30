import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402
from db_bootstrap import ensure_core_schema  # noqa: E402


class PlanChangeTests(unittest.TestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        api_server.DB_PATH = self.db_path
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            ensure_core_schema(conn)
            cur = conn.cursor()
            cur.execute("SELECT id FROM companies WHERE code='default' LIMIT 1")
            self.company_id = int(cur.fetchone()["id"])
        self.admin = {
            "id": 1,
            "codigo": "ADMIN",
            "nome": "ADMIN",
            "company_id": self.company_id,
            "is_admin": True,
            "perfil_app": "ADMIN",
        }

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def _insert_vehicles(self, total):
        with sqlite3.connect(self.db_path) as conn:
            for idx in range(total):
                conn.execute(
                    "INSERT INTO veiculos (placa, modelo, capacidade_cx, company_id) VALUES (?, 'TRUCK', 100, ?)",
                    (f"PLT{idx:04d}", self.company_id),
                )

    def test_admin_can_upgrade_company_plan(self):
        payload = api_server.CompanyPlanChangeIn(plan_code="growth", reason="upgrade teste")

        result = api_server.admin_change_company_plan(self.company_id, payload, admin=self.admin)

        self.assertTrue(result["ok"])
        self.assertEqual(result["subscription"]["plan_code"], "growth")
        with sqlite3.connect(self.db_path) as conn:
            audit_count = conn.execute("SELECT COUNT(*) FROM audit_logs WHERE action='plano_alterado'").fetchone()[0]
        self.assertEqual(audit_count, 1)

    def test_downgrade_is_blocked_when_vehicle_count_exceeds_target_limit(self):
        payload = api_server.CompanyPlanChangeIn(plan_code="professional")
        api_server.admin_change_company_plan(self.company_id, payload, admin=self.admin)
        self._insert_vehicles(6)

        with self.assertRaises(api_server.HTTPException) as ctx:
            api_server.admin_change_company_plan(
                self.company_id,
                api_server.CompanyPlanChangeIn(plan_code="starter"),
                admin=self.admin,
            )

        self.assertEqual(ctx.exception.status_code, 409)
        self.assertIn("Downgrade bloqueado", str(ctx.exception.detail))
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                """
                SELECT p.code
                FROM subscriptions s
                JOIN plans p ON p.id = s.plan_id
                WHERE s.company_id=?
                ORDER BY s.id DESC
                LIMIT 1
                """,
                (self.company_id,),
            ).fetchone()
        self.assertEqual(row[0], "professional")


if __name__ == "__main__":
    unittest.main()
