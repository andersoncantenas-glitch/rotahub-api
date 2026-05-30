import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402
from app.services import saas_admin_service  # noqa: E402
from db_bootstrap import ensure_core_schema  # noqa: E402


class AdminSaaSEndpointsTests(unittest.TestCase):
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

    def test_admin_lists_companies_plans_usage_and_features(self):
        companies = api_server.admin_list_companies(status="", limit=500, admin=self.admin)
        plans = api_server.admin_list_plans(include_inactive=False, admin=self.admin)
        usage = api_server.admin_company_usage(self.company_id, admin=self.admin)
        features = api_server.admin_company_features(self.company_id, admin=self.admin)

        self.assertEqual(len(companies), 1)
        self.assertGreaterEqual(len(plans), 4)
        self.assertEqual(usage["company_id"], self.company_id)
        self.assertEqual(features["plan_code"], "starter")
        self.assertIn("features", features)

    def test_admin_create_and_register_payment_reactivates_company(self):
        api_server.admin_update_company_status(
            self.company_id,
            api_server.CompanyStatusIn(status="suspended", reason="teste"),
            admin=self.admin,
        )
        created = api_server.admin_create_payment(
            api_server.PaymentCreateIn(company_id=self.company_id, amount=10.0, due_date="2026-05-06"),
            admin=self.admin,
        )

        payment_id = int(created["payment"]["id"])
        paid = api_server.admin_register_payment(
            payment_id,
            api_server.PaymentRegisterIn(method="manual", reference="REC-1"),
            admin=self.admin,
        )

        self.assertEqual(paid["payment"]["status"], "paid")
        with sqlite3.connect(self.db_path) as conn:
            company_status = conn.execute("SELECT status FROM companies WHERE id=?", (self.company_id,)).fetchone()[0]
            audit_count = conn.execute("SELECT COUNT(*) FROM audit_logs WHERE action='pagamento_registrado'").fetchone()[0]
        self.assertEqual(company_status, "active")
        self.assertEqual(audit_count, 1)

    def test_overdue_check_suspends_company(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "UPDATE subscriptions SET next_due_date=date('now', '-3 day') WHERE company_id=?",
                (self.company_id,),
            )

        result = api_server.admin_run_overdue_check(api_server.BillingAutomationIn(grace_days=0), admin=self.admin)

        self.assertTrue(result["ok"])
        self.assertEqual(result["summary"]["suspended"], 1)
        with sqlite3.connect(self.db_path) as conn:
            company_status = conn.execute("SELECT status FROM companies WHERE id=?", (self.company_id,)).fetchone()[0]
        self.assertEqual(company_status, "suspended")

    def test_desktop_saas_admin_service_dashboard(self):
        from app.db.connection import configure_connection

        configure_connection(self.db_path)
        result = saas_admin_service.get_dashboard(self.company_id)

        self.assertTrue(result["ok"])
        self.assertEqual(result["data"]["company"]["id"], self.company_id)
        self.assertIn("plans", result["data"])


if __name__ == "__main__":
    unittest.main()
