import os
import sqlite3
import tempfile
import unittest

from app.db.connection import configure_connection
from app.services import audit_service, company_service, payment_service, plan_service, subscription_service
from db_bootstrap import ensure_core_schema
from tests._contract_test_helpers import ContractTestCase


class SaasBootstrapTests(ContractTestCase):
    def setUp(self):
        fd, self.db_path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        configure_connection(self.db_path)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("PRAGMA foreign_keys = ON")
            ensure_core_schema(conn)

    def tearDown(self):
        try:
            os.remove(self.db_path)
        except OSError:
            pass

    def _count(self, table):
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM {table}")
            row = cur.fetchone()
            return int(row[0] if row else 0)

    def test_bootstrap_creates_and_seeds_saas_tables(self):
        self.assertEqual(self._count("companies"), 1)
        self.assertGreaterEqual(self._count("plans"), 4)
        self.assertEqual(self._count("subscriptions"), 1)
        self.assertEqual(self._count("payments"), 0)
        self.assertEqual(self._count("audit_logs"), 1)

        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute("SELECT code, status FROM companies ORDER BY id LIMIT 1")
            self.assertEqual(cur.fetchone(), ("default", "active"))
            cur.execute("SELECT code FROM plans ORDER BY monthly_price ASC, id ASC")
            self.assertIn(("starter",), cur.fetchall())
            cur.execute("SELECT vehicle_limit FROM plans WHERE code = 'corporate_private'")
            self.assertEqual(cur.fetchone(), (50,))

    def test_saas_services_return_contracts(self):
        default_company = company_service.get_default_company()
        self.assert_contract(default_company)
        self.assertTrue(default_company["ok"])
        company_id = int(default_company["data"]["id"])

        plans = plan_service.list_plans()
        self.assert_contract(plans)
        self.assertTrue(plans["ok"])
        self.assertGreaterEqual(len(plans["data"]), 4)

        subscription = subscription_service.get_company_subscription(company_id)
        self.assert_contract(subscription)
        self.assertTrue(subscription["ok"])
        self.assertEqual(subscription["data"]["plan_code"], "starter")

        changed = subscription_service.change_company_plan(company_id, "growth")
        self.assert_contract(changed)
        self.assertTrue(changed["ok"])
        self.assertEqual(changed["data"]["plan_code"], "growth")

        payment = payment_service.create_payment(
            {
                "company_id": company_id,
                "amount": 123.45,
                "due_date": "2026-05-06",
            }
        )
        self.assert_contract(payment)
        self.assertTrue(payment["ok"])

        paid = payment_service.register_payment(payment["data"]["id"], method="manual", reference="TESTE")
        self.assert_contract(paid)
        self.assertTrue(paid["ok"])
        self.assertEqual(paid["data"]["status"], "paid")

        audit = audit_service.record_audit_log(
            company_id=company_id,
            action="test_event",
            entity_type="company",
            entity_id=str(company_id),
            metadata={"source": "unit_test"},
        )
        self.assert_contract(audit)
        self.assertTrue(audit["ok"])


if __name__ == "__main__":
    unittest.main()
