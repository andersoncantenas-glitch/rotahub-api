import os
import sqlite3
import tempfile
import unittest


os.environ.setdefault("ROTA_SECRET", "test-secret")
os.environ["ROTA_DB"] = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name

import api_server  # noqa: E402
from app.services.vehicle_limit_service import check_vehicle_limit  # noqa: E402
from db_bootstrap import ensure_core_schema  # noqa: E402


class VehicleLimitTests(unittest.TestCase):
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

    def tearDown(self):
        try:
            os.unlink(self.db_path)
        except OSError:
            pass

    def _insert_vehicle(self, placa):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO veiculos (placa, modelo, capacidade_cx, company_id) VALUES (?, 'TRUCK', 100, ?)",
                (placa, self.company_id),
            )

    def _create_company_with_plan(self, code: str, plan_code: str = "professional") -> int:
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO companies (code, name, status) VALUES (?, ?, 'active')",
                (code, code.upper()),
            )
            company_id = int(cur.lastrowid)
            cur.execute("SELECT id FROM plans WHERE code=? LIMIT 1", (plan_code,))
            plan_id = int(cur.fetchone()[0])
            cur.execute(
                """
                INSERT INTO subscriptions (
                    company_id, plan_id, status, billing_cycle,
                    current_period_start, current_period_end, next_due_date
                )
                VALUES (?, ?, 'active', 'monthly', date('now'), date('now', '+30 day'), date('now', '+30 day'))
                """,
                (company_id, plan_id),
            )
        return company_id

    def test_vehicle_limit_service_blocks_when_plan_limit_is_reached(self):
        for idx in range(5):
            self._insert_vehicle(f"ABC1D2{idx}")

        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            result = check_vehicle_limit(conn, self.company_id, exclude_placa="NEW1234")

        self.assertFalse(result["ok"])
        self.assertIn("Limite de 5 veiculos atingido", result["error"])
        self.assertEqual(result["data"]["vehicle_count"], 5)

    def test_desktop_vehicle_upsert_blocks_new_vehicle_above_limit_and_audits(self):
        for idx in range(5):
            self._insert_vehicle(f"ABC1D2{idx}")

        payload = api_server.DesktopVeiculoUpsertIn(
            placa="ZZZ9Z99",
            modelo="Truck Novo",
            capacidade_cx=120,
        )

        with self.assertRaises(api_server.HTTPException) as ctx:
            api_server.desktop_veiculos_upsert(payload, _ok=True)

        self.assertEqual(ctx.exception.status_code, 403)
        self.assertIn("Limite de 5 veiculos atingido", str(ctx.exception.detail))
        with sqlite3.connect(self.db_path) as conn:
            count_new = conn.execute("SELECT COUNT(*) FROM veiculos WHERE placa='ZZZ9Z99'").fetchone()[0]
            audit_count = conn.execute(
                "SELECT COUNT(*) FROM audit_logs WHERE action='tentativa_exceder_limite_veiculo'"
            ).fetchone()[0]
        self.assertEqual(count_new, 0)
        self.assertEqual(audit_count, 1)

    def test_desktop_vehicle_upsert_allows_update_when_limit_is_reached(self):
        for idx in range(5):
            self._insert_vehicle(f"ABC1D2{idx}")

        payload = api_server.DesktopVeiculoUpsertIn(
            placa="ABC1D20",
            modelo="Truck Atualizado",
            capacidade_cx=150,
        )

        result = api_server.desktop_veiculos_upsert(payload, _ok=True)

        self.assertTrue(result["ok"])
        self.assertEqual(result["updated"], 1)
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT modelo, capacidade_cx FROM veiculos WHERE placa='ABC1D20'").fetchone()
        self.assertEqual(row, ("TRUCK ATUALIZADO", 150))

    def test_desktop_vehicle_upsert_uses_company_header(self):
        for idx in range(5):
            self._insert_vehicle(f"ABC1D2{idx}")
        other_company_id = self._create_company_with_plan("cliente-b")

        payload = api_server.DesktopVeiculoUpsertIn(
            placa="BBB1B11",
            modelo="Truck Cliente B",
            capacidade_cx=90,
        )

        result = api_server.desktop_veiculos_upsert(payload, _ok=True, x_company_id=str(other_company_id))

        self.assertTrue(result["ok"])
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute("SELECT company_id FROM veiculos WHERE placa='BBB1B11'").fetchone()
        self.assertEqual(int(row[0]), other_company_id)

    def test_desktop_vehicle_upsert_rejects_unknown_company_header(self):
        payload = api_server.DesktopVeiculoUpsertIn(
            placa="XXX1X11",
            modelo="Truck Invalido",
            capacidade_cx=90,
        )

        with self.assertRaises(api_server.HTTPException) as ctx:
            api_server.desktop_veiculos_upsert(payload, _ok=True, x_company_id="99999")

        self.assertEqual(ctx.exception.status_code, 404)


if __name__ == "__main__":
    unittest.main()
