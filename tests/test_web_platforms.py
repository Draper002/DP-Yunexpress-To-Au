import importlib
import os
import tempfile
import unittest
from pathlib import Path

from fastapi.testclient import TestClient


class WebPlatformTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        os.environ["DP_WEB_DATA_DIR"] = cls.temp_dir.name
        os.environ["DP_WEB_ADMIN_EMAIL"] = "admin@test.local"
        os.environ["DP_WEB_ADMIN_PASSWORD"] = "test-password"
        os.environ["DP_WEB_COOKIE_SECURE"] = "0"

        cls.web = importlib.import_module("web.app")
        cls.client_context = TestClient(cls.web.app)
        cls.client = cls.client_context.__enter__()

    @classmethod
    def tearDownClass(cls):
        cls.client_context.__exit__(None, None, None)
        cls.temp_dir.cleanup()

    def login(self, email="admin@test.local", password="test-password"):
        self.client.cookies.clear()
        response = self.client.post(
            "/login",
            data={"email": email, "password": password},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        return response

    def test_home_exposes_both_platforms_for_all_three_steps(self):
        response = self.login()
        self.assertEqual(response.text.count('name="platform" value="sf"'), 3)
        self.assertIn("顺丰国际订单数据", response.text)
        self.assertIn("顺丰国际面单 PDF", response.text)
        self.assertIn("SF INTERNATIONAL", response.text)

    def test_platform_specific_config_requirements(self):
        self.assertEqual(
            self.web.required_config_keys(1, "sf"),
            ("sku", "sf_template"),
        )
        self.assertEqual(
            self.web.required_config_keys(2, "sf"),
            ("dp_template",),
        )
        self.assertEqual(self.web.required_config_keys(3, "sf"), ())
        self.assertEqual(
            self.web.required_config_keys(3, "yunexpress"),
            ("sku",),
        )

    def test_health_endpoint_lists_both_platforms(self):
        response = self.client.get("/healthz")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["platforms"], ["yunexpress", "sf"])

    def test_employee_cannot_download_another_users_result(self):
        owner_email = "owner@test.local"
        other_email = "other@test.local"
        with self.web.connection() as conn:
            for email in (owner_email, other_email):
                conn.execute(
                    "INSERT OR IGNORE INTO users(email,password_hash,role,created_at) VALUES(?,?,?,?)",
                    (email, self.web.hash_password("employee-pass"), "employee", "2026-07-26T20:00:00"),
                )
            owner_id = conn.execute("SELECT id FROM users WHERE email=?", (owner_email,)).fetchone()["id"]
            conn.commit()

        result = Path(self.temp_dir.name) / "private-result.xlsx"
        result.write_bytes(b"private")
        self.web.add_run(owner_id, "步骤 2 · 顺丰国际", "2026-07-26", "SUCCESS", result)
        with self.web.connection() as conn:
            run_id = conn.execute("SELECT max(id) AS id FROM runs").fetchone()["id"]

        self.login(other_email, "employee-pass")
        response = self.client.get(f"/download/{run_id}", follow_redirects=False)
        self.assertEqual(response.status_code, 404)


if __name__ == "__main__":
    unittest.main()
