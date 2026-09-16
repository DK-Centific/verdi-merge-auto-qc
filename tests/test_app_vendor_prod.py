"""App helpers: destination status and Send fail-loud behavior."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import sharepoint_vendor_prod as sp


class StatusAndSendIntentTests(unittest.TestCase):
    def test_status_local_when_no_auth(self):
        import app as appmod

        cfg = sp.SharePointConfig()
        cfg.graph_token = ""
        cfg.tenant_id = ""
        cfg.client_id = ""
        cfg.client_secret = ""
        cfg.url_from_env = False
        cfg.require = False
        with patch.object(sp, "SharePointConfig") as cfg_cls:
            cfg_cls.from_env.return_value = cfg
            payload = appmod.vendor_prod_status_payload()
        self.assertIn(payload["destination_kind"], ("local", "unconfigured"))
        self.assertFalse(payload["sharepoint"]["auth_configured"])

    def test_send_sharepoint_without_auth_raises(self):
        import app as appmod

        cfg = sp.SharePointConfig()
        cfg.url_from_env = True
        cfg.graph_token = ""
        cfg.tenant_id = ""
        cfg.client_id = ""
        cfg.client_secret = ""
        self.assertTrue(sp.sharepoint_write_intended(cfg))
        self.assertFalse(cfg.auth_configured())
        # The route should not call local append when SharePoint is intended.
        with self.assertRaises(sp.SharePointError):
            # mimic the branch in send_vendor_prod
            if not cfg.auth_configured():
                raise sp.SharePointError(
                    "SharePoint is the intended Vendor Prod destination, but Graph auth is missing.",
                    hint=sp.AUTH_HINT,
                    status_code=503,
                )

        called = {"local": False}

        def boom(*_a, **_k):
            called["local"] = True
            raise AssertionError("local send must not run")

        with patch.object(appmod, "_send_local_master", side_effect=boom):
            with patch.object(sp, "sharepoint_write_intended", return_value=True):
                with patch.object(sp, "SharePointConfig") as cfg_cls:
                    cfg_cls.from_env.return_value = cfg
                    with patch.object(appmod, "_load_last_merged_rows", return_value=[{"fileName": "x.wav", "vendorName": "A"}]):
                        from fastapi.testclient import TestClient

                        try:
                            client = TestClient(appmod.app)
                        except ImportError:
                            self.skipTest("httpx not installed")
                        res = client.post("/api/send-vendor-prod")
        self.assertFalse(called["local"])
        self.assertEqual(res.status_code, 503)
        body = res.json()
        self.assertFalse(body["ok"])
        self.assertIn("Graph auth", body["error"])
        self.assertIn("AZURE", body.get("hint", ""))

    def test_run_then_local_send_against_prod_schema(self):
        """Demo run + Send to a temp copy of the Rate Approval Log schema."""
        from fastapi.testclient import TestClient
        import shutil
        import tempfile

        import app as appmod
        import ots_pipeline as pipe

        uploaded = Path("/home/ubuntu/.cursor/projects/workspace/uploads/Vendor_Prod_Files_cde7.xlsx")
        with tempfile.TemporaryDirectory() as td:
            if uploaded.exists():
                master = Path(td) / "Vendor Prod Files.xlsx"
                shutil.copy(uploaded, master)
            else:
                # Minimal schema if the upload is not on this machine
                master = Path(td) / "Vendor Prod Files.xlsx"
                master.write_bytes(
                    __import__("tests.test_vendor_prod", fromlist=["_sample_master_bytes"])._sample_master_bytes()
                )

            names_before, _ = pipe.extract_existing_rate_keys_from_path(master)
            client = TestClient(appmod.app)
            run = client.post("/api/run", data={"demo_mode": "1", "skip_oneforma": "1"})
            self.assertEqual(run.status_code, 200, run.text)
            body = run.json()
            self.assertTrue(body["ok"])
            self.assertGreater(body["merged_count"], 0)
            self.assertIn("unique_count", body)
            self.assertIn("duplicate_count", body)
            self.assertIn("vendor_prod", body)
            self.assertEqual(body["vendor_prod"]["dedupe_key"], pipe.RATE_DEDUPE_KEY)
            # Preview xlsx exists
            vp = client.get("/api/download/vendor-prod-rows")
            self.assertEqual(vp.status_code, 200)

            send = client.post(
                "/api/send-vendor-prod",
                data={"force_local": "1", "master_path": str(master)},
            )
            self.assertEqual(send.status_code, 200, send.text)
            sent = send.json()
            self.assertTrue(sent["ok"])
            self.assertEqual(sent["destination"], "local")
            names_after, _ = pipe.extract_existing_rate_keys_from_path(master)
            self.assertEqual(len(names_after), len(names_before) + sent["appended"])
            # Second send is all duplicates
            send2 = client.post(
                "/api/send-vendor-prod",
                data={"force_local": "1", "master_path": str(master)},
            )
            self.assertEqual(send2.json()["appended"], 0)
            self.assertGreater(send2.json()["skipped_duplicates"], 0)


if __name__ == "__main__":
    unittest.main()
