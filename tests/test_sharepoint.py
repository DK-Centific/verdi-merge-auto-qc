"""SharePoint config + Graph client (mocked HTTP)."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import sharepoint_vendor_prod as sp


class ParseUrlTests(unittest.TestCase):
    def test_parse_live_sharing_url(self):
        parsed = sp.parse_sharepoint_xlsx_url(sp.DEFAULT_SHAREPOINT_VENDOR_PROD_URL)
        self.assertEqual(parsed["hostname"], "digitaltechedge.sharepoint.com")
        self.assertEqual(parsed["site_path"], "sites/Verdi_Juniper")
        self.assertEqual(parsed["library"], "Shared Documents")
        self.assertEqual(parsed["file_path"], "Vendor Audio Production/Vendor Prod Files.xlsx")

    def test_encode_sharing_url_prefix(self):
        encoded = sp.encode_sharing_url("https://example.sharepoint.com/x")
        self.assertTrue(encoded.startswith("u!"))
        self.assertNotIn("/", encoded)
        self.assertNotIn("+", encoded)


class ConfigTests(unittest.TestCase):
    def test_auth_and_intended_flags(self):
        env = {
            "SHAREPOINT_VENDOR_PROD_URL": sp.DEFAULT_SHAREPOINT_VENDOR_PROD_URL,
            "AZURE_TENANT_ID": "",
            "AZURE_CLIENT_ID": "",
            "AZURE_CLIENT_SECRET": "",
            "SHAREPOINT_GRAPH_TOKEN": "",
            "SHAREPOINT_REQUIRE": "",
        }
        with patch.dict("os.environ", env, clear=False):
            # Ensure the keys we care about win even if the runner has leftovers
            for k, v in env.items():
                if v == "":
                    pass
            cfg = sp.SharePointConfig.from_env()
            cfg.tenant_id = ""
            cfg.client_id = ""
            cfg.client_secret = ""
            cfg.graph_token = ""
            cfg.url_from_env = True
            cfg.require = False
            self.assertFalse(cfg.auth_configured())
            self.assertTrue(sp.sharepoint_write_intended(cfg))
            public = cfg.public_status()
            self.assertFalse(public["auth_configured"])
            self.assertNotIn("secret", str(public).lower())

    def test_client_credentials_auth_configured(self):
        cfg = sp.SharePointConfig(
            tenant_id="t",
            client_id="c",
            client_secret="s",
        )
        self.assertTrue(cfg.auth_configured())
        self.assertEqual(cfg.public_status()["auth_mode"], "client_credentials")


class GraphClientTests(unittest.TestCase):
    def test_token_error_is_actionable(self):
        cfg = sp.SharePointConfig(tenant_id="t", client_id="c", client_secret="s")

        class Resp:
            status_code = 401
            text = "denied"

            def json(self):
                return {"error": "invalid_client", "error_description": "bad secret"}

        with patch("sharepoint_vendor_prod.requests.post", return_value=Resp()):
            with self.assertRaises(sp.SharePointError) as ctx:
                sp.acquire_graph_token(cfg)
        self.assertIn("Azure AD refused", str(ctx.exception))
        self.assertIn("AZURE_CLIENT_SECRET", ctx.exception.hint)
        self.assertNotIn("client_secret", str(ctx.exception).lower())
        dumped = str(ctx.exception.to_dict())
        self.assertNotIn("client_secret=s", dumped)

    def test_download_upload_mocked(self):
        cfg = sp.SharePointConfig(graph_token="tok", drive_id="d1", item_id="i1")
        xlsx = b"PK\x03\x04" + b"fake-xlsx"

        class MetaResp:
            status_code = 200
            content = b'{"id":"i1"}'

            def json(self):
                return {
                    "id": "i1",
                    "name": "Vendor Prod Files.xlsx",
                    "eTag": '"abc"',
                    "size": len(xlsx),
                    "webUrl": "https://example/file",
                    "parentReference": {"driveId": "d1"},
                }

        class ContentResp:
            status_code = 200
            content = xlsx

            def json(self):
                return {}

        class PutResp:
            status_code = 200
            content = b"{}"

            def json(self):
                return {
                    "id": "i1",
                    "name": "Vendor Prod Files.xlsx",
                    "eTag": '"def"',
                    "size": len(xlsx),
                    "webUrl": "https://example/file",
                    "parentReference": {"driveId": "d1"},
                }

        def fake_request(method, url, **kwargs):
            if method == "GET" and url.endswith("/items/i1"):
                return MetaResp()
            if method == "GET" and url.endswith("/content"):
                return ContentResp()
            if method == "PUT" and url.endswith("/content"):
                return PutResp()
            raise AssertionError(f"unexpected {method} {url}")

        client = sp.SharePointClient(cfg)
        with patch.object(client, "_request", side_effect=fake_request):
            data, item = client.download()
            self.assertEqual(data, xlsx)
            self.assertEqual(item.drive_id, "d1")
            updated = client.upload(xlsx, item)
            self.assertEqual(updated.item_id, "i1")

    def test_missing_auth_download_fails_loudly(self):
        cfg = sp.SharePointConfig()
        cfg.graph_token = ""
        cfg.tenant_id = ""
        cfg.client_id = ""
        cfg.client_secret = ""
        with self.assertRaises(sp.SharePointError) as ctx:
            sp.download_vendor_prod_workbook(cfg)
        self.assertEqual(ctx.exception.status_code, 503)
        self.assertIn("Graph auth", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
