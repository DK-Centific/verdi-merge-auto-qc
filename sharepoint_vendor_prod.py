"""Microsoft Graph download/upload for the live Vendor Prod Files workbook.

Production destination (Rate Approval Log):
https://digitaltechedge.sharepoint.com/:x:/r/sites/Verdi_Juniper/Shared%20Documents/Vendor%20Audio%20Production/Vendor%20Prod%20Files.xlsx

Auth (any one path):
  1. SHAREPOINT_GRAPH_TOKEN — pre-issued Graph bearer token
  2. AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET — app-only client credentials

Target (optional overrides; defaults match the live workbook):
  SHAREPOINT_VENDOR_PROD_URL
  SHAREPOINT_SITE_HOSTNAME / SHAREPOINT_SITE_PATH / SHAREPOINT_FILE_PATH
  SHAREPOINT_DRIVE_ID / SHAREPOINT_ITEM_ID

This module never logs tokens or secrets.
"""
from __future__ import annotations

import base64
import os
import re
from dataclasses import dataclass, field
from typing import Any, Optional
from urllib.parse import unquote, urlparse

import requests

DEFAULT_SHAREPOINT_VENDOR_PROD_URL = (
    "https://digitaltechedge.sharepoint.com/:x:/r/sites/Verdi_Juniper/"
    "Shared%20Documents/Vendor%20Audio%20Production/Vendor%20Prod%20Files.xlsx"
)
DEFAULT_SITE_HOSTNAME = "digitaltechedge.sharepoint.com"
DEFAULT_SITE_PATH = "sites/Verdi_Juniper"
DEFAULT_FILE_PATH = "Vendor Audio Production/Vendor Prod Files.xlsx"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
LOGIN_ROOT = "https://login.microsoftonline.com"

# Sharing-URL download is attempted only when no Graph auth is configured.
# Most Verdi links are org-restricted and will fail — that is reported loudly.
_ANON_UA = "Verdi-Merge-Auto-QC/1.0"


class SharePointError(Exception):
    """Actionable SharePoint / Graph failure. Never include secrets in `message`."""

    def __init__(
        self,
        message: str,
        *,
        hint: str = "",
        status_code: int = 502,
        details: Optional[dict[str, Any]] = None,
    ):
        super().__init__(message)
        self.hint = hint
        self.status_code = status_code
        self.details = details or {}

    def to_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"error": str(self), "destination": "sharepoint"}
        if self.hint:
            out["hint"] = self.hint
        if self.details:
            out["details"] = self.details
        return out


def _env(name: str, default: str = "") -> str:
    return (os.environ.get(name) or default).strip()


def _truthy(name: str) -> bool:
    return _env(name).lower() in ("1", "true", "yes", "on")


def encode_sharing_url(url: str) -> str:
    """Graph shares API id: u! + base64url of the sharing URL."""
    raw = base64.b64encode(url.encode("utf-8")).decode("ascii")
    raw = raw.rstrip("=").replace("/", "_").replace("+", "-")
    return f"u!{raw}"


def parse_sharepoint_xlsx_url(url: str) -> dict[str, str]:
    """Pull hostname / site path / library-relative file path from a SharePoint :x: link."""
    url = (url or "").strip()
    out = {
        "hostname": "",
        "site_path": "",
        "library": "Shared Documents",
        "file_path": "",
        "url": url,
    }
    if not url:
        return out
    parsed = urlparse(url)
    out["hostname"] = parsed.netloc
    path = unquote(parsed.path or "")
    # /:x:/r/sites/Verdi_Juniper/Shared Documents/Vendor Audio Production/Vendor Prod Files.xlsx
    path = re.sub(r"^/?:[a-z]:/[a-z]/", "/", path, flags=re.I)
    m = re.search(r"/(sites/[^/]+)(?:/(.*))?$", path, flags=re.I)
    if m:
        out["site_path"] = m.group(1).strip("/")
        remainder = (m.group(2) or "").strip("/")
        if remainder:
            parts = remainder.split("/", 1)
            out["library"] = parts[0] or "Shared Documents"
            out["file_path"] = parts[1] if len(parts) > 1 else ""
    return out


@dataclass
class SharePointConfig:
    vendor_prod_url: str = DEFAULT_SHAREPOINT_VENDOR_PROD_URL
    url_from_env: bool = False
    site_hostname: str = DEFAULT_SITE_HOSTNAME
    site_path: str = DEFAULT_SITE_PATH
    file_path: str = DEFAULT_FILE_PATH
    drive_id: str = ""
    item_id: str = ""
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    graph_token: str = ""
    timeout: float = 60.0
    require: bool = False

    @classmethod
    def from_env(cls) -> "SharePointConfig":
        url_env = _env("SHAREPOINT_VENDOR_PROD_URL")
        url = url_env or DEFAULT_SHAREPOINT_VENDOR_PROD_URL
        parsed = parse_sharepoint_xlsx_url(url)
        timeout_raw = _env("SHAREPOINT_TIMEOUT", "60")
        try:
            timeout = float(timeout_raw)
        except ValueError:
            timeout = 60.0
        return cls(
            vendor_prod_url=url,
            url_from_env=bool(url_env),
            site_hostname=_env("SHAREPOINT_SITE_HOSTNAME") or parsed["hostname"] or DEFAULT_SITE_HOSTNAME,
            site_path=_env("SHAREPOINT_SITE_PATH") or parsed["site_path"] or DEFAULT_SITE_PATH,
            file_path=_env("SHAREPOINT_FILE_PATH") or parsed["file_path"] or DEFAULT_FILE_PATH,
            drive_id=_env("SHAREPOINT_DRIVE_ID"),
            item_id=_env("SHAREPOINT_ITEM_ID"),
            tenant_id=_env("AZURE_TENANT_ID") or _env("SHAREPOINT_TENANT_ID"),
            client_id=_env("AZURE_CLIENT_ID") or _env("SHAREPOINT_CLIENT_ID"),
            client_secret=_env("AZURE_CLIENT_SECRET") or _env("SHAREPOINT_CLIENT_SECRET"),
            graph_token=_env("SHAREPOINT_GRAPH_TOKEN") or _env("GRAPH_ACCESS_TOKEN"),
            timeout=timeout,
            require=_truthy("SHAREPOINT_REQUIRE"),
        )

    def auth_configured(self) -> bool:
        if self.graph_token:
            return True
        return bool(self.tenant_id and self.client_id and self.client_secret)

    def public_status(self) -> dict[str, Any]:
        """Safe for /api/vendor-prod/status — no secrets."""
        return {
            "url": self.vendor_prod_url,
            "url_from_env": self.url_from_env,
            "site_hostname": self.site_hostname,
            "site_path": self.site_path,
            "file_path": self.file_path,
            "drive_id_set": bool(self.drive_id),
            "item_id_set": bool(self.item_id),
            "auth_configured": self.auth_configured(),
            "auth_mode": (
                "graph_token"
                if self.graph_token
                else ("client_credentials" if self.auth_configured() else "none")
            ),
            "require": self.require,
        }


def sharepoint_write_intended(cfg: Optional[SharePointConfig] = None) -> bool:
    """True when Send must target SharePoint (not silent local)."""
    cfg = cfg or SharePointConfig.from_env()
    return cfg.auth_configured() or cfg.url_from_env or cfg.require


AUTH_HINT = (
    "Set Azure app-only credentials for Microsoft Graph, then grant the app "
    "Sites.ReadWrite.All (or Files.ReadWrite.All) on the Verdi_Juniper site and admin-consent. "
    "Required env: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET. "
    "Or set SHAREPOINT_GRAPH_TOKEN to a bearer token that can read/write the workbook. "
    "See README → SharePoint (Vendor Prod Files)."
)


def acquire_graph_token(cfg: SharePointConfig) -> str:
    if cfg.graph_token:
        return cfg.graph_token
    if not (cfg.tenant_id and cfg.client_id and cfg.client_secret):
        raise SharePointError(
            "SharePoint is the intended destination, but Microsoft Graph auth is not configured.",
            hint=AUTH_HINT,
            status_code=503,
        )
    token_url = f"{LOGIN_ROOT}/{cfg.tenant_id}/oauth2/v2.0/token"
    try:
        r = requests.post(
            token_url,
            data={
                "client_id": cfg.client_id,
                "client_secret": cfg.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=cfg.timeout,
            headers={"User-Agent": _ANON_UA},
        )
    except requests.RequestException as e:
        raise SharePointError(
            f"Could not reach Azure AD token endpoint ({type(e).__name__}).",
            hint="Check outbound HTTPS to login.microsoftonline.com and AZURE_TENANT_ID.",
            status_code=502,
        ) from e
    if r.status_code >= 400:
        raise SharePointError(
            f"Azure AD refused the client-credentials token request (HTTP {r.status_code}).",
            hint=(
                "Check AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET. "
                "The secret is never logged. " + AUTH_HINT
            ),
            status_code=502,
            details={"azure_status": r.status_code, "azure_error": _safe_error_body(r)},
        )
    token = (r.json() or {}).get("access_token")
    if not token:
        raise SharePointError(
            "Azure AD token response did not include access_token.",
            hint=AUTH_HINT,
            status_code=502,
        )
    return str(token)


def _safe_error_body(r: requests.Response) -> dict[str, Any]:
    try:
        data = r.json()
    except Exception:
        return {"text": (r.text or "")[:300]}
    if not isinstance(data, dict):
        return {"text": str(data)[:300]}
    # Strip anything that might echo secrets
    keep = {}
    for k in ("error", "error_description", "error_codes", "error_uri", "code", "message"):
        if k in data:
            keep[k] = data[k]
    inner = data.get("error")
    if isinstance(inner, dict):
        keep["error"] = {k: inner.get(k) for k in ("code", "message") if k in inner}
    return keep


@dataclass
class DriveItemRef:
    drive_id: str
    item_id: str
    name: str = ""
    etag: str = ""
    web_url: str = ""
    size: int = 0


@dataclass
class SharePointClient:
    cfg: SharePointConfig
    _token: str = field(default="", repr=False)

    def _headers(self) -> dict[str, str]:
        if not self._token:
            self._token = acquire_graph_token(self.cfg)
        return {
            "Authorization": f"Bearer {self._token}",
            "Accept": "application/json",
            "User-Agent": _ANON_UA,
        }

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        kwargs.setdefault("timeout", self.cfg.timeout)
        headers = dict(self._headers())
        extra = kwargs.pop("headers", None) or {}
        headers.update(extra)
        try:
            r = requests.request(method, url, headers=headers, **kwargs)
        except requests.RequestException as e:
            raise SharePointError(
                f"Graph {method} failed ({type(e).__name__}).",
                hint="Check outbound HTTPS to graph.microsoft.com and SharePoint site access.",
                status_code=502,
            ) from e
        if r.status_code in (401, 403):
            raise SharePointError(
                f"Graph {method} returned HTTP {r.status_code} (not authorized for this file).",
                hint=(
                    "The Azure app needs Sites.ReadWrite.All or Files.ReadWrite.All "
                    "(application permission) with admin consent, and access to site "
                    f"{self.cfg.site_hostname}/{self.cfg.site_path}. " + AUTH_HINT
                ),
                status_code=r.status_code,
                details=_safe_error_body(r),
            )
        if r.status_code == 404:
            raise SharePointError(
                "SharePoint file or site was not found via Graph.",
                hint=(
                    f"Check SHAREPOINT_SITE_HOSTNAME={self.cfg.site_hostname}, "
                    f"SHAREPOINT_SITE_PATH={self.cfg.site_path}, "
                    f"SHAREPOINT_FILE_PATH={self.cfg.file_path}. "
                    "Confirm the workbook still lives under Shared Documents / "
                    "Vendor Audio Production / Vendor Prod Files.xlsx."
                ),
                status_code=404,
                details=_safe_error_body(r),
            )
        if r.status_code == 412:
            raise SharePointError(
                "SharePoint file changed while we were appending (eTag mismatch).",
                hint="Someone else saved Vendor Prod Files.xlsx. Retry Send.",
                status_code=409,
            )
        if r.status_code == 423:
            raise SharePointError(
                "SharePoint file is locked (open in Excel / another editor).",
                hint="Close Vendor Prod Files.xlsx in the browser or desktop Excel, then retry Send.",
                status_code=423,
            )
        if r.status_code >= 400:
            raise SharePointError(
                f"Graph {method} returned HTTP {r.status_code}.",
                hint="See details. " + AUTH_HINT,
                status_code=502,
                details=_safe_error_body(r),
            )
        return r

    def resolve_item(self) -> DriveItemRef:
        if self.cfg.drive_id and self.cfg.item_id:
            data = self._request(
                "GET",
                f"{GRAPH_ROOT}/drives/{self.cfg.drive_id}/items/{self.cfg.item_id}",
            ).json()
            return self._item_from_json(data)

        # Prefer sharing URL when present — works with delegated tokens and some app setups
        if self.cfg.vendor_prod_url:
            share_id = encode_sharing_url(self.cfg.vendor_prod_url)
            try:
                data = self._request("GET", f"{GRAPH_ROOT}/shares/{share_id}/driveItem").json()
                return self._item_from_json(data)
            except SharePointError as e:
                # Fall through to site-path resolve unless it was a hard auth failure
                if e.status_code in (401, 403):
                    raise

        site_id = self._resolve_site_id()
        path = self.cfg.file_path.lstrip("/")
        data = self._request(
            "GET",
            f"{GRAPH_ROOT}/sites/{site_id}/drive/root:/{path}",
        ).json()
        return self._item_from_json(data)

    def _resolve_site_id(self) -> str:
        host = self.cfg.site_hostname
        site = self.cfg.site_path if self.cfg.site_path.startswith("/") else f"/{self.cfg.site_path}"
        data = self._request("GET", f"{GRAPH_ROOT}/sites/{host}:{site}").json()
        site_id = data.get("id")
        if not site_id:
            raise SharePointError(
                "Graph site lookup returned no id.",
                hint=f"Confirm the site {host}{site} exists and the app can read it.",
                status_code=502,
            )
        return str(site_id)

    def _item_from_json(self, data: dict[str, Any]) -> DriveItemRef:
        parent = data.get("parentReference") or {}
        drive_id = str(parent.get("driveId") or self.cfg.drive_id or "")
        item_id = str(data.get("id") or self.cfg.item_id or "")
        if not drive_id or not item_id:
            raise SharePointError(
                "Graph driveItem response was missing driveId or item id.",
                hint="Set SHAREPOINT_DRIVE_ID and SHAREPOINT_ITEM_ID, or check the sharing URL.",
                status_code=502,
            )
        return DriveItemRef(
            drive_id=drive_id,
            item_id=item_id,
            name=str(data.get("name") or ""),
            etag=str(data.get("eTag") or data.get("@odata.etag") or ""),
            web_url=str(data.get("webUrl") or ""),
            size=int(data.get("size") or 0),
        )

    def download(self) -> tuple[bytes, DriveItemRef]:
        item = self.resolve_item()
        r = self._request(
            "GET",
            f"{GRAPH_ROOT}/drives/{item.drive_id}/items/{item.item_id}/content",
        )
        data = r.content
        if not data or data[:4] != b"PK\x03\x04":
            raise SharePointError(
                "Downloaded SharePoint item is not an .xlsx workbook.",
                hint="Confirm SHAREPOINT_FILE_PATH points at Vendor Prod Files.xlsx.",
                status_code=502,
            )
        return data, item

    def upload(self, content: bytes, item: DriveItemRef) -> DriveItemRef:
        if not content:
            raise SharePointError("Refusing to upload empty workbook bytes.", status_code=500)
        headers = {}
        if item.etag:
            headers["If-Match"] = item.etag
        # Simple upload is valid under 4 MB; this workbook is typically ~100 KB.
        # Use an upload session for larger payloads.
        if len(content) < 4_000_000:
            r = self._request(
                "PUT",
                f"{GRAPH_ROOT}/drives/{item.drive_id}/items/{item.item_id}/content",
                data=content,
                headers={**headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            )
            return self._item_from_json(r.json()) if r.content else item

        session = self._request(
            "POST",
            f"{GRAPH_ROOT}/drives/{item.drive_id}/items/{item.item_id}/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        ).json()
        upload_url = session.get("uploadUrl")
        if not upload_url:
            raise SharePointError("Graph createUploadSession did not return uploadUrl.", status_code=502)
        total = len(content)
        # uploadUrl is pre-authorized — do not attach the Graph bearer token
        try:
            r = requests.put(
                upload_url,
                data=content,
                headers={
                    "Content-Length": str(total),
                    "Content-Range": f"bytes 0-{total - 1}/{total}",
                    "User-Agent": _ANON_UA,
                },
                timeout=max(self.cfg.timeout, 120.0),
            )
        except requests.RequestException as e:
            raise SharePointError(
                f"Graph upload session failed ({type(e).__name__}).",
                hint="Retry Send. If it keeps failing, the file may be locked in Excel.",
                status_code=502,
            ) from e
        if r.status_code >= 400:
            raise SharePointError(
                f"Graph upload session returned HTTP {r.status_code}.",
                hint="Close the workbook if it is open, then retry Send.",
                status_code=502,
                details=_safe_error_body(r),
            )
        try:
            return self._item_from_json(r.json())
        except Exception:
            return item


def download_vendor_prod_workbook(cfg: Optional[SharePointConfig] = None) -> tuple[bytes, dict[str, Any]]:
    cfg = cfg or SharePointConfig.from_env()
    if not cfg.auth_configured():
        raise SharePointError(
            "Cannot download Vendor Prod Files.xlsx from SharePoint without Graph auth.",
            hint=AUTH_HINT,
            status_code=503,
        )
    client = SharePointClient(cfg)
    data, item = client.download()
    meta = {
        "destination": "sharepoint",
        "name": item.name,
        "web_url": item.web_url,
        "size": item.size,
        "drive_id": item.drive_id,
        "item_id": item.item_id,
        "etag": item.etag,
        "site_hostname": cfg.site_hostname,
        "site_path": cfg.site_path,
        "file_path": cfg.file_path,
    }
    return data, meta


def upload_vendor_prod_workbook(
    content: bytes,
    item: DriveItemRef,
    cfg: Optional[SharePointConfig] = None,
) -> dict[str, Any]:
    cfg = cfg or SharePointConfig.from_env()
    client = SharePointClient(cfg)
    updated = client.upload(content, item)
    return {
        "destination": "sharepoint",
        "name": updated.name,
        "web_url": updated.web_url,
        "size": updated.size or len(content),
        "drive_id": updated.drive_id,
        "item_id": updated.item_id,
        "etag": updated.etag,
    }
