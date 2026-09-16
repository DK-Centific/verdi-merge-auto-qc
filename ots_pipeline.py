"""OTS agency intake: control → vendor sources → consolidate → OneForma QA."""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import parse_qs, urlparse

import openpyxl
import requests
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "template.xlsx"
FIXTURES = ROOT / "fixtures"
SAMPLES = ROOT / "samples"

OTS_COLUMNS = [
    "status",
    "feedback",
    "folderName",
    "fileName",
    "locale",
    "JSONfile",
    "audiofile",
    "ICF_speaker_1",
    "ICF_speaker_2",
    "speaker_1_cumulativeAudioMinutes",
    "speaker_2_cumulativeAudioMinutes",
    "totaldurationMinute",
    "ingestionBatch",
    "durationSeconds",
]

# Header aliases → OTS column
HEADER_ALIASES = {
    "status": "status",
    "feedback": "feedback",
    "foldername": "folderName",
    "folder name": "folderName",
    "folder": "folderName",
    "filename": "fileName",
    "file name": "fileName",
    "file": "fileName",
    "locale": "locale",
    "jsonfile": "JSONfile",
    "json file": "JSONfile",
    "json": "JSONfile",
    "audiofile": "audiofile",
    "audio file": "audiofile",
    "audio": "audiofile",
    "icf_speaker_1": "ICF_speaker_1",
    "icf speaker 1": "ICF_speaker_1",
    "icf_speaker_2": "ICF_speaker_2",
    "icf speaker 2": "ICF_speaker_2",
    "speaker_1_cumulativeaudiominutes": "speaker_1_cumulativeAudioMinutes",
    "speaker_2_cumulativeaudiominutes": "speaker_2_cumulativeAudioMinutes",
    "totaldurationminute": "totaldurationMinute",
    "total duration minute": "totaldurationMinute",
    "totaldurationminutes": "totaldurationMinute",
    "duration": "totaldurationMinute",
    "duration_minutes": "totaldurationMinute",
    "qa_status": "status",
    "qastatus": "status",
    "jsonfileurl": "JSONfile",
    "json file url": "JSONfile",
    "audiofileurl": "audiofile",
    "audio file url": "audiofile",
}

ONEFORMA_BASE = "https://dfsupport.oneforma.com/pipeline-api/api/ots-agency-vendor-blob-files"
LogFn = Callable[[str], None]


def _log(log: Optional[LogFn], msg: str) -> None:
    if log:
        log(msg)


def _norm_header(h: Any) -> str:
    if h is None:
        return ""
    s = str(h).strip().lower().replace("-", "_")
    s = re.sub(r"\s+", " ", s)
    return s


def mask_secret(value: Any) -> str:
    if value is None or str(value).strip() == "":
        return ""
    return "***MASKED***"


def normalize_locale(raw: Any) -> tuple[str, str]:
    """Return (normalized, raw_string)."""
    if raw is None:
        return "", ""
    raw_s = str(raw).strip()
    if not raw_s:
        return "", ""
    # Strip parenthetical / English suffix common in folder spellings
    base = raw_s.split("(")[0].strip()
    base = re.sub(r"_English.*$", "", base, flags=re.I)
    base = re.sub(r"\s*English\s*$", "", base, flags=re.I)
    base = base.strip().replace(" ", "")
    # he- IL → he-IL; el_GR → el-GR
    base = base.replace("__", "_").replace("_", "-")
    base = re.sub(r"-+", "-", base)
    m = re.match(r"^([A-Za-z]{2,3})-([A-Za-z0-9]+)", base)
    if m:
        return f"{m.group(1).lower()}-{m.group(2).upper() if len(m.group(2)) <= 3 else m.group(2)}", raw_s
    # MSA special
    if re.match(r"^ar-?msa$", base, re.I):
        return "ar-MSA", raw_s
    return base, raw_s


def container_to_vendor_direction(container: str) -> tuple[str, str]:
    c = (container or "").strip().lower()
    for direction in ("handoff", "handback"):
        if c.endswith(direction):
            vendor = c[: -len(direction)]
            # known container quirks
            if vendor == "avyaa2":
                vendor = "avyaan"
            elif vendor == "avyaan2":
                vendor = "avyaan"
            return vendor, direction
    return c, ""


def vendor_api_name(vendor: str) -> str:
    v = (vendor or "").strip().lower()
    mapping = {
        "avyaan": "avyaan2",
        "avyaa2": "avyaan2",
        "consultbae": "consultbae",
        "language solution asia": "languagesolutionasia",
    }
    return mapping.get(v, v.replace(" ", ""))


@dataclass
class VendorControl:
    name: str
    container: str = ""
    sas_masked: str = ""
    expire: str = ""
    comments: str = ""
    pocs_masked: str = ""
    drive_link: str = ""
    has_password: bool = False
    # password kept only in-memory during run; never logged/serialized
    _password: str = field(default="", repr=False)


@dataclass
class SourceStatus:
    vendor: str
    status: str  # ok | source_unavailable | needs_live_fetch | fallback_upload | skipped
    detail: str = ""
    rows_loaded: int = 0


@dataclass
class QAFinding:
    severity: str
    vendor: str
    fileName: str
    folderName: str
    locale: str
    issue: str
    detail: str = ""


def parse_control_workbook(data: bytes, log: Optional[LogFn] = None) -> list[VendorControl]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [_norm_header(h) for h in rows[0]]

    def col(*names: str) -> Optional[int]:
        for n in names:
            n = _norm_header(n)
            if n in headers:
                return headers.index(n)
        return None

    i_name = col("vendor name", "vendor", "name")
    i_container = col("vendor blob container name", "blob container name", "container")
    i_sas = col("blob sas key", "sas key", "sas")
    i_expire = col("expire date", "expire", "expiry")
    i_comments = col("comments", "comment")
    i_pocs = col("pocs", "poc")
    i_link = col("link to google drive", "google drive", "link", "drive link")
    i_pw = col("password", "pwd", "pass")
    # Column H fallback (0-based index 7) if header missing
    if i_pw is None and len(headers) >= 8:
        i_pw = 7

    vendors: list[VendorControl] = []
    for row in rows[1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        name = str(row[i_name]).strip() if i_name is not None and row[i_name] else ""
        if not name:
            continue
        pw = ""
        if i_pw is not None and i_pw < len(row) and row[i_pw]:
            pw = str(row[i_pw]).strip()
        vc = VendorControl(
            name=name,
            container=str(row[i_container]).strip() if i_container is not None and row[i_container] else "",
            sas_masked=mask_secret(row[i_sas] if i_sas is not None and i_sas < len(row) else None),
            expire=str(row[i_expire]).strip() if i_expire is not None and row[i_expire] else "",
            comments=str(row[i_comments]).strip() if i_comments is not None and row[i_comments] else "",
            pocs_masked=mask_secret(row[i_pocs] if i_pocs is not None and i_pocs < len(row) else None),
            drive_link=str(row[i_link]).strip() if i_link is not None and row[i_link] else "",
            has_password=bool(pw),
            _password=pw,
        )
        vendors.append(vc)
    _log(log, f"Control sheet: {len(vendors)} vendor row(s)")
    return vendors


def _map_headers(header_row: list[Any]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, h in enumerate(header_row):
        key = HEADER_ALIASES.get(_norm_header(h))
        if key and key not in mapping:
            mapping[key] = idx
        # Also match exact OTS names
        hn = str(h).strip() if h else ""
        if hn in OTS_COLUMNS and hn not in mapping:
            mapping[hn] = idx
    return mapping


def _is_instruction_row(values: list[Any]) -> bool:
    joined = " ".join(str(v).lower() for v in values if v)
    needles = ("you will specify", "just add yes", "we will fill", "instructions")
    return any(n in joined for n in needles)


def extract_ots_rows_from_workbook(data: bytes, vendor: str, log: Optional[LogFn] = None) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    out: list[dict[str, Any]] = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # Find header row (first with fileName/folderName-like cells)
        header_idx = None
        colmap: dict[str, int] = {}
        for i, row in enumerate(rows[:10]):
            cm = _map_headers(list(row))
            if "fileName" in cm or "folderName" in cm or ("locale" in cm and "status" in cm):
                header_idx = i
                colmap = cm
                break
            # Template style: "Column name:" in A, headers in rest
            if row and str(row[0] or "").strip().lower().startswith("column name"):
                cm = _map_headers(list(row))
                if len(cm) >= 3:
                    header_idx = i
                    colmap = cm
                    break
        if header_idx is None:
            continue
        for row in rows[header_idx + 1 :]:
            vals = list(row)
            if all(v is None or str(v).strip() == "" for v in vals):
                continue
            if _is_instruction_row(vals):
                continue
            rec = {c: "" for c in OTS_COLUMNS}
            for col, idx in colmap.items():
                if idx < len(vals) and vals[idx] is not None:
                    rec[col] = vals[idx]
            # Require at least a filename or folder
            if not str(rec.get("fileName") or "").strip() and not str(rec.get("folderName") or "").strip():
                continue
            loc_norm, loc_raw = normalize_locale(rec.get("locale"))
            if loc_norm:
                rec["locale"] = loc_norm
            if loc_raw and loc_raw != loc_norm:
                fb = str(rec.get("feedback") or "")
                note = f"localeRaw={loc_raw}"
                rec["feedback"] = f"{fb}; {note}".strip("; ") if fb else note
            rec["_vendor"] = vendor
            out.append(rec)
    _log(log, f"  Extracted {len(out)} row(s) from workbook for {vendor}")
    return out


def google_sheet_export_urls(link: str) -> list[str]:
    """Build candidate export URLs for a Google Sheets / Drive link."""
    urls = []
    link = (link or "").strip()
    if not link:
        return urls
    # Already an export link
    if "export?" in link or link.endswith(".xlsx"):
        urls.append(link)
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", link)
    if m:
        sid = m.group(1)
        urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx")
        urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv")
    m = re.search(r"/file/d/([a-zA-Z0-9-_]+)", link)
    if m:
        fid = m.group(1)
        urls.append(f"https://drive.google.com/uc?export=download&id={fid}")
    if link not in urls:
        urls.append(link)
    return urls


def fetch_google_source(link: str, log: Optional[LogFn] = None) -> tuple[Optional[bytes], str]:
    session = requests.Session()
    session.headers.update({"User-Agent": "OTS-Agency-QA-Demo/1.0"})
    last_detail = "no urls"
    for url in google_sheet_export_urls(link):
        try:
            _log(log, f"  Trying source URL…")
            r = session.get(url, timeout=20, allow_redirects=True)
            if r.status_code == 404:
                last_detail = "HTTP 404 (file does not exist)"
                continue
            if r.status_code >= 400:
                last_detail = f"HTTP {r.status_code}"
                continue
            ctype = (r.headers.get("Content-Type") or "").lower()
            content = r.content
            if not content:
                last_detail = "empty body"
                continue
            # HTML interstitial / login
            if "text/html" in ctype and not content[:4] == b"PK\x03\x04":
                last_detail = "HTML response (likely login wall or missing file)"
                continue
            return content, f"fetched ({len(content)} bytes)"
        except requests.RequestException as e:
            last_detail = f"request error: {type(e).__name__}"
            continue
    return None, last_detail


def fetch_pangeanic_stub(link: str, has_password: bool, log: Optional[LogFn] = None) -> tuple[Optional[bytes], str]:
    """Demo: do not hang on passworded Pangeanic file manager. Clear stub status."""
    _log(log, "  Pangeanic: passworded file manager — stubbing (needs live fetch)")
    # Quick HEAD/GET to see if reachable, but never submit password or store it
    try:
        r = requests.get(link, timeout=8, allow_redirects=True, headers={"User-Agent": "OTS-Agency-QA-Demo/1.0"})
        reachable = f"HTTP {r.status_code}"
    except requests.RequestException as e:
        reachable = f"unreachable ({type(e).__name__})"
    detail = (
        f"needs_live_fetch; site {reachable}; "
        f"password_present_in_control={has_password} (not used/stored in demo)"
    )
    return None, detail



def resolve_gdrive_sync(vendor_name: str) -> Optional[Path]:
    """Case-insensitive match under samples/gdrive_sync/{vendor}.xlsx."""
    sync_dir = SAMPLES / "gdrive_sync"
    if not sync_dir.is_dir():
        return None
    key = (vendor_name or "").strip().lower()
    if not key:
        return None
    for p in sync_dir.glob("*.xlsx"):
        if p.stem.lower() == key:
            return p
    return None


def load_vendor_sources(
    vendors: list[VendorControl],
    fallback_files: dict[str, bytes],
    log: Optional[LogFn] = None,
) -> tuple[list[dict[str, Any]], list[SourceStatus]]:
    """fallback_files: vendor_name_lower → xlsx bytes."""
    all_rows: list[dict[str, Any]] = []
    statuses: list[SourceStatus] = []

    for vc in vendors:
        key = vc.name.strip().lower()
        _log(log, f"Vendor: {vc.name}")

        if key in fallback_files:
            rows = extract_ots_rows_from_workbook(fallback_files[key], vc.name, log=log)
            all_rows.extend(rows)
            statuses.append(SourceStatus(vc.name, "fallback_upload", "using uploaded vendor xlsx", len(rows)))
            continue

        link = vc.drive_link or ""
        if not link:
            statuses.append(SourceStatus(vc.name, "source_unavailable", "no Column G link and no fallback upload", 0))
            continue

        if "pangeanic" in link.lower() or "cv.pangeanic" in link.lower():
            _data, detail = fetch_pangeanic_stub(link, vc.has_password, log=log)
            # Automatic local sample when live passworded filemgr is unavailable
            sample = SAMPLES / "pangeanic_maple_delivery_tracking.xlsx"
            if sample.exists():
                rows = extract_ots_rows_from_workbook(sample.read_bytes(), vc.name, log=log)
                all_rows.extend(rows)
                statuses.append(
                    SourceStatus(
                        vc.name,
                        "fallback_upload",
                        f"local sample {sample.name} (vendor=pangeanic container=pangeanichandoff); {detail}",
                        len(rows),
                    )
                )
                continue
            statuses.append(SourceStatus(vc.name, "needs_live_fetch", detail, 0))
            continue

        if "docs.google.com" in link or "drive.google.com" in link or "spreadsheets" in link:
            sync_path = resolve_gdrive_sync(vc.name)
            if sync_path is not None:
                rows = extract_ots_rows_from_workbook(sync_path.read_bytes(), vc.name, log=log)
                all_rows.extend(rows)
                statuses.append(
                    SourceStatus(
                        vc.name,
                        "ok",
                        f"source=gdrive_sync; {sync_path.relative_to(ROOT)}",
                        len(rows),
                    )
                )
                continue
            data, detail = fetch_google_source(link, log=log)
            if data is None:
                statuses.append(SourceStatus(vc.name, "source_unavailable", detail, 0))
                continue
            # CSV vs xlsx
            if data[:4] != b"PK\x03\x04":
                # wrap CSV into a minimal workbook
                try:
                    text = data.decode("utf-8-sig", errors="replace")
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    for line in text.splitlines():
                        # crude CSV
                        import csv as _csv

                        pass
                    reader = __import__("csv").reader(io.StringIO(text))
                    for crow in reader:
                        ws.append(crow)
                    buf = io.BytesIO()
                    wb.save(buf)
                    data = buf.getvalue()
                except Exception as e:
                    statuses.append(SourceStatus(vc.name, "source_unavailable", f"csv parse failed: {e}", 0))
                    continue
            rows = extract_ots_rows_from_workbook(data, vc.name, log=log)
            all_rows.extend(rows)
            statuses.append(SourceStatus(vc.name, "ok", detail, len(rows)))
            continue

        # Unknown link type — try GET once
        data, detail = fetch_google_source(link, log=log)
        if data and data[:4] == b"PK\x03\x04":
            rows = extract_ots_rows_from_workbook(data, vc.name, log=log)
            all_rows.extend(rows)
            statuses.append(SourceStatus(vc.name, "ok", detail, len(rows)))
        else:
            statuses.append(SourceStatus(vc.name, "source_unavailable", detail or "unsupported link", 0))

    # Clear passwords from memory
    for vc in vendors:
        vc._password = ""

    return all_rows, statuses


def oneforma_get(path: str, params: Optional[dict] = None, timeout: int = 45) -> dict:
    url = f"{ONEFORMA_BASE}/{path.lstrip('/')}"
    r = requests.get(
        url,
        params=params or {},
        timeout=timeout,
        headers={"User-Agent": "OTS-Agency-QA-Demo/1.0", "Accept": "application/json"},
    )
    r.raise_for_status()
    return r.json()


def fetch_oneforma_files(
    vendors: Optional[list[str]] = None,
    page_size: int = 200,
    log: Optional[LogFn] = None,
    use_fixture_on_fail: bool = True,
) -> list[dict]:
    rows: list[dict] = []
    try:
        if vendors:
            for v in vendors:
                page = 1
                while True:
                    _log(log, f"OneForma fetch vendor={v} page={page}")
                    data = oneforma_get(
                        "files",
                        {
                            "vendor": v,
                            "page": page,
                            "pageSize": page_size,
                            "includeNonAudio": "true",
                        },
                    )
                    batch = data.get("rows") or []
                    rows.extend(batch)
                    total = int(data.get("total") or 0)
                    if page * page_size >= total or not batch:
                        break
                    page += 1
        else:
            page = 1
            while True:
                _log(log, f"OneForma fetch all page={page}")
                data = oneforma_get(
                    "files",
                    {"page": page, "pageSize": page_size, "includeNonAudio": "true"},
                )
                batch = data.get("rows") or []
                rows.extend(batch)
                total = int(data.get("total") or 0)
                if page * page_size >= total or not batch:
                    break
                page += 1
        _log(log, f"OneForma: {len(rows)} blob file row(s)")
        return rows
    except Exception as e:
        _log(log, f"OneForma live fetch failed: {e}")
        if use_fixture_on_fail:
            fix = FIXTURES / "oneforma_alchemy_page1.json"
            if fix.exists():
                data = json.loads(fix.read_text())
                rows = data.get("rows") or []
                _log(log, f"Using fixture {fix.name}: {len(rows)} row(s)")
                return rows
        raise


def _minutes_from_seconds(sec: Any) -> Optional[float]:
    try:
        if sec is None or sec == "":
            return None
        return round(float(sec) / 60.0, 2)
    except (TypeError, ValueError):
        return None


def _to_float(v: Any) -> Optional[float]:
    try:
        if v is None or v == "":
            return None
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None



def _parse_ingestion_batch(blob: Optional[dict] = None, folder_name: Any = "") -> str:
    """Return YYYY-MM-DD ingestion batch date string, or empty."""
    blob = blob or {}
    date_iso = str(blob.get("dateIso") or "").strip()
    if date_iso:
        # Accept YYYY-MM-DD already
        if re.match(r"^\d{4}-\d{2}-\d{2}", date_iso):
            return date_iso[:10]
    date_raw = str(blob.get("date") or "").strip()
    if re.match(r"^\d{8}$", date_raw):
        return f"{date_raw[0:4]}-{date_raw[4:6]}-{date_raw[6:8]}"
    # Derive from folder trailing /YYYYMMDD (optionally before /passed|/failed)
    folder = str(folder_name or blob.get("folder") or "").strip().replace("\\", "/")
    m = re.search(r"(?:^|/)(\d{8})(?:/|$)", folder)
    if m:
        d = m.group(1)
        return f"{d[0:4]}-{d[4:6]}-{d[6:8]}"
    return ""


def _duration_seconds_for_row(blob: Optional[dict], totalduration_minute: Any) -> Optional[float]:
    blob = blob or {}
    raw = blob.get("durationSeconds")
    if raw is not None and raw != "":
        try:
            return float(raw)
        except (TypeError, ValueError):
            pass
    mins = _to_float(totalduration_minute)
    if mins is not None:
        return round(mins * 60.0, 3)
    return None


def _enrich_row_from_blob(r: dict[str, Any], blob: Optional[dict], vendor: str) -> None:
    """Stash Rate Approval Log fields + ingestionBatch / durationSeconds on annotated row."""
    blob = blob or {}
    # vendorName: prefer tracker vendor, else blob vendor
    r["vendorName"] = vendor or str(blob.get("vendor") or "") or str(r.get("vendorName") or "")
    wf = str(blob.get("workflow") or "").strip() or "Maple"
    r["projectCode"] = wf
    r["workflow"] = wf

    # locale: prefer normalized OneForma locale when present
    blob_loc = str(blob.get("locale") or "").strip()
    if blob_loc:
        b_norm, _ = normalize_locale(blob_loc)
        if b_norm:
            r["locale"] = b_norm

    folder = str(r.get("folderName") or "").strip()
    fn = str(r.get("fileName") or "").strip()
    full = str(blob.get("fullPath") or "").strip()
    if full:
        r["filePath"] = full
    elif folder and fn and ("/" in folder or "\\" in folder or folder.lower().startswith("maple")):
        r["filePath"] = f"{folder.rstrip('/')}/{fn}"
    else:
        r["filePath"] = r.get("filePath") or ""

    batch = _parse_ingestion_batch(blob, folder or blob.get("folder"))
    r["ingestionBatch"] = batch

    dur = _duration_seconds_for_row(blob if blob else None, r.get("totaldurationMinute"))
    r["durationSeconds"] = dur if dur is not None else ""


def run_qa(
    tracker_rows: list[dict[str, Any]],
    blob_rows: list[dict],
    source_statuses: list[SourceStatus],
    log: Optional[LogFn] = None,
) -> tuple[list[dict[str, Any]], list[QAFinding]]:
    """Annotate tracker rows + produce findings (incl. blob-only misses)."""
    findings: list[QAFinding] = []

    for st in source_statuses:
        if st.status in ("source_unavailable", "needs_live_fetch"):
            findings.append(
                QAFinding(
                    severity="warning",
                    vendor=st.vendor,
                    fileName="",
                    folderName="",
                    locale="",
                    issue=st.status,
                    detail=st.detail,
                )
            )

    # Index blobs by filename (lower) and by (folder lower, filename lower)
    by_name: dict[str, list[dict]] = {}
    by_folder_name: dict[tuple[str, str], list[dict]] = {}
    for b in blob_rows:
        fn = str(b.get("filename") or "").strip().lower()
        folder = str(b.get("folder") or "").strip().lower()
        if fn:
            by_name.setdefault(fn, []).append(b)
            by_folder_name.setdefault((folder, fn), []).append(b)

    matched_blob_keys: set[tuple[str, str]] = set()
    annotated: list[dict[str, Any]] = []

    for rec in tracker_rows:
        r = dict(rec)
        vendor = str(r.pop("_vendor", "") or "")
        fn = str(r.get("fileName") or "").strip()
        folder = str(r.get("folderName") or "").strip()
        loc = str(r.get("locale") or "").strip()
        fn_l = fn.lower()
        folder_l = folder.lower()

        status_bits: list[str] = []
        feedback_bits: list[str] = []
        if r.get("feedback"):
            feedback_bits.append(str(r["feedback"]))
        if r.get("status"):
            status_bits.append(str(r["status"]))

        blob = None
        if (folder_l, fn_l) in by_folder_name:
            blob = by_folder_name[(folder_l, fn_l)][0]
        elif fn_l in by_name:
            blob = by_name[fn_l][0]
            # folder mismatch if folders differ
            bfolder = str(blob.get("folder") or "")
            if folder and bfolder and folder_l != bfolder.lower():
                findings.append(
                    QAFinding(
                        "error",
                        vendor,
                        fn,
                        folder,
                        loc,
                        "folder_mismatch",
                        f"tracker={folder} blob={bfolder}",
                    )
                )
                feedback_bits.append(f"folder_mismatch blob={bfolder}")
                status_bits.append("QA:folder_mismatch")

        if blob is None:
            findings.append(
                QAFinding(
                    "error",
                    vendor,
                    fn,
                    folder,
                    loc,
                    "in_tracker_missing_on_blob",
                    "No matching filename on OneForma blob list",
                )
            )
            feedback_bits.append("in_tracker_missing_on_blob")
            status_bits.append("QA:missing_on_blob")
        else:
            matched_blob_keys.add(
                (str(blob.get("folder") or "").lower(), str(blob.get("filename") or "").lower())
            )
            # locale
            blob_loc = str(blob.get("locale") or "")
            blob_raw = str(blob.get("localeRaw") or "")
            t_norm, _ = normalize_locale(loc)
            b_norm, _ = normalize_locale(blob_loc or blob_raw)
            if t_norm and b_norm and t_norm.lower() != b_norm.lower():
                findings.append(
                    QAFinding(
                        "warning",
                        vendor,
                        fn,
                        folder,
                        loc,
                        "locale_mismatch",
                        f"tracker={loc} blob={blob_loc} localeRaw={blob_raw}",
                    )
                )
                feedback_bits.append(f"locale_mismatch blob={blob_loc}; localeRaw={blob_raw}")
                status_bits.append("QA:locale_mismatch")
            elif blob_raw and loc and loc != blob_raw and t_norm.lower() == b_norm.lower():
                # normalized match but raw differed — keep localeRaw note
                if f"localeRaw={blob_raw}" not in ";".join(feedback_bits):
                    feedback_bits.append(f"localeRaw={blob_raw}")

            # filename exact already matched; still flag if case/path quirks
            # duration
            api_min = _minutes_from_seconds(blob.get("durationSeconds"))
            tr_min = _to_float(r.get("totaldurationMinute"))
            if api_min is not None and tr_min is not None:
                if abs(api_min - tr_min) > 0.15:
                    findings.append(
                        QAFinding(
                            "warning",
                            vendor,
                            fn,
                            folder,
                            loc,
                            "duration_mismatch",
                            f"tracker={tr_min} min vs blob={api_min} min ({blob.get('durationSeconds')}s)",
                        )
                    )
                    feedback_bits.append(f"duration_mismatch tracker={tr_min} blob={api_min}")
                    status_bits.append("QA:duration_mismatch")
            elif api_min is not None and tr_min is None:
                r["totaldurationMinute"] = api_min
                feedback_bits.append(f"duration_filled_from_blob={api_min}")

            # vendor/direction from container
            bvendor = str(blob.get("vendor") or "")
            if vendor and bvendor:
                vn = vendor_api_name(vendor)
                if vn != bvendor.lower() and vendor.lower() not in bvendor.lower():
                    findings.append(
                        QAFinding(
                            "warning",
                            vendor,
                            fn,
                            folder,
                            loc,
                            "vendor_mismatch",
                            f"tracker_vendor={vendor} blob_vendor={bvendor}",
                        )
                    )

        # Enrich for Merged sheet + Rate Approval Log mapping
        _enrich_row_from_blob(r, blob, vendor)

        r["status"] = " | ".join(status_bits) if status_bits else r.get("status") or ""
        r["feedback"] = "; ".join(feedback_bits)
        r["_vendor"] = vendor
        annotated.append(r)

    # Blob present, missing in tracker — only for vendors that contributed tracker rows
    # (avoids flooding the demo with hundreds of info rows for untouched vendors).
    tracker_names = {str(r.get("fileName") or "").strip().lower() for r in tracker_rows}
    vendors_with_tracker = {
        vendor_api_name(str(r.get("_vendor") or ""))
        for r in tracker_rows
        if r.get("_vendor")
    }
    # Also include raw lower names
    vendors_with_tracker |= {
        str(r.get("_vendor") or "").strip().lower() for r in tracker_rows if r.get("_vendor")
    }
    for b in blob_rows:
        fn = str(b.get("filename") or "").strip()
        folder = str(b.get("folder") or "").strip()
        key = (folder.lower(), fn.lower())
        if key in matched_blob_keys:
            continue
        if fn.lower() in tracker_names:
            continue  # matched by name already
        bvendor = str(b.get("vendor") or "").strip().lower()
        if vendors_with_tracker and bvendor not in vendors_with_tracker:
            continue
        findings.append(
            QAFinding(
                "info",
                str(b.get("vendor") or ""),
                fn,
                folder,
                str(b.get("locale") or ""),
                "on_blob_missing_in_tracker",
                f"container={b.get('container')}",
            )
        )

    _log(log, f"QA findings: {len(findings)}")
    return annotated, findings


def write_result_workbook(
    annotated_rows: list[dict[str, Any]],
    findings: list[QAFinding],
    source_statuses: list[SourceStatus],
    vendors: list[VendorControl],
    vendor_prod_preview: Optional[dict[str, Any]] = None,
) -> bytes:
    """Build Result_Merged.xlsx: Merged + Vendor_Prod_Rows + QA_Findings + Source_Status + Control_Masked."""
    wb = openpyxl.Workbook()

    # --- Merged sheet from template shape ---
    ws = wb.active
    ws.title = "Merged"
    has_label = False
    data_cols = list(OTS_COLUMNS)
    if TEMPLATE_PATH.exists():
        twb = openpyxl.load_workbook(TEMPLATE_PATH)
        tws = twb[twb.sheetnames[0]]
        header_vals = [tws.cell(1, c).value for c in range(1, tws.max_column + 1)]
        if header_vals and str(header_vals[0] or "").lower().startswith("column name"):
            # Keep label cell + template headers, then append any new OTS cols missing from template
            base = [str(h) for h in header_vals[1:] if h]
            for extra in OTS_COLUMNS:
                if extra not in base:
                    base.append(extra)
            data_cols = base
            ws.append([header_vals[0]] + data_cols)
            has_label = True
        else:
            # Use OTS names from template row if present, else defaults
            cleaned = [str(h) for h in header_vals if h and not str(h).lower().startswith("column")]
            data_cols = cleaned if cleaned else list(OTS_COLUMNS)
            for extra in OTS_COLUMNS:
                if extra not in data_cols:
                    data_cols.append(extra)
            ws.append(data_cols)
    else:
        ws.append(data_cols)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    last = ws.max_column + 1
    ws.cell(1, last, "vendor")
    ws.cell(1, last).font = Font(bold=True)

    for rec in annotated_rows:
        if has_label:
            row_out = [""]
            for h in data_cols:
                row_out.append(rec.get(h, ""))
        else:
            row_out = [rec.get(h, "") for h in data_cols]
        row_out.append(rec.get("_vendor", ""))
        ws.append(row_out)

    # --- Vendor_Prod_Rows (Rate Approval Log fill columns A/B/E/F/G/T/U/V) ---
    wvp = wb.create_sheet("Vendor_Prod_Rows")
    wvp.append(VENDOR_PROD_SHEET_HEADERS)
    for cell in wvp[1]:
        cell.font = Font(bold=True)
    preview = vendor_prod_preview or preview_vendor_prod(annotated_rows)
    ordered_vp = (
        list(preview.get("unique_rows") or [])
        + list(preview.get("duplicate_rows") or [])
        + list(preview.get("incomplete_rows") or [])
    )
    if not ordered_vp:
        for rec in to_rate_approval_rows(annotated_rows):
            wvp.append(
                [
                    rec.get("vendorName", ""),
                    rec.get("projectCode", ""),
                    rec.get("workflow", ""),
                    rec.get("locale", ""),
                    rec.get("ingestionBatch", ""),
                    rec.get("fileName", ""),
                    rec.get("filePath", ""),
                    rec.get("durationSeconds", ""),
                    "",
                    "",
                ]
            )
    else:
        for rec in ordered_vp:
            wvp.append(
                [
                    rec.get("vendorName", ""),
                    rec.get("projectCode", ""),
                    rec.get("workflow", ""),
                    rec.get("locale", ""),
                    rec.get("ingestionBatch", ""),
                    rec.get("fileName", ""),
                    rec.get("filePath", ""),
                    rec.get("durationSeconds", ""),
                    rec.get("dedupeStatus", ""),
                    rec.get("dedupeReason", ""),
                ]
            )

    # --- QA_Findings ---
    wq = wb.create_sheet("QA_Findings")
    wq.append(["severity", "vendor", "fileName", "folderName", "locale", "issue", "detail"])
    for cell in wq[1]:
        cell.font = Font(bold=True)
    for f in findings:
        wq.append([f.severity, f.vendor, f.fileName, f.folderName, f.locale, f.issue, f.detail])

    # --- Source_Status ---
    ws2 = wb.create_sheet("Source_Status")
    ws2.append(["vendor", "status", "detail", "rows_loaded"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    for s in source_statuses:
        ws2.append([s.vendor, s.status, s.detail, s.rows_loaded])

    # --- Control_Masked (never secrets) ---
    wc = wb.create_sheet("Control_Masked")
    wc.append(
        [
            "Vendor Name",
            "Vendor Blob Container Name",
            "Blob SAS Key",
            "Expire date",
            "Comments",
            "POCs",
            "Link to Google Drive",
            "password",
        ]
    )
    for cell in wc[1]:
        cell.font = Font(bold=True)
    for v in vendors:
        wc.append(
            [
                v.name,
                v.container,
                v.sas_masked or "",
                v.expire,
                v.comments,
                v.pocs_masked or "",
                v.drive_link,
                "***SET***" if v.has_password else "",
            ]
        )

    # Autosize lightly
    for sheet in wb.worksheets:
        for col in range(1, min(sheet.max_column, 14) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run_pipeline(
    control_bytes: bytes,
    fallback_files: dict[str, bytes],
    log: Optional[LogFn] = None,
    oneforma_vendors: Optional[list[str]] = None,
    skip_oneforma: bool = False,
    existing_rate_names: Optional[set[str]] = None,
    existing_rate_composite: Optional[set[str]] = None,
) -> dict[str, Any]:
    vendors = parse_control_workbook(control_bytes, log=log)
    rows, statuses = load_vendor_sources(vendors, fallback_files, log=log)

    blob_rows: list[dict] = []
    if not skip_oneforma:
        # Prefer API vendor names from control containers / names
        api_vendors = []
        for v in vendors:
            api_vendors.append(vendor_api_name(v.name))
        # Dedupe; drop unknown empties
        api_vendors = sorted(set(x for x in api_vendors if x))
        # OneForma known set — filter to ones that exist when possible
        try:
            filters = oneforma_get("filters")
            known = set(filters.get("vendors") or [])
            filtered = [v for v in api_vendors if v in known]
            if filtered:
                api_vendors = filtered
            _log(log, f"OneForma vendors to query: {api_vendors}")
        except Exception as e:
            _log(log, f"filters endpoint failed ({e}); querying without filter list")

        if oneforma_vendors:
            api_vendors = oneforma_vendors
        blob_rows = fetch_oneforma_files(vendors=api_vendors or None, log=log)
    else:
        _log(log, "Skipping OneForma (skip_oneforma=1)")

    annotated, findings = run_qa(rows, blob_rows, statuses, log=log)

    # Drop internal-only keys from public merged_rows but keep enrichment fields
    public_rows: list[dict[str, Any]] = []
    for r in annotated:
        pr = {k: v for k, v in r.items()}
        public_rows.append(pr)

    vendor_prod = preview_vendor_prod(
        public_rows,
        existing_rate_names,
        existing_rate_composite,
    )
    _log(
        log,
        "Vendor Prod Rate Approval rows: "
        f"built={vendor_prod['rate_rows_built']} "
        f"unique={vendor_prod['unique_count']} "
        f"duplicate={vendor_prod['duplicate_count']} "
        f"incomplete={vendor_prod['incomplete_count']} "
        f"key={RATE_DEDUPE_KEY}",
    )
    xlsx = write_result_workbook(
        annotated,
        findings,
        statuses,
        vendors,
        vendor_prod_preview=vendor_prod,
    )

    return {
        "xlsx_bytes": xlsx,
        "merged_rows": public_rows,
        "vendor_prod": vendor_prod,
        "findings": [asdict(f) for f in findings],
        "source_statuses": [asdict(s) for s in statuses],
        "merged_count": len(annotated),
        "blob_count": len(blob_rows),
        "vendors": [
            {
                "name": v.name,
                "container": v.container,
                "sas": v.sas_masked,
                "pocs": v.pocs_masked,
                "link": v.drive_link,
                "has_password": v.has_password,
            }
            for v in vendors
        ],
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }




RATE_APPROVAL_SHEET = "Rate Approval Log"
DEFAULT_MASTER_PATH = ROOT / "masters" / "Vendor Prod Files.xlsx"
# File Name (col T) is the primary Rate Approval Log dedupe key.
# Composite A|B|F|T|U is also checked so a blank-T historical row cannot collide later.
RATE_DEDUPE_KEY = "fileName (column T); also composite vendor|project|locale|fileName|filePath"
VENDOR_PROD_FILL_COLUMNS = (
    "vendorName",
    "projectCode",
    "workflow",
    "locale",
    "ingestionBatch",
    "fileName",
    "filePath",
    "durationSeconds",
)
VENDOR_PROD_SHEET_HEADERS = [
    "Vendor Name",
    "Project Code",
    "Workflow",
    "Locale",
    "Ingestion Batch",
    "File Name",
    "File path",
    "Durations (seconds)",
    "dedupeStatus",
    "dedupeReason",
]


def _batch_to_excel_date(batch: Any):
    """Convert YYYY-MM-DD (or datetime) to datetime.date/datetime for Excel."""
    from datetime import date, datetime as dt

    if batch is None or batch == "":
        return None
    if isinstance(batch, (dt, date)):
        return batch
    s = str(batch).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        try:
            return dt.strptime(s, "%Y-%m-%d")
        except ValueError:
            return s
    if re.match(r"^\d{8}$", s):
        try:
            return dt.strptime(s, "%Y%m%d")
        except ValueError:
            return s
    return s


def to_rate_approval_rows(merged_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Map enriched merged rows → Rate Approval Log field dicts (filled cols only)."""
    out: list[dict[str, Any]] = []
    for r in merged_rows or []:
        vendor = str(r.get("vendorName") or r.get("_vendor") or "").strip()
        file_name = str(r.get("fileName") or "").strip()
        project = str(r.get("projectCode") or r.get("workflow") or "Maple").strip() or "Maple"
        workflow = str(r.get("workflow") or project or "Maple").strip() or "Maple"
        locale = str(r.get("locale") or "").strip()
        batch = r.get("ingestionBatch") or ""
        file_path = str(r.get("filePath") or "").strip()
        dur = r.get("durationSeconds")
        if dur == "":
            dur = None
        row = {
            "vendorName": vendor,
            "projectCode": project,
            "workflow": workflow,
            "locale": locale,
            "ingestionBatch": batch,
            "fileName": file_name,
            "filePath": file_path,
            "durationSeconds": dur,
        }
        out.append(row)
    return out


def rate_approval_keys(row: dict[str, Any]) -> tuple[str, str]:
    """Return (file_name_key, composite_key) for Rate Approval Log dedupe."""
    vendor = str(row.get("vendorName") or "").strip()
    project = str(row.get("projectCode") or "Maple").strip() or "Maple"
    locale = str(row.get("locale") or "").strip()
    file_name = str(row.get("fileName") or "").strip()
    file_path = str(row.get("filePath") or "").strip()
    name_key = file_name.lower()
    composite = f"{vendor.lower()}|{project.lower()}|{locale.lower()}|{name_key}|{file_path.lower()}"
    return name_key, composite


def extract_existing_rate_keys(ws) -> tuple[set[str], set[str]]:
    """Scan Rate Approval Log for existing File Name (T) and composite A|B|F|T|U keys."""
    existing_names: set[str] = set()
    existing_composite: set[str] = set()
    notes_start = _find_notes_start(ws)
    scan_limit = (notes_start - 1) if notes_start else ws.max_row
    for r in range(2, scan_limit + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        b = str(ws.cell(r, 2).value or "").strip()
        f = str(ws.cell(r, 6).value or "").strip()
        t = str(ws.cell(r, 20).value or "").strip()
        u = str(ws.cell(r, 21).value or "").strip()
        if t:
            existing_names.add(t.lower())
        if a or t:
            existing_composite.add(f"{a.lower()}|{b.lower()}|{f.lower()}|{t.lower()}|{u.lower()}")
    return existing_names, existing_composite


def extract_existing_rate_keys_from_bytes(data: bytes) -> tuple[set[str], set[str]]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    if RATE_APPROVAL_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet {RATE_APPROVAL_SHEET!r} missing from workbook")
    return extract_existing_rate_keys(wb[RATE_APPROVAL_SHEET])


def extract_existing_rate_keys_from_path(path: Path | str) -> tuple[set[str], set[str]]:
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    if RATE_APPROVAL_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet {RATE_APPROVAL_SHEET!r} missing from {path.name}")
    return extract_existing_rate_keys(wb[RATE_APPROVAL_SHEET])


def classify_rate_approval_rows(
    new_rows: list[dict[str, Any]],
    existing_names: Optional[set[str]] = None,
    existing_composite: Optional[set[str]] = None,
) -> dict[str, Any]:
    """Split rows into unique / duplicate / incomplete using the Rate Approval Log key.

    Mutates copies of the input sets so intra-batch File Name repeats count as duplicates.
    """
    names = set(existing_names or set())
    composites = set(existing_composite or set())
    unique_rows: list[dict[str, Any]] = []
    duplicate_rows: list[dict[str, Any]] = []
    incomplete_rows: list[dict[str, Any]] = []

    for src in new_rows or []:
        row = {k: src.get(k) for k in VENDOR_PROD_FILL_COLUMNS}
        vendor = str(row.get("vendorName") or "").strip()
        file_name = str(row.get("fileName") or "").strip()
        if not vendor or not file_name:
            tagged = {
                **row,
                "dedupeStatus": "incomplete",
                "dedupeReason": "missing Vendor Name and/or File Name",
            }
            incomplete_rows.append(tagged)
            continue
        name_key, composite = rate_approval_keys(row)
        if name_key in names:
            tagged = {
                **row,
                "dedupeStatus": "duplicate",
                "dedupeReason": f"File Name already on Rate Approval Log: {file_name}",
            }
            duplicate_rows.append(tagged)
            continue
        if composite in composites:
            tagged = {
                **row,
                "dedupeStatus": "duplicate",
                "dedupeReason": "composite vendor|project|locale|fileName|filePath already present",
            }
            duplicate_rows.append(tagged)
            continue
        names.add(name_key)
        composites.add(composite)
        tagged = {**row, "dedupeStatus": "unique", "dedupeReason": "not on Rate Approval Log"}
        unique_rows.append(tagged)

    return {
        "unique_rows": unique_rows,
        "duplicate_rows": duplicate_rows,
        "incomplete_rows": incomplete_rows,
        "unique_count": len(unique_rows),
        "duplicate_count": len(duplicate_rows),
        "incomplete_count": len(incomplete_rows),
        "rate_rows_built": len(new_rows or []),
        "dedupe_key": RATE_DEDUPE_KEY,
    }


def preview_vendor_prod(
    merged_rows: list[dict[str, Any]],
    existing_names: Optional[set[str]] = None,
    existing_composite: Optional[set[str]] = None,
) -> dict[str, Any]:
    """Map merged/enriched rows → Rate Approval Log format and classify vs destination keys."""
    rate_rows = to_rate_approval_rows(merged_rows)
    classified = classify_rate_approval_rows(rate_rows, existing_names, existing_composite)
    classified["rate_rows"] = rate_rows
    classified["unique_file_names"] = [
        str(r.get("fileName") or "") for r in classified["unique_rows"]
    ]
    classified["duplicate_file_names"] = [
        str(r.get("fileName") or "") for r in classified["duplicate_rows"]
    ]
    return classified


def write_vendor_prod_preview_xlsx(preview: dict[str, Any]) -> bytes:
    """Small workbook of Rate Approval Log fill columns + dedupe status."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendor_Prod_Rows"
    ws.append(VENDOR_PROD_SHEET_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ordered = (
        list(preview.get("unique_rows") or [])
        + list(preview.get("duplicate_rows") or [])
        + list(preview.get("incomplete_rows") or [])
    )
    for rec in ordered:
        ws.append(
            [
                rec.get("vendorName", ""),
                rec.get("projectCode", ""),
                rec.get("workflow", ""),
                rec.get("locale", ""),
                rec.get("ingestionBatch", ""),
                rec.get("fileName", ""),
                rec.get("filePath", ""),
                rec.get("durationSeconds", ""),
                rec.get("dedupeStatus", ""),
                rec.get("dedupeReason", ""),
            ]
        )
    for col in range(1, len(VENDOR_PROD_SHEET_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _find_last_data_row(ws) -> int:
    """Last row with Vendor Name (A) or File Name (T) filled in the data area (before notes)."""
    notes_start = None
    for r in range(1, ws.max_row + 1):
        o_val = ws.cell(r, 15).value  # col O
        if o_val and "notes" in str(o_val).lower():
            notes_start = r
            break
    limit = (notes_start - 1) if notes_start else ws.max_row
    last = 1
    for r in range(2, limit + 1):
        a = ws.cell(r, 1).value
        t = ws.cell(r, 20).value
        if (a is not None and str(a).strip()) or (t is not None and str(t).strip()):
            last = r
    return last


def _find_notes_start(ws) -> Optional[int]:
    for r in range(1, ws.max_row + 1):
        o_val = ws.cell(r, 15).value
        if o_val and "notes" in str(o_val).lower():
            return r
    return None


def append_rate_approval_log(
    master_path: Path | str,
    new_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """Append unique Rate Approval Log rows; atomic write via temp file.

    Returns {appended, skipped_duplicates, skipped_incomplete, next_row, master_path}.
    """
    master_path = Path(master_path)
    if not master_path.exists():
        raise FileNotFoundError(f"Master workbook not found: {master_path}")

    wb = openpyxl.load_workbook(master_path)
    if RATE_APPROVAL_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet {RATE_APPROVAL_SHEET!r} missing from {master_path.name}")
    classified, last_data = _append_rate_rows_to_workbook(wb, new_rows)
    _atomic_save_workbook(wb, master_path)
    return _append_summary(classified, last_data, str(master_path))


def append_rate_approval_log_bytes(
    data: bytes,
    new_rows: list[dict[str, Any]],
) -> tuple[bytes, dict[str, Any]]:
    """Append unique Rate Approval Log rows in memory; return (xlsx_bytes, summary)."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    if RATE_APPROVAL_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet {RATE_APPROVAL_SHEET!r} missing from workbook")
    classified, last_data = _append_rate_rows_to_workbook(wb, new_rows)
    buf = io.BytesIO()
    wb.save(buf)
    summary = _append_summary(classified, last_data, "in-memory")
    return buf.getvalue(), summary


def _append_summary(classified: dict[str, Any], last_data: int, master_path: str) -> dict[str, Any]:
    appended = classified["unique_count"]
    return {
        "appended": appended,
        "skipped_duplicates": classified["duplicate_count"],
        "skipped_incomplete": classified["incomplete_count"],
        "next_row": last_data + 1 + appended,
        "master_path": master_path,
        "last_data_row_before": last_data,
        "dedupe_key": RATE_DEDUPE_KEY,
        "unique_file_names": [str(r.get("fileName") or "") for r in classified["unique_rows"]],
        "duplicate_file_names": [str(r.get("fileName") or "") for r in classified["duplicate_rows"]],
    }


def _atomic_save_workbook(wb, master_path: Path) -> None:
    import os
    import tempfile

    master_path = Path(master_path)
    master_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(master_path.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(tmp_path)
        tmp_path.replace(master_path)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise


def _append_rate_rows_to_workbook(
    wb,
    new_rows: list[dict[str, Any]],
) -> tuple[dict[str, Any], int]:
    """Mutate workbook: append only unique Rate Approval Log rows. Returns (classified, last_data)."""
    if RATE_APPROVAL_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet {RATE_APPROVAL_SHEET!r} missing from workbook")
    ws = wb[RATE_APPROVAL_SHEET]
    existing_names, existing_composite = extract_existing_rate_keys(ws)
    classified = classify_rate_approval_rows(new_rows, existing_names, existing_composite)
    to_write = classified["unique_rows"]
    last_data = _find_last_data_row(ws)

    # Make room before notes if needed. openpyxl insert_rows does not reliably
    # shift merged ranges, and note merges (P:AP) cover File Name (T) — unmerge first.
    need = len(to_write)
    if need:
        notes_start = _find_notes_start(ws)
        first_write = last_data + 1

        def _unmerge_from(row_start: int) -> list:
            doomed = [
                str(mr)
                for mr in list(ws.merged_cells.ranges)
                if mr.max_row >= row_start
            ]
            for ref in doomed:
                try:
                    ws.unmerge_cells(ref)
                except ValueError:
                    pass
            return doomed

        if notes_start is not None:
            _unmerge_from(notes_start)
            free = notes_start - first_write
            if free < need:
                insert_at = notes_start
                amount = need - free
                ws.insert_rows(insert_at, amount=amount)
                # Clear stale merges that openpyxl may leave behind pointing at data rows
                _unmerge_from(first_write)
        else:
            _unmerge_from(first_write)

        for i, row in enumerate(to_write):
            r = first_write + i
            ws.cell(r, 1).value = str(row.get("vendorName") or "").strip()
            ws.cell(r, 2).value = str(row.get("projectCode") or "Maple").strip() or "Maple"
            # C, D left empty (do not overwrite / no formulas on new rows)
            ws.cell(r, 5).value = (
                str(row.get("workflow") or row.get("projectCode") or "Maple").strip() or "Maple"
            )
            ws.cell(r, 6).value = str(row.get("locale") or "").strip()
            batch_val = _batch_to_excel_date(row.get("ingestionBatch"))
            if batch_val is not None:
                ws.cell(r, 7).value = batch_val
            ws.cell(r, 20).value = str(row.get("fileName") or "").strip()
            ws.cell(r, 21).value = str(row.get("filePath") or "").strip()
            dur = row.get("durationSeconds")
            if dur is not None and dur != "":
                try:
                    ws.cell(r, 22).value = float(dur)
                except (TypeError, ValueError):
                    ws.cell(r, 22).value = dur

    return classified, last_data


def build_demo_control_bytes() -> bytes:
    path = SAMPLES / "demo_control.xlsx"
    return path.read_bytes()


def demo_fallback_files() -> dict[str, bytes]:
    """Demo local samples. Alchemy/arcca use gdrive_sync via Column G when present."""
    out: dict[str, bytes] = {}
    for name, fname in [
        ("alchemy", "demo_vendor_alchemy.xlsx"),
        ("arcca", "demo_vendor_arcca.xlsx"),
        ("pangeanic", "pangeanic_maple_delivery_tracking.xlsx"),
    ]:
        if name != "pangeanic" and resolve_gdrive_sync(name) is not None:
            continue
        p = SAMPLES / fname
        if p.exists():
            out[name] = p.read_bytes()
    return out
