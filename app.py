"""Verdi Merge + Auto QC — Centific (OTS Agency Intake + OneForma QA)."""
from __future__ import annotations

import json
import re
import traceback
import uuid
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import ots_pipeline as pipe
import sharepoint_vendor_prod as sp

ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
STATIC = ROOT / "static"
STATIC.mkdir(exist_ok=True)
MASTERS = ROOT / "masters"
DEFAULT_LOCAL_MASTER = MASTERS / "Vendor Prod Files.xlsx"

app = FastAPI(title="Verdi Merge + Auto QC", version="0.1.0")
app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")

# In-memory last run artifacts (demo scale)
_LAST: dict[str, Any] = {}
_LOGS: list[str] = []


def _append_log(msg: str) -> None:
    _LOGS.append(msg)
    if len(_LOGS) > 500:
        del _LOGS[: len(_LOGS) - 500]


@app.get("/health")
def health():
    return {"ok": True, "service": "ots-agency-qa"}


@app.get("/")
def index():
    html_path = STATIC / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/api/logs")
def get_logs():
    return {"logs": list(_LOGS)}


@app.post("/api/clear-logs")
def clear_logs():
    _LOGS.clear()
    return {"ok": True}


@app.get("/api/oneforma/filters")
def proxy_filters():
    try:
        data = pipe.oneforma_get("filters")
        return JSONResponse(data)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=502)


@app.get("/api/oneforma/files")
def proxy_files(
    vendor: Optional[str] = None,
    direction: Optional[str] = None,
    workflow: Optional[str] = None,
    locale: Optional[str] = None,
    qcStatus: Optional[str] = None,
    dateFrom: Optional[str] = None,
    dateTo: Optional[str] = None,
    search: Optional[str] = None,
    includeNonAudio: Optional[str] = "true",
    failedProbeOnly: Optional[str] = None,
    page: int = 1,
    pageSize: int = 50,
    sortBy: Optional[str] = None,
    sortOrder: Optional[str] = None,
):
    params = {
        k: v
        for k, v in {
            "vendor": vendor,
            "direction": direction,
            "workflow": workflow,
            "locale": locale,
            "qcStatus": qcStatus,
            "dateFrom": dateFrom,
            "dateTo": dateTo,
            "search": search,
            "includeNonAudio": includeNonAudio,
            "failedProbeOnly": failedProbeOnly,
            "page": page,
            "pageSize": pageSize,
            "sortBy": sortBy,
            "sortOrder": sortOrder,
        }.items()
        if v is not None and v != ""
    }
    try:
        data = pipe.oneforma_get("files", params)
        return JSONResponse(data)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=502)


def _safe_vendor_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (name or "").strip().lower())


def _parse_vendor_tag(filename: str, form_vendor: Optional[str]) -> str:
    if form_vendor and form_vendor.strip():
        return form_vendor.strip().lower()
    # filename patterns: alchemy.xlsx, demo_vendor_alchemy.xlsx, vendor_alchemy_source.xlsx
    stem = Path(filename or "").stem.lower()
    m = re.search(r"(?:vendor[_\-]?)?([a-z0-9]+)(?:[_\-]?(?:source|tracker|demo))?$", stem)
    if "alchemy" in stem:
        return "alchemy"
    if "arcca" in stem:
        return "arcca"
    if "pangeanic" in stem:
        return "pangeanic"
    if "senserv" in stem:
        return "senserv"
    if "avyaan" in stem or "avyaa" in stem:
        return "avyaan"
    if "language" in stem:
        return "languagesolutionasia"
    if "consult" in stem:
        return "consultbae"
    if m:
        return m.group(1)
    return stem


@app.post("/api/run")
async def run_merge_qa(
    control: Optional[UploadFile] = File(None),
    source_sheet: Optional[UploadFile] = File(None),
    source_vendor: Optional[str] = Form(None),
    vendor_files: list[UploadFile] = File(default=[]),
    demo_mode: Optional[str] = Form(None),
    skip_oneforma: Optional[str] = Form(None),
):
    _LOGS.clear()
    run_id = uuid.uuid4().hex[:10]
    _append_log(f"=== Run {run_id} starting ===")

    try:
        is_demo = (demo_mode or "").lower() in ("1", "true", "yes", "on")
        skip_of = (skip_oneforma or "").lower() in ("1", "true", "yes", "on")

        if is_demo or control is None or not control.filename:
            _append_log("Using demo control workbook (no real SAS/passwords)")
            control_bytes = pipe.build_demo_control_bytes()
        else:
            control_bytes = await control.read()
            # Persist upload without secrets logging
            dest = UPLOADS / f"{run_id}_control.xlsx"
            dest.write_bytes(control_bytes)
            _append_log(f"Loaded control upload ({len(control_bytes)} bytes)")

        fallback: dict[str, bytes] = {}
        if is_demo:
            fallback.update(pipe.demo_fallback_files())
            _append_log(f"Demo fallback vendors: {list(fallback.keys())}")

        # Optional single source sheet
        if source_sheet and source_sheet.filename:
            data = await source_sheet.read()
            vkey = _parse_vendor_tag(source_sheet.filename, source_vendor)
            fallback[vkey] = data
            (UPLOADS / f"{run_id}_vendor_{vkey}.xlsx").write_bytes(data)
            _append_log(f"Fallback source for vendor={vkey} ({len(data)} bytes)")

        # Multiple vendor files
        for vf in vendor_files or []:
            if not vf.filename:
                continue
            data = await vf.read()
            vkey = _parse_vendor_tag(vf.filename, None)
            fallback[vkey] = data
            (UPLOADS / f"{run_id}_vendor_{vkey}.xlsx").write_bytes(data)
            _append_log(f"Fallback source for vendor={vkey} ({len(data)} bytes)")

        dest_preview = _load_destination_keys_for_preview()
        result = pipe.run_pipeline(
            control_bytes,
            fallback,
            log=_append_log,
            skip_oneforma=skip_of,
            existing_rate_names=dest_preview.get("existing_names"),
            existing_rate_composite=dest_preview.get("existing_composite"),
        )

        vendor_prod = result.get("vendor_prod") or {}
        vendor_prod = {
            **vendor_prod,
            "destination": dest_preview.get("destination") or {},
        }

        out_path = UPLOADS / f"{run_id}_Result_Merged.xlsx"
        out_path.write_bytes(result["xlsx_bytes"])
        # Also stable name for download convenience
        stable = UPLOADS / "Result_Merged.xlsx"
        stable.write_bytes(result["xlsx_bytes"])
        vp_preview_path = UPLOADS / "Vendor_Prod_Rows.xlsx"
        vp_preview_path.write_bytes(pipe.write_vendor_prod_preview_xlsx(vendor_prod))

        merged_rows = result.get("merged_rows") or []
        # Persist last run (incl. enriched + Vendor Prod rows) for Send
        (UPLOADS / "last_run.json").write_text(
            json.dumps(
                {
                    "run_id": run_id,
                    "merged_count": result["merged_count"],
                    "blob_count": result["blob_count"],
                    "generated_at": result["generated_at"],
                    "merged_rows": merged_rows,
                    "vendor_prod": vendor_prod,
                    "findings": result["findings"],
                    "source_statuses": result["source_statuses"],
                    "vendors": result["vendors"],
                },
                default=str,
            ),
            encoding="utf-8",
        )

        _LAST.clear()
        _LAST.update(
            {
                "run_id": run_id,
                "path": str(stable),
                "merged_rows": merged_rows,
                "vendor_prod": vendor_prod,
                "findings": result["findings"],
                "source_statuses": result["source_statuses"],
                "merged_count": result["merged_count"],
                "blob_count": result["blob_count"],
                "vendors": result["vendors"],
                "generated_at": result["generated_at"],
            }
        )
        dest_kind = (dest_preview.get("destination") or {}).get("kind") or "none"
        _append_log(
            f"Done. merged={result['merged_count']} blobs={result['blob_count']} "
            f"findings={len(result['findings'])} "
            f"vendor_prod_unique={vendor_prod.get('unique_count')} "
            f"vendor_prod_duplicate={vendor_prod.get('duplicate_count')} "
            f"destination={dest_kind}"
        )
        _append_log(f"=== Run {run_id} complete ===")

        return {
            "ok": True,
            "run_id": run_id,
            "download_url": "/api/download/result",
            "vendor_prod_download_url": "/api/download/vendor-prod-rows",
            "merged_count": result["merged_count"],
            "blob_count": result["blob_count"],
            "unique_count": vendor_prod.get("unique_count"),
            "duplicate_count": vendor_prod.get("duplicate_count"),
            "incomplete_count": vendor_prod.get("incomplete_count"),
            "rate_rows_built": vendor_prod.get("rate_rows_built"),
            "dedupe_key": vendor_prod.get("dedupe_key"),
            "vendor_prod": _public_vendor_prod(vendor_prod),
            "findings": result["findings"],
            "source_statuses": result["source_statuses"],
            "vendors": result["vendors"],
            "logs": list(_LOGS),
        }
    except Exception as e:
        _append_log(f"ERROR: {e}")
        _append_log(traceback.format_exc())
        return JSONResponse(
            {"ok": False, "error": str(e), "logs": list(_LOGS)},
            status_code=500,
        )


@app.get("/api/download/result")
def download_result():
    path = UPLOADS / "Result_Merged.xlsx"
    if not path.exists():
        return JSONResponse({"error": "No result yet — run merge + QA first"}, status_code=404)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Result_Merged.xlsx",
    )


@app.get("/api/last")
def last_result():
    if not _LAST:
        return {"ok": False, "message": "No run yet"}
    return {"ok": True, **{k: v for k, v in _LAST.items() if k != "path"}}




def _load_last_merged_rows() -> list:
    if _LAST.get("merged_rows"):
        return list(_LAST["merged_rows"])
    last_path = UPLOADS / "last_run.json"
    if last_path.exists():
        try:
            data = json.loads(last_path.read_text(encoding="utf-8"))
            return list(data.get("merged_rows") or [])
        except Exception:
            return []
    return []


def _public_vendor_prod(vendor_prod: dict[str, Any]) -> dict[str, Any]:
    """JSON-safe Vendor Prod preview: counts + short row lists (not a full workbook)."""
    unique_rows = vendor_prod.get("unique_rows") or []
    duplicate_rows = vendor_prod.get("duplicate_rows") or []
    incomplete_rows = vendor_prod.get("incomplete_rows") or []
    return {
        "rate_rows_built": vendor_prod.get("rate_rows_built"),
        "unique_count": vendor_prod.get("unique_count"),
        "duplicate_count": vendor_prod.get("duplicate_count"),
        "incomplete_count": vendor_prod.get("incomplete_count"),
        "dedupe_key": vendor_prod.get("dedupe_key") or pipe.RATE_DEDUPE_KEY,
        "destination": vendor_prod.get("destination") or {},
        "unique_file_names": vendor_prod.get("unique_file_names") or [
            r.get("fileName") for r in unique_rows
        ],
        "duplicate_file_names": vendor_prod.get("duplicate_file_names") or [
            r.get("fileName") for r in duplicate_rows
        ],
        "unique_rows": unique_rows[:50],
        "duplicate_rows": duplicate_rows[:50],
        "incomplete_rows": incomplete_rows[:20],
        "rate_rows": (vendor_prod.get("rate_rows") or [])[:50],
    }


def _local_master_path(master_path: Optional[str] = None) -> Path:
    path = Path(master_path) if master_path else DEFAULT_LOCAL_MASTER
    if not path.is_absolute():
        path = (ROOT / path).resolve()
    return path


def vendor_prod_status_payload() -> dict[str, Any]:
    cfg = sp.SharePointConfig.from_env()
    local = DEFAULT_LOCAL_MASTER
    local_exists = local.exists()
    intended = sp.sharepoint_write_intended(cfg)
    if cfg.auth_configured():
        kind = "sharepoint"
        message = (
            "Send will download the live SharePoint Rate Approval Log, append unique rows, "
            "and upload the workbook back."
        )
    elif cfg.url_from_env or cfg.require:
        kind = "sharepoint_unauthenticated"
        message = (
            "SharePoint is the intended destination, but Graph auth is not configured. "
            "Send will fail instead of writing only a local file."
        )
    elif local_exists:
        kind = "local"
        message = (
            "SharePoint auth is not set. Send will use the local demo master under masters/."
        )
    else:
        kind = "unconfigured"
        message = (
            "No SharePoint Graph auth and no local masters/Vendor Prod Files.xlsx. "
            "Demo Send needs a local master; production Send needs Azure/Graph env vars."
        )
    return {
        "ok": True,
        "destination_kind": kind,
        "sharepoint_intended": intended,
        "local_master_exists": local_exists,
        "local_master_path": str(local) if local_exists else None,
        "sharepoint": cfg.public_status(),
        "dedupe_key": pipe.RATE_DEDUPE_KEY,
        "sheet": pipe.RATE_APPROVAL_SHEET,
        "fill_columns": "A Vendor Name, B Project Code, E Workflow, F Locale, G Ingestion Batch, T File Name, U File path, V Durations (seconds)",
        "message": message,
    }


def _load_destination_keys_for_preview() -> dict[str, Any]:
    """Best-effort Rate Approval Log keys for after-run unique/duplicate KPIs.

    Never fails the merge run. SharePoint errors are recorded on destination.
    """
    cfg = sp.SharePointConfig.from_env()
    if cfg.auth_configured():
        try:
            data, meta = sp.download_vendor_prod_workbook(cfg)
            names, composites = pipe.extract_existing_rate_keys_from_bytes(data)
            _append_log(
                f"Vendor Prod KPI: loaded {len(names)} File Name key(s) from SharePoint "
                f"({meta.get('name') or 'Vendor Prod Files.xlsx'})"
            )
            return {
                "existing_names": names,
                "existing_composite": composites,
                "destination": {
                    "kind": "sharepoint",
                    "label": "SharePoint Rate Approval Log",
                    "existing_file_names": len(names),
                    **{k: meta.get(k) for k in ("web_url", "name", "file_path", "site_path")},
                },
            }
        except Exception as e:
            _append_log(f"Vendor Prod KPI: SharePoint download failed ({e}); unique/dup are in-batch only")
            return {
                "existing_names": set(),
                "existing_composite": set(),
                "destination": {
                    "kind": "sharepoint_error",
                    "label": "SharePoint (download failed — in-batch File Name dups only)",
                    "error": str(e),
                    "hint": getattr(e, "hint", ""),
                    "existing_file_names": 0,
                },
            }

    local = DEFAULT_LOCAL_MASTER
    if local.exists():
        try:
            names, composites = pipe.extract_existing_rate_keys_from_path(local)
            _append_log(
                f"Vendor Prod KPI: loaded {len(names)} File Name key(s) from local master"
            )
            return {
                "existing_names": names,
                "existing_composite": composites,
                "destination": {
                    "kind": "local",
                    "label": "Local masters/Vendor Prod Files.xlsx",
                    "path": str(local),
                    "existing_file_names": len(names),
                },
            }
        except Exception as e:
            _append_log(f"Vendor Prod KPI: local master unreadable ({e})")
            return {
                "existing_names": set(),
                "existing_composite": set(),
                "destination": {
                    "kind": "local_error",
                    "label": "Local master unreadable — in-batch File Name dups only",
                    "error": str(e),
                    "existing_file_names": 0,
                },
            }

    _append_log(
        "Vendor Prod KPI: no destination loaded (no Graph auth, no local master) — "
        "unique/duplicate counts are in-batch File Name only"
    )
    return {
        "existing_names": set(),
        "existing_composite": set(),
        "destination": {
            "kind": "none",
            "label": "No Rate Approval Log loaded — in-batch File Name dups only",
            "existing_file_names": 0,
        },
    }


@app.get("/api/vendor-prod/status")
def vendor_prod_status():
    return vendor_prod_status_payload()


@app.post("/api/send-vendor-prod")
async def send_vendor_prod(
    master_path: Optional[str] = Form(None),
    force_local: Optional[str] = Form(None),
):
    """Dedupe against the destination Rate Approval Log, then append only unique rows.

    Default: live SharePoint workbook when Graph auth (or an explicit SharePoint URL) is set.
    Demo fallback: local masters/Vendor Prod Files.xlsx when SharePoint is not configured.
    If SharePoint is intended but auth/upload fails, this errors out — it does not silently
    write only a local file.
    """
    rows = _load_last_merged_rows()
    if not rows:
        return JSONResponse(
            {"ok": False, "error": "No merge result yet — run merge + QA first"},
            status_code=400,
        )

    use_local = (force_local or "").lower() in ("1", "true", "yes", "on")
    cfg = sp.SharePointConfig.from_env()
    rate_rows = pipe.to_rate_approval_rows(rows)

    try:
        if use_local:
            summary = _send_local_master(rate_rows, master_path)
        elif sp.sharepoint_write_intended(cfg):
            if not cfg.auth_configured():
                raise sp.SharePointError(
                    "SharePoint is the intended Vendor Prod destination, but Graph auth is missing.",
                    hint=sp.AUTH_HINT,
                    status_code=503,
                )
            summary = _send_sharepoint(cfg, rate_rows)
        else:
            summary = _send_local_master(rate_rows, master_path)

        (UPLOADS / "last_rate_append.json").write_text(
            json.dumps({**summary, "rate_rows_built": len(rate_rows)}, default=str, indent=2),
            encoding="utf-8",
        )
        _append_log(
            f"Send Vendor Prod: destination={summary.get('destination')} "
            f"appended={summary['appended']} "
            f"skipped_duplicates={summary['skipped_duplicates']} "
            f"skipped_incomplete={summary['skipped_incomplete']} "
            f"next_row={summary.get('next_row')}"
        )
        return {
            "ok": True,
            "updated_copy": "/api/download/vendor-prod-updated",
            "rate_rows_built": len(rate_rows),
            "unique_count": summary.get("appended"),
            "duplicate_count": summary.get("skipped_duplicates"),
            "incomplete_count": summary.get("skipped_incomplete"),
            **summary,
            "logs": list(_LOGS),
        }
    except sp.SharePointError as e:
        _append_log(f"Send Vendor Prod ERROR: {e}")
        if e.hint:
            _append_log(f"HINT: {e.hint}")
        payload = {
            "ok": False,
            **e.to_dict(),
            "logs": list(_LOGS),
        }
        return JSONResponse(payload, status_code=e.status_code or 502)
    except FileNotFoundError as e:
        _append_log(f"Send Vendor Prod ERROR: {e}")
        return JSONResponse(
            {
                "ok": False,
                "error": str(e),
                "hint": (
                    "For a local demo, place a copy of Vendor Prod Files.xlsx under masters/ "
                    "(gitignored). For production, set Azure/Graph env vars so Send can "
                    "download and update the SharePoint workbook."
                ),
                "destination": "local",
                "logs": list(_LOGS),
            },
            status_code=404,
        )
    except Exception as e:
        _append_log(f"Send Vendor Prod ERROR: {e}")
        _append_log(traceback.format_exc())
        return JSONResponse(
            {"ok": False, "error": str(e), "logs": list(_LOGS)},
            status_code=500,
        )


def _send_local_master(rate_rows: list, master_path: Optional[str]) -> dict[str, Any]:
    path = _local_master_path(master_path)
    if not path.exists():
        raise FileNotFoundError(f"Local master not found: {path}")
    _append_log(f"Send Vendor Prod: local fallback {path}")
    summary = pipe.append_rate_approval_log(path, rate_rows)
    updated_copy = UPLOADS / "Vendor_Prod_Files_updated.xlsx"
    updated_copy.write_bytes(path.read_bytes())
    summary["destination"] = "local"
    summary["destination_label"] = f"Local master {path.name}"
    summary["master_path"] = str(path)
    summary["sharepoint_uploaded"] = False
    return summary


def _send_sharepoint(cfg: sp.SharePointConfig, rate_rows: list) -> dict[str, Any]:
    _append_log("Send Vendor Prod: downloading live SharePoint Rate Approval Log…")
    data, meta = sp.download_vendor_prod_workbook(cfg)
    _append_log(
        f"Send Vendor Prod: downloaded {meta.get('name') or 'workbook'} "
        f"({meta.get('size') or len(data)} bytes) — checking duplicates on Rate Approval Log"
    )
    item = sp.DriveItemRef(
        drive_id=str(meta.get("drive_id") or ""),
        item_id=str(meta.get("item_id") or ""),
        name=str(meta.get("name") or ""),
        etag=str(meta.get("etag") or ""),
        web_url=str(meta.get("web_url") or ""),
        size=int(meta.get("size") or 0),
    )
    new_bytes, summary = pipe.append_rate_approval_log_bytes(data, rate_rows)
    _append_log(
        f"Send Vendor Prod: after dedupe unique={summary['appended']} "
        f"duplicate={summary['skipped_duplicates']} incomplete={summary['skipped_incomplete']}"
    )
    if summary["appended"]:
        _append_log("Send Vendor Prod: uploading updated workbook to SharePoint…")
        uploaded = sp.upload_vendor_prod_workbook(new_bytes, item, cfg)
        summary["sharepoint_uploaded"] = True
        summary["web_url"] = uploaded.get("web_url") or meta.get("web_url")
        _append_log("Send Vendor Prod: SharePoint upload complete")
    else:
        summary["sharepoint_uploaded"] = False
        summary["web_url"] = meta.get("web_url")
        _append_log("Send Vendor Prod: nothing unique to append — SharePoint file left unchanged")

    updated_copy = UPLOADS / "Vendor_Prod_Files_updated.xlsx"
    updated_copy.write_bytes(new_bytes)
    # Local cache of the post-send workbook (ephemeral on Render; not the live destination)
    try:
        MASTERS.mkdir(parents=True, exist_ok=True)
        (MASTERS / "Vendor Prod Files.xlsx").write_bytes(new_bytes)
    except OSError:
        pass
    summary["destination"] = "sharepoint"
    summary["destination_label"] = "SharePoint Rate Approval Log"
    summary["sharepoint"] = {
        "web_url": summary.get("web_url"),
        "name": meta.get("name"),
        "site_path": meta.get("site_path"),
        "file_path": meta.get("file_path"),
        "uploaded": summary["sharepoint_uploaded"],
    }
    return summary


@app.get("/api/download/vendor-prod-updated")
def download_vendor_prod_updated():
    path = UPLOADS / "Vendor_Prod_Files_updated.xlsx"
    if not path.exists():
        return JSONResponse({"error": "No updated master yet — use Send first"}, status_code=404)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Vendor_Prod_Files_updated.xlsx",
    )


@app.get("/api/download/vendor-prod-rows")
def download_vendor_prod_rows():
    path = UPLOADS / "Vendor_Prod_Rows.xlsx"
    if not path.exists():
        return JSONResponse(
            {"error": "No Vendor Prod rows yet — run merge + QA first"},
            status_code=404,
        )
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Vendor_Prod_Rows.xlsx",
    )


@app.get("/api/samples/{name}")
def get_sample(name: str):
    allowed = {
        "demo_control.xlsx",
        "demo_vendor_alchemy.xlsx",
        "demo_vendor_arcca.xlsx",
        "pangeanic_maple_delivery_tracking.xlsx",
        "template.xlsx",
    }
    if name == "template.xlsx":
        path = ROOT / "template.xlsx"
    else:
        if name not in allowed:
            return JSONResponse({"error": "unknown sample"}, status_code=404)
        path = ROOT / "samples" / name
    if not path.exists():
        return JSONResponse({"error": "missing"}, status_code=404)
    return FileResponse(path, filename=name)


if __name__ == "__main__":
    import os
    import uvicorn

    port = int(os.environ.get("PORT", "8876"))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
