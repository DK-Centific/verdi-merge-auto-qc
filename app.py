"""Verdi Merge + Auto QC — Centific (OTS Agency Intake + OneForma QA)."""
from __future__ import annotations

import json
import re
import traceback
import uuid
from pathlib import Path
from typing import Any, Optional

import requests
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

import ots_pipeline as pipe

ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
STATIC = ROOT / "static"
STATIC.mkdir(exist_ok=True)

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

        result = pipe.run_pipeline(
            control_bytes,
            fallback,
            log=_append_log,
            skip_oneforma=skip_of,
        )

        out_path = UPLOADS / f"{run_id}_Result_Merged.xlsx"
        out_path.write_bytes(result["xlsx_bytes"])
        # Also stable name for download convenience
        stable = UPLOADS / "Result_Merged.xlsx"
        stable.write_bytes(result["xlsx_bytes"])

        merged_rows = result.get("merged_rows") or []
        # Persist last run (incl. enriched rows) for Send → Vendor Prod
        (UPLOADS / "last_run.json").write_text(
            json.dumps(
                {
                    "run_id": run_id,
                    "merged_count": result["merged_count"],
                    "blob_count": result["blob_count"],
                    "generated_at": result["generated_at"],
                    "merged_rows": merged_rows,
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
                "findings": result["findings"],
                "source_statuses": result["source_statuses"],
                "merged_count": result["merged_count"],
                "blob_count": result["blob_count"],
                "vendors": result["vendors"],
                "generated_at": result["generated_at"],
            }
        )
        _append_log(
            f"Done. merged={result['merged_count']} blobs={result['blob_count']} "
            f"findings={len(result['findings'])}"
        )
        _append_log(f"=== Run {run_id} complete ===")

        return {
            "ok": True,
            "run_id": run_id,
            "download_url": "/api/download/result",
            "merged_count": result["merged_count"],
            "blob_count": result["blob_count"],
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


@app.post("/api/send-vendor-prod")
async def send_vendor_prod(master_path: Optional[str] = Form(None)):
    """Append last merge rows into Rate Approval Log (local master copy) with dedupe."""
    rows = _load_last_merged_rows()
    if not rows:
        return JSONResponse(
            {"ok": False, "error": "No merge result yet — run merge + QA first"},
            status_code=400,
        )

    path = Path(master_path) if master_path else (ROOT / "masters" / "Vendor Prod Files.xlsx")
    if not path.is_absolute():
        path = (ROOT / path).resolve()
    if not path.exists():
        return JSONResponse({"ok": False, "error": f"Master not found: {path}"}, status_code=404)

    try:
        rate_rows = pipe.to_rate_approval_rows(rows)
        summary = pipe.append_rate_approval_log(path, rate_rows)
        # Write append summary + copy updated master for download convenience
        (UPLOADS / "last_rate_append.json").write_text(
            json.dumps({**summary, "rate_rows_built": len(rate_rows)}, default=str, indent=2),
            encoding="utf-8",
        )
        updated_copy = UPLOADS / "Vendor_Prod_Files_updated.xlsx"
        updated_copy.write_bytes(path.read_bytes())
        _append_log(
            f"Send Vendor Prod: appended={summary['appended']} "
            f"skipped_duplicates={summary['skipped_duplicates']} "
            f"skipped_incomplete={summary['skipped_incomplete']} "
            f"next_row={summary['next_row']}"
        )
        return {
            "ok": True,
            "master_path": str(path),
            "updated_copy": "/api/download/vendor-prod-updated",
            "rate_rows_built": len(rate_rows),
            **summary,
            "logs": list(_LOGS),
        }
    except Exception as e:
        _append_log(f"Send Vendor Prod ERROR: {e}")
        _append_log(traceback.format_exc())
        return JSONResponse(
            {"ok": False, "error": str(e), "logs": list(_LOGS)},
            status_code=500,
        )


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
