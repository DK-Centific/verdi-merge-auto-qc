# Verdi Merge + Auto QC — Centific


> **Not** the Agency Excel Consolidator (`excel-merger` / Centaurus). This is the Verdi OTS Agency Intake → Merge + OneForma QA tool.
Small FastAPI app that:

1. Loads a SharePoint-style **CONTROL** workbook (upload or demo sample).
2. For each vendor, follows Column G (“Link to Google Drive”) **or** uses an uploaded vendor source `.xlsx` fallback.
3. Extracts / normalizes rows into the **OTS template** columns (plus `ingestionBatch` + `durationSeconds`) and consolidates into one workbook.
4. QAs against the live **OneForma Vendor Blob File List** APIs (server-side proxy, no browser CORS).
5. Optionally **Send data to Vendor Prod Files** — appends unique rows into the **Rate Approval Log** sheet of the local master workbook (`masters/Vendor Prod Files.xlsx`, synced from SharePoint). Dedupe prefers File Name (col T); also skips exact A+B+F+T+U matches.

This is **not** the Centaurus 30-col excel-merger.

## Quick start (Mac / local)

```bash
cd /Users/davidk/Desktop/ots-agency-qa   # or /workspace/ots-agency-qa on the box
.venv/bin/uvicorn app:app --host 127.0.0.1 --port 8876
```

Open **http://127.0.0.1:8876**.

Health check: `GET /health` → `{"ok": true}`.

If the venv is missing:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## UI buttons

| Control | What it does |
|--------|----------------|
| **Control workbook** | Upload Verdi/Juniper-style control sheet. Column H passwords are read in-memory only and never displayed. |
| **Optional single source sheet** | Fallback vendor tracker xlsx when Google links are dead. Optionally tag the vendor name. |
| **Extra vendor source files** | Multiple fallbacks; filename should include the vendor (e.g. `alchemy.xlsx`). |
| **Demo mode** | Uses `samples/demo_control.xlsx` (no real SAS/passwords) + alchemy/arcca/pangeanic sample trackers, then live OneForma QA. |
| **Skip OneForma** | Offline merge only (no blob QA). |
| **Run merge + QA** | Executes the pipeline; fills progress log + findings table. |
| **Probe OneForma filters** | Proxies `GET …/filters` for a quick connectivity check. |
| **Download Result_Merged.xlsx** | Merged OTS sheet + `QA_Findings` + `Source_Status` + `Control_Masked`. |
| **Send data to Vendor Prod Files** | Builds Rate Approval Log rows from the last merge and appends into the local master (dedupe). Writes `uploads/last_rate_append.json` and a copy at `uploads/Vendor_Prod_Files_updated.xlsx`. |

## OTS template columns

Copied from the real template to `template.xlsx`, plus enrichment columns:

`status`, `feedback`, `folderName`, `fileName`, `locale`, `JSONfile`, `audiofile`, `ICF_speaker_1`, `ICF_speaker_2`, `speaker_1_cumulativeAudioMinutes`, `speaker_2_cumulativeAudioMinutes`, `totaldurationMinute`, **`ingestionBatch`**, **`durationSeconds`**

- `ingestionBatch` — OneForma `dateIso` / parsed `date`, else trailing `/YYYYMMDD` from folder.
- `durationSeconds` — OneForma `durationSeconds`, else `totaldurationMinute × 60`.

Row 2 instructions from the template are **not** copied into output data rows.

## Rate Approval Log mapping (Send)

Filled columns only (C, D, H–S, W–Z left empty so existing formulas stay intact on prior rows):

| Col | Field | Source |
|-----|--------|--------|
| A | Vendor Name | vendorName |
| B | Project Code | projectCode (default Maple) |
| E | Workflow | workflow |
| F | Locale | locale (prefer OneForma) |
| G | Ingestion Batch | ingestionBatch |
| T | File Name | fileName |
| U | File path | filePath |
| V | Durations (seconds) | durationSeconds |

**Send updates the local master copy** under `masters/` (kept in sync with a SharePoint download). Live SharePoint push via Graph/browser upload is a follow-up if not implemented yet.

## OneForma APIs (proxied)

- Page: https://dfsupport.oneforma.com/ots-agency-vendor-blob-file-list
- `GET /api/oneforma/filters` → upstream filters
- `GET /api/oneforma/files` → upstream files (query params forwarded)
- `POST /api/send-vendor-prod` → append Rate Approval Log on local master

Durations from the API are **seconds**; the template also keeps **minutes** (`totaldurationMinute`).

## Known gaps (expected in this demo)

- **Google Sheets** Column G links currently **404** → vendor marked `source_unavailable`; use fallback uploads or Demo mode.
- **Pangeanic** (`cv.pangeanic.ai`) is a passworded file manager → stubbed / local sample fallback (password from Column H is never logged/stored/displayed).
- **ConsultBae** may not appear in OneForma vendor filters yet.
- Secrets (SAS keys, passwords, POC emails) are **masked** in UI and in `Control_Masked` sheet. Do not commit real SharePoint downloads (`masters/*.xlsx` is gitignored).

## Layout

```
ots-agency-qa/
  app.py              # FastAPI server
  ots_pipeline.py     # parse / fetch / merge / QA / Rate Approval Log append
  template.xlsx       # real OTS headers + instructions
  static/index.html   # single-page UI
  masters/            # local Vendor Prod Files.xlsx (SharePoint sync copy; gitignored)
  samples/            # synthetic demo control + vendor trackers
  fixtures/           # offline OneForma snapshot (alchemy)
  uploads/            # runtime uploads + Result_Merged.xlsx (gitignored)
  requirements.txt
  .gitignore
  README.md
```