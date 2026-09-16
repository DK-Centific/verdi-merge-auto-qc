# Verdi Merge + Auto QC tool — Centific

> **Not** the Agency Excel Consolidator (`excel-merger` / Centaurus). This is the Verdi OTS Agency Intake → Merge + OneForma QA tool.

Small FastAPI app that:

1. Loads a SharePoint-style **CONTROL** workbook (upload or demo sample).
2. For each vendor, follows Column G (“Link to Google Drive”) **or** uses an uploaded vendor source `.xlsx` fallback.
3. Extracts / normalizes rows into the **OTS template** columns (plus `ingestionBatch` + `durationSeconds`) and consolidates into one workbook.
4. QAs against the live **OneForma Vendor Blob File List** APIs (server-side proxy, no browser CORS).
5. After each run, builds **Vendor Prod Files / Rate Approval Log** rows (columns A/B/E/F/G/T/U/V) and shows **unique vs duplicate** counts against the destination log.
6. **Send data to Vendor Prod Files** downloads the live SharePoint workbook (when Graph auth is configured), skips File Names already on **Rate Approval Log**, appends only unique rows, and uploads the file back.

This is **not** the Centaurus 30-col excel-merger.

## Quick start (Mac / local)

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/uvicorn app:app --host 127.0.0.1 --port 8876
```

Open **http://127.0.0.1:8876**.

Health check: `GET /health` → `{"ok": true}`.

Local Send without SharePoint: put a copy of `Vendor Prod Files.xlsx` in `masters/` (gitignored), or check **Send to local master only (demo fallback)** after a run.

## UI buttons

| Control | What it does |
|--------|----------------|
| **Control workbook** | Upload Verdi/Juniper-style control sheet. Column H passwords are read in-memory only and never displayed. |
| **Optional single source sheet** | Fallback vendor tracker xlsx when Google links are dead. Optionally tag the vendor name. |
| **Extra vendor source files** | Multiple fallbacks; filename should include the vendor (e.g. `alchemy.xlsx`). |
| **Demo mode** | Uses `samples/demo_control.xlsx` (no real SAS/passwords) + alchemy/arcca/pangeanic sample trackers, then live OneForma QA. |
| **Skip OneForma** | Offline merge only (no blob QA). |
| **Run merge + QA** | Executes the pipeline; fills progress log, unique/duplicate KPI cards, and Vendor Prod preview. |
| **Probe OneForma filters** | Proxies `GET …/filters` for a quick connectivity check. |
| **Download Result_Merged.xlsx** | Merged OTS sheet + `Vendor_Prod_Rows` + `QA_Findings` + `Source_Status` + `Control_Masked`. |
| **Download Vendor Prod rows** | Just the Rate Approval Log fill columns plus unique/duplicate status. |
| **Send data to Vendor Prod Files** | Dedupe against the destination Rate Approval Log, append unique rows only. SharePoint when Graph auth is set; otherwise local `masters/` fallback. |

## OTS template columns

Copied from the real template to `template.xlsx`, plus enrichment columns:

`status`, `feedback`, `folderName`, `fileName`, `locale`, `JSONfile`, `audiofile`, `ICF_speaker_1`, `ICF_speaker_2`, `speaker_1_cumulativeAudioMinutes`, `speaker_2_cumulativeAudioMinutes`, `totaldurationMinute`, **`ingestionBatch`**, **`durationSeconds`**

- `ingestionBatch` — OneForma `dateIso` / parsed `date`, else trailing `/YYYYMMDD` from folder.
- `durationSeconds` — OneForma `durationSeconds`, else `totaldurationMinute × 60`.

Row 2 instructions from the template are **not** copied into output data rows.

## Rate Approval Log mapping (Vendor Prod Files)

Live destination of truth:

https://digitaltechedge.sharepoint.com/:x:/r/sites/Verdi_Juniper/Shared%20Documents/Vendor%20Audio%20Production/Vendor%20Prod%20Files.xlsx

Sheet: **Rate Approval Log**. Filled columns only (C, D, H–S, W–Z left empty so existing formulas stay intact):

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

**Dedupe key:** File Name (column T), case-insensitive. A row is also skipped if the composite `vendor|project|locale|fileName|filePath` already exists. After **Run merge + QA**, the summary cards show unique vs duplicate counts against that key (SharePoint log when auth works; otherwise local `masters/` if present; otherwise in-batch File Name repeats only).

## SharePoint (Vendor Prod Files) — ops notes

Send **does not scrape** the SharePoint UI. It uses Microsoft Graph:

1. Download `Vendor Prod Files.xlsx`
2. Check **Rate Approval Log** for existing File Names
3. Append only unique A/B/E/F/G/T/U/V cells
4. Upload the workbook back (eTag / lock failures are returned as errors)

### When Send uses SharePoint vs local

| Situation | Send target |
|-----------|-------------|
| `AZURE_TENANT_ID` + `AZURE_CLIENT_ID` + `AZURE_CLIENT_SECRET` set, **or** `SHAREPOINT_GRAPH_TOKEN` set | Live SharePoint workbook |
| `SHAREPOINT_VENDOR_PROD_URL` set (or `SHAREPOINT_REQUIRE=1`) but **no** Graph auth | **Fails loudly** — does not write only local |
| Nothing SharePoint-related configured | Local `masters/Vendor Prod Files.xlsx` (demo) |
| UI checkbox **Send to local master only** | Local master even if Graph auth exists |

Render Blueprint (`render.yaml`) sets the production SharePoint URL. Fill the Azure secrets in the Render Dashboard or Send will error instead of quietly updating a local file (Render’s disk is ephemeral anyway).

### Azure app registration (click path)

1. Open [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it (example: `verdi-merge-vendor-prod`). Supported account types: **this directory only**. Register.
3. Copy **Application (client) ID** and **Directory (tenant) ID**.
4. **Certificates & secrets** → **New client secret** → copy the **Value** once.
5. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → add **`Sites.ReadWrite.All`** (or `Files.ReadWrite.All`) → **Grant admin consent**.
6. If your tenant uses **Sites.Selected** instead of tenant-wide Sites.ReadWrite.All, grant this app write on site `Verdi_Juniper` (SharePoint admin / Graph `sites/{id}/permissions`).
7. On Render: **Dashboard** → service **verdi-merge-auto-qc** → **Environment** → set:

   - `AZURE_TENANT_ID`
   - `AZURE_CLIENT_ID`
   - `AZURE_CLIENT_SECRET`

   Optional overrides (already defaulted in `render.yaml`):

   - `SHAREPOINT_VENDOR_PROD_URL`
   - `SHAREPOINT_SITE_HOSTNAME=digitaltechedge.sharepoint.com`
   - `SHAREPOINT_SITE_PATH=sites/Verdi_Juniper`
   - `SHAREPOINT_FILE_PATH=Vendor Audio Production/Vendor Prod Files.xlsx`

Laptop test without a client secret: sign in with a user that can edit the workbook, copy a Graph token, set `SHAREPOINT_GRAPH_TOKEN` in `.env` (never commit it). Tokens expire.

### Check destination

`GET /api/vendor-prod/status` returns whether Send will use SharePoint or local (no secrets in the response).

### Common Send errors

| Error | What to do |
|-------|------------|
| Graph auth is not configured | Set the three `AZURE_*` env vars (or a token). |
| HTTP 401 / 403 | Admin-consent the Graph permission; confirm the app can access `Verdi_Juniper`. |
| HTTP 404 | File moved. Update `SHAREPOINT_FILE_PATH` / URL. |
| File is locked (423) | Close the workbook in Excel / browser, retry Send. |
| eTag mismatch | Someone else saved at the same time. Retry Send. |

## OneForma APIs (proxied)

- Page: https://dfsupport.oneforma.com/ots-agency-vendor-blob-file-list
- `GET /api/oneforma/filters` → upstream filters
- `GET /api/oneforma/files` → upstream files (query params forwarded)
- `POST /api/send-vendor-prod` → dedupe + append Rate Approval Log (SharePoint or local)
- `GET /api/download/vendor-prod-rows` → Rate Approval Log-shaped preview from the last run

Durations from the API are **seconds**; the template also keeps **minutes** (`totaldurationMinute`).

## Known gaps (expected in this demo)

- **Google Sheets** Column G links currently **404** → vendor marked `source_unavailable`; use fallback uploads or Demo mode.
- **Pangeanic** (`cv.pangeanic.ai`) is a passworded file manager → stubbed / local sample fallback (password from Column H is never logged/stored/displayed).
- **ConsultBae** may not appear in OneForma vendor filters yet.
- Secrets (SAS keys, passwords, POC emails, Azure client secrets) are **masked** / never committed. Do not commit real SharePoint downloads (`masters/*.xlsx` is gitignored).

## Layout

```
verdi-merge-auto-qc/
  app.py                      # FastAPI server
  ots_pipeline.py             # parse / fetch / merge / QA / Rate Approval Log append
  sharepoint_vendor_prod.py   # Microsoft Graph download/upload
  template.xlsx               # real OTS headers + instructions
  static/index.html           # single-page UI
  masters/                    # optional local Vendor Prod copy (gitignored)
  samples/                    # synthetic demo control + vendor trackers
  fixtures/                   # offline OneForma snapshot (alchemy)
  uploads/                    # runtime uploads + Result_Merged.xlsx (gitignored)
  tests/                      # unit tests
  .env.example                # SharePoint / Azure env names (no secrets)
  requirements.txt
  .gitignore
  README.md
```
