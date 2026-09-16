# Mac deploy (parent agent)

Tarball ready on box: `/workspace/ots-agency-qa-full.tgz`

MachineId: `255009a4-0a5f-4f94-ab9f-ae1975ea7ce8`
Dest: `/Users/davidk/Desktop/ots-agency-qa`

## Steps
1. `CopyFromBox` box_path=`/workspace/ots-agency-qa-full.tgz` computer_path=`/Users/davidk/Desktop/ots-agency-qa-full.tgz` machineId=`255009a4-0a5f-4f94-ab9f-ae1975ea7ce8`
2. Shell (machineId=255009a4…):
```bash
cd /Users/davidk/Desktop
mkdir -p ots-agency-qa
tar -xzf ots-agency-qa-full.tgz
# tarball extracts as ./ots-agency-qa/... — if nested, flatten:
# rsync -a ots-agency-qa/ots-agency-qa/ ots-agency-qa/ && rm -rf ots-agency-qa/ots-agency-qa
cd ots-agency-qa
# keep existing .venv if present; else create
if [ ! -x .venv/bin/uvicorn ]; then python3 -m venv .venv && .venv/bin/pip install -r requirements.txt; fi
# kill old server on 8876
pkill -f 'uvicorn app:app.*8876' 2>/dev/null || true
lsof -tiTCP:8876 -sTCP:LISTEN | xargs -r kill 2>/dev/null || true
sleep 1
nohup .venv/bin/uvicorn app:app --host 127.0.0.1 --port 8876 > /tmp/ots-agency-qa.log 2>&1 &
sleep 2
curl -sS http://127.0.0.1:8876/health
curl -sS http://127.0.0.1:8876/ | grep -E 'Send data to Vendor Prod|Download demo control|Download OTS template'
# demo + send
curl -sS -X POST http://127.0.0.1:8876/api/run -F demo_mode=1 -o /tmp/ots_run.json
python3 -c 'import json;d=json.load(open("/tmp/ots_run.json"));print(d["ok"],d["merged_count"])'
curl -sS -X POST http://127.0.0.1:8876/api/send-vendor-prod -o /tmp/ots_send1.json
curl -sS -X POST http://127.0.0.1:8876/api/send-vendor-prod -o /tmp/ots_send2.json
python3 -c 'import json;a=json.load(open("/tmp/ots_send1.json"));b=json.load(open("/tmp/ots_send2.json"));print("send1",a);print("send2",b)'
```

## Box-verified numbers (already done on box @ 127.0.0.1:8876)
- Demo run: merged=82, blobs=503, findings=85
- Pangeanic 76/76 have ingestionBatch + durationSeconds
- Merged sheet includes ingestionBatch + durationSeconds columns
- Send #1: appended=6, skipped_duplicates=76, skipped_incomplete=0, next_row=391
- Send #2: appended=0, skipped_duplicates=82
- UI: Send button present; Download demo control / Download OTS template removed
- SharePoint live upload: not done (no Computer Use in this subagent; stop at local master)
