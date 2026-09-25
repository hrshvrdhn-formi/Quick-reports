# Quick reports

A small FastAPI service that runs Metabase queries every 5 minutes and appends
the results to Excel workbooks on SharePoint/OneDrive (via Microsoft Graph).

## Live trackers (`SNAPSHOT_JOBS` in `main.py`)

| Job | Metabase DB | Workbook → tab | Query |
|---|---|---|---|
| `final-tracker` | V2_Production (3) | Chola Tracker → FINAL 2509 TRACKER | `REDIAL_SUMMARY_SQL` (workflows 93, 463, 94) |
| `cashify-retention` | alpha-database (4) | Cashify Tracker → Sheet1 | `CASHIFY_RETENTION_SQL` (workflow 38) |
| `redial` | V2_Production (3) | Chola Tracker → redial based numbers | `REDIAL_BREAKDOWN_SQL`, one column per run |

Each run adds one row (a new column for `redial`) with an IST timestamp.

## Cost tracker

Per-agent voice-call cost from `voice_call_metrics` (Metabase Production
database) goes to the Cost Tracker workbook, one row per agent with calls:

| Command | Metabase question | Sheet | Period |
|---|---|---|---|
| `python cost_tracker.py push hourly` | 57 | Per 1 hour | last full IST hour; run at the top of each hour |
| `python cost_tracker.py push daily` | 58 | Per day | yesterday (IST); run each morning |

"Connected" means the customer spoke. Per-interaction costs are all calls' cost
divided by connected interactions, and INR uses `USD_INR_RATE` (default 96.0).
Re-running a period replaces its rows. `cards` creates/updates the questions from
the SQL in `cost_tracker.py`, and `preview hourly|daily` prints the rows without
writing. These run as scheduled jobs outside this service.

## Endpoints

- `GET /chola/interactions/today`
- `GET /chola/schedules/today`
- `GET /chola/redials/today`

## Setup

```bash
pip install -r requirements.txt
export METABASE_API_KEY=mb_...
export AZURE_TENANT_ID=<tenant id> AZURE_CLIENT_ID=<app client id>
python sharepoint_excel.py login   # one-time Microsoft device-code sign-in
python -m uvicorn main:app --port 8000
```

The Azure app needs the delegated Microsoft Graph `Files.ReadWrite` permission
with "Allow public client flows" enabled (or set `AZURE_CLIENT_SECRET` for an
app-only setup, see `sharepoint_excel.py`). Run a single worker; each worker
would append its own copy of every row.
