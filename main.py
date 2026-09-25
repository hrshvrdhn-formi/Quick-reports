"""Chola interaction summary for today (IST), queried through Metabase.

Run:
    pip install -r requirements.txt
    set METABASE_API_KEY=mb_...        (Git Bash: export METABASE_API_KEY=mb_...)
    uvicorn main:app --port 8000

GET /chola/interactions/today

While running, every SNAPSHOT_INTERVAL_SECONDS (default 300 = 5 min) each job in
SNAPSHOT_JOBS runs its query and appends one row to its SharePoint worksheet.
See sharepoint_excel.py for the Microsoft auth setup.
"""
import asyncio
import logging
import os
from contextlib import asynccontextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

import httpx
from fastapi import FastAPI, HTTPException

from sharepoint_excel import SharePointSheetAppender

METABASE_URL = os.getenv("METABASE_URL", "https://metabase-internal.formi.co.in")
METABASE_API_KEY = os.environ["METABASE_API_KEY"]
METABASE_DATABASE_ID = int(os.getenv("METABASE_DATABASE_ID", "3"))  # V2_Production
ALPHA_DATABASE_ID = 4  # alpha-database
SNAPSHOT_INTERVAL_SECONDS = int(os.getenv("SNAPSHOT_INTERVAL_SECONDS", "300"))

IST = timezone(timedelta(hours=5, minutes=30))

log = logging.getLogger("chola-interactions")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)

# Latest dial attempt per lead from IST midnight until now, bucketed by outcome.
TODAY_SQL = """
WITH bounds AS (
    SELECT
        (date_trunc('day', now() AT TIME ZONE 'Asia/Kolkata') AT TIME ZONE 'Asia/Kolkata') AS day_start,
        now() AS day_end
),
chola_dials AS (
    SELECT
        da.lead_id,
        da.outcome,
        da.raw_provider_status,
        ROW_NUMBER() OVER (
            PARTITION BY da.lead_id
            ORDER BY da.created_at DESC, da.id DESC
        ) AS rn
    FROM public.dial_attempt da
    JOIN public.organization o ON o.id = da.organization_id
    CROSS JOIN bounds b
    WHERE o.organization_id = 'chola'
      AND da.created_at >= b.day_start
      AND da.created_at <  b.day_end
),
latest_status AS (
    SELECT * FROM chola_dials WHERE rn = 1
)
SELECT
    to_char(now() AT TIME ZONE 'Asia/Kolkata', 'YYYY-MM-DD') AS activity_date_ist,
    to_char(now() AT TIME ZONE 'Asia/Kolkata', 'YYYY-MM-DD"T"HH24:MI:SS+05:30') AS as_of_ist,
    COUNT(DISTINCT lead_id) FILTER (WHERE outcome = 'completed') AS complete,
    COUNT(DISTINCT lead_id) FILTER (WHERE outcome = 'dialing' OR raw_provider_status = 'dialing') AS dialing,
    COUNT(DISTINCT lead_id) FILTER (WHERE outcome = 'no_connect') AS did_not_pick,
    COUNT(DISTINCT lead_id) FILTER (WHERE outcome = 'queued' OR raw_provider_status = 'queued') AS queued,
    COUNT(DISTINCT lead_id) FILTER (
        WHERE outcome IN ('telephony_failed', 'telephony_failed_exhausted')
    ) AS telephony_failed,
    COUNT(DISTINCT lead_id) AS total
FROM latest_status
"""

# Per-schedule dials that ended today (IST), plus each schedule's current queue.
SCHEDULES_TODAY_SQL = """
WITH today AS (
    SELECT ha.schedule_id,
           count(*)                                         AS dialed,
           count(*) FILTER (WHERE da.outcome = 'completed') AS connected
    FROM dial_attempt da
    JOIN ph_holder_activity ha ON ha.id = da.activity_id
    WHERE da.organization_id = 3  -- chola
      AND da.outcome IS NOT NULL
      AND ha.schedule_id IS NOT NULL
      AND (da.ended_at AT TIME ZONE 'Asia/Kolkata')::date
          = (now()     AT TIME ZONE 'Asia/Kolkata')::date
    GROUP BY ha.schedule_id
),
backlog AS (
    SELECT schedule_id,
           count(*) FILTER (WHERE dispatch_state = 'queued') AS queued
    FROM ph_holder_activity
    WHERE organization_id = 3
      AND schedule_id IS NOT NULL
    GROUP BY schedule_id
)
SELECT
    count(*)                         AS schedules,
    sum(t.dialed)                    AS dialed,
    sum(t.connected)                 AS connected,
    sum(t.dialed) - sum(t.connected) AS did_not_pick,
    sum(coalesce(b.queued, 0))       AS still_queued
FROM today t
LEFT JOIN backlog b ON b.schedule_id = t.schedule_id
"""

# Redial ledger breakdown for schedules that ran or changed today (IST).
REDIAL_BREAKDOWN_SQL = """
SELECT rl.origin_schedule_id AS schedule_id,
       s.status              AS schedule_status,
       rl.status,
       a.dispatch_state,
       count(*)
FROM redial_ledger rl
JOIN schedule s ON s.id = rl.origin_schedule_id
LEFT JOIN ph_holder_activity a ON a.id = rl.activity_id
WHERE s.run_date = (now() AT TIME ZONE 'Asia/Kolkata')::date
   OR (s.updated_at AT TIME ZONE 'Asia/Kolkata')::date = (now() AT TIME ZONE 'Asia/Kolkata')::date
GROUP BY rl.origin_schedule_id, s.status, rl.status, a.dispatch_state
ORDER BY rl.origin_schedule_id, count(*) DESC
"""

# Redial + dial + queue totals for workflows 93, 463, 94 (the Chola campaigns).
REDIAL_SUMMARY_SQL = """
WITH breakdown AS (
    SELECT rl.origin_schedule_id AS schedule_id,
           s.status              AS schedule_status,
           rl.status,
           a.dispatch_state,
           count(*)
    FROM redial_ledger rl
    JOIN schedule s ON s.id = rl.origin_schedule_id
    LEFT JOIN ph_holder_activity a ON a.id = rl.activity_id
    WHERE rl.workflow_id IN (93, 463, 94)
      AND (   s.run_date = (now() AT TIME ZONE 'Asia/Kolkata')::date
           OR (s.updated_at AT TIME ZONE 'Asia/Kolkata')::date = (now() AT TIME ZONE 'Asia/Kolkata')::date)
    GROUP BY rl.origin_schedule_id, s.status, rl.status, a.dispatch_state
)
SELECT
    count(DISTINCT schedule_id) FILTER (WHERE schedule_status = 'dialing')                              AS schedules_running,
    coalesce(sum(count) FILTER (WHERE status = 'scheduled' AND dispatch_state = 'done'), 0)             AS scheduled_done,
    coalesce(sum(count) FILTER (WHERE status = 'scheduled' AND dispatch_state = 'queued'), 0)           AS scheduled_queued,
    coalesce(sum(count) FILTER (WHERE status = 'failed' OR dispatch_state = 'dispatch_failed'), 0)      AS failed,
    (SELECT count(*) FROM redial_ledger
      WHERE status = 'scheduled'
        AND workflow_id IN (93, 463, 94)
        AND (created_at AT TIME ZONE 'Asia/Kolkata')::date = (now() AT TIME ZONE 'Asia/Kolkata')::date) AS total_redials_scheduled,
    (SELECT count(*) FROM redial_ledger rl2
       JOIN ph_holder_activity a2 ON a2.id = rl2.activity_id
      WHERE rl2.status = 'scheduled' AND a2.dispatch_state = 'done'
        AND rl2.workflow_id IN (93, 463, 94)
        AND (rl2.created_at AT TIME ZONE 'Asia/Kolkata')::date = (now() AT TIME ZONE 'Asia/Kolkata')::date) AS total_redials_done,
    (SELECT count(*) FROM dial_attempt d
       JOIN ph_policy_holder ph ON ph.id = d.lead_id
      WHERE ph.workflow_id IN (93, 463, 94)
        AND d.outcome IN ('completed', 'no_connect')
        AND (d.ended_at AT TIME ZONE 'Asia/Kolkata')::date = (now() AT TIME ZONE 'Asia/Kolkata')::date) AS total_dials_done,
    (SELECT count(*) FROM ph_holder_activity q
       JOIN ph_policy_holder ph2 ON ph2.id = q.holder_id
      WHERE ph2.workflow_id IN (93, 463, 94)
        AND q.dispatch_state = 'queued')                                                                AS total_in_queue,
    (SELECT count(*) FROM ph_holder_activity q
       JOIN ph_policy_holder ph2 ON ph2.id = q.holder_id
       LEFT JOIN schedule s2 ON s2.id = q.schedule_id
      WHERE ph2.workflow_id IN (93, 463, 94)
        AND q.dispatch_state = 'queued'
        AND (s2.id IS NULL OR s2.status <> 'paused'))                                                   AS in_queue_active,
    (SELECT count(*) FROM ph_holder_activity q
       JOIN ph_policy_holder ph2 ON ph2.id = q.holder_id
       JOIN schedule s2 ON s2.id = q.schedule_id
      WHERE ph2.workflow_id IN (93, 463, 94)
        AND q.dispatch_state = 'queued'
        AND s2.status = 'paused')                                                                       AS in_queue_paused
FROM breakdown
"""

CHOLA_TRACKER_URL = (
    "https://agenticuniverse-my.sharepoint.com/:x:/g/personal/harshavardhan_agenticuniverse_ai/"
    "IQCFHQBo2e4SR58e6O0Q9uiSAdvNblmbJtkUACP4uw-zMmE"
)
REDIAL_WORKSHEET = "redial based numbers"


@dataclass
class SnapshotJob:
    name: str
    sql: str
    share_url: str
    worksheet: str
    # Sheet header -> result column. The first column is always "Timestamp" (IST).
    columns: list[tuple[str, str]]
    database: int = METABASE_DATABASE_ID


# Same tracker as REDIAL_SUMMARY_SQL, for Cashify's Retention workflow on alpha.
CASHIFY_RETENTION_SQL = REDIAL_SUMMARY_SQL.replace("IN (93, 463, 94)", "= 38")

CASHIFY_TRACKER_URL = (
    "https://agenticuniverse-my.sharepoint.com/:x:/g/personal/harshavardhan_agenticuniverse_ai/"
    "IQC-EUIajrtoSoDa8eaJlR19AbuNRp8ezye3A9EFU-THOU8"
)

# The Sheet1, "Chola - Hindi + English" and old "2509 Redial Tracker" jobs were
# retired on 2026-09-25; REDIAL_SUMMARY_SQL now covers their numbers.
SNAPSHOT_JOBS = [
    SnapshotJob(
        name="final-tracker",
        sql=REDIAL_SUMMARY_SQL,
        share_url=CHOLA_TRACKER_URL,
        worksheet="FINAL 2509 TRACKER",
        columns=[
            ("Total Schedules Running", "schedules_running"),
            ("Total Scheduled - Done", "scheduled_done"),
            ("Total - Scheduled Queued", "scheduled_queued"),
            ("Failed", "failed"),
            ("Total Redials Scheduled", "total_redials_scheduled"),
            ("Total Redials Done", "total_redials_done"),
            ("Total Dials Done", "total_dials_done"),
            ("Total In Queue", "total_in_queue"),
            ("In Queue - Active", "in_queue_active"),
            ("In Queue - Paused", "in_queue_paused"),
        ],
    ),
    SnapshotJob(
        name="cashify-retention",
        sql=CASHIFY_RETENTION_SQL,
        share_url=CASHIFY_TRACKER_URL,
        worksheet="Sheet1",
        columns=[
            ("Total Schedules Running", "schedules_running"),
            ("Total Scheduled - Done", "scheduled_done"),
            ("Total - Scheduled Queued", "scheduled_queued"),
            ("Failed", "failed"),
            ("Total Redials Scheduled", "total_redials_scheduled"),
            ("Total Redials Done", "total_redials_done"),
            ("Total Dials Done", "total_dials_done"),
            ("Total In Queue", "total_in_queue"),
            ("In Queue - Active", "in_queue_active"),
            ("In Queue - Paused", "in_queue_paused"),
        ],
        database=ALPHA_DATABASE_ID,
    ),
]

class MetabaseError(Exception):
    pass


async def run_metabase_query(sql: str, database: int = METABASE_DATABASE_ID) -> dict:
    """Run a native query and return its first row as {column: value}."""
    rows = await run_metabase_query_rows(sql, database)
    if not rows:
        raise MetabaseError("Metabase returned no rows")
    return rows[0]


async def run_metabase_query_rows(sql: str, database: int = METABASE_DATABASE_ID) -> list[dict]:
    """Run a native query and return every row as {column: value}."""
    payload = {
        "database": database,
        "type": "native",
        "native": {"query": sql},
    }
    async with httpx.AsyncClient(timeout=60) as client:
        try:
            resp = await client.post(
                f"{METABASE_URL}/api/dataset",
                json=payload,
                headers={"x-api-key": METABASE_API_KEY},
            )
            resp.raise_for_status()
        except httpx.HTTPError as exc:
            raise MetabaseError(f"Metabase request failed: {exc}")

    result = resp.json()
    if result.get("status") != "completed":
        raise MetabaseError(result.get("error", "Metabase query failed"))

    cols = [c["name"] for c in result["data"]["cols"]]
    return [dict(zip(cols, row)) for row in result["data"]["rows"]]


async def fetch_today_summary() -> dict:
    return await run_metabase_query(TODAY_SQL)


def make_appender(job: SnapshotJob) -> SharePointSheetAppender:
    return SharePointSheetAppender(
        job.share_url,
        job.worksheet,
        ["Timestamp"] + [header for header, _ in job.columns],
        number_formats=["yyyy-mm-dd hh:mm:ss"] + ["General"] * len(job.columns),
    )


async def run_job(job: SnapshotJob, appender: SharePointSheetAppender) -> None:
    result = await run_metabase_query(job.sql, job.database)
    taken_at = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    # Aggregates come back NULL when nothing matched yet today; write 0 instead.
    row = [taken_at] + [result[key] if result[key] is not None else 0 for _, key in job.columns]
    row_number = await asyncio.to_thread(appender.append, row)
    log.info("[%s] Wrote snapshot %s to %s row %d", job.name, taken_at, job.worksheet, row_number)


def redial_status_header(schedule_id) -> str:
    return f"Sch {schedule_id} status"


def redial_count_header(schedule_id, status, dispatch_state) -> str:
    return f"Sch {schedule_id} · {status} · {dispatch_state or 'no activity'}"


async def run_redial_job(appender: SharePointSheetAppender) -> None:
    """One column per run (labels in column A): each schedule's status, then a
    count per (redial status, dispatch state) combination. New combinations
    add rows at the bottom of the label column."""
    rows = await run_metabase_query_rows(REDIAL_BREAKDOWN_SQL)
    taken_at = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    record = {"Timestamp": taken_at}
    for r in rows:  # ordered by schedule, then count desc
        record.setdefault(redial_status_header(r["schedule_id"]), r["schedule_status"])
        record[redial_count_header(r["schedule_id"], r["status"], r["dispatch_state"])] = r["count"]
    present = {f"Sch {r['schedule_id']} " for r in rows}

    def fill_missing(header: str):
        # A combination that drained to zero for a schedule still in today's
        # results is 0; schedules that dropped out of the query are left blank.
        is_count = " · " in header
        return 0 if is_count and any(header.startswith(p) for p in present) else ""

    column = await asyncio.to_thread(
        appender.append_keyed_column,
        record,
        fill_missing,
        {"Timestamp": "yyyy-mm-dd hh:mm:ss"},
        group_of=lambda label: label.split(" ")[1] if label.startswith("Sch ") else None,
    )
    log.info("[redial] Wrote snapshot %s to %s column %s", taken_at, REDIAL_WORKSHEET, column)


async def snapshot_loop() -> None:
    appenders: dict[str, SharePointSheetAppender] = {}
    while True:
        for job in SNAPSHOT_JOBS:
            try:
                if job.name not in appenders:
                    appenders[job.name] = make_appender(job)
                await run_job(job, appenders[job.name])
            except Exception:
                log.exception("[%s] Snapshot failed; will retry next interval", job.name)
        try:
            if "redial" not in appenders:
                appenders["redial"] = SharePointSheetAppender(CHOLA_TRACKER_URL, REDIAL_WORKSHEET, [])
            await run_redial_job(appenders["redial"])
        except Exception:
            log.exception("[redial] Snapshot failed; will retry next interval")
        await asyncio.sleep(SNAPSHOT_INTERVAL_SECONDS)


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(snapshot_loop())
    yield
    task.cancel()


app = FastAPI(title="Chola Interactions", lifespan=lifespan)


@app.get("/chola/interactions/today")
async def chola_interactions_today():
    try:
        return await fetch_today_summary()
    except MetabaseError as exc:
        raise HTTPException(status_code=502, detail=str(exc))


@app.get("/chola/redials/today")
async def chola_redials_today():
    try:
        return await run_metabase_query_rows(REDIAL_BREAKDOWN_SQL)
    except MetabaseError as exc:
        raise HTTPException(status_code=502, detail=str(exc))


@app.get("/chola/schedules/today")
async def chola_schedules_today():
    try:
        return await run_metabase_query(SCHEDULES_TODAY_SQL)
    except MetabaseError as exc:
        raise HTTPException(status_code=502, detail=str(exc))
