"""Cost tracker: per-agent voice-call cost for the last hour and the last day.

Two saved Metabase questions over voice_call_metrics (Production database) feed
the two sheets of the Cost Tracker workbook, one row per agent that had calls:

  * "Per 1 hour": the last full IST hour; run at the top of every hour
  * "Per day":    yesterday (IST); run once each morning

"Connected" means the customer spoke (is_zero_utterance is false and the call
has a duration). Per-interaction costs spread the cost of every call, connected
or not, over the connected interactions. INR = USD x USD_INR_RATE.

    export METABASE_API_KEY=mb_...          # push also needs AZURE_*, see sharepoint_excel.py
    python cost_tracker.py cards            # create/update both Metabase questions
    python cost_tracker.py preview hourly   # or daily: print what push would write
    python cost_tracker.py push hourly      # or daily; re-running replaces that period's rows
"""
import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime

import httpx

from sharepoint_excel import SharePointSheetAppender

METABASE_URL = os.getenv("METABASE_URL", "https://metabase-internal.formi.co.in")
COST_DATABASE_ID = 2  # "Production", where voice_call_metrics lives
USD_INR_RATE = float(os.getenv("USD_INR_RATE", "96.0"))

COST_TRACKER_URL = (
    "https://agenticuniverse-my.sharepoint.com/:x:/g/personal/harshavardhan_agenticuniverse_ai/"
    "IQBA8JWWiympRbBPDvSSLK2rAcaxxFXRUAjMHGlHPckptdc"
)

# One row per agent for the IST period [{from_ist}, {to_ist}), both
# `timestamp without time zone` expressions in IST.
METRICS_SQL = """
WITH params AS (
    SELECT {from_ist} AS from_ist,
           {to_ist}   AS to_ist
),
calls AS (
    SELECT
        m.agent_id,
        COALESCE(a.name, '?')                                            AS agent,
        COALESCE(m.exotel_call_duration_s, m.total_call_ms / 1000.0, 0) AS duration_s,
        NOT COALESCE(m.is_zero_utterance, false)
            AND COALESCE(m.exotel_call_duration_s, m.total_call_ms / 1000.0, 0) > 0 AS connected,
        COALESCE(m.llm_cost_usd, 0)               AS llm_usd,
        COALESCE(m.llm_cache_storage_cost_usd, 0) AS llm_cache_storage_usd,
        COALESCE(m.tts_cost_usd, 0)               AS tts_usd,
        COALESCE(m.stt_cost_usd, 0)               AS stt_usd,
        COALESCE(m.exotel_cost_usd, 0)            AS telephony_usd
    FROM voice_call_metrics m
    LEFT JOIN agent a ON a.id = m.agent_id
    CROSS JOIN params p
    WHERE m.agent_id IS NOT NULL
      AND m.created_at >= p.from_ist AT TIME ZONE 'Asia/Kolkata'
      AND m.created_at <  p.to_ist   AT TIME ZONE 'Asia/Kolkata'
),
per_agent AS (
    SELECT
        agent_id,
        agent,
        count(*)                                 AS total_calls,
        count(*) FILTER (WHERE connected)        AS connected,
        avg(duration_s) FILTER (WHERE connected) AS avg_duration_s,
        sum(llm_usd)                             AS llm_usd,
        sum(llm_cache_storage_usd)               AS llm_cache_storage_usd,
        sum(tts_usd)                             AS tts_usd,
        sum(stt_usd)                             AS stt_usd,
        sum(telephony_usd)                       AS telephony_usd,
        sum(llm_usd + llm_cache_storage_usd + tts_usd + stt_usd + telephony_usd) AS total_usd
    FROM calls
    GROUP BY agent_id, agent
)
SELECT
    to_char(p.from_ist, 'YYYY-MM-DD"T"HH24:MI:SS')                          AS period_start_ist,
    to_char(p.to_ist,   'YYYY-MM-DD"T"HH24:MI:SS')                          AS period_end_ist,
    agent,
    agent_id,
    connected                                                               AS connected_interactions,
    ROUND(avg_duration_s::numeric, 1)                                       AS avg_duration_s,
    ROUND((llm_usd               / NULLIF(connected, 0))::numeric, 6)       AS llm_usd_per_interaction,
    ROUND((tts_usd               / NULLIF(connected, 0))::numeric, 6)       AS tts_usd_per_interaction,
    ROUND((stt_usd               / NULLIF(connected, 0))::numeric, 6)       AS stt_usd_per_interaction,
    ROUND((telephony_usd         / NULLIF(connected, 0))::numeric, 6)       AS telephony_usd_per_interaction,
    ROUND((llm_cache_storage_usd / NULLIF(connected, 0))::numeric, 6)       AS llm_cache_storage_usd_per_interaction,
    ROUND((total_usd             / NULLIF(connected, 0))::numeric, 6)       AS total_usd_per_interaction,
    total_calls,
    ROUND(total_usd::numeric, 4)                                            AS total_usd
FROM per_agent
CROSS JOIN params p
ORDER BY total_usd DESC
"""


@dataclass
class Period:
    name: str
    card_name: str
    worksheet: str
    from_ist: str
    to_ist: str
    suffix: str  # appended to the headers of the columns this adds to the sheet


PERIODS = {
    "hourly": Period(
        name="hourly",
        card_name="Cost tracker - per agent, last full hour (IST)",
        worksheet="Per 1 hour",
        from_ist="date_trunc('hour', now() AT TIME ZONE 'Asia/Kolkata') - INTERVAL '1 hour'",
        to_ist="date_trunc('hour', now() AT TIME ZONE 'Asia/Kolkata')",
        suffix=" (last 1 hour)",
    ),
    "daily": Period(
        name="daily",
        card_name="Cost tracker - per agent, yesterday (IST)",
        worksheet="Per day",
        from_ist="date_trunc('day', now() AT TIME ZONE 'Asia/Kolkata') - INTERVAL '1 day'",
        to_ist="date_trunc('day', now() AT TIME ZONE 'Asia/Kolkata')",
        suffix="",
    ),
}

USD = "0.000000"
EXCEL_EPOCH = datetime(1899, 12, 30)


def excel_serial(value: datetime) -> float:
    """Excel date serial (days since 1899-12-30), so the cell is a real date."""
    return (value - EXCEL_EPOCH).total_seconds() / 86400


def hour_label(value: datetime) -> str:
    return f"{value.hour % 12 or 12}{'am' if value.hour < 12 else 'pm'}"


def sheet_columns(period: Period) -> list[tuple[str, str, object]]:
    """(header, number format, row -> value). The first columns match the headers
    already in the workbook; the rest are added to the right of them."""
    s = period.suffix
    start = lambda r: datetime.fromisoformat(r["period_start_ist"])
    end = lambda r: datetime.fromisoformat(r["period_end_ist"])
    inr = lambda key: lambda r: None if r[key] is None else round(r[key] * USD_INR_RATE, 4)
    if period.name == "hourly":
        lead = [
            ("Timestamp", "yyyy-mm-dd h:mm AM/PM", lambda r: excel_serial(end(r))),
            ("Timeframe of Analysis", "@", lambda r: f"{hour_label(start(r))} to {hour_label(end(r))}"),
        ]
    else:
        lead = [("Date", "yyyy-mm-dd", lambda r: excel_serial(start(r)))]
    field = lambda key: lambda r: r[key]
    return lead + [
        ("Agent", "General", field("agent")),
        ("Agent ID", "0", field("agent_id")),
        ("No. of Connected Interactions (last 1 hour)", "0", field("connected_interactions")),
        ("Average Duration per interaction (last 1 hour)", "0.0", field("avg_duration_s")),
        ("LLM - Average Cost per Interaction (last 1 hour)", USD, field("llm_usd_per_interaction")),
        ("TTS - Average Cost per Interaction (last 1 hour)", USD, field("tts_usd_per_interaction")),
        ("STT - Average Cost per Interaction (last 1 hour)", USD, field("stt_usd_per_interaction")),
        # Added columns (not in the original layout).
        (f"Telephony - Average Cost per Interaction{s}", USD, field("telephony_usd_per_interaction")),
        (f"LLM Cache Storage - Average Cost per Interaction{s}", USD, field("llm_cache_storage_usd_per_interaction")),
        (f"Total - Average Cost per Interaction (USD){s}", USD, field("total_usd_per_interaction")),
        (f"Total - Average Cost per Interaction (INR){s}", "0.0000", inr("total_usd_per_interaction")),
        (f"Total Calls{s}", "0", field("total_calls")),
        (f"Total Cost (USD){s}", "0.0000", field("total_usd")),
        (f"Total Cost (INR){s}", "0.00", inr("total_usd")),
        ("USD to INR Rate", "0.00", lambda r: USD_INR_RATE),
    ]


def _client() -> httpx.Client:
    key = os.environ.get("METABASE_API_KEY")
    if not key:
        sys.exit("METABASE_API_KEY is not set")
    return httpx.Client(base_url=METABASE_URL, headers={"x-api-key": key}, timeout=120)


def find_card(client: httpx.Client, name: str) -> dict | None:
    resp = client.get("/api/search", params={"q": name, "models": "card"})
    resp.raise_for_status()
    body = resp.json()
    results = body["data"] if isinstance(body, dict) else body
    matches = [r for r in results if r["name"] == name and not r.get("archived")]
    return matches[0] if matches else None


def upsert_card(client: httpx.Client, period: Period, database: int, collection: int | None) -> int:
    """Create the period's saved question, or update its SQL if it exists."""
    dataset_query = {
        "database": database,
        "type": "native",
        "native": {
            "query": METRICS_SQL.format(from_ist=period.from_ist, to_ist=period.to_ist).strip(),
            "template-tags": {},
        },
    }
    description = (
        f"Per-agent voice_call_metrics cost, written to the Cost Tracker workbook's "
        f"'{period.worksheet}' sheet by Quick-reports/cost_tracker.py; edit the SQL there."
    )
    existing = find_card(client, period.card_name)
    if existing:
        body = {"dataset_query": dataset_query, "description": description}
        if collection is not None:
            body["collection_id"] = collection
        client.put(f"/api/card/{existing['id']}", json=body).raise_for_status()
        return existing["id"]
    resp = client.post(
        "/api/card",
        json={
            "name": period.card_name,
            "description": description,
            "display": "table",
            "visualization_settings": {},
            "collection_id": collection,
            "dataset_query": dataset_query,
        },
    )
    resp.raise_for_status()
    return resp.json()["id"]


def run_card(client: httpx.Client, period: Period) -> list[dict]:
    card = find_card(client, period.card_name)
    if not card:
        sys.exit(f"No Metabase question named {period.card_name!r}; run `python cost_tracker.py cards`")
    resp = client.post(f"/api/card/{card['id']}/query", json={})
    resp.raise_for_status()
    result = resp.json()
    if result.get("status") != "completed":
        sys.exit(f"Query failed: {result.get('error')}")
    cols = [c["name"] for c in result["data"]["cols"]]
    return [dict(zip(cols, row)) for row in result["data"]["rows"]]


def sheet_rows(period: Period, results: list[dict]) -> list[list]:
    return [[value(r) for _, _, value in sheet_columns(period)] for r in results]


def push(client: httpx.Client, period: Period) -> str:
    """Write the period's rows, replacing any rows already there for it."""
    results = run_card(client, period)
    if not results:
        return f"No calls in the period; nothing written to {period.worksheet}"
    columns = sheet_columns(period)
    appender = SharePointSheetAppender(
        COST_TRACKER_URL,
        period.worksheet,
        [header for header, _, _ in columns],
        number_formats=[fmt for _, fmt, _ in columns],
    )
    first, replaced = appender.replace_rows(sheet_rows(period, results))
    span = f"{results[0]['period_start_ist']} to {results[0]['period_end_ist']}"
    verb = f"Replaced {replaced} rows with" if replaced else "Wrote"
    return f"{verb} {len(results)} rows for {span} at '{period.worksheet}'!A{first}"


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    cards = sub.add_parser("cards", help="create or update both Metabase questions")
    cards.add_argument("--database", type=int, default=COST_DATABASE_ID)
    cards.add_argument("--collection", type=int, help="Metabase collection id (default: root)")
    for cmd, help_text in (("preview", "print the rows push would write"), ("push", "write the rows to the workbook")):
        sub.add_parser(cmd, help=help_text).add_argument("period", choices=PERIODS)
    args = parser.parse_args()

    with _client() as client:
        if args.cmd == "cards":
            for period in PERIODS.values():
                card_id = upsert_card(client, period, args.database, args.collection)
                print(f"{period.name}: card {card_id} {METABASE_URL}/question/{card_id}")
        elif args.cmd == "push":
            print(push(client, PERIODS[args.period]))
        else:
            period = PERIODS[args.period]
            headers = [header for header, _, _ in sheet_columns(period)]
            for row in sheet_rows(period, run_card(client, period)):
                print(json.dumps(dict(zip(headers, row))))


if __name__ == "__main__":
    main()
