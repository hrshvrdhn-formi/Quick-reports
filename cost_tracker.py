"""Daily voice-call cost by agent, as a saved Metabase question.

Power Automate runs this question every morning and writes the rows into the
Cost tracker workbook (see power_automate/README.md). This file holds the SQL
and a small CLI to create/update the question and preview its output:

    export METABASE_API_KEY=mb_...
    python cost_tracker.py card [--database 2] [--collection <id>]
    python cost_tracker.py preview <card id>
"""
import argparse
import json
import os
import sys

import httpx

METABASE_URL = os.getenv("METABASE_URL", "https://metabase-internal.formi.co.in")
COST_DATABASE_ID = 2  # "Production", where voice_call_metrics lives

CARD_NAME = "Voice cost by agent - yesterday (IST)"

# Yesterday (IST midnight to midnight) from voice_call_metrics, one row per
# agent plus a TOTAL row last. report_date tags every row with the day covered.
# Calls with no agent_id are grouped under '(no agent_id)'.
COST_BY_AGENT_SQL = """
WITH params AS (
    SELECT (date_trunc('day', now() AT TIME ZONE 'Asia/Kolkata') - INTERVAL '1 day') AT TIME ZONE 'Asia/Kolkata' AS t_from,
           (date_trunc('day', now() AT TIME ZONE 'Asia/Kolkata'))                    AT TIME ZONE 'Asia/Kolkata' AS t_to
),
per_agent AS (
    SELECT
        m.agent_id,
        COALESCE(m.agent_id || ' - ' || COALESCE(a.name, '?'), '(no agent_id)') AS agent,
        COUNT(*)                                                              AS calls,
        SUM(COALESCE(m.exotel_call_duration_s, m.total_call_ms / 1000.0, 0)) / 60.0 AS minutes,
        100.0 * COUNT(*) FILTER (WHERE m.llm_token_usage_source = 'actual')
              / NULLIF(COUNT(*), 0)                                           AS token_coverage_pct,
        SUM(COALESCE(m.llm_cost_usd, 0))                                      AS llm_usd,
        SUM(COALESCE(m.llm_cache_storage_cost_usd, 0))                        AS llm_cache_storage_usd,
        SUM(COALESCE(m.tts_cost_usd, 0))                                      AS tts_usd,
        SUM(COALESCE(m.stt_cost_usd, 0))                                      AS stt_usd,
        SUM(COALESCE(m.exotel_cost_usd, 0))                                   AS telephony_usd,
        SUM(COALESCE(m.total_cost_usd, 0))                                    AS total_stored_usd,
        (bool_or(m.llm_cost_usd IS NULL))::int
      + (bool_or(m.llm_cache_storage_cost_usd IS NULL))::int
      + (bool_or(m.tts_cost_usd IS NULL))::int
      + (bool_or(m.stt_cost_usd IS NULL))::int
      + (bool_or(m.exotel_cost_usd IS NULL))::int                             AS unpriced_services
    FROM voice_call_metrics m
    LEFT JOIN agent a ON a.id = m.agent_id
    CROSS JOIN params p
    WHERE m.created_at >= p.t_from
      AND m.created_at <  p.t_to
    GROUP BY m.agent_id, a.name
),
rows AS (
    SELECT 0 AS sort_key, agent, calls, minutes, token_coverage_pct,
           llm_usd, llm_cache_storage_usd, tts_usd, stt_usd, telephony_usd,
           total_stored_usd, unpriced_services
    FROM per_agent
    UNION ALL
    SELECT 1, 'TOTAL', SUM(calls), SUM(minutes),
           SUM(token_coverage_pct * calls) / NULLIF(SUM(calls), 0),
           SUM(llm_usd), SUM(llm_cache_storage_usd), SUM(tts_usd), SUM(stt_usd),
           SUM(telephony_usd), SUM(total_stored_usd), SUM(unpriced_services)
    FROM per_agent
)
SELECT
    to_char(p.t_from AT TIME ZONE 'Asia/Kolkata', 'YYYY-MM-DD')  AS report_date,
    agent,
    calls,
    ROUND(minutes::numeric, 2)                                   AS minutes,
    ROUND(token_coverage_pct::numeric, 1)                        AS token_coverage_pct,
    ROUND(llm_usd::numeric, 4)                                   AS llm_usd,
    ROUND(llm_cache_storage_usd::numeric, 4)                     AS llm_cache_storage_usd,
    ROUND(tts_usd::numeric, 4)                                   AS tts_usd,
    ROUND(stt_usd::numeric, 4)                                   AS stt_usd,
    ROUND(telephony_usd::numeric, 4)                             AS telephony_usd,
    ROUND(total_usd::numeric, 4)                                 AS total_usd,
    ROUND((total_usd / NULLIF(calls, 0))::numeric, 4)            AS cost_per_call_usd,
    ROUND((total_usd / NULLIF(minutes, 0))::numeric, 4)          AS cost_per_minute_usd,
    ROUND(total_stored_usd::numeric, 4)                          AS total_stored_usd,
    ROUND((total_usd - total_stored_usd)::numeric, 6)            AS drift_usd,
    unpriced_services
FROM (
    SELECT r.*, llm_usd + llm_cache_storage_usd + tts_usd + stt_usd + telephony_usd AS total_usd
    FROM rows r
) x
CROSS JOIN params p
ORDER BY sort_key, total_usd DESC
"""


def _client() -> httpx.Client:
    key = os.environ.get("METABASE_API_KEY")
    if not key:
        sys.exit("METABASE_API_KEY is not set")
    return httpx.Client(base_url=METABASE_URL, headers={"x-api-key": key}, timeout=60)


def find_card(client: httpx.Client) -> dict | None:
    resp = client.get("/api/search", params={"q": CARD_NAME, "models": "card"})
    resp.raise_for_status()
    body = resp.json()
    results = body["data"] if isinstance(body, dict) else body
    matches = [r for r in results if r["name"] == CARD_NAME and not r.get("archived")]
    return matches[0] if matches else None


def upsert_card(client: httpx.Client, database: int, collection: int | None) -> int:
    """Create the saved question, or update its SQL if one with CARD_NAME exists."""
    dataset_query = {
        "database": database,
        "type": "native",
        "native": {"query": COST_BY_AGENT_SQL.strip(), "template-tags": {}},
    }
    description = (
        "Yesterday's (IST) voice_call_metrics cost per agent plus a TOTAL row. "
        "Read daily by the Cost tracker Power Automate flow; managed from "
        "Quick-reports/cost_tracker.py, so edit the SQL there."
    )
    existing = find_card(client)
    if existing:
        body = {"dataset_query": dataset_query, "description": description}
        if collection is not None:
            body["collection_id"] = collection
        client.put(f"/api/card/{existing['id']}", json=body).raise_for_status()
        return existing["id"]
    resp = client.post(
        "/api/card",
        json={
            "name": CARD_NAME,
            "description": description,
            "display": "table",
            "visualization_settings": {},
            "collection_id": collection,
            "dataset_query": dataset_query,
        },
    )
    resp.raise_for_status()
    return resp.json()["id"]


def run_card(client: httpx.Client, card_id: int) -> dict:
    """The same call the Power Automate HTTP step makes."""
    resp = client.post(f"/api/card/{card_id}/query", json={})
    resp.raise_for_status()
    result = resp.json()
    if result.get("status") != "completed":
        sys.exit(f"Query failed: {result.get('error')}")
    return result["data"]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    card = sub.add_parser("card", help="create or update the saved question")
    card.add_argument("--database", type=int, default=COST_DATABASE_ID)
    card.add_argument("--collection", type=int, help="Metabase collection id (default: root)")
    preview = sub.add_parser("preview", help="run the saved question and print its rows")
    preview.add_argument("card_id", type=int)
    args = parser.parse_args()

    with _client() as client:
        if args.cmd == "card":
            card_id = upsert_card(client, args.database, args.collection)
            print(f"Card {card_id}: {METABASE_URL}/question/{card_id}")
            print(f"Power Automate URI: {METABASE_URL}/api/card/{card_id}/query")
        else:
            data = run_card(client, args.card_id)
            cols = [c["name"] for c in data["cols"]]
            for row in data["rows"]:
                print(json.dumps(dict(zip(cols, row))))


if __name__ == "__main__":
    main()
