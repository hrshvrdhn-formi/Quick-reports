# Cost tracker (Power Automate)

Every morning a Power Automate flow runs the Metabase question
**Voice cost by agent - yesterday (IST)** and writes its rows into the
[Cost tracker workbook](https://agenticuniverse-my.sharepoint.com/:x:/g/personal/harshavardhan_agenticuniverse_ai/IQBA8JWWiympRbBPDvSSLK2rAcaxxFXRUAjMHGlHPckptdc):
one row per agent plus a TOTAL row, tagged with the day they cover.

```
Recurrence (07:00 IST) ─▶ HTTP: POST Metabase /api/card/<id>/query ─▶ status = completed?
                                                                     ├─ yes ─▶ Excel "Run script" VoiceCostTracker
                                                                     └─ no  ─▶ Terminate (Failed)
```

| Piece | Where |
|---|---|
| SQL + card create/update CLI | `cost_tracker.py` (`COST_BY_AGENT_SQL`) |
| Office Script that writes the table | `power_automate/VoiceCostTracker.ts` |
| Output | Sheet **Daily Cost**, table **VoiceCost** (the script creates both on first run) |

Table columns: Date, Agent, Calls, Minutes, Token Coverage %, LLM USD,
LLM Cache Storage USD, TTS USD, STT USD, Telephony USD, Total USD,
Cost per Call USD, Cost per Minute USD, Total Stored USD, Drift USD,
Unpriced Services, Pulled At (IST).

## Before you start

- **Metabase has to be reachable from Power Automate.** The flow runs in
  Microsoft's cloud, not on the office network. If `metabase-internal.formi.co.in`
  is behind a VPN or an IP allowlist, allow the
  [Power Platform outbound IPs](https://learn.microsoft.com/power-automate/ip-address-configuration)
  for your region, or the HTTP step will time out.
- **Premium licence.** The HTTP action is a premium connector.
- **Office Scripts** must be enabled for the tenant (Excel for the web shows an
  **Automate** tab).
- Build the flow as the workbook's owner. "Run script" uses scripts from the
  flow owner's OneDrive.

## 1. Create the Metabase question

```bash
pip install -r requirements.txt
export METABASE_API_KEY=mb_...
python cost_tracker.py card --database 3        # V2_Production; add --collection <id> to file it
# Card 123: https://metabase-internal.formi.co.in/question/123
python cost_tracker.py preview 123              # the rows the flow will get
```

Running `card` again updates the saved SQL in place, so edit the query in
`cost_tracker.py` and re-run instead of editing it in Metabase. To do it by
hand instead: **New → SQL query**, pick the database, paste
`COST_BY_AGENT_SQL`, save it as `Voice cost by agent - yesterday (IST)`, and take
the id from the question's URL.

The API key's group needs native-query access to the database and view access
to the question's collection.

## 2. Add the Office Script to the workbook

1. Open the Cost tracker workbook in Excel for the web.
2. **Automate → New Script**, replace the editor contents with
   `power_automate/VoiceCostTracker.ts`, rename it **VoiceCostTracker**, then
   **Save script**.

You don't need to create the sheet or table. The first run adds a **Daily Cost**
sheet with a **VoiceCost** table. (If a **Daily Cost** sheet already has data
but no such table, the script stops rather than overwrite it.)

## 3. Build the flow

**Create → Scheduled cloud flow**, name it `Cost tracker - daily`.

1. **Recurrence**: interval `1` / `Day`. Under the advanced options, set
   **Time zone** to `(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi` and
   **At these hours** to `7`, **At these minutes** to `0`. The query covers
   the previous IST day, so this leaves 7 hours for late cost writes.

2. **HTTP** (rename it to `Query Metabase`, because the expressions below use this name):
   - Method `POST`
   - URI `https://metabase-internal.formi.co.in/api/card/<card id>/query`
   - Headers: `x-api-key` = your Metabase key, `Content-Type` = `application/json`
   - Body `{}`
   - **Settings**: turn **Asynchronous pattern** off (Metabase replies `202`),
     and turn **Secure inputs** on so the key stays out of run history.

3. **Condition**: left side expression `body('Query_Metabase')?['status']`,
   *is equal to*, right side `completed`.
   - **If no**: **Terminate**, status `Failed`, message expression
     `string(body('Query_Metabase')?['error'])`.
   - **If yes**: **Excel Online (Business) → Run script**:
     - Location `OneDrive for Business`, Document Library `OneDrive`
     - File: the Cost tracker workbook
     - Script: `VoiceCostTracker`
     - **metabaseData**: expression `string(body('Query_Metabase')?['data'])`

4. Optional: add a **Send an email (V2)** or Teams message after the Condition
   with **Configure run after → has failed / has timed out** checked, so a
   missed day gets noticed.

**Save**, then **Test → Manually** to run it once. The Run script step's
`result` output reads like `Added 14 rows for 2026-09-24`.

## Re-runs and gaps

- Re-running writes **yesterday** again and replaces that day's rows where they
  sit, so running it twice or after a partial failure never duplicates rows.
- The question always covers *yesterday*, so a day the flow missed can't be
  filled by running it later. Run it before midnight IST the same day, or
  backfill by hand.
- A day with no calls writes a single blank TOTAL row.
