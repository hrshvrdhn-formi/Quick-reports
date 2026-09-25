/**
 * Office Script for the Cost tracker workbook. The Power Automate flow calls it
 * ("Run script") with the `data` object of the Metabase card query response:
 *   { "cols": [{ "name": "report_date", ... }, ...], "rows": [["2026-09-24", ...], ...] }
 *
 * Writes one table row per result row (agents, then TOTAL). Rows already in the
 * table for the same Date are replaced in place, so re-running the flow for a
 * day never duplicates it.
 */

const SHEET_NAME = "Daily Cost";
const TABLE_NAME = "VoiceCost";
const PULLED_AT = "Pulled At (IST)";

// [table header, Metabase column (cost_tracker.py COST_BY_AGENT_SQL), number format]
const COLUMNS: [string, string, string][] = [
  ["Date", "report_date", "yyyy-mm-dd"],
  ["Agent", "agent", "General"],
  ["Calls", "calls", "0"],
  ["Minutes", "minutes", "0.00"],
  ["Token Coverage %", "token_coverage_pct", "0.0"],
  ["LLM USD", "llm_usd", "0.0000"],
  ["LLM Cache Storage USD", "llm_cache_storage_usd", "0.0000"],
  ["TTS USD", "tts_usd", "0.0000"],
  ["STT USD", "stt_usd", "0.0000"],
  ["Telephony USD", "telephony_usd", "0.0000"],
  ["Total USD", "total_usd", "0.0000"],
  ["Cost per Call USD", "cost_per_call_usd", "0.0000"],
  ["Cost per Minute USD", "cost_per_minute_usd", "0.0000"],
  ["Total Stored USD", "total_stored_usd", "0.0000"],
  ["Drift USD", "drift_usd", "0.000000"],
  ["Unpriced Services", "unpriced_services", "0"],
];
const HEADERS = COLUMNS.map((c) => c[0]).concat([PULLED_AT]);
const FORMATS = COLUMNS.map((c) => c[2]).concat(["yyyy-mm-dd hh:mm"]);

interface MetabaseColumn {
  name: string;
}
interface MetabaseData {
  cols: MetabaseColumn[];
  rows: (string | number | boolean | null)[][];
}

function main(workbook: ExcelScript.Workbook, metabaseData: string): string {
  const data = JSON.parse(metabaseData) as MetabaseData;
  const colIndex: { [name: string]: number } = {};
  data.cols.forEach((col, i) => {
    colIndex[col.name] = i;
  });
  for (const [, key] of COLUMNS) {
    if (!(key in colIndex)) {
      throw new Error(`Metabase result has no "${key}" column; was the card's SQL changed?`);
    }
  }
  if (data.rows.length === 0) {
    return "Metabase returned no rows; nothing written";
  }

  // Excel date serials: days since 1899-12-30, the time of day as a fraction.
  const excelEpoch = Date.UTC(1899, 11, 30);
  const dateSerial = (iso: string): number => {
    const [y, m, d] = iso.split("-").map(Number);
    return (Date.UTC(y, m - 1, d) - excelEpoch) / 86400000;
  };
  const pulledAt = (Date.now() + 5.5 * 3600000 - excelEpoch) / 86400000; // IST

  const values: (string | number | boolean)[][] = data.rows.map((row) => {
    const out: (string | number | boolean)[] = COLUMNS.map(([, key]) => {
      const v = row[colIndex[key]];
      if (v === null || v === undefined) return ""; // e.g. every column but Agent on a day with no calls
      return key === "report_date" ? dateSerial(String(v)) : v;
    });
    out.push(pulledAt);
    return out;
  });
  const day = values[0][0] as number;
  const dayText = String(data.rows[0][colIndex["report_date"]]);

  let table = workbook.getTable(TABLE_NAME);
  if (!table) {
    // First run: build the sheet and the table from the header plus this day's rows.
    let sheet = workbook.getWorksheet(SHEET_NAME);
    if (!sheet) {
      sheet = workbook.addWorksheet(SHEET_NAME);
    } else if (sheet.getUsedRange(true)) {
      throw new Error(`Sheet "${SHEET_NAME}" already has data but no "${TABLE_NAME}" table; clear it or rename it`);
    }
    const range = sheet.getRangeByIndexes(0, 0, values.length + 1, HEADERS.length);
    range.setValues([HEADERS as (string | number | boolean)[]].concat(values));
    table = workbook.addTable(range, true);
    table.setName(TABLE_NAME);
    applyFormats(table);
    sheet.getRange().getFormat().autofitColumns();
    return `Created ${TABLE_NAME} with ${values.length} rows for ${dayText}`;
  }

  const header = table.getHeaderRowRange().getValues()[0].map(String);
  if (header.join("|") !== HEADERS.join("|")) {
    throw new Error(`${TABLE_NAME} headers changed; expected: ${HEADERS.join(", ")}`);
  }

  // Drop rows already written for this day (a re-run) and put the new ones in their place.
  let insertAt = -1;
  let replaced = 0;
  const count = table.getRowCount();
  if (count > 0) {
    const dates = table.getColumnByName("Date").getRangeBetweenHeaderAndTotal().getValues();
    for (let i = count - 1; i >= 0; i--) {
      if (dates[i][0] === day) {
        table.deleteRowsAt(i, 1);
        insertAt = i;
        replaced++;
      }
    }
  }
  table.addRows(insertAt, values);
  applyFormats(table);
  return replaced
    ? `Replaced ${replaced} rows for ${dayText} with ${values.length}`
    : `Added ${values.length} rows for ${dayText}`;
}

function applyFormats(table: ExcelScript.Table) {
  HEADERS.forEach((header, i) => {
    table.getColumnByName(header).getRangeBetweenHeaderAndTotal().setNumberFormat(FORMATS[i]);
  });
}
