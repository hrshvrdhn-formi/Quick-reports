"""Append rows to a worksheet in a SharePoint/OneDrive workbook via Microsoft Graph.

Auth (Azure AD app registration; AZURE_TENANT_ID + AZURE_CLIENT_ID always required):
  * App-only: set AZURE_CLIENT_SECRET. The app needs Files.ReadWrite.All
    (application permission, admin-consented).
  * Delegated: leave AZURE_CLIENT_SECRET unset. The app needs Files.ReadWrite
    (delegated) with "Allow public client flows" enabled. Sign in once with
    `python sharepoint_excel.py login`; the token is cached in MSAL_CACHE_PATH
    and refreshed silently afterwards.

Writes go through the Excel API, so they land even while the workbook is open
in Excel Online / desktop (co-authoring), unlike overwriting the file.
"""
import base64
import os
import re
import threading
from pathlib import Path
from urllib.parse import quote

import httpx
import msal

GRAPH = "https://graph.microsoft.com/v1.0"
DELEGATED_SCOPES = ["Files.ReadWrite"]
APP_SCOPES = ["https://graph.microsoft.com/.default"]

TENANT_ID = os.getenv("AZURE_TENANT_ID", "")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
MSAL_CACHE_PATH = Path(os.getenv("MSAL_CACHE_PATH", ".msal_cache.json"))


class GraphAuthError(Exception):
    pass


class _TokenProvider:
    def __init__(self):
        if not TENANT_ID or not CLIENT_ID:
            raise GraphAuthError("AZURE_TENANT_ID and AZURE_CLIENT_ID must be set")
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        self._lock = threading.Lock()
        if CLIENT_SECRET:
            self._app = msal.ConfidentialClientApplication(
                CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
            )
            self._cache = None
        else:
            self._cache = msal.SerializableTokenCache()
            if MSAL_CACHE_PATH.exists():
                self._cache.deserialize(MSAL_CACHE_PATH.read_text())
            self._app = msal.PublicClientApplication(
                CLIENT_ID, authority=authority, token_cache=self._cache
            )

    def _persist(self):
        if self._cache is not None and self._cache.has_state_changed:
            MSAL_CACHE_PATH.write_text(self._cache.serialize())

    def get_token(self) -> str:
        with self._lock:
            if CLIENT_SECRET:
                result = self._app.acquire_token_for_client(scopes=APP_SCOPES)
            else:
                accounts = self._app.get_accounts()
                result = (
                    self._app.acquire_token_silent(DELEGATED_SCOPES, account=accounts[0])
                    if accounts
                    else None
                )
                if not result:
                    raise GraphAuthError(
                        "No cached Microsoft sign-in; run `python sharepoint_excel.py login`"
                    )
            self._persist()
        if "access_token" not in result:
            raise GraphAuthError(result.get("error_description", "Token acquisition failed"))
        return result["access_token"]

    def device_login(self):
        flow = self._app.initiate_device_flow(scopes=DELEGATED_SCOPES)
        if "user_code" not in flow:
            raise GraphAuthError(flow.get("error_description", "Device flow failed"))
        print(flow["message"])
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise GraphAuthError(result.get("error_description", "Sign-in failed"))
        self._persist()
        print(f"Signed in as {result.get('id_token_claims', {}).get('preferred_username')}")


class SharePointSheetAppender:
    """Appends one row per call below the last used row of `worksheet`."""

    def __init__(
        self, share_url: str, worksheet: str, headers: list[str], number_formats: list[str] | None = None
    ):
        self._tokens = _TokenProvider()
        self._share_url = share_url
        self._sheet_name = worksheet
        self._sheet = quote(worksheet.replace("'", "''"), safe="")
        self._sheet_checked = False
        self._headers = headers
        self._number_formats = number_formats  # per column, applied to data rows
        self._item_path = None  # /drives/{driveId}/items/{itemId}, resolved lazily

    def _client(self) -> httpx.Client:
        return httpx.Client(
            timeout=30, headers={"Authorization": f"Bearer {self._tokens.get_token()}"}
        )

    def _resolve_item(self, client: httpx.Client) -> str:
        if self._item_path is None:
            encoded = base64.urlsafe_b64encode(self._share_url.encode()).decode().rstrip("=")
            resp = client.get(f"{GRAPH}/shares/u!{encoded}/driveItem")
            resp.raise_for_status()
            item = resp.json()
            self._item_path = f"/drives/{item['parentReference']['driveId']}/items/{item['id']}"
        return self._item_path

    def _ws(self, client: httpx.Client) -> str:
        return f"{GRAPH}{self._resolve_item(client)}/workbook/worksheets('{self._sheet}')"

    def _ensure_worksheet(self, client: httpx.Client):
        if self._sheet_checked:
            return
        resp = client.get(f"{self._ws(client)}?$select=name")
        if resp.status_code == 404:
            client.post(
                f"{GRAPH}{self._resolve_item(client)}/workbook/worksheets/add",
                json={"name": self._sheet_name},
            ).raise_for_status()
        else:
            resp.raise_for_status()
        self._sheet_checked = True

    @staticmethod
    def _col(n: int) -> str:
        letters = ""
        while n:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    def _last_used_row(self, client: httpx.Client, width: int) -> int:
        # Only look at the columns we write, so formulas the user fills down in
        # columns to the right don't push new rows further down the sheet.
        cols = f"A:{self._col(width)}"
        resp = client.get(
            f"{self._ws(client)}/range(address='{cols}')/usedRange(valuesOnly=true)?$select=address"
        )
        if resp.status_code == 404:  # no values in these columns yet
            return 0
        resp.raise_for_status()
        # e.g. "'Chola - Hindi + English'!A1:E12" or "...!A1" for a single cell
        match = re.search(r"(\d+)$", resp.json()["address"])
        return int(match.group(1)) if match else 0

    def _write_row(self, client: httpx.Client, row: int, values: list, number_formats=None):
        address = f"A{row}:{self._col(len(values))}{row}"
        body = {"values": [values]}
        if number_formats:
            body["numberFormat"] = [number_formats]
        resp = client.patch(f"{self._ws(client)}/range(address='{address}')", json=body)
        resp.raise_for_status()

    def _ensure_headers(self, client: httpx.Client):
        address = f"A1:{self._col(len(self._headers))}1"
        resp = client.get(f"{self._ws(client)}/range(address='{address}')?$select=values")
        resp.raise_for_status()
        current = resp.json()["values"][0]
        # Fill only blank or "+" placeholder header cells; never overwrite real ones.
        merged = [
            want if str(have).strip() in ("", "+") else have
            for have, want in zip(current, self._headers)
        ]
        if merged != current:
            self._write_row(client, 1, merged)

    def append(self, values: list) -> int:
        with self._client() as client:
            self._ensure_worksheet(client)
            self._ensure_headers(client)
            row = max(self._last_used_row(client, len(values)), 1) + 1
            self._write_row(client, row, values, self._number_formats)
            return row

    def replace_rows(self, rows: list[list]) -> tuple[int, int]:
        """Write `rows` in place of the existing rows whose column A equals the
        first row's column A (e.g. a date being re-run), or below the last used
        row if there are none. Returns (first row number written, rows replaced)."""
        key, width = rows[0][0], len(rows[0])
        with self._client() as client:
            self._ensure_worksheet(client)
            self._ensure_headers(client)
            last = max(self._last_used_row(client, width), 1)
            matches = []
            if last >= 2:
                resp = client.get(f"{self._ws(client)}/range(address='A2:A{last}')?$select=values")
                resp.raise_for_status()
                matches = [i + 2 for i, (v,) in enumerate(resp.json()["values"]) if v == key]
            # Delete matching runs of rows bottom-up so earlier row numbers stay valid.
            runs: list[list[int]] = []
            for r in matches:
                if runs and runs[-1][1] == r - 1:
                    runs[-1][1] = r
                else:
                    runs.append([r, r])
            for start, end in reversed(runs):
                client.post(
                    f"{self._ws(client)}/range(address='{start}:{end}')/delete", json={"shift": "Up"}
                ).raise_for_status()
            if matches:
                first = matches[0]
                client.post(
                    f"{self._ws(client)}/range(address='{first}:{first + len(rows) - 1}')/insert",
                    json={"shift": "Down"},
                ).raise_for_status()
            else:
                first = last + 1
            address = f"A{first}:{self._col(width)}{first + len(rows) - 1}"
            body = {"values": [["" if v is None else v for v in row] for row in rows]}
            if self._number_formats:
                body["numberFormat"] = [self._number_formats] * len(rows)
            client.patch(f"{self._ws(client)}/range(address='{address}')", json=body).raise_for_status()
            return first, len(matches)

    @staticmethod
    def _col_index(letters: str) -> int:
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch) - 64
        return n

    def append_keyed_column(
        self, record: dict, fill_missing, number_formats: dict | None = None, group_of=None
    ) -> str:
        """Transposed append_keyed: labels live in column A, each call adds a new
        column to the right. A new label is inserted as a whole new row right
        after the last existing label with the same `group_of(label)`, or at the
        bottom if its group is new. Returns the column letter written."""
        with self._client() as client:
            self._ensure_worksheet(client)
            resp = client.get(f"{self._ws(client)}/range(address='A:A')/usedRange?$select=values")
            if resp.status_code == 404:  # label column is still empty
                labels = []
            else:
                resp.raise_for_status()
                labels = [str(r[0]) for r in resp.json()["values"]]
            while labels and labels[-1] == "":
                labels.pop()
            for key in (k for k in record if k not in labels):
                group = group_of(key) if group_of else None
                same_group = [i for i, lbl in enumerate(labels) if group is not None and group_of(lbl) == group]
                if same_group and same_group[-1] + 1 < len(labels):
                    index = same_group[-1] + 1
                    # Shift every row below down so earlier columns stay aligned.
                    client.post(
                        f"{self._ws(client)}/range(address='{index + 1}:{index + 1}')/insert",
                        json={"shift": "Down"},
                    ).raise_for_status()
                else:
                    index = len(labels)
                labels.insert(index, key)
                client.patch(
                    f"{self._ws(client)}/range(address='A{index + 1}')",
                    json={"values": [[key]]},
                ).raise_for_status()

            # Next free column, looking only at the label rows so anything the
            # user puts below the block doesn't matter.
            resp = client.get(
                f"{self._ws(client)}/range(address='1:{len(labels)}')/usedRange(valuesOnly=true)?$select=address"
            )
            resp.raise_for_status()
            last_col = re.search(r"([A-Z]+)\d+$", resp.json()["address"]).group(1)
            col = self._col(max(self._col_index(last_col), 1) + 1)

            values = [[record[k] if k in record else fill_missing(k)] for k in labels]
            formats = [[(number_formats or {}).get(k, "General")] for k in labels]
            client.patch(
                f"{self._ws(client)}/range(address='{col}1:{col}{len(labels)}')",
                json={"values": values, "numberFormat": formats},
            ).raise_for_status()
            return col

    def append_keyed(self, record: dict, fill_missing, number_formats: dict | None = None) -> int:
        """Append `record` ({header: value}) under a header row that grows.

        Existing header columns keep their position; keys not yet in the header
        are added to the right in `record` order. Header columns absent from
        `record` get `fill_missing(header)`. `number_formats` maps header -> format.
        """
        with self._client() as client:
            self._ensure_worksheet(client)
            resp = client.get(f"{self._ws(client)}/range(address='1:1')/usedRange?$select=values")
            if resp.status_code == 404:  # header row is still empty
                header = []
            else:
                resp.raise_for_status()
                header = [str(v) for v in resp.json()["values"][0]]
            while header and header[-1] == "":
                header.pop()
            new_keys = [k for k in record if k not in header]
            if new_keys:
                header += new_keys
                self._write_row(client, 1, header)
            values = [record[h] if h in record else fill_missing(h) for h in header]
            formats = [(number_formats or {}).get(h, "General") for h in header]
            row = max(self._last_used_row(client, len(header)), 1) + 1
            self._write_row(client, row, values, formats)
            return row


if __name__ == "__main__":
    import sys

    if sys.argv[1:] == ["login"]:
        _TokenProvider().device_login()
    else:
        print("usage: python sharepoint_excel.py login")
