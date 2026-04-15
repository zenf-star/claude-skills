#!/usr/bin/env python3
"""
Refreshes the 'Product Areas 2.0' sheet in the Release Calendar spreadsheet.

Usage:
    python3 skills/google-sheets/refresh_product_areas.py

What it does:
1. Reads base squad data (A, B, C, D - human-entered) from 'Product Areas 2.0'
2. Reads product areas from 'ProductAreaList' (source of truth)
3. Matches squads to teams using smart fuzzy matching
4. Rebuilds the sheet: expands rows (one per product area), fills F, adds G formulas
5. Sets dropdown validation on Column F
6. Merges A, B, C, D for consecutive same-squad rows
"""

from __future__ import annotations

import base64
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding

SA_KEY = os.environ.get("GOOGLE_SA_KEY_PATH", os.path.expanduser("~/Desktop/agents-492903-b819951769b4.json"))
BASE = "https://sheets.googleapis.com/v4/spreadsheets"
SPREADSHEET_ID = "1n2bFkZXGs975Hmgyz5pwarVA_NW6jmsxplMf5nVK77Y"
SHEET_NAME = "Product Areas 2.0"
PAL_SHEET_NAME = "ProductAreaList"


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_access_token() -> str:
    with open(SA_KEY) as f:
        sa = json.load(f)
    now = int(time.time())
    header = base64.urlsafe_b64encode(
        json.dumps({"alg": "RS256", "typ": "JWT"}).encode()
    ).rstrip(b"=")
    claims = base64.urlsafe_b64encode(json.dumps({
        "iss": sa["client_email"],
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now, "exp": now + 3600,
    }).encode()).rstrip(b"=")
    signing_input = header + b"." + claims
    key = serialization.load_pem_private_key(sa["private_key"].encode(), password=None)
    sig = base64.urlsafe_b64encode(
        key.sign(signing_input, padding.PKCS1v15(), hashes.SHA256())
    ).rstrip(b"=")
    jwt = (signing_input + b"." + sig).decode()
    data = urllib.parse.urlencode({
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "assertion": jwt,
    }).encode()
    resp = urllib.request.urlopen(
        urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    )
    return json.loads(resp.read())["access_token"]


# ── API helpers ───────────────────────────────────────────────────────────────

def api(url: str, token: str, method: str = "GET", body: dict | None = None):
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    if body is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(body).encode()
    return json.loads(urllib.request.urlopen(req).read())


# ── Matching ──────────────────────────────────────────────────────────────────

def get_initials(s: str) -> str:
    """Extract first letter of each word. 'S&I' -> 'SI', 'Savings & Investment' -> 'SI'."""
    words = [w for w in re.split(r"[\s&/\-]+", s) if w]
    return "".join(w[0].upper() for w in words)


def match_squad_to_team(squad: str, team_names: list[str]) -> str | None:
    """Match a squad name to a ProductAreaList team name using multiple strategies."""
    sq = squad.lower().strip()

    # 1. Exact match
    for t in team_names:
        if t.lower().strip() == sq:
            return t

    # 2. Squad is a substring of exactly one team name
    hits = [t for t in team_names if sq in t.lower()]
    if len(hits) == 1:
        return hits[0]

    # 3. Team core (after "Payments - " etc.) contained in squad or vice versa
    hits = []
    for t in team_names:
        core = t.split(" - ", 1)[-1].lower().strip()
        if core in sq or sq in core:
            hits.append(t)
    if len(hits) == 1:
        return hits[0]

    # 4. Initial-letter match (e.g. S&I -> Savings & Investment)
    sq_init = get_initials(squad)
    if len(sq_init) >= 2:
        for t in team_names:
            if get_initials(t) == sq_init:
                return t
            parts = t.split(" - ", 1)
            if len(parts) > 1 and get_initials(parts[-1]) == sq_init:
                return t

    return None


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Refreshing Product Areas 2.0")
    print("=" * 60)

    token = get_access_token()
    print("\n[1/8] Authenticated with service account")

    # ── Read both sheets + metadata ───────────────────────────────────────
    pa2 = api(f"{BASE}/{SPREADSHEET_ID}/values/{urllib.parse.quote(SHEET_NAME)}?majorDimension=ROWS", token)
    pal = api(f"{BASE}/{SPREADSHEET_ID}/values/{urllib.parse.quote(PAL_SHEET_NAME)}?majorDimension=ROWS", token)
    meta = api(f"{BASE}/{SPREADSHEET_ID}?fields=sheets(properties(sheetId,title),merges)", token)

    pa2_rows = pa2.get("values", [])
    pal_rows = pal.get("values", [])
    sheet_meta = next(s for s in meta["sheets"] if s["properties"]["title"] == SHEET_NAME)
    sheet_id = sheet_meta["properties"]["sheetId"]
    existing_merges = sheet_meta.get("merges", [])

    print(f"[2/8] Read {len(pa2_rows)} rows from '{SHEET_NAME}', {len(pal_rows)} rows from '{PAL_SHEET_NAME}'")

    # ── Build team -> areas mapping ───────────────────────────────────────
    team_to_areas: dict[str, list[str]] = {}
    for row in pal_rows[1:]:
        if len(row) >= 3 and row[0].strip() and row[2].strip():
            team_to_areas.setdefault(row[2].strip(), []).append(row[0].strip())
    team_names = sorted(team_to_areas.keys())

    print(f"[3/8] ProductAreaList: {len(team_names)} teams, {sum(len(v) for v in team_to_areas.values())} product areas")

    # ── Extract base squad data (collapse previous expansion rows) ────────
    base_entries: list[dict] = []
    current_squad = None

    for row in pa2_rows[1:]:
        c_val = row[2].strip() if len(row) > 2 else ""
        if not c_val:
            base_entries.append({"type": "separator"})
            current_squad = None
        elif c_val == current_squad:
            continue  # skip expansion row from previous run
        else:
            current_squad = c_val
            base_entries.append({
                "type": "squad",
                "a": row[0] if len(row) > 0 and row[0].strip() else "",
                "b": row[1] if len(row) > 1 and row[1].strip() else "",
                "c": c_val,
                "d": row[3] if len(row) > 3 and row[3].strip() else "",
            })

    squads = [e for e in base_entries if e["type"] == "squad"]
    print(f"[4/8] Extracted {len(squads)} base squads")

    # ── Match squads to teams ─────────────────────────────────────────────
    squad_to_areas: dict[str, list[str]] = {}
    matched = []
    unmatched = []

    for entry in squads:
        squad = entry["c"]
        team = match_squad_to_team(squad, team_names)
        if team:
            areas = sorted(team_to_areas.get(team, []))
            squad_to_areas[squad] = areas
            matched.append(f"  {squad:25s} -> {team:25s} ({len(areas)} areas)")
        else:
            squad_to_areas[squad] = []
            unmatched.append(squad)

    print(f"\n  Matched ({len(matched)}):")
    for m in matched:
        print(m)
    if unmatched:
        print(f"\n  Unmatched ({len(unmatched)}):")
        for u in unmatched:
            print(f"  {u:25s} -> (no match in ProductAreaList)")

    # ── Build new dataset ─────────────────────────────────────────────────
    new_rows: list[list[str]] = []
    merge_ranges: list[tuple[int, int]] = []  # (start_0idx, end_0idx) inclusive
    validation_entries: list[tuple[int, list[str]]] = []  # (row_0idx, values)

    row_idx = 1  # 0-indexed, row 0 is header

    for entry in base_entries:
        if entry["type"] == "separator":
            new_rows.append([""] * 7)
            row_idx += 1
            continue

        squad = entry["c"]
        areas = squad_to_areas.get(squad, [])
        n = max(len(areas), 1)
        start = row_idx
        dropdown = areas if areas else ["(no product areas mapped)"]

        for i in range(n):
            sheet_row = row_idx + 1  # 1-indexed for formulas
            f_val = areas[i] if i < len(areas) else ""
            g_val = ""
            if f_val:
                g_val = f'=IFERROR(INDEX(ProductAreaList!B:B, MATCH(F{sheet_row}, ProductAreaList!A:A, 0)), "")'

            new_rows.append([
                entry["a"] if i == 0 else "",
                entry["b"] if i == 0 else "",
                squad,
                entry["d"] if i == 0 else "",
                "",  # Column E spacer
                f_val,
                g_val,
            ])
            validation_entries.append((row_idx, dropdown))
            row_idx += 1

        end = row_idx - 1
        if end > start:
            merge_ranges.append((start, end))

    print(f"\n[5/8] Built {len(new_rows)} rows, {len(merge_ranges)} merge groups, {len(validation_entries)} dropdowns")

    # ── Clear old sheet ───────────────────────────────────────────────────
    clear_requests = []
    for m in existing_merges:
        clear_requests.append({"unmergeCells": {"range": {
            "sheetId": sheet_id,
            "startRowIndex": m["startRowIndex"], "endRowIndex": m["endRowIndex"],
            "startColumnIndex": m["startColumnIndex"], "endColumnIndex": m["endColumnIndex"],
        }}})
    # Clear all Column F validations
    clear_requests.append({"setDataValidation": {"range": {
        "sheetId": sheet_id, "startRowIndex": 1,
        "startColumnIndex": 5, "endColumnIndex": 6,
    }}})

    if clear_requests:
        api(f"{BASE}/{SPREADSHEET_ID}:batchUpdate", token, method="POST", body={"requests": clear_requests})
    print(f"[6/8] Cleared {len(existing_merges)} old merges and validations")

    # Clear content below header
    clear_range = f"{SHEET_NAME}!A2:G"
    api(f"{BASE}/{SPREADSHEET_ID}/values/{urllib.parse.quote(clear_range)}:clear", token, method="POST", body={})
    print(f"       Cleared data range {clear_range}")

    # ── Write new data ────────────────────────────────────────────────────
    write_range = f"{SHEET_NAME}!A2:G{len(new_rows) + 1}"
    api(
        f"{BASE}/{SPREADSHEET_ID}/values/{urllib.parse.quote(write_range)}?valueInputOption=USER_ENTERED",
        token, method="PUT",
        body={"range": write_range, "majorDimension": "ROWS", "values": new_rows},
    )
    print(f"[7/8] Wrote {len(new_rows)} rows")

    # ── Set validations and merges ────────────────────────────────────────
    fmt_requests = []

    for row_0, vals in validation_entries:
        fmt_requests.append({"setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_0, "endRowIndex": row_0 + 1,
                "startColumnIndex": 5, "endColumnIndex": 6,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": v} for v in vals],
                },
                "showCustomUi": True,
                "strict": False,
            },
        }})

    for start, end in merge_ranges:
        for col in range(4):  # A, B, C, D
            fmt_requests.append({"mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start, "endRowIndex": end + 1,
                    "startColumnIndex": col, "endColumnIndex": col + 1,
                },
                "mergeType": "MERGE_ALL",
            }})

    if fmt_requests:
        api(f"{BASE}/{SPREADSHEET_ID}:batchUpdate", token, method="POST", body={"requests": fmt_requests})
    print(f"[8/8] Applied {len(validation_entries)} dropdowns and {len(merge_ranges) * 4} merges")

    print("\n" + "=" * 60)
    print("  DONE")
    print("=" * 60)


if __name__ == "__main__":
    main()
