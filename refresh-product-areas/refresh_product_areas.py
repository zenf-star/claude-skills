#!/usr/bin/env python3
"""
Refreshes the 'Product Areas 2.0' sheet in the Release Calendar spreadsheet.

Builds hierarchical merges:
  - Column A (L0): merged across ALL squads in the L0 group
  - Column B (L1/Tribe): merged across all squads in the same tribe
  - Column C (Squad): merged across product areas within a squad
  - Column D (PO): merged across product areas within a squad

Applies consistent color formatting and per-squad dropdown validations.
No blank rows between squads within the same L0 group; 1 blank row between
L0 groups.
"""

from __future__ import annotations

import base64
import json
import os
import re
import time
import urllib.parse
import urllib.request

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding

SA_KEY = os.environ.get(
    "GOOGLE_SA_KEY_PATH",
    os.path.expanduser("~/Desktop/agents-492903-b819951769b4.json"),
)
BASE = "https://sheets.googleapis.com/v4/spreadsheets"
SPREADSHEET_ID = "1n2bFkZXGs975Hmgyz5pwarVA_NW6jmsxplMf5nVK77Y"
SHEET_NAME = "Product Areas 2.0"
PAL_SHEET_NAME = "ProductAreaList"

# Colors (RGB 0-1 scale) — gradient from dark (L0) to light (Squad)
COL_A_BG = {"red": 0.58, "green": 0.77, "blue": 0.49}
COL_B_BG = {"red": 0.71, "green": 0.84, "blue": 0.66}
COL_C_BG = {"red": 0.85, "green": 0.92, "blue": 0.83}
COL_D_BG = {"red": 0.94, "green": 0.94, "blue": 0.94}  # light grey
WHITE = {"red": 1, "green": 1, "blue": 1}

# Border styles
WHITE_BORDER = {"style": "SOLID", "width": 1, "color": WHITE}
DOTTED_BORDER = {
    "style": "DOTTED", "width": 1,
    "color": {"red": 0.75, "green": 0.75, "blue": 0.75},
}
NO_BORDER = {"style": "NONE"}
WHITE_BORDERS = {
    "top": WHITE_BORDER, "bottom": WHITE_BORDER,
    "left": WHITE_BORDER, "right": WHITE_BORDER,
}
DOTTED_BORDERS = {
    "top": DOTTED_BORDER, "bottom": DOTTED_BORDER,
    "left": DOTTED_BORDER, "right": DOTTED_BORDER,
}
NO_BORDERS = {
    "top": NO_BORDER, "bottom": NO_BORDER,
    "left": NO_BORDER, "right": NO_BORDER,
}


# ── Auth ─────────────────────────────────────────────────────────────────────

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
    key = serialization.load_pem_private_key(
        sa["private_key"].encode(), password=None
    )
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


# ── API helpers ──────────────────────────────────────────────────────────────

def api(url: str, token: str, method: str = "GET", body: dict | None = None):
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    if body is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(body).encode()
    return json.loads(urllib.request.urlopen(req).read())


def quote(s: str) -> str:
    return urllib.parse.quote(s)


# ── Matching ─────────────────────────────────────────────────────────────────

def get_initials(s: str) -> str:
    """Extract first letter of each word. 'S&I' -> 'SI'."""
    words = [w for w in re.split(r"[\s&/\-]+", s) if w]
    return "".join(w[0].upper() for w in words)


def match_squad_to_team(squad: str, team_names: list[str]) -> str | None:
    """Match a squad name to a ProductAreaList team name."""
    sq = squad.lower().strip()

    # 1. Exact match
    for t in team_names:
        if t.lower().strip() == sq:
            return t

    # 2. Squad is a substring of exactly one team name
    hits = [t for t in team_names if sq in t.lower()]
    if len(hits) == 1:
        return hits[0]

    # 3. Team core (after "- ") contained in squad or vice versa
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


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Refreshing Product Areas 2.0")
    print("=" * 60)

    token = get_access_token()
    print("\n[1/9] Authenticated with service account")

    # ── Read data + metadata ──────────────────────────────────────────────
    pa2 = api(
        f"{BASE}/{SPREADSHEET_ID}/values/{quote(SHEET_NAME)}?majorDimension=ROWS",
        token,
    )
    pal = api(
        f"{BASE}/{SPREADSHEET_ID}/values/{quote(PAL_SHEET_NAME)}?majorDimension=ROWS",
        token,
    )
    meta = api(
        f"{BASE}/{SPREADSHEET_ID}?fields=sheets(properties(sheetId,title),merges)",
        token,
    )

    pa2_rows = pa2.get("values", [])
    pal_rows = pal.get("values", [])
    sheet_meta = next(
        s for s in meta["sheets"] if s["properties"]["title"] == SHEET_NAME
    )
    sheet_id = sheet_meta["properties"]["sheetId"]
    existing_merges = sheet_meta.get("merges", [])

    print(
        f"[2/9] Read {len(pa2_rows)} rows from '{SHEET_NAME}', "
        f"{len(pal_rows)} rows from '{PAL_SHEET_NAME}'"
    )

    # ── Build team -> areas mapping from ProductAreaList ──────────────────
    team_to_areas: dict[str, list[str]] = {}
    all_areas: list[str] = []
    for row in pal_rows[1:]:
        if len(row) >= 3 and row[0].strip() and row[2].strip():
            team_to_areas.setdefault(row[2].strip(), []).append(row[0].strip())
            all_areas.append(row[0].strip())
    team_names = sorted(team_to_areas.keys())
    all_areas = sorted(set(all_areas))

    print(
        f"[3/9] ProductAreaList: {len(team_names)} teams, "
        f"{sum(len(v) for v in team_to_areas.values())} product areas"
    )

    # ── Extract unique squads with hierarchy ──────────────────────────────
    # L0/L1 values carry forward until a new value appears in col A/B
    squads: list[dict] = []
    cur_l0 = ""
    cur_l1 = ""
    seen: set[str] = set()

    for row in pa2_rows[1:]:
        a = row[0].strip() if len(row) > 0 else ""
        b = row[1].strip() if len(row) > 1 else ""
        c = row[2].strip() if len(row) > 2 else ""
        d = row[3].strip() if len(row) > 3 else ""

        if a:
            cur_l0 = a
        if b:
            cur_l1 = b
        if not c or c in seen:
            continue

        seen.add(c)
        squads.append({"l0": cur_l0, "l1": cur_l1, "squad": c, "po": d})

    print(f"[4/9] Extracted {len(squads)} unique squads")

    # ── Match squads to teams ─────────────────────────────────────────────
    squad_areas: dict[str, list[str]] = {}
    matched_log: list[str] = []
    unmatched_log: list[str] = []

    for sq in squads:
        team = match_squad_to_team(sq["squad"], team_names)
        if team:
            areas = sorted(team_to_areas.get(team, []))
            squad_areas[sq["squad"]] = areas
            matched_log.append(
                f"  {sq['squad']:25s} -> {team:25s} ({len(areas)} areas)"
            )
        else:
            squad_areas[sq["squad"]] = []
            unmatched_log.append(sq["squad"])

    print(f"\n  Matched ({len(matched_log)}):")
    for m in matched_log:
        print(m)
    if unmatched_log:
        print(f"\n  Unmatched ({len(unmatched_log)}):")
        for u in unmatched_log:
            print(f"  {u:25s} -> (no match in ProductAreaList)")

    # ── Group squads by L0 ────────────────────────────────────────────────
    l0_groups: list[dict] = []
    cur_name = None

    for sq in squads:
        if sq["l0"] != cur_name:
            cur_name = sq["l0"]
            l0_groups.append({"l0": cur_name, "squads": [sq]})
        else:
            l0_groups[-1]["squads"].append(sq)

    # ── Build new dataset with hierarchical tracking ──────────────────────
    new_rows: list[list[str]] = []
    l0_merges: list[tuple[int, int]] = []     # col A
    l1_merges: list[tuple[int, int]] = []     # col B
    cd_merges: list[tuple[int, int]] = []     # cols C + D
    validations: list[tuple[int, list[str]]] = []
    l0_row_ranges: list[tuple[int, int]] = []  # for coloring

    row_idx = 1  # 0-indexed; row 0 is the header

    for gi, l0g in enumerate(l0_groups):
        # One blank separator row between L0 groups
        if gi > 0:
            new_rows.append([""] * 7)
            row_idx += 1

        l0_start = row_idx
        prev_l1: str | None = None
        l1_start: int | None = None

        for sq in l0g["squads"]:
            # ── Handle L1 transitions ─────────────────────────────────
            if sq["l1"] != prev_l1:
                # Close previous L1 merge if it spans multiple rows
                if l1_start is not None and row_idx > l1_start + 1:
                    l1_merges.append((l1_start, row_idx - 1))
                prev_l1 = sq["l1"]
                l1_start = row_idx

            # ── Expand squad into product-area rows ───────────────────
            areas = squad_areas.get(sq["squad"], [])
            n = max(len(areas), 1)
            sq_start = row_idx
            # Dropdown: squad's areas if matched, otherwise all areas
            dropdown = areas if areas else all_areas

            for i in range(n):
                sheet_row = row_idx + 1  # 1-indexed for formulas
                f_val = areas[i] if i < len(areas) else ""
                g_val = (
                    f'=IFERROR(INDEX(ProductAreaList!B:B, '
                    f'MATCH(F{sheet_row}, ProductAreaList!A:A, 0)), "")'
                )

                new_rows.append([
                    sq["l0"] if row_idx == l0_start else "",
                    sq["l1"] if row_idx == l1_start else "",
                    sq["squad"] if i == 0 else "",
                    sq["po"] if i == 0 else "",
                    "",  # Column E spacer
                    f_val,
                    g_val,
                ])
                validations.append((row_idx, dropdown))
                row_idx += 1

            sq_end = row_idx - 1
            if sq_end > sq_start:
                cd_merges.append((sq_start, sq_end))

        # Close last L1 in this L0 group
        if l1_start is not None and row_idx > l1_start + 1:
            l1_merges.append((l1_start, row_idx - 1))

        l0_end = row_idx - 1
        if l0_end > l0_start:
            l0_merges.append((l0_start, l0_end))
        l0_row_ranges.append((l0_start, l0_end))

    total_rows = len(new_rows)
    print(
        f"\n[5/9] Built {total_rows} rows | "
        f"L0 merges: {len(l0_merges)}, L1 merges: {len(l1_merges)}, "
        f"C/D merges: {len(cd_merges)}, dropdowns: {len(validations)}"
    )

    # ── Clear old sheet ───────────────────────────────────────────────────
    clear_reqs: list[dict] = []

    # Unmerge all existing merges
    for m in existing_merges:
        clear_reqs.append({"unmergeCells": {"range": {
            "sheetId": sheet_id,
            "startRowIndex": m["startRowIndex"],
            "endRowIndex": m["endRowIndex"],
            "startColumnIndex": m["startColumnIndex"],
            "endColumnIndex": m["endColumnIndex"],
        }}})

    # Clear all Column F validations
    clear_reqs.append({"setDataValidation": {"range": {
        "sheetId": sheet_id,
        "startRowIndex": 1,
        "startColumnIndex": 5,
        "endColumnIndex": 6,
    }}})

    # Reset background colors + formatting in data area to white/default
    clear_end = max(len(pa2_rows) + 5, total_rows + 10)
    clear_reqs.append({"repeatCell": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 1,
            "endRowIndex": clear_end,
            "startColumnIndex": 0,
            "endColumnIndex": 7,
        },
        "cell": {"userEnteredFormat": {
            "backgroundColor": WHITE,
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": False},
            "borders": NO_BORDERS,
        }},
        "fields": (
            "userEnteredFormat.backgroundColor,"
            "userEnteredFormat.verticalAlignment,"
            "userEnteredFormat.textFormat.bold,"
            "userEnteredFormat.borders"
        ),
    }})

    if clear_reqs:
        api(
            f"{BASE}/{SPREADSHEET_ID}:batchUpdate",
            token, method="POST",
            body={"requests": clear_reqs},
        )

    # Clear content below header
    clear_range = f"{SHEET_NAME}!A2:G"
    api(
        f"{BASE}/{SPREADSHEET_ID}/values/{quote(clear_range)}:clear",
        token, method="POST", body={},
    )
    print(
        f"[6/9] Cleared {len(existing_merges)} merges, validations, "
        f"formatting, and data"
    )

    # ── Write data ────────────────────────────────────────────────────────
    write_range = f"{SHEET_NAME}!A2:G{total_rows + 1}"
    api(
        f"{BASE}/{SPREADSHEET_ID}/values/{quote(write_range)}"
        f"?valueInputOption=USER_ENTERED",
        token, method="PUT",
        body={
            "range": write_range,
            "majorDimension": "ROWS",
            "values": new_rows,
        },
    )
    # Write footer attribution: column A, 2 rows below last data, blue text, overflow
    footer_row_1idx = total_rows + 3  # header=1, data=2..N+1, gap, footer=N+3
    footer_row_0idx = footer_row_1idx - 1
    footer_range = f"{SHEET_NAME}!A{footer_row_1idx}"
    api(
        f"{BASE}/{SPREADSHEET_ID}/values/{quote(footer_range)}"
        f"?valueInputOption=USER_ENTERED",
        token, method="PUT",
        body={
            "range": footer_range,
            "majorDimension": "ROWS",
            "values": [[
                "Built with Claude Code --> "
                "https://github.com/zenf-star/claude-skills/tree/main/refresh-product-areas"
            ]],
        },
    )
    # Format footer: blue text, overflow wrap
    api(
        f"{BASE}/{SPREADSHEET_ID}:batchUpdate",
        token, method="POST",
        body={"requests": [{"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": footer_row_0idx,
                "endRowIndex": footer_row_0idx + 1,
                "startColumnIndex": 0, "endColumnIndex": 1,
            },
            "cell": {"userEnteredFormat": {
                "textFormat": {"foregroundColor": {
                    "red": 0.0, "green": 0.0, "blue": 1.0,
                }},
                "wrapStrategy": "OVERFLOW_CELL",
            }},
            "fields": (
                "userEnteredFormat.textFormat.foregroundColor,"
                "userEnteredFormat.wrapStrategy"
            ),
        }}]},
    )
    print(f"[7/9] Wrote {total_rows} rows + footer (row {footer_row_1idx})")

    # ── Apply formatting (colors + borders) per L0 group ────────────────
    fmt_reqs: list[dict] = []

    FMT_FIELDS_FULL = (
        "userEnteredFormat.backgroundColor,"
        "userEnteredFormat.verticalAlignment,"
        "userEnteredFormat.textFormat.bold,"
        "userEnteredFormat.borders"
    )
    FMT_FIELDS_BG_BORDER = (
        "userEnteredFormat.backgroundColor,"
        "userEnteredFormat.verticalAlignment,"
        "userEnteredFormat.borders"
    )
    FMT_FIELDS_BORDER = "userEnteredFormat.borders"

    for start, end in l0_row_ranges:
        # Column A — darkest green, bold, white gridlines
        fmt_reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 0, "endColumnIndex": 1,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": COL_A_BG,
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True},
                "borders": WHITE_BORDERS,
            }},
            "fields": FMT_FIELDS_FULL,
        }})
        # Column B — medium green, bold, white gridlines
        fmt_reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 1, "endColumnIndex": 2,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": COL_B_BG,
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True},
                "borders": WHITE_BORDERS,
            }},
            "fields": FMT_FIELDS_FULL,
        }})
        # Column C — lightest green, white gridlines
        fmt_reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 2, "endColumnIndex": 3,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": COL_C_BG,
                "verticalAlignment": "MIDDLE",
                "borders": WHITE_BORDERS,
            }},
            "fields": FMT_FIELDS_BG_BORDER,
        }})
        # Column D — light grey, white gridlines
        fmt_reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 3, "endColumnIndex": 4,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": COL_D_BG,
                "verticalAlignment": "MIDDLE",
                "borders": WHITE_BORDERS,
            }},
            "fields": FMT_FIELDS_BG_BORDER,
        }})
        # Columns F-G — dotted gridlines
        fmt_reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 5, "endColumnIndex": 7,
            },
            "cell": {"userEnteredFormat": {
                "borders": DOTTED_BORDERS,
            }},
            "fields": FMT_FIELDS_BORDER,
        }})

    if fmt_reqs:
        api(
            f"{BASE}/{SPREADSHEET_ID}:batchUpdate",
            token, method="POST",
            body={"requests": fmt_reqs},
        )
    print(f"[8/9] Applied formatting ({len(fmt_reqs)} ranges: colors + borders)")

    # ── Apply merges + validations ────────────────────────────────────────
    apply_reqs: list[dict] = []

    # L0 merges — Column A
    for start, end in l0_merges:
        apply_reqs.append({"mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 0, "endColumnIndex": 1,
            },
            "mergeType": "MERGE_ALL",
        }})

    # L1 merges — Column B
    for start, end in l1_merges:
        apply_reqs.append({"mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start, "endRowIndex": end + 1,
                "startColumnIndex": 1, "endColumnIndex": 2,
            },
            "mergeType": "MERGE_ALL",
        }})

    # Squad merges — Columns C and D
    for start, end in cd_merges:
        for col in [2, 3]:
            apply_reqs.append({"mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start, "endRowIndex": end + 1,
                    "startColumnIndex": col, "endColumnIndex": col + 1,
                },
                "mergeType": "MERGE_ALL",
            }})

    # Dropdown validations — Column F
    for row_0, vals in validations:
        apply_reqs.append({"setDataValidation": {
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

    if apply_reqs:
        api(
            f"{BASE}/{SPREADSHEET_ID}:batchUpdate",
            token, method="POST",
            body={"requests": apply_reqs},
        )

    n_merges = len(l0_merges) + len(l1_merges) + len(cd_merges) * 2
    print(
        f"[9/9] Applied {n_merges} merges "
        f"(L0: {len(l0_merges)}, L1: {len(l1_merges)}, "
        f"C/D: {len(cd_merges) * 2}) and {len(validations)} dropdowns"
    )

    # ── Summary ───────────────────────────────────────────────────────────
    print("\n  Layout summary:")
    for start, end in l0_row_ranges:
        # Find the L0 name from the first row in this range
        row_data = new_rows[start - 1]  # -1 because new_rows is 0-indexed from row 1
        l0_name = row_data[0] or "?"
        print(f"    {l0_name:12s} rows {start + 1}-{end + 1} ({end - start + 1} rows)")

    print("\n" + "=" * 60)
    print("  DONE")
    print("=" * 60)


if __name__ == "__main__":
    main()
