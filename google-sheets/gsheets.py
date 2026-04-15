#!/usr/bin/env python3
"""Google Sheets CLI — read, write, append, list via service account."""

from __future__ import annotations

import argparse
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


def get_access_token(readonly: bool = True) -> str:
    scope = "https://www.googleapis.com/auth/spreadsheets"
    if readonly:
        scope += ".readonly"
    with open(SA_KEY) as f:
        sa = json.load(f)
    now = int(time.time())
    header = base64.urlsafe_b64encode(
        json.dumps({"alg": "RS256", "typ": "JWT"}).encode()
    ).rstrip(b"=")
    claims = base64.urlsafe_b64encode(
        json.dumps(
            {
                "iss": sa["client_email"],
                "scope": scope,
                "aud": "https://oauth2.googleapis.com/token",
                "iat": now,
                "exp": now + 3600,
            }
        ).encode()
    ).rstrip(b"=")
    signing_input = header + b"." + claims
    key = serialization.load_pem_private_key(sa["private_key"].encode(), password=None)
    sig = base64.urlsafe_b64encode(
        key.sign(signing_input, padding.PKCS1v15(), hashes.SHA256())
    ).rstrip(b"=")
    jwt_token = (signing_input + b"." + sig).decode()
    data = urllib.parse.urlencode(
        {
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": jwt_token,
        }
    ).encode()
    resp = urllib.request.urlopen(
        urllib.request.Request("https://oauth2.googleapis.com/token", data=data)
    )
    return json.loads(resp.read())["access_token"]


def api_request(url: str, token: str, method: str = "GET", body: dict | None = None):
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    if body is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(body).encode()
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read())


def parse_url(url: str) -> tuple[str, str | None]:
    """Extract spreadsheet ID and gid from a Google Sheets URL."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not m:
        print(f"Error: could not parse spreadsheet ID from URL", file=sys.stderr)
        sys.exit(1)
    spreadsheet_id = m.group(1)
    gid = None
    gid_match = re.search(r"[?&#]gid=(\d+)", url)
    if gid_match:
        gid = gid_match.group(1)
    return spreadsheet_id, gid


def resolve_sheet_name(spreadsheet_id: str, gid: str, token: str) -> str:
    """Resolve a gid to a sheet name."""
    meta = api_request(
        f"{BASE}/{spreadsheet_id}?fields=sheets(properties(sheetId,title))", token
    )
    for sheet in meta["sheets"]:
        if str(sheet["properties"]["sheetId"]) == gid:
            return sheet["properties"]["title"]
    print(f"Error: no sheet found with gid={gid}", file=sys.stderr)
    sys.exit(1)


def list_sheets(spreadsheet_id: str, token: str):
    meta = api_request(
        f"{BASE}/{spreadsheet_id}?fields=properties.title,sheets(properties(sheetId,title,index))",
        token,
    )
    print(f"Spreadsheet: {meta['properties']['title']}\n")
    print(f"{'Index':<6} {'GID':<12} {'Sheet Name'}")
    print("-" * 40)
    for sheet in meta["sheets"]:
        p = sheet["properties"]
        print(f"{p['index']:<6} {p['sheetId']:<12} {p['title']}")


def read_sheet(spreadsheet_id: str, sheet_name: str, cell_range: str | None, token: str):
    range_str = sheet_name
    if cell_range:
        range_str = f"{sheet_name}!{cell_range}"
    url = f"{BASE}/{spreadsheet_id}/values/{urllib.parse.quote(range_str)}"
    data = api_request(url, token)
    rows = data.get("values", [])
    if not rows:
        print("(empty)")
        return
    # Print as TSV for easy consumption
    for row in rows:
        print("\t".join(str(c) for c in row))


def write_sheet(
    spreadsheet_id: str,
    sheet_name: str,
    cell_range: str,
    values: list,
    token: str,
):
    range_str = f"{sheet_name}!{cell_range}"
    url = f"{BASE}/{spreadsheet_id}/values/{urllib.parse.quote(range_str)}?valueInputOption=USER_ENTERED"
    body = {"range": range_str, "majorDimension": "ROWS", "values": values}
    result = api_request(url, token, method="PUT", body=body)
    print(f"Updated {result.get('updatedCells', 0)} cells in {result.get('updatedRange', range_str)}")


def append_sheet(
    spreadsheet_id: str,
    sheet_name: str,
    values: list,
    token: str,
):
    range_str = f"{sheet_name}!A:A"
    url = f"{BASE}/{spreadsheet_id}/values/{urllib.parse.quote(range_str)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
    body = {"majorDimension": "ROWS", "values": values}
    result = api_request(url, token, method="POST", body=body)
    updates = result.get("updates", {})
    print(f"Appended {updates.get('updatedRows', 0)} rows at {updates.get('updatedRange', '?')}")


def main():
    parser = argparse.ArgumentParser(description="Google Sheets CLI via service account")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--url", help="Google Sheets URL (extracts ID and gid)")
    source.add_argument("--id", help="Spreadsheet ID directly")

    parser.add_argument("--sheet", help="Sheet name (required with --id, auto-resolved from gid with --url)")
    parser.add_argument("--range", help="Cell range, e.g. A1:D10")
    parser.add_argument("--values", help="JSON array of arrays for write/append")

    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--list", action="store_true", help="List all sheets")
    action.add_argument("--read", action="store_true", help="Read sheet data")
    action.add_argument("--write", action="store_true", help="Write to range (overwrites)")
    action.add_argument("--append", action="store_true", help="Append rows after last data")

    args = parser.parse_args()

    # Resolve spreadsheet ID and sheet name
    if args.url:
        spreadsheet_id, gid = parse_url(args.url)
    else:
        spreadsheet_id = args.id
        gid = None

    is_write = args.write or args.append
    token = get_access_token(readonly=not is_write)

    if args.list:
        list_sheets(spreadsheet_id, token)
        return

    # Resolve sheet name
    sheet_name = args.sheet
    if not sheet_name:
        if gid:
            sheet_name = resolve_sheet_name(spreadsheet_id, gid, token)
        else:
            print("Error: --sheet is required when using --id without a gid", file=sys.stderr)
            sys.exit(1)

    if args.read:
        read_sheet(spreadsheet_id, sheet_name, args.range, token)
    elif args.write:
        if not args.range or not args.values:
            print("Error: --write requires --range and --values", file=sys.stderr)
            sys.exit(1)
        values = json.loads(args.values)
        write_sheet(spreadsheet_id, sheet_name, args.range, values, token)
    elif args.append:
        if not args.values:
            print("Error: --append requires --values", file=sys.stderr)
            sys.exit(1)
        values = json.loads(args.values)
        append_sheet(spreadsheet_id, sheet_name, values, token)


if __name__ == "__main__":
    main()
