---
name: google-sheets
description: Use when needing to read from or write to Google Sheets. Authenticates via a Google Cloud service account with JWT/OAuth2. Covers fetching sheet data by URL or ID, writing/updating cells, and listing available sheets.
---

# Google Sheets Read/Write

## Overview

Access Google Sheets via the Google Sheets API v4 using a service account for authentication. No user interaction needed — the service account handles auth automatically.

## Prerequisites

- **Service account key:** Set `GOOGLE_SA_KEY_PATH` env var, or default `~/Desktop/agents-492903-b819951769b4.json`
- **Service account:** `prd-jira-agent@agents-492903.iam.gserviceaccount.com`
- **Python 3** with `cryptography` library (pre-installed)
- The target spreadsheet must be shared with the service account email (Editor for writes, Viewer for reads)

## Quick Reference

| Operation | Script Flag | Example |
|-----------|------------|---------|
| List sheets | `--list` | `python3 gsheets.py --url "..." --list` |
| Read sheet | `--read` | `python3 gsheets.py --url "..." --read` |
| Read range | `--read --range "A1:D10"` | `python3 gsheets.py --url "..." --read --range "A1:D10"` |
| Write cells | `--write` | `python3 gsheets.py --url "..." --write --range "A1" --values '[["a","b"],["c","d"]]'` |
| Append rows | `--append` | `python3 gsheets.py --url "..." --append --values '[["new","row"]]'` |

## Usage

### From a Google Sheets URL

```bash
python3 google-sheets/gsheets.py \
  --url "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit?gid=SHEET_GID" \
  --read
```

### By spreadsheet ID directly

```bash
python3 google-sheets/gsheets.py \
  --id "SPREADSHEET_ID" \
  --sheet "Sheet1" \
  --read
```

### Writing data

Values are passed as a JSON array of arrays (rows of cells):

```bash
python3 google-sheets/gsheets.py \
  --url "https://docs.google.com/spreadsheets/d/.../edit?gid=0" \
  --write --range "A1" --values '[["Name","Score"],["Alice",95],["Bob",87]]'
```

### Appending rows

Appends after the last row with data in the sheet:

```bash
python3 google-sheets/gsheets.py \
  --url "https://docs.google.com/spreadsheets/d/.../edit?gid=0" \
  --append --values '[["Charlie",72]]'
```

## Inline Usage (without script)

If you need to do something the script doesn't cover, use the authentication pattern directly in Python:

```python
import json, time, os, urllib.request, urllib.parse, base64
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding

SA_KEY = os.environ.get("GOOGLE_SA_KEY_PATH", os.path.expanduser("~/Desktop/agents-492903-b819951769b4.json"))

def get_access_token(scope="https://www.googleapis.com/auth/spreadsheets"):
    with open(SA_KEY) as f:
        sa = json.load(f)
    now = int(time.time())
    header = base64.urlsafe_b64encode(json.dumps({"alg":"RS256","typ":"JWT"}).encode()).rstrip(b'=')
    claims = base64.urlsafe_b64encode(json.dumps({
        "iss": sa["client_email"], "scope": scope,
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now, "exp": now + 3600
    }).encode()).rstrip(b'=')
    signing_input = header + b'.' + claims
    key = serialization.load_pem_private_key(sa["private_key"].encode(), password=None)
    sig = base64.urlsafe_b64encode(key.sign(signing_input, padding.PKCS1v15(), hashes.SHA256())).rstrip(b'=')
    jwt_token = (signing_input + b'.' + sig).decode()
    data = urllib.parse.urlencode({"grant_type":"urn:ietf:params:oauth:grant-type:jwt-bearer","assertion":jwt_token}).encode()
    resp = urllib.request.urlopen(urllib.request.Request("https://oauth2.googleapis.com/token", data=data))
    return json.loads(resp.read())["access_token"]
```

Then call any Sheets API v4 endpoint with `Authorization: Bearer {token}`.

## Common Mistakes

- **401 Unauthorized**: The spreadsheet isn't shared with the service account email. Share it with `prd-jira-agent@agents-492903.iam.gserviceaccount.com`.
- **403 Forbidden**: The Sheets API isn't enabled in the GCP project. Enable it at `console.cloud.google.com/apis/library/sheets.googleapis.com`.
- **Using readonly scope for writes**: Read uses `spreadsheets.readonly`, writes need `spreadsheets`. The script handles this automatically.
- **gid vs sheet name**: The API uses sheet names, not gid numbers. The script resolves gid to name automatically.
