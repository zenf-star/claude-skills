---
name: refresh-product-areas
description: Use when refreshing the "Product Areas 2.0" sheet in the Release Calendar Google Sheet. Rebuilds product area mappings, dropdowns, descriptions, and cell merges from ProductAreaList. Run each quarter after a human updates Columns A-D.
---

# Refresh Product Areas 2.0

## Overview

Rebuilds the "Product Areas 2.0" sheet from scratch using "ProductAreaList" as the single source of truth. Handles fuzzy squad-to-team name matching, row expansion, dropdown validation, description formulas, and cell merging.

**Depends on:** `google-sheets` skill (same service account auth)

## When to Run

- Each quarter after a human has updated Columns A-D (L0, Tribe, Squad, PO) with one row per squad
- When ProductAreaList is updated with new or renamed product areas
- When squad-to-team mappings need refreshing

## Single Command

```bash
python3 refresh-product-areas/refresh_product_areas.py
```

## What It Does

1. Reads base squad data (A, B, C, D) from "Product Areas 2.0"
2. Reads all product areas from "ProductAreaList" (Column A = area, Column C = team)
3. Matches squads to teams using smart fuzzy matching:
   - Exact match: `Onboard` -> `Onboard`
   - Substring: `Cards` -> `Payments - Cards`
   - Bidirectional substring: `Stables` -> `Payments - Stable`
   - Initial-letter: `S&I` -> `Savings & Investment`
4. Clears existing merges, validations, and data below the header
5. Expands rows: one per product area per squad
6. Fills Column F (product area name) with dropdown validation
7. Adds INDEX/MATCH formula in Column G for descriptions from ProductAreaList
8. Merges Columns A, B, C, D for consecutive same-squad rows
9. Reports unmatched squads (squads with no corresponding team in ProductAreaList)

## Assumptions

- Columns A (L0), B (Tribe), C (Squad), D (PO) are pre-filled by a human
- Column E is an intentional spacer (left empty)
- ProductAreaList is the source of truth for product areas and descriptions
- Unmatched squads get a "(no product areas mapped)" placeholder dropdown
- The script is idempotent: running it twice produces the same result

## Spreadsheet Details

- **Spreadsheet:** Release Calendar (`1n2bFkZXGs975Hmgyz5pwarVA_NW6jmsxplMf5nVK77Y`)
- **Target sheet:** Product Areas 2.0 (gid `35953061`)
- **Source sheet:** ProductAreaList (gid `2077508963`)

## Prerequisites

- `GOOGLE_SA_KEY_PATH` env var set (or key at `~/Desktop/agents-492903-b819951769b4.json`)
- Python 3 with `cryptography` library
- Spreadsheet shared with `prd-jira-agent@agents-492903.iam.gserviceaccount.com` (Editor)
